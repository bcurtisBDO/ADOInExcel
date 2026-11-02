import {
  ADO_API_VERSION,
  ATTACHMENT_LINK,
  HYPERLINK_LINK,
  IDENTITY_FIELDS,
  MAX_IDENTITY_PICKLIST,
  PARENT_LINK_REVERSE,
  WORK_ITEM_BATCH_SIZE
} from "./constants";
import { allowedWorkItemTypeNames, flattenClassificationPaths, isTreePathField } from "./columns";
import {
  applyParentTitles,
  chunk,
  encodeBasicPat,
  formatIdentityCell,
  identityLabelsFromPicks,
  idsFromQueryResult,
  isDateTimeField,
  mapConcurrent,
  missingParentIds,
  normalizeOrgUrl,
  parentIdFromRelations,
  titleMapFromItems,
  vsspsOrigin
} from "./fields";
import { isHtmlField } from "./html";
import { RelationTypeInfo } from "./linkTypes";
import {
  AdoCredential,
  AdoProject,
  ClassificationNode,
  IdentityPick,
  JsonPatchOperation,
  ProcessCatalog,
  ProcessFieldInfo,
  QueryResultLike,
  SavedQuery,
  WorkItemDetails,
  WorkItemLike,
  WorkItemTypeInfo
} from "./types";
import { commentHtml, mapWorkItemComments, workItemDetailsFromSources } from "./workItemDetails";

export interface QueryNode {
  id?: string;
  name?: string;
  path?: string;
  isFolder?: boolean;
  hasChildren?: boolean;
  queryType?: string;
  children?: QueryNode[];
}

export class AdoClientError extends Error {
  constructor(
    message: string,
    readonly status?: number,
    readonly body?: string
  ) {
    super(message);
    this.name = "AdoClientError";
  }
}

/** The `message` Azure DevOps put in an error response, e.g. "TF201036: …", when there is one. */
export function adoErrorDetail(error: unknown): string | undefined {
  if (!(error instanceof AdoClientError) || !error.body) {
    return undefined;
  }
  try {
    const parsed = JSON.parse(error.body) as { message?: unknown };
    return typeof parsed.message === "string" && parsed.message.trim() ? parsed.message.trim() : undefined;
  } catch {
    return undefined;
  }
}

const REQUEST_TIMEOUT_MS = 60_000;
const UPLOAD_TIMEOUT_MS = 300_000;

export class AdoWorkItemClient {
  private readonly orgUrl: string;
  private readonly authorization: string;

  constructor(orgUrl: string, credential: AdoCredential) {
    this.orgUrl = normalizeOrgUrl(orgUrl);
    if (typeof credential === "string") {
      this.authorization = `Basic ${encodeBasicPat(credential)}`;
    } else if (credential.type === "bearer") {
      this.authorization = `Bearer ${credential.value}`;
    } else {
      this.authorization = `Basic ${encodeBasicPat(credential.value)}`;
    }
  }

  get organizationUrl(): string {
    return this.orgUrl;
  }

  async listProjects(): Promise<AdoProject[]> {
    const payload = await this.get<{ value: AdoProject[] }>("/_apis/projects?$top=1000");
    return (payload.value || []).slice().sort((left, right) => compareNames(left.name, right.name));
  }

  async getQuery(project: string, queryId: string): Promise<QueryNode> {
    return this.get<QueryNode>(
      `/${encodeURIComponent(project)}/_apis/wit/queries/${queryId}?$expand=wiql`
    );
  }

  async runQuery(project: string, queryId: string): Promise<QueryResultLike> {
    return this.get<QueryResultLike>(
      `/${encodeURIComponent(project)}/_apis/wit/wiql/${queryId}`
    );
  }

  async listQueries(project: string): Promise<SavedQuery[]> {
    const saved: SavedQuery[] = [];
    const seen = new Set<string>();
    const pending: string[] = [];
    const errors: unknown[] = [];

    try {
      const roots = await this.get<{ value?: QueryNode[] }>(
        `/${encodeURIComponent(project)}/_apis/wit/queries?$depth=1&$expand=all`
      );
      for (const folder of roots.value || []) {
        collectQueryTree(folder, saved, seen, pending);
      }
    } catch (error) {
      errors.push(error);
      pending.push("My Queries", "Shared Queries");
    }

    if (pending.length === 0 && saved.length === 0) {
      pending.push("My Queries", "Shared Queries");
    }

    while (pending.length > 0) {
      const id = pending.shift()!;
      if (seen.has(`fetch:${id}`)) {
        continue;
      }
      seen.add(`fetch:${id}`);
      try {
        // WIT $depth is 0–2 per request. Deeper Shared Queries / My Queries
        // folder trees (6+ parent/child levels) are walked by fetching the next
        // unexpanded folders until the queue is empty.
        const tree = await this.get<QueryNode>(
          `/${encodeURIComponent(project)}/_apis/wit/queries/${encodeURIComponent(id)}?$depth=2&$expand=all`
        );
        collectQueryTree(tree, saved, seen, pending);
      } catch (error) {
        errors.push(error);
      }
    }

    if (saved.length === 0 && errors.length > 0) {
      const first = errors[0];
      const status = first instanceof AdoClientError ? first.status : undefined;
      if (status === 401 || status === 403) {
        throw new AdoClientError(
          "Connected, but listing queries was denied. Grant the PAT Work Items (Read) and try again.",
          status
        );
      }
      throw first instanceof Error
        ? first
        : new AdoClientError("Could not list saved queries.");
    }

    return saved;
  }

  async listWorkItemTypes(project: string): Promise<WorkItemTypeInfo[]> {
    const payload = await this.get<{ value?: Array<Record<string, unknown>> }>(
      `/${encodeURIComponent(project)}/_apis/wit/workitemtypes`
    );
    return (payload.value || [])
      .map((raw) => ({
        name: String(raw.name || "").trim(),
        referenceName: typeof raw.referenceName === "string" ? raw.referenceName : undefined,
        description: typeof raw.description === "string" ? raw.description : undefined,
        isDisabled: Boolean(raw.isDisabled)
      }))
      .filter((type) => type.name);
  }

  async listWorkItemTypeStates(project: string, typeName: string): Promise<string[]> {
    try {
      const payload = await this.get<{ value?: Array<{ name?: string }> }>(
        `/${encodeURIComponent(project)}/_apis/wit/workitemtypes/${encodeURIComponent(typeName)}/states`
      );
      return (payload.value || []).map((state) => String(state.name || "").trim()).filter(Boolean);
    } catch {
      return [];
    }
  }

  async loadProcessCatalog(project: string): Promise<ProcessCatalog> {
    const typeInfos = (await this.listWorkItemTypes(project)).filter((type) => !type.isDisabled);
    const types = allowedWorkItemTypeNames(typeInfos);
    const [projectFields, areaRoot, iterationRoot, identities] = await Promise.all([
      this.listProjectFields(project),
      this.getClassificationNode(project, "areas"),
      this.getClassificationNode(project, "iterations"),
      this.listProjectIdentities(project)
    ]);

    const statesByType: Record<string, string[]> = {};
    const transitionsByType: Record<string, Record<string, string[]>> = {};
    const used = new Map<string, ProcessFieldInfo>();
    const picklists: Record<string, string[]> = {};
    const suggested = new Set<string>();
    const htmlFields = new Set<string>();
    const identityFields = new Set<string>(IDENTITY_FIELDS);

    await Promise.all(
      types.map(async (typeName) => {
        const [states, typeFields, transitions] = await Promise.all([
          this.listWorkItemTypeStates(project, typeName),
          this.listWorkItemTypeFields(project, typeName),
          this.listWorkItemTypeTransitions(project, typeName)
        ]);
        statesByType[typeName] = states;
        if (Object.keys(transitions).length > 0) {
          transitionsByType[typeName] = transitions;
        }
        for (const field of typeFields) {
          const catalogField = projectFields.get(field.referenceName) || field;
          const merged: ProcessFieldInfo = {
            ...catalogField,
            ...field,
            allowedValues: unionStringList(catalogField.allowedValues, field.allowedValues),
            isHtml: catalogField.isHtml || field.isHtml || isHtmlField(field.referenceName),
            isIdentity: catalogField.isIdentity || field.isIdentity,
            isDate: catalogField.isDate || field.isDate,
            isPicklistSuggested: catalogField.isPicklistSuggested || field.isPicklistSuggested
          };
          used.set(field.referenceName, merged);
          if (merged.isHtml) {
            htmlFields.add(merged.referenceName);
          }
          if (merged.isIdentity) {
            identityFields.add(merged.referenceName);
          }
          if (merged.isPicklistSuggested) {
            suggested.add(merged.referenceName);
          }
          if (
            merged.allowedValues &&
            merged.allowedValues.length > 0 &&
            !isTreePathField(merged.referenceName) &&
            merged.referenceName !== "System.State" &&
            merged.referenceName !== "System.WorkItemType"
          ) {
            picklists[merged.referenceName] = unionStringList(
              picklists[merged.referenceName],
              merged.allowedValues
            );
          }
        }
      })
    );

    return {
      types,
      statesByType,
      transitionsByType,
      picklists,
      suggestedPicklists: Array.from(suggested),
      htmlFields: Array.from(htmlFields),
      identityFields: Array.from(identityFields),
      identities,
      areaPaths: flattenClassificationPaths(areaRoot),
      iterationPaths: flattenClassificationPaths(iterationRoot),
      fields: Array.from(used.values()).sort((left, right) => left.name.localeCompare(right.name))
    };
  }

  async listProcessFields(project: string): Promise<ProcessFieldInfo[]> {
    return (await this.loadProcessCatalog(project)).fields;
  }

  private async listWorkItemTypeTransitions(
    project: string,
    typeName: string
  ): Promise<Record<string, string[]>> {
    try {
      const payload = await this.get<{ transitions?: Record<string, unknown> }>(
        `/${encodeURIComponent(project)}/_apis/wit/workitemtypes/${encodeURIComponent(typeName)}`
      );
      return mapTransitions(payload.transitions);
    } catch {
      return {};
    }
  }

  private async listProjectIdentities(project: string): Promise<string[]> {
    try {
      const teams = await this.get<{ value?: Array<{ id?: string; name?: string }> }>(
        `/_apis/projects/${encodeURIComponent(project)}/teams?$top=25`
      );
      const picks: IdentityPick[] = [];
      const seen = new Set<string>();
      for (const team of (teams.value || []).slice(0, 25)) {
        const teamKey = team.id || team.name;
        if (!teamKey) {
          continue;
        }
        try {
          const members = await this.get<{
            value?: Array<{ identity?: Record<string, unknown> }>;
          }>(`/_apis/projects/${encodeURIComponent(project)}/teams/${encodeURIComponent(teamKey)}/members?$top=200`);
          for (const member of members.value || []) {
            if (!member.identity) {
              continue;
            }
            const pick = mapIdentity(member.identity);
            if (!pick.displayName) {
              continue;
            }
            const key = (pick.uniqueName || pick.displayName).toLowerCase();
            if (seen.has(key)) {
              continue;
            }
            seen.add(key);
            picks.push(pick);
            if (picks.length >= MAX_IDENTITY_PICKLIST) {
              return identityLabelsFromPicks(picks);
            }
          }
        } catch {
          // Skip teams the identity APIs refuse; Assigned To still uses people already on the sheet.
        }
      }
      return identityLabelsFromPicks(picks);
    } catch {
      return [];
    }
  }

  private async getClassificationNode(
    project: string,
    structure: "areas" | "iterations"
  ): Promise<ClassificationNode | undefined> {
    try {
      return await this.get<ClassificationNode>(
        `/${encodeURIComponent(project)}/_apis/wit/classificationnodes/${structure}?$depth=14`
      );
    } catch {
      return undefined;
    }
  }

  private async listProjectFields(project: string): Promise<Map<string, ProcessFieldInfo>> {
    const payload = await this.get<{ value?: Array<Record<string, unknown>> }>(
      `/${encodeURIComponent(project)}/_apis/wit/fields`
    );
    const map = new Map<string, ProcessFieldInfo>();
    for (const raw of payload.value || []) {
      const field = mapProcessField(raw);
      if (field) {
        map.set(field.referenceName, field);
      }
    }
    return map;
  }

  private async listWorkItemTypeFields(project: string, typeName: string): Promise<ProcessFieldInfo[]> {
    try {
      const payload = await this.get<{ value?: Array<Record<string, unknown>>; fields?: Array<Record<string, unknown>> }>(
        `/${encodeURIComponent(project)}/_apis/wit/workitemtypes/${encodeURIComponent(typeName)}/fields`
      );
      return (payload.value || payload.fields || []).map(mapProcessField).filter((field): field is ProcessFieldInfo => Boolean(field));
    } catch {
      return [];
    }
  }

  async getWorkItems(
    project: string,
    ids: number[],
    fields: string[],
    expand?: "Relations" | "None"
  ): Promise<WorkItemLike[]> {
    const unique = Array.from(new Set(ids.filter((id) => Number.isFinite(id) && id > 0)));
    const groups = chunk(unique, WORK_ITEM_BATCH_SIZE);
    const batches = await mapConcurrent(groups, 3, async (group) => {
      const payload = await this.post<{ value?: WorkItemLike[] }>(
        `/${encodeURIComponent(project)}/_apis/wit/workitemsbatch`,
        {
          ids: group,
          fields,
          errorPolicy: "omit",
          $expand: expand || "None"
        }
      );
      return payload.value || [];
    });
    return batches.flat();
  }

  async enrichParentTitles(project: string, items: WorkItemLike[]): Promise<WorkItemLike[]> {
    const titles = titleMapFromItems(items);
    const missing = missingParentIds(items, titles);
    if (missing.length > 0) {
      const parents = await this.getWorkItems(project, missing, ["System.Id", "System.Title"]);
      for (const parent of parents) {
        if (parent.id) {
          titles.set(parent.id, String(parent.fields?.["System.Title"] || "").trim());
        }
      }
    }
    return applyParentTitles(items, titles);
  }

  async getWorkItem(project: string, id: number): Promise<WorkItemLike> {
    return this.get<WorkItemLike>(
      `/${encodeURIComponent(project)}/_apis/wit/workitems/${id}?$expand=relations`
    );
  }

  async loadWorkItemDetails(project: string, id: number): Promise<WorkItemDetails> {
    const [item, comments] = await Promise.all([
      this.getWorkItem(project, id),
      this.listWorkItemComments(project, id)
    ]);
    const parentId = parentIdFromRelations(item);
    let parent: { id: number; title: string } | undefined;
    if (parentId) {
      try {
        const parents = await this.getWorkItems(project, [parentId], ["System.Id", "System.Title"]);
        parent = {
          id: parentId,
          title: String(parents[0]?.fields?.["System.Title"] || "").trim()
        };
      } catch {
        parent = { id: parentId, title: "" };
      }
    }
    return workItemDetailsFromSources(item, this.orgUrl, project, comments, parent);
  }

  async listWorkItemComments(project: string, id: number): Promise<WorkItemDetails["comments"]> {
    try {
      const url = `${this.orgUrl}/${encodeURIComponent(project)}/_apis/wit/workItems/${id}/comments?$top=50&order=desc&api-version=7.1-preview.4`;
      const payload = await this.request<{ comments?: Array<{ id?: number; text?: string; createdBy?: { displayName?: string }; createdDate?: string }> }>(
        url,
        { method: "GET" }
      );
      return mapWorkItemComments(payload.comments || []);
    } catch {
      return [];
    }
  }

  async addWorkItemComment(project: string, id: number, text: string): Promise<void> {
    const url = `${this.orgUrl}/${encodeURIComponent(project)}/_apis/wit/workItems/${id}/comments?api-version=7.1-preview.4`;
    await this.request(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text: commentHtml(text) })
    });
  }

  async addWorkItemLink(id: number, rel: string, targetId: number): Promise<void> {
    await this.updateWorkItem(id, [
      {
        op: "add",
        path: "/relations/-",
        value: {
          rel,
          url: `${this.orgUrl}/_apis/wit/workItems/${targetId}`
        }
      }
    ]);
  }

  async addWorkItemHyperlink(id: number, url: string, comment?: string): Promise<void> {
    const text = comment?.trim();
    await this.updateWorkItem(id, [
      {
        op: "add",
        path: "/relations/-",
        value: {
          rel: HYPERLINK_LINK,
          url,
          ...(text ? { attributes: { comment: text } } : {})
        }
      }
    ]);
  }

  async listRelationTypes(): Promise<RelationTypeInfo[]> {
    const payload = await this.get<{ value?: RelationTypeInfo[] }>("/_apis/wit/workitemrelationtypes");
    return payload.value || [];
  }

  /** With expectedRev, Azure DevOps rejects the removal if the work item changed, so a shifted index can't remove the wrong relation. */
  async removeWorkItemRelation(id: number, index: number, expectedRev?: number): Promise<void> {
    const operations: JsonPatchOperation[] = [];
    if (expectedRev !== undefined) {
      operations.push({ op: "test", path: "/rev", value: expectedRev });
    }
    operations.push({ op: "remove", path: `/relations/${index}` });
    await this.updateWorkItem(id, operations);
  }

  async addWorkItemAttachment(project: string, id: number, fileName: string, data: ArrayBuffer): Promise<void> {
    const uploadUrl = `${this.orgUrl}/${encodeURIComponent(project)}/_apis/wit/attachments?fileName=${encodeURIComponent(fileName)}&api-version=${ADO_API_VERSION}`;
    const uploaded = await this.request<{ url?: string }>(
      uploadUrl,
      {
        method: "POST",
        headers: { "Content-Type": "application/octet-stream" },
        body: data
      },
      UPLOAD_TIMEOUT_MS
    );
    if (!uploaded.url) {
      throw new AdoClientError("Azure DevOps did not return an attachment URL.");
    }
    await this.updateWorkItem(id, [
      {
        op: "add",
        path: "/relations/-",
        value: {
          rel: ATTACHMENT_LINK,
          url: uploaded.url,
          attributes: { comment: "", name: fileName }
        }
      }
    ]);
  }

  async updateWorkItem(id: number, operations: JsonPatchOperation[]): Promise<WorkItemLike> {
    return this.patch<WorkItemLike>(`/_apis/wit/workitems/${id}`, operations);
  }

  async createWorkItem(
    project: string,
    type: string,
    operations: JsonPatchOperation[]
  ): Promise<WorkItemLike> {
    return this.post<WorkItemLike>(
      `/${encodeURIComponent(project)}/_apis/wit/workitems/$${encodeURIComponent(type)}`,
      operations,
      "application/json-patch+json"
    );
  }

  async searchIdentities(term: string): Promise<IdentityPick[]> {
    const query = term.trim();
    if (query.length < 2) {
      return [];
    }

    try {
      const fromIdentities = await this.searchIdentitiesApi(query);
      if (fromIdentities.length > 0) {
        return fromIdentities;
      }
    } catch {
      // Fall through to Graph subject query.
    }

    return this.searchGraphSubjects(query);
  }

  async setParentLink(childId: number, parentId: number | null, previousParentId?: number | null): Promise<void> {
    if ((parentId || null) === (previousParentId || null)) {
      return;
    }

    const item = await this.get<WorkItemLike>(`/_apis/wit/workitems/${childId}?$expand=relations`);
    const relations = item.relations || [];
    const operations: JsonPatchOperation[] = [];

    for (let index = relations.length - 1; index >= 0; index -= 1) {
      if (relations[index].rel === PARENT_LINK_REVERSE) {
        operations.push({ op: "remove", path: `/relations/${index}` });
      }
    }

    if (parentId) {
      operations.push({
        op: "add",
        path: "/relations/-",
        value: {
          rel: PARENT_LINK_REVERSE,
          url: `${this.orgUrl}/_apis/wit/workItems/${parentId}`
        }
      });
    }

    if (operations.length > 0) {
      await this.updateWorkItem(childId, operations);
    }
  }

  private async searchIdentitiesApi(query: string): Promise<IdentityPick[]> {
    const url = `${vsspsOrigin(this.orgUrl)}/_apis/identities?searchFilter=General&filterValue=${encodeURIComponent(query)}&queryMembership=None&api-version=${ADO_API_VERSION}`;
    const payload = await this.request<{ value?: Array<Record<string, unknown>> }>(url, { method: "GET" });
    return (payload.value || []).map(mapIdentity).filter((identity) => identity.displayName);
  }

  private async searchGraphSubjects(query: string): Promise<IdentityPick[]> {
    const url = `${vsspsOrigin(this.orgUrl)}/_apis/graph/subjectquery?api-version=7.1-preview.1`;
    const payload = await this.request<{ value?: Array<Record<string, unknown>> }>(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ query, subjectKind: ["User"] })
    });
    return (payload.value || []).map(mapGraphSubject).filter((identity) => identity.displayName);
  }

  private url(path: string): string {
    const separator = path.includes("?") ? "&" : "?";
    return `${this.orgUrl}${path}${separator}api-version=${ADO_API_VERSION}`;
  }

  private async get<T>(path: string): Promise<T> {
    return this.request<T>(this.url(path), { method: "GET" });
  }

  private async post<T>(path: string, body: unknown, contentType = "application/json"): Promise<T> {
    return this.request<T>(this.url(path), {
      method: "POST",
      headers: { "Content-Type": contentType },
      body: JSON.stringify(body)
    });
  }

  private async patch<T>(path: string, body: unknown): Promise<T> {
    return this.request<T>(this.url(path), {
      method: "PATCH",
      headers: { "Content-Type": "application/json-patch+json" },
      body: JSON.stringify(body)
    });
  }

  private async request<T>(url: string, init: RequestInit, timeoutMs = REQUEST_TIMEOUT_MS): Promise<T> {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    let response: Response;
    try {
      response = await fetch(url, {
        ...init,
        signal: controller.signal,
        headers: {
          Accept: "application/json",
          Authorization: this.authorization,
          "X-TFS-FedAuthRedirect": "Suppress",
          ...(init.headers || {})
        }
      });
    } catch (error) {
      if (controller.signal.aborted) {
        throw new AdoClientError(
          `Azure DevOps did not respond within ${Math.round(timeoutMs / 1000)} seconds (${this.orgUrl}). Check the organization URL and your network connection.`
        );
      }
      throw new AdoClientError(
        `Could not reach ${this.orgUrl}: ${error instanceof Error ? error.message : String(error)}. Check the organization URL and your network connection.`
      );
    } finally {
      clearTimeout(timer);
    }

    if (!response.ok) {
      const body = await response.text();
      throw new AdoClientError(
        `Azure DevOps request failed (${response.status}): ${body.slice(0, 400)}`,
        response.status,
        body
      );
    }

    if (response.status === 204) {
      return undefined as T;
    }

    return (await response.json()) as T;
  }
}

function propertyValue(properties: unknown, key: string): string | undefined {
  if (!properties || typeof properties !== "object") {
    return undefined;
  }
  const value = (properties as Record<string, unknown>)[key];
  if (typeof value === "string") {
    return value;
  }
  if (value && typeof value === "object" && typeof (value as { $value?: unknown }).$value === "string") {
    return (value as { $value: string }).$value;
  }
  return undefined;
}

function mapIdentity(raw: Record<string, unknown>): IdentityPick {
  const displayName = String(raw.providerDisplayName || raw.customDisplayName || raw.displayName || "").trim();
  const uniqueName =
    propertyValue(raw.properties, "Account") ||
    propertyValue(raw.properties, "Mail") ||
    (typeof raw.uniqueName === "string" ? raw.uniqueName : undefined);
  return { displayName, uniqueName, descriptor: typeof raw.descriptor === "string" ? raw.descriptor : undefined };
}

function mapGraphSubject(raw: Record<string, unknown>): IdentityPick {
  return {
    displayName: String(raw.displayName || "").trim(),
    uniqueName: String(raw.mailAddress || raw.principalName || "").trim() || undefined,
    descriptor: typeof raw.descriptor === "string" ? raw.descriptor : undefined
  };
}

export function isQueryFolder(node: QueryNode): boolean {
  return node.isFolder === true || (Boolean(node.hasChildren) && !node.queryType);
}

export function collectQueryTree(
  node: QueryNode | undefined,
  acc: SavedQuery[],
  seen: Set<string>,
  foldersToFetch: string[]
): void {
  if (!node) {
    return;
  }

  if (!isQueryFolder(node)) {
    if (node.id && node.name && !seen.has(node.id)) {
      seen.add(node.id);
      acc.push({
        id: node.id,
        name: node.name,
        path: node.path || node.name
      });
    }
    return;
  }

  // Truncated $depth=2 responses often omit hasChildren on folder stubs.
  // Queue every folder we have not already fetched so trees of depth 6+ load.
  if (node.id && !seen.has(`fetch:${node.id}`)) {
    foldersToFetch.push(node.id);
  }

  for (const child of node.children || []) {
    collectQueryTree(child, acc, seen, foldersToFetch);
  }
}

export function flattenQueries(node: QueryNode, acc: SavedQuery[]): void {
  collectQueryTree(node, acc, new Set(), []);
}

const NAME_COMPARE: Intl.CollatorOptions = { sensitivity: "base", numeric: true };

export function compareNames(a: string, b: string): number {
  return a.localeCompare(b, undefined, NAME_COMPARE);
}

export function queryPathParts(path: string): string[] {
  return path
    .split(/[/\\]/)
    .map((part) => part.trim())
    .filter(Boolean);
}

export function compareQueryPaths(a: string, b: string): number {
  const left = queryPathParts(a);
  const right = queryPathParts(b);
  const count = Math.min(left.length, right.length);
  for (let index = 0; index < count; index += 1) {
    const ordered = compareNames(left[index], right[index]);
    if (ordered !== 0) {
      return ordered;
    }
  }
  return left.length - right.length;
}

export interface QueryPickerNode {
  name: string;
  id: string;
  folder: boolean;
  children: QueryPickerNode[];
}

export interface QueryPickerRow {
  id: string;
  label: string;
  disabled: boolean;
}

export function queryPickerTree(queries: SavedQuery[]): QueryPickerNode[] {
  type Mutable = { name: string; id: string; folder: boolean; children: Map<string, Mutable> };
  const root = new Map<string, Mutable>();
  const sorted = queries.slice().sort((left, right) =>
    compareQueryPaths(left.path || left.name, right.path || right.name)
  );
  for (const query of sorted) {
    const parts = queryPathParts(query.path || query.name);
    let level = root;
    for (let index = 0; index < parts.length; index += 1) {
      const name = parts[index];
      const isLeaf = index === parts.length - 1;
      const key = name.toLowerCase();
      if (!level.has(key)) {
        level.set(key, {
          name,
          id: isLeaf ? query.id : "",
          folder: !isLeaf,
          children: new Map()
        });
      }
      const node = level.get(key)!;
      if (isLeaf) {
        node.id = query.id;
      } else {
        node.folder = true;
      }
      level = node.children;
    }
  }
  const toArray = (nodes: Map<string, Mutable>): QueryPickerNode[] =>
    Array.from(nodes.values())
      .sort((left, right) => compareNames(left.name, right.name))
      .map((node) => ({
        name: node.name,
        id: node.id,
        folder: node.folder || node.children.size > 0,
        children: toArray(node.children)
      }));
  return toArray(root);
}

export function queryPickerRows(queries: SavedQuery[]): QueryPickerRow[] {
  const sorted = queries.slice().sort((left, right) =>
    compareQueryPaths(left.path || left.name, right.path || right.name)
  );
  const rows: QueryPickerRow[] = [];
  const folders = new Set<string>();
  for (const query of sorted) {
    const parts = queryPathParts(query.path || query.name);
    const leaf = parts[parts.length - 1] || query.name;
    for (let depth = 0; depth < parts.length - 1; depth += 1) {
      const folderPath = parts.slice(0, depth + 1).join("/");
      if (folders.has(folderPath.toLowerCase())) {
        continue;
      }
      folders.add(folderPath.toLowerCase());
      rows.push({
        id: "",
        label: `${"  ".repeat(depth)}${parts[depth]}`,
        disabled: true
      });
    }
    rows.push({
      id: query.id,
      label: `${"  ".repeat(Math.max(parts.length - 1, 0))}${leaf}`,
      disabled: false
    });
  }
  return rows;
}

export function idsFromSavedQueryResult(result: QueryResultLike): number[] {
  return idsFromQueryResult(result);
}

export function identityCellValue(identity: IdentityPick): string {
  return formatIdentityCell(identity);
}

function mapProcessField(raw: Record<string, unknown> | undefined): ProcessFieldInfo | undefined {
  if (!raw) {
    return undefined;
  }
  const referenceName = String(raw.referenceName || "").trim();
  const name = String(raw.name || "").trim();
  if (!referenceName || !name) {
    return undefined;
  }
  const type = typeof raw.type === "string" ? raw.type : undefined;
  const typeName = (type || "").toLowerCase();
  return {
    referenceName,
    name,
    type,
    readOnly: Boolean(raw.readOnly),
    isIdentity: Boolean(raw.isIdentity) || typeName === "identity",
    isHtml: typeName === "html" || isHtmlField(referenceName),
    isDate: isDateTimeField(referenceName, type),
    isPicklistSuggested: Boolean(raw.isPicklistSuggested),
    allowedValues: stringList(raw.allowedValues)
  };
}

function stringList(raw: unknown): string[] {
  if (!raw) {
    return [];
  }
  if (Array.isArray(raw)) {
    return raw
      .map((item) => {
        if (typeof item === "string" || typeof item === "number" || typeof item === "boolean") {
          return String(item).trim();
        }
        if (item && typeof item === "object" && typeof (item as { name?: unknown }).name === "string") {
          return String((item as { name: string }).name).trim();
        }
        return "";
      })
      .filter(Boolean);
  }
  if (typeof raw === "object") {
    const record = raw as Record<string, unknown>;
    if (Array.isArray(record.$values)) {
      return stringList(record.$values);
    }
    if (Array.isArray(record.value)) {
      return stringList(record.value);
    }
  }
  return [];
}

function unionStringList(left?: string[], right?: string[]): string[] {
  const seen = new Set<string>();
  const result: string[] = [];
  for (const value of [...(left || []), ...(right || [])]) {
    const trimmed = value.trim();
    if (!trimmed || seen.has(trimmed.toLowerCase())) {
      continue;
    }
    seen.add(trimmed.toLowerCase());
    result.push(trimmed);
  }
  return result;
}

function mapTransitions(raw: Record<string, unknown> | undefined): Record<string, string[]> {
  if (!raw) {
    return {};
  }
  const transitions: Record<string, string[]> = {};
  for (const [from, value] of Object.entries(raw)) {
    const states = stringList(value);
    if (states.length > 0) {
      transitions[from] = states;
    }
  }
  return transitions;
}
