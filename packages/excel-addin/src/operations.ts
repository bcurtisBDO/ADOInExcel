import {
  AdoClientError,
  AdoWorkItemClient,
  DEFAULT_COLUMNS,
  DEFAULT_LINK_TYPES,
  HYPERLINK_LINK,
  IdentityPick,
  LinkTypeOption,
  RelationTypeInfo,
  MAX_WORK_ITEMS,
  MAX_TREE_LEVELS,
  MIN_TREE_LEVELS,
  PARENT_TITLE_FIELD,
  ProcessCatalog,
  ProcessFieldInfo,
  QueryResultLike,
  TableRow,
  TreeLayout,
  WorkItemDetails,
  WorkbookColumn,
  WorkItemLike,
  WorkbookMetadata,
  addableLinkTypes,
  addableProcessFields,
  adoErrorDetail,
  applyCatalogToColumns,
  applyRevisionConflicts,
  buildColumnLayout,
  buildMetadata,
  buildPublishPlan,
  buildTreeFromParentField,
  buildTreeLayout,
  catalogFields,
  catalogFromMetadata,
  defaultStateForType,
  emptyProcessCatalog,
  fieldReferenceNames,
  fieldToCell,
  fillParentTitlesOnRows,
  hasParentTitleColumn,
  identityCellValue,
  idsFromQueryResult,
  insertWorkItemRow,
  isHierarchyLink,
  isLockedColumn,
  isSupportedQueryId,
  mapConcurrent,
  markIdentityColumns,
  markDateColumns,
  mergeColumns,
  normalizeHyperlinkUrl,
  normalizeOrgUrl,
  orderItems,
  organizationKey,
  queryWebUrl,
  relationTypeNames,
  sameOrganization,
  shiftTreeLevel,
  singleTitleColumns,
  withRelationNames,
  workItemToSnapshot,
  workItemWebUrl
} from "@ado-open-in-excel/shared";
import { ConnectionSettings, isGuid, loadSettings, saveSettings } from "./auth";
import { acquireAdoToken, currentEntraUsername, signInWithEntra, signOutEntra } from "./entra";
import {
  getSheetSelection,
  readMetadata,
  readTableRows,
  reapplyProcessValidation,
  replaceCreatedIds,
  restoreTreePresentation,
  setClassificationPaths,
  setIdentityLabels,
  setSelectedCell,
  writeMetadata,
  writeRows,
  writeWorkItems
} from "./excelIo";

const PUBLISH_CONCURRENCY = 5;

let processCatalogCache: { orgKey: string; project: string; catalog: ProcessCatalog } | undefined;

function cachedCatalog(orgUrl: string, project: string): ProcessCatalog | undefined {
  if (processCatalogCache?.orgKey === organizationKey(orgUrl) && processCatalogCache.project === project) {
    return processCatalogCache.catalog;
  }
  return undefined;
}

let relationTypeCache: { orgKey: string; types: RelationTypeInfo[] } | undefined;

async function relationTypesFor(client: AdoWorkItemClient): Promise<RelationTypeInfo[]> {
  const orgKey = organizationKey(client.organizationUrl);
  if (relationTypeCache?.orgKey === orgKey) {
    return relationTypeCache.types;
  }
  const types = await client.listRelationTypes();
  relationTypeCache = { orgKey, types };
  return types;
}

export interface OperationResult {
  message: string;
  details?: string[];
  ok?: boolean;
}

export type ProgressFn = (message: string) => void;

async function clientFromSettings(): Promise<{ client: AdoWorkItemClient; orgUrl: string; settings: ConnectionSettings }> {
  const settings = await loadSettings();
  if (!settings?.orgUrl) {
    throw new Error("Connect to Azure DevOps first.");
  }
  const orgUrl = normalizeOrgUrl(settings.orgUrl);
  if (settings.authKind === "entra") {
    if (!settings.entraClientId || !isGuid(settings.entraClientId)) {
      throw new Error("Enter your Entra application (client) ID, then sign in with Microsoft.");
    }
    const token = await acquireAdoToken(settings.entraClientId);
    return {
      client: new AdoWorkItemClient(orgUrl, { type: "bearer", value: token }),
      orgUrl,
      settings
    };
  }
  if (!settings.pat) {
    throw new Error("Enter a PAT, or sign in with Microsoft.");
  }
  return {
    client: new AdoWorkItemClient(orgUrl, { type: "pat", value: settings.pat }),
    orgUrl,
    settings
  };
}

/** Like clientFromSettings, but refuses to talk to a different org than the one the workbook was loaded from. */
async function clientForWorkbook(
  metadata: WorkbookMetadata
): Promise<{ client: AdoWorkItemClient; orgUrl: string; settings: ConnectionSettings }> {
  const connection = await clientFromSettings();
  if (metadata.orgUrl && !sameOrganization(metadata.orgUrl, connection.orgUrl)) {
    const workbookOrg = normalizeOrgUrl(metadata.orgUrl);
    throw new Error(
      `This workbook was loaded from ${workbookOrg}, but you are connected to ${connection.orgUrl}. ` +
        `Connect to ${workbookOrg} to work with it, or load a query from ${connection.orgUrl} to replace the sheet.`
    );
  }
  return connection;
}

export async function connectWithPat(orgUrl: string, pat: string): Promise<OperationResult> {
  if (!orgUrl.trim() || !pat.trim()) {
    throw new Error("Organization URL and PAT are required.");
  }
  const client = new AdoWorkItemClient(orgUrl, { type: "pat", value: pat.trim() });
  const projects = await client.listProjects();
  const current = (await loadSettings()) || { orgUrl: normalizeOrgUrl(orgUrl), authKind: "pat" as const };
  await saveSettings({
    ...current,
    orgUrl: normalizeOrgUrl(orgUrl),
    authKind: "pat",
    pat: pat.trim()
  });
  processCatalogCache = undefined;
  relationTypeCache = undefined;
  return connectedMessage(projects.length, projects.map((project) => project.name));
}

export async function connectWithEntra(orgUrl: string, clientId: string): Promise<OperationResult> {
  if (!orgUrl.trim()) {
    throw new Error("Organization URL is required.");
  }
  if (!isGuid(clientId)) {
    throw new Error("Enter a valid Entra application (client) ID.");
  }
  const signedIn = await signInWithEntra(clientId.trim());
  const token = await acquireAdoToken(clientId.trim());
  const client = new AdoWorkItemClient(orgUrl, { type: "bearer", value: token });
  const projects = await client.listProjects();
  const current = (await loadSettings()) || { orgUrl: normalizeOrgUrl(orgUrl), authKind: "entra" as const };
  await saveSettings({
    ...current,
    orgUrl: normalizeOrgUrl(orgUrl),
    authKind: "entra",
    entraClientId: clientId.trim(),
    entraUsername: signedIn.username,
    pat: undefined
  });
  processCatalogCache = undefined;
  relationTypeCache = undefined;
  const names = projects.map((project) => project.name);
  const preview = names.slice(0, 8).join(", ");
  return {
    message: `Signed in as ${signedIn.username}. Found ${projects.length} project(s)${preview ? `: ${preview}` : "."}`,
    ok: true
  };
}

export async function disconnectEntra(): Promise<OperationResult> {
  const settings = await loadSettings();
  if (settings?.entraClientId) {
    await signOutEntra(settings.entraClientId);
  }
  if (settings) {
    await saveSettings({ ...settings, authKind: "pat", entraUsername: undefined });
  }
  return { message: "Signed out of Microsoft. You can still connect with a PAT.", ok: true };
}

function connectedMessage(count: number, names: string[]): OperationResult {
  const preview = names.slice(0, 8).join(", ");
  return {
    message: `Connected. Found ${count} project(s)${preview ? `: ${preview}` : "."}`,
    ok: true
  };
}

export async function listProjects() {
  const { client } = await clientFromSettings();
  return client.listProjects();
}

export async function listQueries(project: string) {
  const { client } = await clientFromSettings();
  const queries = await client.listQueries(project);
  return queries.filter((query) => isSupportedQueryId(query.id));
}

export async function searchIdentities(term: string): Promise<IdentityPick[]> {
  const { client } = await clientFromSettings();
  return client.searchIdentities(term);
}

export async function applyIdentityToSelection(identity: IdentityPick): Promise<OperationResult> {
  const selection = await getSheetSelection();
  const metadata = await readMetadata();
  if (!selection || !metadata) {
    throw new Error("Select an Assigned To (or other identity) cell in the Work Items sheet.");
  }
  const column = metadata.columns.find(
    (item) => item.name === selection.columnName || item.referenceName === selection.columnName
  );
  if (!column?.identity && selection.columnName !== "Assigned To" && column?.referenceName !== "System.AssignedTo") {
    throw new Error(`"${selection.columnName}" is not an identity field. Select Assigned To or another person field, then apply.`);
  }
  await setSelectedCell(identityCellValue(identity));
  return { message: `Assigned ${identity.displayName}. Publish to save the change.`, ok: true };
}

async function loadItems(
  client: AdoWorkItemClient,
  orgUrl: string,
  project: string,
  action: WorkbookMetadata["action"],
  columns: WorkbookColumn[],
  ids: number[],
  queryId?: string,
  queryName?: string,
  queryResult?: QueryResultLike,
  forceTree?: boolean
): Promise<OperationResult> {
  if (ids.length === 0) {
    throw new Error("The query or selection returned no work items.");
  }
  if (ids.length > MAX_WORK_ITEMS) {
    throw new Error(`Too many work items (${ids.length}). Narrow the query below ${MAX_WORK_ITEMS}.`);
  }

  let layout = buildTreeLayout(queryResult);
  let displayColumns = markIdentityColumns(columns);
  const fields = fieldReferenceNames(displayColumns);
  const [fetched, catalog] = await Promise.all([
    client.getWorkItems(project, ids, fields),
    loadCatalog(client, project).catch(() => emptyProcessCatalog())
  ]);
  if (catalog.types.length > 0) {
    rememberCatalog(client.organizationUrl, project, catalog);
  }
  if (!layout && forceTree) {
    layout = buildTreeFromParentField(fetched);
  }
  if (layout) {
    displayColumns = markDateColumns(
      applyCatalogToColumns(
        markIdentityColumns(singleTitleColumns(columns), fetched),
        catalog
      ),
      fetched
    );
  } else {
    displayColumns = markDateColumns(
      applyCatalogToColumns(markIdentityColumns(columns, fetched), catalog),
      fetched
    );
  }
  let items = layout ? orderItems(fetched, layout) : fetched;
  if (hasParentTitleColumn(displayColumns)) {
    items = await client.enrichParentTitles(project, items);
  }
  const metadata = buildMetadata({
    orgUrl,
    project,
    action,
    columns: displayColumns,
    items,
    queryId,
    queryName,
    workItemIds: action === "OpenItems" ? ids : undefined,
    listType: layout ? "tree" : "flat",
    layout,
    catalog
  });
  await writeWorkItems(items, displayColumns, metadata, layout);
  const treeNote = layout
    ? " Tree list indents Title and groups rows for parent/child. Indent a row to nest it."
    : "";
  return {
    message: `Loaded ${items.length} work item(s) into Excel.${treeNote}`,
    ok: true
  };
}

export async function loadQuery(project: string, queryId: string, queryName?: string): Promise<OperationResult> {
  const { client, orgUrl } = await clientFromSettings();
  const result = await client.runQuery(project, queryId);
  const columns = mergeColumns(result.columns, undefined);
  return loadItems(
    client,
    orgUrl,
    project,
    "OpenQuery",
    columns,
    idsFromQueryResult(result),
    queryId,
    queryName,
    result
  );
}

export async function refresh(): Promise<OperationResult> {
  const metadata = await readMetadata();
  if (!metadata) {
    throw new Error("This workbook is not bound to Azure DevOps. Load a query first, or open a workbook downloaded from Azure DevOps.");
  }

  const { client, orgUrl } = await clientForWorkbook(metadata);
  let ids = metadata.workItemIds || [];
  let queryResult: QueryResultLike | undefined;
  if (metadata.action === "OpenQuery" && metadata.queryId) {
    queryResult = await client.runQuery(metadata.project, metadata.queryId);
    ids = idsFromQueryResult(queryResult);
  } else {
    const rows = await readTableRows(metadata.columns);
    ids = rows.map((row) => row.id).filter((id): id is number => Boolean(id));
  }

  const baseColumns = metadata.columns.some((column) => column.titleLevel)
    ? mergeColumns(metadata.columns.filter((column) => !column.titleLevel))
    : metadata.columns.length
      ? metadata.columns
      : DEFAULT_COLUMNS;

  return loadItems(
    client,
    orgUrl || metadata.orgUrl,
    metadata.project,
    metadata.action,
    baseColumns,
    ids,
    metadata.queryId,
    metadata.queryName,
    queryResult,
    metadata.listType === "tree"
  );
}

export async function publish(onProgress?: ProgressFn): Promise<OperationResult> {
  const metadata = await requireWorkbook();
  const { client } = await clientForWorkbook(metadata);
  onProgress?.("Reading workbook…");
  const rows = await readTableRows(metadata.columns);
  let catalog: ProcessCatalog;
  try {
    catalog = await loadCatalog(client, metadata.project);
  } catch {
    catalog = catalogFromMetadata(metadata);
  }
  rememberCatalog(client.organizationUrl, metadata.project, catalog);
  const merged: WorkbookMetadata = {
    ...metadata,
    ...catalogFields(catalog),
    workItemTypes: catalog.types.length ? catalog.types : metadata.workItemTypes,
    statesByType: Object.keys(catalog.statesByType).length ? catalog.statesByType : metadata.statesByType,
    areaPaths: catalog.areaPaths.length ? catalog.areaPaths : metadata.areaPaths,
    iterationPaths: catalog.iterationPaths.length ? catalog.iterationPaths : metadata.iterationPaths
  };
  let plan = buildPublishPlan(rows, merged);
  if (plan.errors.length > 0) {
    return {
      message: "Fix the tree list, then publish again.",
      details: plan.errors,
      ok: false
    };
  }

  const updateIds = plan.changes.filter((change) => change.kind === "update" && change.id).map((change) => change.id!);

  if (updateIds.length > 0) {
    onProgress?.("Checking revisions…");
    const current = await client.getWorkItems(metadata.project, updateIds, ["System.Id", "System.Rev"]);
    const serverRevisions: Record<string, number> = {};
    for (const item of current) {
      if (item.id && typeof item.rev === "number") {
        serverRevisions[String(item.id)] = item.rev;
      }
    }
    plan = applyRevisionConflicts(plan, metadata, serverRevisions);
  }

  const details: string[] = [];
  const created: Array<{ rowIndex: number; id: number }> = [];
  const idByRow = new Map<number, number>();
  let published = 0;
  let failed = 0;

  const skipChanges = plan.changes.filter((change) => change.kind === "skip");
  for (const change of skipChanges) {
    if (change.id) {
      idByRow.set(change.rowIndex, change.id);
    }
  }

  for (const change of plan.changes) {
    if (change.kind === "conflict" || change.kind === "error") {
      failed += 1;
      details.push(`Row ${change.rowIndex + 2}: ${change.message}`);
    }
  }

  const publishChanges = plan.changes.filter((change) => change.kind === "create" || change.kind === "update");
  const totalOps = publishChanges.length;
  let finishedOps = 0;
  onProgress?.(totalOps ? `Publishing 0/${totalOps}…` : "Publishing…");

  await mapConcurrent(publishChanges, PUBLISH_CONCURRENCY, async (change) => {
    try {
      if (change.kind === "create" && change.workItemType && change.operations) {
        const createdItem = await client.createWorkItem(metadata.project, change.workItemType, change.operations);
        if (createdItem.id) {
          created.push({ rowIndex: change.rowIndex, id: createdItem.id });
          idByRow.set(change.rowIndex, createdItem.id);
        }
        published += 1;
        details.push(`Row ${change.rowIndex + 2}: created ${createdItem.id ?? ""}`.trim());
      } else if (change.kind === "update" && change.id) {
        if (change.operations && change.operations.length > 0) {
          await client.updateWorkItem(change.id, change.operations);
        }
        idByRow.set(change.rowIndex, change.id);
        published += 1;
        details.push(`#${change.id}: ${change.message}`);
      } else {
        failed += 1;
        details.push(`Row ${change.rowIndex + 2}: ${change.message || "Unable to publish this row."}`);
      }
    } catch (error) {
      failed += 1;
      details.push(`Row ${change.rowIndex + 2}: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      finishedOps += 1;
      onProgress?.(`Publishing ${finishedOps}/${totalOps}…`);
    }
  });

  if (metadata.listType === "tree" || (metadata.treeLevels || 0) > 1) {
    const parentJobs: Array<{ childId: number; parentId: number | null; previousParentId?: number | null }> = [];
    for (const change of plan.changes) {
      const childId = change.id || idByRow.get(change.rowIndex);
      if (!childId || (change.kind !== "create" && change.kind !== "update")) {
        continue;
      }
      let parentId = change.desiredParentId ?? null;
      if (!parentId && change.desiredParentRowIndex !== undefined) {
        parentId = idByRow.get(change.desiredParentRowIndex) ?? null;
      }
      parentJobs.push({ childId, parentId, previousParentId: change.previousParentId });
    }

    if (parentJobs.length > 0) {
      onProgress?.(`Updating parent links (0/${parentJobs.length})…`);
      let done = 0;
      await mapConcurrent(parentJobs, PUBLISH_CONCURRENCY, async (job) => {
        try {
          await client.setParentLink(job.childId, job.parentId, job.previousParentId);
        } catch (error) {
          failed += 1;
          details.push(
            `#${job.childId}: could not update parent (${error instanceof Error ? error.message : String(error)})`
          );
        } finally {
          done += 1;
          onProgress?.(`Updating parent links (${done}/${parentJobs.length})…`);
        }
      });
    }
  }

  await replaceCreatedIds(created);

  if (published > 0 && failed === 0) {
    onProgress?.("Syncing workbook…");
    await syncAfterPublish(client, merged, plan, idByRow);
  } else if (published > 0 && failed > 0) {
    onProgress?.("Updating revisions after partial publish…");
    await syncAfterPublish(client, merged, plan, idByRow);
  }

  const skipped = plan.skipCount;
  return {
    message: `Publish finished. ${published} saved, ${skipped} unchanged, ${failed} issue(s).`,
    details,
    ok: failed === 0
  };
}

async function syncAfterPublish(
  client: AdoWorkItemClient,
  metadata: WorkbookMetadata,
  plan: ReturnType<typeof buildPublishPlan>,
  idByRow: Map<number, number>
): Promise<void> {
  const publishedIds = Array.from(new Set(Array.from(idByRow.values()).filter((id) => id > 0)));
  const fields = fieldReferenceNames(metadata.columns);
  const items =
    publishedIds.length > 0 ? await client.getWorkItems(metadata.project, publishedIds, fields) : [];
  const revisions = { ...metadata.revisions };
  const snapshot = { ...metadata.snapshot };
  for (const item of items) {
    if (!item.id) {
      continue;
    }
    const key = String(item.id);
    if (typeof item.rev === "number") {
      revisions[key] = item.rev;
    }
    snapshot[key] = workItemToSnapshot(item, metadata.columns);
  }

  const parents = { ...metadata.parents };
  const isTree = metadata.listType === "tree" || (metadata.treeLevels || 0) > 1;
  if (isTree) {
    for (const change of plan.changes) {
      const childId = change.id || idByRow.get(change.rowIndex);
      if (!childId) {
        continue;
      }
      let parentId = change.desiredParentId ?? null;
      if (!parentId && change.desiredParentRowIndex !== undefined) {
        parentId = idByRow.get(change.desiredParentRowIndex) ?? null;
      }
      if (parentId) {
        parents[String(childId)] = parentId;
      } else {
        delete parents[String(childId)];
      }
    }
  }

  let rows = await readTableRows(metadata.columns);
  const byId = new Map(items.filter((item) => item.id).map((item) => [item.id as number, item]));
  rows = rows.map((row) => {
    if (!row.id || !byId.has(row.id)) {
      return row;
    }
    const item = byId.get(row.id)!;
    const nextValues = { ...row.values };
    if (typeof item.rev === "number") {
      nextValues["System.Rev"] = item.rev;
    }
    return { ...row, values: nextValues, treeLevel: row.treeLevel || (isTree ? 1 : undefined) };
  });

  if (hasParentTitleColumn(metadata.columns)) {
    if (isTree) {
      rows = fillParentTitlesOnRows(rows);
    } else {
      const titles = new Map(
        rows.filter((row) => row.id).map((row) => [row.id as number, String(row.values["System.Title"] ?? "").trim()])
      );
      rows = rows.map((row) => {
        const parentId = row.id ? parents[String(row.id)] : undefined;
        return {
          ...row,
          values: {
            ...row.values,
            [PARENT_TITLE_FIELD]: parentId ? titles.get(parentId) || "" : ""
          }
        };
      });
    }
  }

  const nextMeta: WorkbookMetadata = {
    ...metadata,
    revisions,
    snapshot,
    parents,
    loadedAt: new Date().toISOString()
  };
  await writeRows(rows, metadata.columns, nextMeta, { skipAutofit: true });
}

export async function convertToTree(): Promise<OperationResult> {
  const metadata = await readMetadata();
  if (!metadata) {
    throw new Error("Load a workbook first.");
  }
  const rows = await readTableRows(metadata.columns);
  const columns = markIdentityColumns(singleTitleColumns(stripTitleLevels(metadata.columns)));
  const nextRows: TableRow[] = rows.map((row) => ({ ...row, treeLevel: row.treeLevel || 1 }));
  const nextMeta: WorkbookMetadata = {
    ...metadata,
    listType: "tree",
    treeLevels: Math.max(MIN_TREE_LEVELS, metadata.treeLevels || MIN_TREE_LEVELS),
    columns
  };
  await writeRows(nextRows, columns, nextMeta);
  return { message: "Converted to a tree list. Indent a row to nest it under the item above.", ok: true };
}

export async function addTreeLevel(): Promise<OperationResult> {
  const metadata = await readMetadata();
  if (!metadata) {
    throw new Error("Load a workbook first.");
  }
  const rows = await readTableRows(metadata.columns);
  const levels = Math.min(
    MAX_TREE_LEVELS,
    Math.max(MIN_TREE_LEVELS, metadata.treeLevels || MIN_TREE_LEVELS) + 1
  );
  const columns = markIdentityColumns(singleTitleColumns(stripTitleLevels(metadata.columns)));
  const nextMeta: WorkbookMetadata = { ...metadata, listType: "tree", treeLevels: levels, columns };
  await writeRows(
    rows.map((row) => ({ ...row, treeLevel: row.treeLevel || 1 })),
    columns,
    nextMeta
  );
  return { message: `You can indent up to ${levels} levels.`, ok: true };
}

export async function indentSelection(delta: number): Promise<OperationResult> {
  const metadata = await readMetadata();
  if (!metadata) {
    throw new Error("Load a workbook first.");
  }
  if (metadata.listType !== "tree") {
    await convertToTree();
  }
  const current = (await readMetadata())!;
  const selection = await getSheetSelection();
  if (!selection) {
    throw new Error("Select a cell on the work item row you want to indent or outdent.");
  }
  let rows = await readTableRows(current.columns);
  const columns = markIdentityColumns(singleTitleColumns(stripTitleLevels(current.columns)));
  let levels = Math.max(
    MIN_TREE_LEVELS,
    current.treeLevels || MIN_TREE_LEVELS,
    ...rows.map((row) => row.treeLevel || 1)
  );
  const currentLevel = rows[selection.rowIndex]?.treeLevel || 1;
  if (delta > 0 && currentLevel >= levels) {
    levels = Math.min(MAX_TREE_LEVELS, levels + 1);
  }
  rows = rows.map((row, index) =>
    index === selection.rowIndex ? shiftTreeLevel(row, delta, levels) : { ...row, treeLevel: row.treeLevel || 1 }
  );
  if (hasParentTitleColumn(columns)) {
    rows = fillParentTitlesOnRows(rows);
  }
  const nextMeta: WorkbookMetadata = { ...current, listType: "tree", treeLevels: levels, columns };
  await writeRows(rows, columns, nextMeta, { skipAutofit: true, skipHyperlinks: true, skipValidation: true });
  return { message: delta > 0 ? "Indented the selected work item." : "Outdented the selected work item.", ok: true };
}

async function requireWorkbook(): Promise<WorkbookMetadata> {
  const metadata = await readMetadata();
  if (!metadata) {
    throw new Error("Load a query first, or open a workbook downloaded from Azure DevOps.");
  }
  return metadata;
}

async function ensureProcessCatalog(metadata: WorkbookMetadata): Promise<ProcessCatalog> {
  const cached = metadata.orgUrl ? cachedCatalog(metadata.orgUrl, metadata.project) : undefined;
  if (cached) {
    return cached;
  }
  const { client } = await clientForWorkbook(metadata);
  const catalog = await loadCatalog(client, metadata.project);
  const next: WorkbookMetadata = {
    ...metadata,
    ...catalogFields(catalog),
    identities: catalog.identities.length ? catalog.identities : metadata.identities,
    areaPaths: catalog.areaPaths.length ? catalog.areaPaths : metadata.areaPaths,
    iterationPaths: catalog.iterationPaths.length ? catalog.iterationPaths : metadata.iterationPaths
  };
  await writeMetadata(next);
  try {
    await reapplyProcessValidation(next);
  } catch {
    // Dropdowns are best-effort.
  }
  return catalog;
}

async function loadCatalog(client: AdoWorkItemClient, project: string): Promise<ProcessCatalog> {
  const cached = cachedCatalog(client.organizationUrl, project);
  if (cached) {
    return cached;
  }
  const catalog = await client.loadProcessCatalog(project);
  rememberCatalog(client.organizationUrl, project, catalog);
  return catalog;
}

function rememberCatalog(orgUrl: string, project: string, catalog: ProcessCatalog): void {
  processCatalogCache = { orgKey: organizationKey(orgUrl), project, catalog };
  setClassificationPaths({ areaPaths: catalog.areaPaths, iterationPaths: catalog.iterationPaths });
  setIdentityLabels(catalog.identities);
}

export async function listWorkbookWorkItemTypes(): Promise<string[]> {
  const metadata = await readMetadata();
  if (!metadata) {
    return [];
  }
  try {
    return (await ensureProcessCatalog(metadata)).types;
  } catch {
    return metadata.workItemTypes || [];
  }
}

export async function listCurrentColumns(): Promise<Array<{ referenceName: string; name: string; locked: boolean }>> {
  const metadata = await readMetadata();
  if (!metadata) {
    return [];
  }
  return metadata.columns
    .filter((column) => !column.titleLevel)
    .map((column) => ({
      referenceName: column.referenceName,
      name: column.name,
      locked: isLockedColumn(column)
    }));
}

export async function listAddableColumns(): Promise<ProcessFieldInfo[]> {
  const metadata = await requireWorkbook();
  const catalog = await ensureProcessCatalog(metadata);
  return addableProcessFields(catalog.fields, metadata.columns);
}

export async function insertWorkItem(workItemType: string): Promise<OperationResult> {
  const type = workItemType.trim();
  if (!type) {
    throw new Error("Choose a work item type from this project's process template.");
  }
  const metadata = await requireWorkbook();
  const catalog = await ensureProcessCatalog(metadata);
  if (!catalog.types.some((name) => name.toLowerCase() === type.toLowerCase())) {
    throw new Error(`"${type}" is not a work item type in this project's process template.`);
  }
  const rows = await readTableRows(metadata.columns);
  const selection = await getSheetSelection();
  const atIndex = selection ? selection.rowIndex + 1 : rows.length;
  const next = insertWorkItemRow(
    rows,
    type,
    metadata.columns,
    atIndex,
    metadata.listType === "tree" || (metadata.treeLevels || 0) > 1,
    defaultStateForType(type, catalog.statesByType)
  );
  const nextMeta: WorkbookMetadata = {
    ...metadata,
    ...catalogFields(catalog),
    areaPaths: catalog.areaPaths.length ? catalog.areaPaths : metadata.areaPaths,
    iterationPaths: catalog.iterationPaths.length ? catalog.iterationPaths : metadata.iterationPaths
  };
  await writeRows(next, metadata.columns, nextMeta);
  return { message: `Inserted a new ${type} row. Set Title, then publish.`, ok: true };
}

/**
 * Applies a staged column layout with one sheet rewrite. New columns are filled from Azure DevOps
 * and recorded in the snapshot so Publish doesn't treat the fetched values as edits.
 */
export async function applyColumnLayout(order: string[]): Promise<OperationResult> {
  const metadata = await requireWorkbook();
  let catalog: ProcessCatalog;
  try {
    catalog = await ensureProcessCatalog(metadata);
  } catch {
    catalog = catalogFromMetadata(metadata);
  }
  const change = buildColumnLayout(metadata.columns, order, addableProcessFields(catalog.fields, metadata.columns));
  if (change.added.length === 0 && change.removed.length === 0 && !change.reordered) {
    return { message: "No column changes to apply.", ok: true };
  }

  const rows = await readTableRows(metadata.columns);
  const snapshot = { ...metadata.snapshot };
  let fetched: WorkItemLike[] = [];
  const addedFields = change.added.map((column) => column.referenceName).filter((name) => name !== PARENT_TITLE_FIELD);
  const addsParentTitle = change.added.some((column) => column.referenceName === PARENT_TITLE_FIELD);
  const ids = rows.map((row) => row.id).filter((id): id is number => Boolean(id && id > 0));

  if ((addedFields.length > 0 && ids.length > 0) || addsParentTitle) {
    const { client } = await clientForWorkbook(metadata);
    if (addedFields.length > 0 && ids.length > 0) {
      fetched = await client.getWorkItems(metadata.project, ids, ["System.Id", ...addedFields]);
    }
    if (addsParentTitle) {
      await fillParentTitleColumn(client, metadata.project, rows);
    }
  }

  const htmlFields = change.added.filter((column) => column.html).map((column) => column.referenceName);
  const byId = new Map(fetched.filter((item) => item.id).map((item) => [item.id as number, item]));
  for (const row of rows) {
    const item = row.id ? byId.get(row.id) : undefined;
    for (const referenceName of addedFields) {
      const value = item ? fieldToCell(item.fields?.[referenceName], referenceName, htmlFields) : "";
      row.values[referenceName] = value;
      if (row.id && snapshot[String(row.id)]) {
        snapshot[String(row.id)] = { ...snapshot[String(row.id)], [referenceName]: value };
      }
    }
    for (const column of change.removed) {
      delete row.values[column.referenceName];
    }
  }

  const columns = markDateColumns(
    applyCatalogToColumns(markIdentityColumns(change.columns, fetched), catalog),
    fetched
  );
  await writeRows(rows, columns, {
    ...metadata,
    ...catalogFields(catalog),
    identities: catalog.identities.length ? catalog.identities : metadata.identities,
    areaPaths: catalog.areaPaths.length ? catalog.areaPaths : metadata.areaPaths,
    iterationPaths: catalog.iterationPaths.length ? catalog.iterationPaths : metadata.iterationPaths,
    columns,
    snapshot
  });

  const parts: string[] = [];
  if (change.added.length > 0) {
    parts.push(`added ${change.added.length}`);
  }
  if (change.removed.length > 0) {
    parts.push(`removed ${change.removed.length}`);
  }
  if (change.reordered) {
    parts.push("reordered");
  }
  return { message: `Columns updated: ${parts.join(", ")}.`, ok: true };
}

async function fillParentTitleColumn(
  client: AdoWorkItemClient,
  project: string,
  rows: TableRow[]
): Promise<void> {
  const ids = rows.map((row) => row.id).filter((id): id is number => Boolean(id && id > 0));
  if (ids.length === 0) {
    for (const row of rows) {
      row.values[PARENT_TITLE_FIELD] = "";
    }
    return;
  }
  const fetched = await client.getWorkItems(project, ids, ["System.Id", "System.Parent", "System.Title"]);
  const enriched = await client.enrichParentTitles(project, fetched);
  const byId = new Map(enriched.filter((item) => item.id).map((item) => [item.id as number, item]));
  for (const row of rows) {
    if (!row.id) {
      row.values[PARENT_TITLE_FIELD] = "";
      continue;
    }
    const item = byId.get(row.id);
    row.values[PARENT_TITLE_FIELD] = String(item?.fields?.[PARENT_TITLE_FIELD] || "");
  }
}

export async function listClassificationPaths(): Promise<{ areaPaths: string[]; iterationPaths: string[] }> {
  const metadata = await readMetadata();
  if (!metadata) {
    return { areaPaths: [], iterationPaths: [] };
  }
  try {
    const catalog = await ensureProcessCatalog(metadata);
    return { areaPaths: catalog.areaPaths, iterationPaths: catalog.iterationPaths };
  } catch {
    return { areaPaths: metadata.areaPaths || [], iterationPaths: metadata.iterationPaths || [] };
  }
}

export async function applyPathToSelection(kind: "area" | "iteration", path: string): Promise<OperationResult> {
  const value = path.trim();
  if (!value) {
    throw new Error(kind === "area" ? "Choose an area path." : "Choose an iteration path.");
  }
  const selection = await getSheetSelection();
  const metadata = await readMetadata();
  if (!selection || !metadata) {
    throw new Error("Select an Area Path or Iteration Path cell in the Work Items sheet.");
  }
  const column = metadata.columns.find(
    (item) => item.name === selection.columnName || item.referenceName === selection.columnName
  );
  const expected =
    kind === "area"
      ? ["Area Path", "System.AreaPath"]
      : ["Iteration Path", "System.IterationPath"];
  if (!column || (!expected.includes(column.name) && !expected.includes(column.referenceName))) {
    throw new Error(
      kind === "area"
        ? `Select an Area Path cell, then apply.`
        : `Select an Iteration Path cell, then apply.`
    );
  }
  await setSelectedCell(value);
  return { message: `Set ${column.name} to ${value}. Publish to save the change.`, ok: true };
}

export async function workbookSummary(): Promise<string> {
  const metadata = await readMetadata();
  if (!metadata) {
    return "No Azure DevOps workbook is loaded.";
  }
  const query = metadata.queryName ? `Query: ${metadata.queryName}. ` : "";
  const tree = metadata.listType === "tree" ? "Tree list. " : "Flat list. ";
  const org = metadata.orgUrl ? normalizeOrgUrl(metadata.orgUrl) : "";
  const settings = await loadSettings();
  const mismatch =
    org && settings?.orgUrl && !sameOrganization(org, settings.orgUrl)
      ? ` You are connected to ${normalizeOrgUrl(settings.orgUrl)}, so Refresh and Publish are blocked for this sheet until you connect to ${org} or load a query.`
      : "";
  const orgLabel = org ? `Organization: ${org}. ` : "";
  return `${query}${tree}${orgLabel}Project: ${metadata.project}. Last loaded: ${new Date(metadata.loadedAt).toLocaleString()}.${mismatch}`;
}

export async function ensureTreePresentation(): Promise<void> {
  const metadata = await readMetadata();
  if (!metadata || metadata.listType !== "tree") {
    return;
  }
  try {
    await restoreTreePresentation(metadata);
  } catch {
    // Indent and outline are best-effort.
  }
}

export async function selectedIdentityHint(): Promise<string> {
  const metadata = await readMetadata();
  const selection = await getSheetSelection();
  if (!metadata || !selection) {
    return "Select an Assigned To cell, search, then apply.";
  }
  const column = metadata.columns.find(
    (item) => item.name === selection.columnName || item.referenceName === selection.columnName
  );
  if (column?.identity || selection.columnName === "Assigned To") {
    return `Selected ${selection.columnName}. Use the cell dropdown, or search below if the person is not listed.`;
  }
  if (column?.referenceName === "System.AreaPath" || selection.columnName === "Area Path") {
    return `Selected Area Path. Choose a node in the tree and apply.`;
  }
  if (column?.referenceName === "System.IterationPath" || selection.columnName === "Iteration Path") {
    return `Selected Iteration Path. Choose a node in the tree and apply.`;
  }
  return `Selected ${selection.columnName}.`;
}

function stripTitleLevels(columns: WorkbookColumn[]): WorkbookColumn[] {
  const rest = columns.filter((column) => !column.titleLevel);
  if (!rest.some((column) => column.referenceName === "System.Title")) {
    return [
      ...rest.slice(0, 3),
      { referenceName: "System.Title", name: "Title" },
      ...rest.slice(3)
    ];
  }
  return rest;
}

export async function signedInLabel(): Promise<string> {
  const settings = await loadSettings();
  if (!settings) {
    return "Not connected.";
  }
  if (settings.authKind === "entra") {
    const user =
      settings.entraUsername ||
      (settings.entraClientId ? await currentEntraUsername(settings.entraClientId) : undefined);
    return `${user ? `Microsoft: ${user}` : "Microsoft: signed in"} (${normalizeOrgUrl(settings.orgUrl)})`;
  }
  return `Connected to ${normalizeOrgUrl(settings.orgUrl)} with PAT.`;
}

export function openExternalUrl(url: string): void {
  const target = url.trim();
  if (!target) {
    throw new Error("Nothing to open.");
  }
  if (typeof Office !== "undefined" && Office.context?.ui?.openBrowserWindow) {
    Office.context.ui.openBrowserWindow(target);
    return;
  }
  window.open(target, "_blank", "noopener,noreferrer");
}

export async function selectedWorkItemId(): Promise<number> {
  const metadata = await requireWorkbook();
  const selection = await getSheetSelection();
  if (!selection) {
    throw new Error("Select a cell on the work item row you want to open.");
  }
  const rows = await readTableRows(metadata.columns);
  const id = rows[selection.rowIndex]?.id;
  if (!id) {
    throw new Error("Select a row that already has a work item ID. Publish new rows first.");
  }
  return id;
}

function workItemUrl(metadata: WorkbookMetadata, id: number, orgUrl?: string): string {
  return workItemWebUrl(orgUrl || metadata.orgUrl, metadata.project, id);
}

export async function workbookQueryLink(): Promise<{ name: string; url: string } | undefined> {
  const metadata = await readMetadata();
  if (!metadata?.queryId || !isSupportedQueryId(metadata.queryId) || !metadata.orgUrl || !metadata.project) {
    return undefined;
  }
  return {
    name: metadata.queryName || "the query",
    url: queryWebUrl(metadata.orgUrl, metadata.project, metadata.queryId)
  };
}

export async function openWorkbookQuery(): Promise<OperationResult> {
  const link = await workbookQueryLink();
  if (!link) {
    throw new Error("This sheet wasn't loaded from a saved query, so there is no query to open.");
  }
  openExternalUrl(link.url);
  return { message: `Opened ${link.name} in Azure DevOps.`, ok: true };
}

export async function openSelectedWorkItem(): Promise<OperationResult> {
  const metadata = await requireWorkbook();
  const id = await selectedWorkItemId();
  let orgUrl = metadata.orgUrl;
  if (!orgUrl) {
    try {
      orgUrl = (await clientFromSettings()).orgUrl;
    } catch {
      // Falls through to the "connect first" error below.
    }
  }
  if (!orgUrl) {
    throw new Error("Connect to Azure DevOps first, or open a workbook downloaded from Azure DevOps.");
  }
  openExternalUrl(workItemUrl(metadata, id, orgUrl));
  return { message: `Opened work item ${id} in Azure DevOps.`, ok: true };
}

export async function loadSelectedWorkItemDetails(): Promise<WorkItemDetails> {
  const metadata = await requireWorkbook();
  const { client } = await clientForWorkbook(metadata);
  const id = await selectedWorkItemId();
  const [details, types] = await Promise.all([
    client.loadWorkItemDetails(metadata.project, id),
    relationTypesFor(client).catch((): RelationTypeInfo[] => [])
  ]);
  return withRelationNames(details, relationTypeNames(types));
}

/** Link types for the Add link picker, from the connected organization; the basic three if they can't be loaded. */
export async function listLinkTypes(): Promise<LinkTypeOption[]> {
  try {
    const { client } = await clientFromSettings();
    return addableLinkTypes(await relationTypesFor(client));
  } catch {
    return DEFAULT_LINK_TYPES.slice();
  }
}

function linkRejected(error: unknown, what: string): unknown {
  if (error instanceof AdoClientError && error.status !== undefined && error.status >= 400 && error.status < 500) {
    return new Error(`Azure DevOps didn't add the ${what}: ${adoErrorDetail(error) || error.message}`);
  }
  return error;
}

/** The selected row's work item, refusing if the task pane is showing details for a different one. */
async function detailsTargetId(shownId?: number): Promise<number> {
  const id = await selectedWorkItemId();
  if (shownId !== undefined && id !== shownId) {
    throw new Error(
      `The selected row is #${id}, but the details shown are for #${shownId}. Select the row for #${shownId}, or click Load details to work with #${id}.`
    );
  }
  return id;
}

export async function addRelatedLink(targetId: number, rel: string, shownId?: number): Promise<OperationResult> {
  if (!Number.isFinite(targetId) || targetId <= 0) {
    throw new Error("Enter a valid work item ID to link.");
  }
  if (!rel) {
    throw new Error("Choose a link type.");
  }
  if (isHierarchyLink(rel)) {
    throw new Error("Parent and child links come from the sheet's hierarchy. Use Indent/Outdent, then Publish.");
  }
  const id = await detailsTargetId(shownId);
  if (targetId === id) {
    throw new Error("A work item cannot be linked to itself.");
  }
  const { client } = await clientForWorkbook(await requireWorkbook());
  try {
    await client.addWorkItemLink(id, rel, targetId);
  } catch (error) {
    throw linkRejected(error, `link to #${targetId}`);
  }
  return { message: `Linked #${targetId} to #${id}.`, ok: true };
}

export async function addWorkItemHyperlink(url: string, comment: string, shownId?: number): Promise<OperationResult> {
  const target = normalizeHyperlinkUrl(url);
  const id = await detailsTargetId(shownId);
  const { client } = await clientForWorkbook(await requireWorkbook());
  try {
    await client.addWorkItemHyperlink(id, target, comment);
  } catch (error) {
    throw linkRejected(error, "hyperlink");
  }
  return { message: `Added a hyperlink to #${id}.`, ok: true };
}

export async function removeWorkItemRelation(
  index: number,
  shown: { id: number; rev?: number }
): Promise<OperationResult> {
  if (!Number.isFinite(index) || index < 0) {
    throw new Error("Choose a related link or attachment to remove.");
  }
  const id = await detailsTargetId(shown.id);
  const { client } = await clientForWorkbook(await requireWorkbook());
  try {
    await client.removeWorkItemRelation(id, index, shown.rev);
  } catch (error) {
    if (shown.rev !== undefined && error instanceof AdoClientError && error.status !== undefined && error.status < 500) {
      throw new Error(
        `#${id} may have changed since its details were loaded, so nothing was removed. Click Load details and try again. (${error.message})`
      );
    }
    throw error;
  }
  return { message: `Removed the selected link from #${id}.`, ok: true };
}

export async function addWorkItemAttachment(fileName: string, data: ArrayBuffer, shownId?: number): Promise<OperationResult> {
  const name = fileName.trim();
  if (!name) {
    throw new Error("Choose a file to attach.");
  }
  const metadata = await requireWorkbook();
  const id = await detailsTargetId(shownId);
  const { client } = await clientForWorkbook(metadata);
  await client.addWorkItemAttachment(metadata.project, id, name, data);
  return { message: `Attached ${name} to #${id}.`, ok: true };
}

export async function addWorkItemComment(text: string, shownId?: number): Promise<OperationResult> {
  const comment = text.trim();
  if (!comment) {
    throw new Error("Enter a comment.");
  }
  const metadata = await requireWorkbook();
  const id = await detailsTargetId(shownId);
  const { client } = await clientForWorkbook(metadata);
  await client.addWorkItemComment(metadata.project, id, comment);
  return { message: `Added a comment on #${id}.`, ok: true };
}

export async function openWorkItemLink(url: string, rel?: string): Promise<OperationResult> {
  if (!/^https?:\/\//i.test(url.trim())) {
    throw new Error(
      rel === HYPERLINK_LINK
        ? "Only web (http or https) hyperlinks can be opened from Excel. Open the work item in Azure DevOps to follow this one."
        : "This link points to a commit, branch, pull request, build, or other Azure DevOps artifact, which can't be opened from Excel. Open the work item in Azure DevOps to follow it."
    );
  }
  openExternalUrl(url);
  return { message: "Opened in the browser.", ok: true };
}
