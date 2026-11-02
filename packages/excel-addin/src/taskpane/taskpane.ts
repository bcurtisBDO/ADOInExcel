import { loadSettings } from "../auth";
import {
  addTreeLevel,
  applyColumnLayout,
  applyIdentityToSelection,
  applyPathToSelection,
  connectWithEntra,
  connectWithPat,
  convertToTree,
  disconnectEntra,
  ensureTreePresentation,
  indentSelection,
  insertWorkItem,
  addRelatedLink,
  addWorkItemAttachment,
  addWorkItemComment,
  addWorkItemHyperlink,
  listLinkTypes,
  listProjects,
  listQueries,
  listAddableColumns,
  listClassificationPaths,
  listCurrentColumns,
  listWorkbookWorkItemTypes,
  loadQuery,
  loadSelectedWorkItemDetails,
  openSelectedWorkItem,
  openWorkbookQuery,
  openWorkItemLink,
  publish,
  refresh,
  removeWorkItemRelation,
  searchIdentities,
  selectedIdentityHint,
  signedInLabel,
  workbookQueryLink,
  workbookSummary
} from "../operations";
import {
  DEFAULT_LINK_TYPES,
  IdentityPick,
  LinkTypeOption,
  QueryPickerNode,
  WorkItemDetails,
  WorkItemLinkInfo,
  indentPathLabel,
  queryPickerTree
} from "@ado-open-in-excel/shared";

let identityHits: IdentityPick[] = [];
let searchTimer: number | undefined;
let selectedQuery: { id: string; name: string } | undefined;
let areaPaths: string[] = [];
let iterationPaths: string[] = [];
let selectedDetails: WorkItemDetails | null = null;

function $(id: string): HTMLElement {
  const element = document.getElementById(id);
  if (!element) {
    throw new Error(`Missing element #${id}`);
  }
  return element;
}

function input(id: string): HTMLInputElement {
  return $(id) as HTMLInputElement;
}

function select(id: string): HTMLSelectElement {
  return $(id) as HTMLSelectElement;
}

const COLLAPSED_SECTIONS_KEY = "adoOpenInExcel.collapsedSections";

function loadCollapsedSections(): Set<string> {
  try {
    const raw = window.localStorage.getItem(COLLAPSED_SECTIONS_KEY);
    const parsed = raw ? (JSON.parse(raw) as unknown) : [];
    return new Set(Array.isArray(parsed) ? parsed.filter((id) => typeof id === "string") : []);
  } catch {
    return new Set();
  }
}

function saveCollapsedSections(): void {
  const collapsed = Array.from(document.querySelectorAll<HTMLElement>(".card.collapsed[data-section]"))
    .map((card) => card.dataset.section)
    .filter((id): id is string => Boolean(id));
  window.localStorage.setItem(COLLAPSED_SECTIONS_KEY, JSON.stringify(collapsed));
}

function setCardCollapsed(card: HTMLElement, collapsed: boolean): void {
  const button = card.querySelector<HTMLButtonElement>(".card-toggle");
  card.classList.toggle("collapsed", collapsed);
  if (button) {
    button.setAttribute("aria-expanded", collapsed ? "false" : "true");
  }
}

function setupCollapsibleCards(): void {
  const collapsed = loadCollapsedSections();
  document.querySelectorAll<HTMLElement>(".card[data-section]").forEach((card) => {
    const id = card.dataset.section;
    const button = card.querySelector<HTMLButtonElement>(".card-toggle");
    if (!id || !button) {
      return;
    }
    setCardCollapsed(card, collapsed.has(id));
    button.addEventListener("click", () => {
      setCardCollapsed(card, !card.classList.contains("collapsed"));
      saveCollapsedSections();
    });
  });
}

function setStatus(message: string, kind: "ok" | "error" | "" = "", details: string[] = []): void {
  const status = $("status");
  status.textContent = message;
  status.className = `status ${kind}`.trim();
  const list = $("details") as HTMLUListElement;
  list.innerHTML = "";
  const shown = details.slice(0, 100);
  for (const detail of shown) {
    const item = document.createElement("li");
    item.textContent = detail;
    list.appendChild(item);
  }
  if (details.length > shown.length) {
    const more = document.createElement("li");
    more.className = "muted";
    more.textContent = `…and ${details.length - shown.length} more`;
    list.appendChild(more);
  }
}

async function withBusy(ids: string[], work: () => Promise<void>): Promise<void> {
  const buttons = ids.map((id) => $(id) as HTMLButtonElement);
  buttons.forEach((button) => {
    button.disabled = true;
  });
  try {
    await work();
  } catch (error) {
    setStatus(error instanceof Error ? error.message : String(error), "error");
  } finally {
    buttons.forEach((button) => {
      button.disabled = false;
    });
  }
}

function fillSelect(element: HTMLSelectElement, items: Array<{ value: string; label: string; disabled?: boolean }>): void {
  const previous = element.value;
  element.innerHTML = "";
  if (items.length === 0) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "None found";
    element.appendChild(option);
    return;
  }
  for (const item of items) {
    const option = document.createElement("option");
    option.value = item.disabled ? "" : item.value;
    option.textContent = item.label;
    option.disabled = Boolean(item.disabled);
    element.appendChild(option);
  }
  if (previous && items.some((item) => !item.disabled && item.value === previous)) {
    element.value = previous;
    return;
  }
  const first = items.find((item) => !item.disabled && item.value);
  if (first) {
    element.value = first.value;
  }
}

function fillPathSelect(element: HTMLSelectElement, paths: string[], filter: string): void {
  const needle = filter.trim().toLowerCase();
  const matched = needle ? paths.filter((path) => path.toLowerCase().includes(needle)) : paths;
  fillSelect(
    element,
    matched.map((path) => ({ value: path, label: indentPathLabel(path) }))
  );
}

function fillLinkTypeSelect(options: LinkTypeOption[]): void {
  const element = select("relatedLinkType");
  const previous = element.value;
  element.innerHTML = "";
  const groups = new Map<string, HTMLOptGroupElement>();
  for (const option of options) {
    let group = groups.get(option.group);
    if (!group) {
      group = document.createElement("optgroup");
      group.label = option.group;
      groups.set(option.group, group);
      element.appendChild(group);
    }
    const item = document.createElement("option");
    item.value = option.rel;
    item.textContent = option.label;
    group.appendChild(item);
  }
  if (previous && options.some((option) => option.rel === previous)) {
    element.value = previous;
  }
}

async function refreshLinkTypes(generation = connectionGeneration): Promise<void> {
  const options = await listLinkTypes();
  if (generation === connectionGeneration) {
    fillLinkTypeSelect(options);
  }
}

function linkOptionLabel(link: WorkItemLinkInfo): string {
  if (link.id) {
    return `${link.label} #${link.id}`;
  }
  if (link.name && link.name !== link.label) {
    return `${link.label}: ${link.name}`;
  }
  return link.label;
}

function renderSelectedDetails(details: WorkItemDetails | null): void {
  selectedDetails = details;
  const heading = $("selectedWorkItem");
  const parentLine = $("selectedParent");
  if (!details) {
    heading.textContent =
      "Select a row with an ID. Click an ID cell, use View in Browser on the Home ribbon or the cell menu, or load details here.";
    parentLine.hidden = true;
    parentLine.textContent = "";
    fillSelect(select("relatedLinks"), []);
    fillSelect(select("attachments"), []);
    $("discussionList").innerHTML = "";
    return;
  }
  heading.textContent = `#${details.id} ${details.title || "(no title)"}`;
  if (details.parentId) {
    const name = details.parentTitle?.trim() || "(no title)";
    parentLine.textContent = `Parent: #${details.parentId} ${name}`;
    parentLine.hidden = false;
  } else {
    parentLine.textContent = "Parent: none";
    parentLine.hidden = false;
  }
  fillSelect(
    select("relatedLinks"),
    details.links.map((link) => ({ value: String(link.index), label: linkOptionLabel(link) }))
  );
  fillSelect(
    select("attachments"),
    details.attachments.map((link) => ({ value: String(link.index), label: linkOptionLabel(link) }))
  );
  const list = $("discussionList");
  list.innerHTML = "";
  if (details.comments.length === 0) {
    const empty = document.createElement("li");
    empty.className = "muted";
    empty.textContent = "No comments yet.";
    list.appendChild(empty);
    return;
  }
  for (const comment of details.comments) {
    const item = document.createElement("li");
    const meta = document.createElement("div");
    meta.className = "meta";
    const when = comment.createdDate ? new Date(comment.createdDate).toLocaleString() : "";
    meta.textContent = when ? `${comment.createdBy} · ${when}` : comment.createdBy;
    const body = document.createElement("p");
    body.textContent = comment.text;
    item.appendChild(meta);
    item.appendChild(body);
    list.appendChild(item);
  }
}

async function reloadSelectedDetails(): Promise<void> {
  const details = await loadSelectedWorkItemDetails();
  renderSelectedDetails(details);
}

function selectedRelationIndex(selectId: string): number {
  const raw = select(selectId).value;
  if (!raw) {
    throw new Error("Choose an item in the list.");
  }
  const value = Number(raw);
  if (!Number.isFinite(value) || value < 0) {
    throw new Error("Choose an item in the list.");
  }
  return value;
}

function shownDetails(): WorkItemDetails {
  if (!selectedDetails) {
    throw new Error("Click Load details for the selected work item first.");
  }
  return selectedDetails;
}

function selectedLink(kind: "links" | "attachments"): WorkItemLinkInfo {
  const index = selectedRelationIndex(kind === "links" ? "relatedLinks" : "attachments");
  const link = selectedDetails?.[kind].find((item) => item.index === index);
  if (!link) {
    throw new Error("Load details for the selected work item first.");
  }
  return link;
}

/** Bumped on every Connect / Sign in so slower lookups for the previous org are discarded. */
let connectionGeneration = 0;
let queryListRequest = 0;

function clearProjectList(): void {
  fillSelect(select("project"), []);
  selectedQuery = undefined;
  $("queryTree").innerHTML = "";
  fillLinkTypeSelect(DEFAULT_LINK_TYPES);
}

async function refreshProjectList(generation = connectionGeneration): Promise<boolean> {
  const projects = await listProjects();
  if (generation !== connectionGeneration) {
    return false;
  }
  fillSelect(
    select("project"),
    projects.map((project) => ({ value: project.name, label: project.name }))
  );
  void refreshLinkTypes(generation);
  await refreshQueryList(generation);
  return generation === connectionGeneration;
}

async function refreshQueryList(generation = connectionGeneration): Promise<void> {
  const request = ++queryListRequest;
  const project = select("project").value;
  selectedQuery = undefined;
  const tree = $("queryTree");
  tree.innerHTML = "";
  if (!project) {
    return;
  }
  const queries = await listQueries(project);
  if (request !== queryListRequest || generation !== connectionGeneration) {
    return;
  }
  const nodes = queryPickerTree(queries);
  const list = document.createElement("ul");
  for (const node of nodes) {
    list.appendChild(renderQueryNode(node, true));
  }
  tree.appendChild(list);
}

function renderQueryNode(node: QueryPickerNode, expanded: boolean): HTMLLIElement {
  const item = document.createElement("li");
  item.setAttribute("role", "treeitem");
  if (node.folder) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "tree-folder";
    button.setAttribute("aria-expanded", expanded ? "true" : "false");
    button.textContent = node.name;
    const children = document.createElement("ul");
    children.hidden = !expanded;
    children.setAttribute("role", "group");
    for (const child of node.children) {
      children.appendChild(renderQueryNode(child, false));
    }
    button.addEventListener("click", () => {
      const open = button.getAttribute("aria-expanded") === "true";
      button.setAttribute("aria-expanded", open ? "false" : "true");
      children.hidden = open;
    });
    item.appendChild(button);
    item.appendChild(children);
    return item;
  }
  const button = document.createElement("button");
  button.type = "button";
  button.className = "tree-query";
  button.dataset.queryId = node.id;
  button.textContent = node.name;
  button.addEventListener("click", () => {
    selectedQuery = { id: node.id, name: node.name };
    document.querySelectorAll(".tree-query").forEach((element) => element.classList.remove("selected"));
    button.classList.add("selected");
  });
  item.appendChild(button);
  return item;
}

async function refreshWorkbookStatus(): Promise<void> {
  $("workbookStatus").textContent = await workbookSummary();
  const queryLink = await workbookQueryLink().catch(() => undefined);
  const openQueryButton = $("openQuery") as HTMLButtonElement;
  openQueryButton.hidden = !queryLink;
  openQueryButton.title = queryLink ? `Open ${queryLink.name} in Azure DevOps` : "";
  $("authStatus").textContent = await signedInLabel();
  $("identityHint").textContent = await selectedIdentityHint();
  await refreshChooserLists();
}

interface DraftColumn {
  referenceName: string;
  name: string;
  locked: boolean;
}

interface FieldChoice {
  referenceName: string;
  name: string;
}

/** Column edits are staged here and written to the sheet only on Apply changes. */
let savedColumns: DraftColumn[] = [];
let draftColumns: DraftColumn[] = [];
let addableFields: FieldChoice[] = [];

function sameColumnOrder(left: DraftColumn[], right: DraftColumn[]): boolean {
  return left.length === right.length && left.every((column, index) => column.referenceName === right[index].referenceName);
}

function columnDraftPending(): boolean {
  return !sameColumnOrder(draftColumns, savedColumns);
}

function syncColumnDraft(current: DraftColumn[], addable: FieldChoice[]): void {
  const sheetChanged = !sameColumnOrder(current, savedColumns);
  savedColumns = current;
  addableFields = addable;
  if (sheetChanged || !columnDraftPending()) {
    draftColumns = current.map((column) => ({ ...column }));
  }
  renderColumnDraft();
}

function availableForDraft(): FieldChoice[] {
  const inDraft = new Set(draftColumns.map((column) => column.referenceName));
  const removedFromSheet = savedColumns
    .filter((column) => !column.locked && !inDraft.has(column.referenceName))
    .map(({ referenceName, name }) => ({ referenceName, name }));
  return [...addableFields.filter((field) => !inDraft.has(field.referenceName)), ...removedFromSheet].sort((left, right) =>
    left.name.localeCompare(right.name)
  );
}

function columnDraftSummary(): string {
  const savedRefs = new Set(savedColumns.map((column) => column.referenceName));
  const draftRefs = new Set(draftColumns.map((column) => column.referenceName));
  const added = draftColumns.filter((column) => !savedRefs.has(column.referenceName)).length;
  const removed = savedColumns.filter((column) => !draftRefs.has(column.referenceName)).length;
  const keptDraft = draftColumns.filter((column) => savedRefs.has(column.referenceName));
  const keptSaved = savedColumns.filter((column) => draftRefs.has(column.referenceName));
  const parts: string[] = [];
  if (added > 0) {
    parts.push(`${added} added`);
  }
  if (removed > 0) {
    parts.push(`${removed} removed`);
  }
  if (!sameColumnOrder(keptDraft, keptSaved)) {
    parts.push("order changed");
  }
  return parts.length > 0
    ? `Pending: ${parts.join(", ")}. Click Apply changes to update the sheet.`
    : "No pending column changes.";
}

function renderAddableFields(): void {
  const element = select("addableColumns");
  const chosen = new Set(Array.from(element.selectedOptions).map((option) => option.value));
  const needle = input("addableFilter").value.trim().toLowerCase();
  const fields = availableForDraft().filter((field) => !needle || field.name.toLowerCase().includes(needle));
  element.innerHTML = "";
  for (const field of fields) {
    const option = document.createElement("option");
    option.value = field.referenceName;
    option.textContent = field.name;
    option.selected = chosen.has(field.referenceName);
    element.appendChild(option);
  }
}

function renderColumnDraft(selectReference?: string): void {
  const savedRefs = new Set(savedColumns.map((column) => column.referenceName));
  const list = select("currentColumns");
  fillSelect(
    list,
    draftColumns.map((column) => ({
      value: column.referenceName,
      label: `${column.name}${column.locked ? " (required)" : ""}${savedRefs.has(column.referenceName) ? "" : " (new)"}`
    }))
  );
  if (selectReference) {
    list.value = selectReference;
  }
  renderAddableFields();
  const status = $("columnDraftStatus");
  status.textContent = columnDraftSummary();
  status.classList.toggle("pending", columnDraftPending());
}

function selectedDraftIndex(): number {
  const index = draftColumns.findIndex((column) => column.referenceName === select("currentColumns").value);
  if (index < 0) {
    throw new Error("Choose a column in Columns in this list.");
  }
  return index;
}

function moveDraftColumn(target: (index: number) => number): void {
  const index = selectedDraftIndex();
  const next = Math.max(0, Math.min(draftColumns.length - 1, target(index)));
  if (next === index) {
    return;
  }
  const [column] = draftColumns.splice(index, 1);
  draftColumns.splice(next, 0, column);
  renderColumnDraft(column.referenceName);
}

function addDraftColumns(): void {
  const chosen = Array.from(select("addableColumns").selectedOptions).map((option) => option.value);
  const available = new Map(availableForDraft().map((field) => [field.referenceName, field]));
  const columns = chosen
    .map((referenceName) => available.get(referenceName))
    .filter((field): field is FieldChoice => Boolean(field))
    .map((field) => ({ ...field, locked: false }));
  if (columns.length === 0) {
    throw new Error("Choose one or more fields in Available fields.");
  }
  const anchor = draftColumns.findIndex((column) => column.referenceName === select("currentColumns").value);
  draftColumns.splice(anchor >= 0 ? anchor + 1 : draftColumns.length, 0, ...columns);
  renderColumnDraft(columns[columns.length - 1].referenceName);
}

function removeDraftColumn(): void {
  const index = selectedDraftIndex();
  const column = draftColumns[index];
  if (column.locked) {
    throw new Error(`The ${column.name} column is required and cannot be removed.`);
  }
  draftColumns.splice(index, 1);
  renderColumnDraft(draftColumns[Math.min(index, draftColumns.length - 1)]?.referenceName);
}

function draftAction(work: () => void): void {
  try {
    work();
  } catch (error) {
    setStatus(error instanceof Error ? error.message : String(error), "error");
  }
}

async function refreshChooserLists(): Promise<void> {
  try {
    const types = await listWorkbookWorkItemTypes();
    fillSelect(
      select("workItemType"),
      types.map((type) => ({ value: type, label: type }))
    );
  } catch {
    fillSelect(select("workItemType"), []);
  }
  let current: DraftColumn[] = [];
  let addable: FieldChoice[] = [];
  try {
    current = await listCurrentColumns();
  } catch {
    current = [];
  }
  try {
    addable = (await listAddableColumns()).map((field) => ({ referenceName: field.referenceName, name: field.name }));
  } catch {
    addable = [];
  }
  syncColumnDraft(current, addable);
  try {
    const paths = await listClassificationPaths();
    areaPaths = paths.areaPaths;
    iterationPaths = paths.iterationPaths;
    fillPathSelect(select("areaPath"), areaPaths, input("areaFilter").value);
    fillPathSelect(select("iterationPath"), iterationPaths, input("iterationFilter").value);
  } catch {
    fillSelect(select("areaPath"), []);
    fillSelect(select("iterationPath"), []);
  }
}

const busyIds = [
  "connect",
  "entraSignIn",
  "entraSignOut",
  "loadQuery",
  "refresh",
  "publish",
  "convertTree",
  "addLevel",
  "indent",
  "outdent",
  "applyIdentity",
  "insertWorkItem",
  "addColumn",
  "removeColumn",
  "moveColumnTop",
  "moveColumnUp",
  "moveColumnDown",
  "moveColumnBottom",
  "applyColumns",
  "discardColumns",
  "applyAreaPath",
  "applyIterationPath",
  "openInAdo",
  "loadWorkItemDetails",
  "openRelatedLink",
  "removeRelatedLink",
  "addRelatedLink",
  "addHyperlink",
  "openAttachment",
  "removeAttachment",
  "addAttachment",
  "addComment"
];

Office.onReady(async (info) => {
  setupCollapsibleCards();

  if (info.host !== Office.HostType.Excel) {
    setStatus("This add-in runs in Microsoft Excel.", "error");
    return;
  }

  const settings = await loadSettings();
  if (settings) {
    input("orgUrl").value = settings.orgUrl;
    if (settings.pat) {
      input("pat").value = settings.pat;
    }
    if (settings.entraClientId) {
      input("clientId").value = settings.entraClientId;
    }
  }

  fillLinkTypeSelect(DEFAULT_LINK_TYPES);
  renderSelectedDetails(null);

  $("connect").addEventListener("click", async () => {
    await withBusy(busyIds, async () => {
      const generation = ++connectionGeneration;
      clearProjectList();
      renderSelectedDetails(null);
      setStatus(`Connecting to ${input("orgUrl").value.trim()}…`);
      const result = await connectWithPat(input("orgUrl").value, input("pat").value);
      await refreshProjectList(generation);
      await refreshWorkbookStatus();
      setStatus(result.message, "ok");
    });
  });

  $("entraSignIn").addEventListener("click", async () => {
    await withBusy(busyIds, async () => {
      const generation = ++connectionGeneration;
      clearProjectList();
      renderSelectedDetails(null);
      const result = await connectWithEntra(input("orgUrl").value, input("clientId").value);
      await refreshProjectList(generation);
      await refreshWorkbookStatus();
      setStatus(result.message, "ok");
    });
  });

  $("entraSignOut").addEventListener("click", async () => {
    await withBusy(["entraSignOut"], async () => {
      const result = await disconnectEntra();
      await refreshWorkbookStatus();
      setStatus(result.message, "ok");
    });
  });

  $("project").addEventListener("change", async () => {
    await withBusy(["loadQuery"], async () => {
      await refreshQueryList();
    });
  });

  $("loadQuery").addEventListener("click", async () => {
    await withBusy(busyIds, async () => {
      const project = select("project").value;
      if (!project || !selectedQuery?.id) {
        throw new Error("Choose a project and a saved query.");
      }
      setStatus("Loading query…");
      renderSelectedDetails(null);
      const result = await loadQuery(project, selectedQuery.id, selectedQuery.name);
      await refreshWorkbookStatus();
      setStatus(result.message, "ok");
    });
  });

  $("refresh").addEventListener("click", async () => {
    await withBusy(busyIds, async () => {
      setStatus("Refreshing…");
      renderSelectedDetails(null);
      const result = await refresh();
      await refreshWorkbookStatus();
      setStatus(result.message, "ok");
    });
  });

  $("publish").addEventListener("click", async () => {
    await withBusy(busyIds, async () => {
      renderSelectedDetails(null);
      const result = await publish((message) => setStatus(message));
      await refreshWorkbookStatus();
      setStatus(result.message, result.ok === false ? "error" : "ok", result.details);
    });
  });

  $("openQuery").addEventListener("click", async () => {
    await withBusy(["openQuery"], async () => {
      const result = await openWorkbookQuery();
      setStatus(result.message, "ok");
    });
  });

  $("openInAdo").addEventListener("click", async () => {
    await withBusy(["openInAdo"], async () => {
      const result = await openSelectedWorkItem();
      setStatus(result.message, "ok");
    });
  });

  $("loadWorkItemDetails").addEventListener("click", async () => {
    await withBusy(["loadWorkItemDetails"], async () => {
      await reloadSelectedDetails();
      setStatus(`Loaded details for #${selectedDetails?.id}.`, "ok");
    });
  });

  $("openRelatedLink").addEventListener("click", async () => {
    await withBusy(["openRelatedLink"], async () => {
      const link = selectedLink("links");
      const result = await openWorkItemLink(link.url, link.rel);
      setStatus(result.message, "ok");
    });
  });

  $("removeRelatedLink").addEventListener("click", async () => {
    await withBusy(["removeRelatedLink", "loadWorkItemDetails"], async () => {
      const result = await removeWorkItemRelation(selectedRelationIndex("relatedLinks"), shownDetails());
      await reloadSelectedDetails();
      setStatus(result.message, "ok");
    });
  });

  $("addRelatedLink").addEventListener("click", async () => {
    await withBusy(["addRelatedLink", "loadWorkItemDetails"], async () => {
      const result = await addRelatedLink(
        Number(input("relatedLinkId").value),
        select("relatedLinkType").value,
        selectedDetails?.id
      );
      input("relatedLinkId").value = "";
      await reloadSelectedDetails();
      setStatus(result.message, "ok");
    });
  });

  $("addHyperlink").addEventListener("click", async () => {
    await withBusy(["addHyperlink", "loadWorkItemDetails"], async () => {
      const result = await addWorkItemHyperlink(input("hyperlinkUrl").value, input("hyperlinkComment").value, selectedDetails?.id);
      input("hyperlinkUrl").value = "";
      input("hyperlinkComment").value = "";
      await reloadSelectedDetails();
      setStatus(result.message, "ok");
    });
  });

  $("openAttachment").addEventListener("click", async () => {
    await withBusy(["openAttachment"], async () => {
      const link = selectedLink("attachments");
      const target = link.url || selectedDetails?.webUrl;
      if (!target) {
        throw new Error("Load details for the selected work item first.");
      }
      await openWorkItemLink(target);
      setStatus(
        link.url ? `Opened attachment ${link.label}.` : `Opened #${selectedDetails?.id} in Azure DevOps.`,
        "ok"
      );
    });
  });

  $("removeAttachment").addEventListener("click", async () => {
    await withBusy(["removeAttachment", "loadWorkItemDetails"], async () => {
      const result = await removeWorkItemRelation(selectedRelationIndex("attachments"), shownDetails());
      await reloadSelectedDetails();
      setStatus(result.message, "ok");
    });
  });

  $("addAttachment").addEventListener("click", async () => {
    await withBusy(["addAttachment", "loadWorkItemDetails"], async () => {
      const file = ( $("attachmentFile") as HTMLInputElement ).files?.[0];
      if (!file) {
        throw new Error("Choose a file to attach.");
      }
      const result = await addWorkItemAttachment(file.name, await file.arrayBuffer(), selectedDetails?.id);
      ($("attachmentFile") as HTMLInputElement).value = "";
      await reloadSelectedDetails();
      setStatus(result.message, "ok");
    });
  });

  $("addComment").addEventListener("click", async () => {
    await withBusy(["addComment", "loadWorkItemDetails"], async () => {
      const result = await addWorkItemComment(($("commentText") as HTMLTextAreaElement).value, selectedDetails?.id);
      ($("commentText") as HTMLTextAreaElement).value = "";
      await reloadSelectedDetails();
      setStatus(result.message, "ok");
    });
  });

  $("convertTree").addEventListener("click", async () => {
    await withBusy(busyIds, async () => {
      const result = await convertToTree();
      await refreshWorkbookStatus();
      setStatus(result.message, "ok");
    });
  });

  $("addLevel").addEventListener("click", async () => {
    await withBusy(busyIds, async () => {
      const result = await addTreeLevel();
      await refreshWorkbookStatus();
      setStatus(result.message, "ok");
    });
  });

  $("indent").addEventListener("click", async () => {
    await withBusy(busyIds, async () => {
      const result = await indentSelection(1);
      await refreshWorkbookStatus();
      setStatus(result.message, "ok");
    });
  });

  $("outdent").addEventListener("click", async () => {
    await withBusy(busyIds, async () => {
      const result = await indentSelection(-1);
      await refreshWorkbookStatus();
      setStatus(result.message, "ok");
    });
  });

  $("insertWorkItem").addEventListener("click", async () => {
    await withBusy(busyIds, async () => {
      const result = await insertWorkItem(select("workItemType").value);
      await refreshWorkbookStatus();
      setStatus(result.message, "ok");
    });
  });

  $("moveColumnTop").addEventListener("click", () => draftAction(() => moveDraftColumn(() => 0)));
  $("moveColumnUp").addEventListener("click", () => draftAction(() => moveDraftColumn((index) => index - 1)));
  $("moveColumnDown").addEventListener("click", () => draftAction(() => moveDraftColumn((index) => index + 1)));
  $("moveColumnBottom").addEventListener("click", () =>
    draftAction(() => moveDraftColumn(() => draftColumns.length - 1))
  );
  $("currentColumns").addEventListener("keydown", (event) => {
    if (!event.altKey || (event.key !== "ArrowUp" && event.key !== "ArrowDown")) {
      return;
    }
    event.preventDefault();
    const delta = event.key === "ArrowUp" ? -1 : 1;
    draftAction(() => moveDraftColumn((index) => index + delta));
  });
  $("removeColumn").addEventListener("click", () => draftAction(removeDraftColumn));
  $("addColumn").addEventListener("click", () => draftAction(addDraftColumns));
  $("addableColumns").addEventListener("dblclick", () => draftAction(addDraftColumns));
  $("addableFilter").addEventListener("input", () => renderAddableFields());

  $("discardColumns").addEventListener("click", () => {
    draftColumns = savedColumns.map((column) => ({ ...column }));
    renderColumnDraft();
    setStatus("Discarded pending column changes.", "ok");
  });

  $("applyColumns").addEventListener("click", async () => {
    await withBusy(busyIds, async () => {
      if (!columnDraftPending()) {
        setStatus("No pending column changes.", "ok");
        return;
      }
      setStatus("Applying column changes…");
      const result = await applyColumnLayout(draftColumns.map((column) => column.referenceName));
      await refreshWorkbookStatus();
      setStatus(result.message, "ok");
    });
  });

  $("areaFilter").addEventListener("input", () => {
    fillPathSelect(select("areaPath"), areaPaths, input("areaFilter").value);
  });

  $("iterationFilter").addEventListener("input", () => {
    fillPathSelect(select("iterationPath"), iterationPaths, input("iterationFilter").value);
  });

  $("applyAreaPath").addEventListener("click", async () => {
    await withBusy(["applyAreaPath"], async () => {
      const result = await applyPathToSelection("area", select("areaPath").value);
      setStatus(result.message, "ok");
    });
  });

  $("applyIterationPath").addEventListener("click", async () => {
    await withBusy(["applyIterationPath"], async () => {
      const result = await applyPathToSelection("iteration", select("iterationPath").value);
      setStatus(result.message, "ok");
    });
  });

  $("identitySearch").addEventListener("input", () => {
    window.clearTimeout(searchTimer);
    searchTimer = window.setTimeout(async () => {
      const term = input("identitySearch").value;
      if (term.trim().length < 2) {
        fillSelect(select("identityResults"), []);
        identityHits = [];
        return;
      }
      try {
        identityHits = await searchIdentities(term);
        fillSelect(
          select("identityResults"),
          identityHits.map((identity, index) => ({
            value: String(index),
            label: identity.uniqueName ? `${identity.displayName} <${identity.uniqueName}>` : identity.displayName
          }))
        );
      } catch (error) {
        setStatus(error instanceof Error ? error.message : String(error), "error");
      }
    }, 300);
  });

  $("applyIdentity").addEventListener("click", async () => {
    await withBusy(["applyIdentity"], async () => {
      const index = Number(select("identityResults").value);
      const identity = identityHits[index];
      if (!identity) {
        throw new Error("Search for a person, then select them in the list.");
      }
      const result = await applyIdentityToSelection(identity);
      setStatus(result.message, "ok");
    });
  });

  // Handlers are registered first so Connect still works if the saved org is slow or unreachable.
  if (settings) {
    void restoreSavedConnection(settings.orgUrl);
  }
  void refreshWorkbookStatus()
    .catch(() => undefined)
    .then(() => ensureTreePresentation())
    .catch(() => undefined);
});

async function restoreSavedConnection(orgUrl: string): Promise<void> {
  const generation = connectionGeneration;
  setStatus(`Reconnecting to ${orgUrl}…`);
  try {
    if (await refreshProjectList(generation)) {
      setStatus(`Connected to ${orgUrl}. You can refresh, publish, or load a query.`, "ok");
    }
  } catch (error) {
    if (generation === connectionGeneration) {
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`Could not reconnect to ${orgUrl}: ${message} Enter a different organization URL or token and click Connect.`, "error");
    }
  }
}
