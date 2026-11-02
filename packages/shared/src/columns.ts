import {
  CHOOSER_EXCLUDED_FIELDS,
  IDENTITY_FIELDS,
  LOCKED_COLUMN_FIELDS,
  MAX_CLASSIFICATION_PATHS,
  PARENT_TITLE_COLUMN,
  PARENT_TITLE_FIELD,
  TREE_PATH_FIELDS
} from "./constants";
import { isDateTimeField, treeLevelCount } from "./fields";
import { isHtmlField } from "./html";
import {
  ClassificationNode,
  ProcessCatalog,
  ProcessFieldInfo,
  TableRow,
  WorkbookColumn,
  WorkItemTypeInfo
} from "./types";

export const PARENT_TITLE_PROCESS_FIELD: ProcessFieldInfo = {
  referenceName: PARENT_TITLE_COLUMN.referenceName,
  name: PARENT_TITLE_COLUMN.name,
  readOnly: true
};

export function allowedWorkItemTypeNames(types: WorkItemTypeInfo[]): string[] {
  return types
    .filter((type) => type.name && !type.isDisabled)
    .map((type) => type.name)
    .sort((left, right) => left.localeCompare(right));
}

export function isLockedColumn(column: WorkbookColumn): boolean {
  return Boolean(column.titleLevel) || LOCKED_COLUMN_FIELDS.has(column.referenceName);
}

export function canRemoveColumn(column: WorkbookColumn): boolean {
  return !isLockedColumn(column);
}

export function addableProcessFields(fields: ProcessFieldInfo[], columns: WorkbookColumn[]): ProcessFieldInfo[] {
  const present = new Set(columns.map((column) => column.referenceName));
  const addable = fields
    .filter(
      (field) =>
        field.referenceName &&
        field.name &&
        field.referenceName !== PARENT_TITLE_FIELD &&
        !present.has(field.referenceName) &&
        !CHOOSER_EXCLUDED_FIELDS.has(field.referenceName)
    )
    .sort((left, right) => left.name.localeCompare(right.name));
  if (!present.has(PARENT_TITLE_FIELD)) {
    return [PARENT_TITLE_PROCESS_FIELD, ...addable].sort((left, right) => left.name.localeCompare(right.name));
  }
  return addable;
}

export function columnFromField(field: ProcessFieldInfo): WorkbookColumn {
  return {
    referenceName: field.referenceName,
    name: field.name,
    identity: field.isIdentity,
    html: field.isHtml || isHtmlField(field.referenceName),
    date: field.isDate || isDateTimeField(field.referenceName, field.type)
  };
}

export function addWorkbookColumn(columns: WorkbookColumn[], field: ProcessFieldInfo): WorkbookColumn[] {
  if (columns.some((column) => column.referenceName === field.referenceName && !column.titleLevel)) {
    return columns;
  }
  return [...columns, columnFromField(field)];
}

export interface ColumnLayoutChange {
  columns: WorkbookColumn[];
  added: WorkbookColumn[];
  removed: WorkbookColumn[];
  reordered: boolean;
}

/**
 * Turns a staged column order (reference names, tree Title levels excluded) into workbook columns.
 * Names not already on the sheet must come from `addable`; required columns cannot be dropped.
 */
export function buildColumnLayout(
  current: WorkbookColumn[],
  order: string[],
  addable: ProcessFieldInfo[]
): ColumnLayoutChange {
  const titles = current.filter((column) => column.titleLevel);
  const existing = current.filter((column) => !column.titleLevel);
  const byReference = new Map(existing.map((column) => [column.referenceName, column]));
  const fields = new Map(addable.map((field) => [field.referenceName, field]));
  const seen = new Set<string>();
  const rest: WorkbookColumn[] = [];
  const added: WorkbookColumn[] = [];

  for (const referenceName of order) {
    if (seen.has(referenceName)) {
      throw new Error(`The ${byReference.get(referenceName)?.name || referenceName} column is listed twice.`);
    }
    seen.add(referenceName);
    const column = byReference.get(referenceName);
    if (column) {
      rest.push(column);
      continue;
    }
    const field = fields.get(referenceName);
    if (!field) {
      throw new Error(`${referenceName} is not available on this project's process template, so it can't be added.`);
    }
    const created = columnFromField(field);
    rest.push(created);
    added.push(created);
  }

  const removed = existing.filter((column) => !seen.has(column.referenceName));
  const required = removed.find((column) => !canRemoveColumn(column));
  if (required) {
    throw new Error(`The ${required.name} column is required and cannot be removed.`);
  }

  const keptOrder = rest.filter((column) => byReference.has(column.referenceName)).map((column) => column.referenceName);
  const originalOrder = existing.filter((column) => seen.has(column.referenceName)).map((column) => column.referenceName);
  const reordered = keptOrder.some((referenceName, index) => referenceName !== originalOrder[index]);

  if (titles.length === 0) {
    return { columns: rest, added, removed, reordered };
  }
  const typeIndex = rest.findIndex((column) => column.referenceName === "System.WorkItemType");
  const columns =
    typeIndex >= 0
      ? [...rest.slice(0, typeIndex + 1), ...titles, ...rest.slice(typeIndex + 1)]
      : [...titles, ...rest];
  return { columns, added, removed, reordered };
}

export function removeWorkbookColumn(columns: WorkbookColumn[], referenceName: string): WorkbookColumn[] {
  const target = columns.find((column) => column.referenceName === referenceName && !column.titleLevel);
  if (!target || !canRemoveColumn(target)) {
    throw new Error(`The ${target?.name || referenceName} column is required and cannot be removed.`);
  }
  return columns.filter((column) => column.referenceName !== referenceName || Boolean(column.titleLevel));
}

export function insertWorkItemRow(
  rows: TableRow[],
  workItemType: string,
  columns: WorkbookColumn[],
  atIndex: number,
  isTree: boolean,
  defaultState = ""
): TableRow[] {
  const index = Math.max(0, Math.min(atIndex, rows.length));
  const values: TableRow["values"] = {};
  for (const column of columns) {
    if (column.titleLevel) {
      continue;
    }
    if (column.referenceName === "System.WorkItemType") {
      values[column.referenceName] = workItemType;
    } else if (column.referenceName === "System.State") {
      values[column.referenceName] = defaultState;
    } else {
      values[column.referenceName] = "";
    }
  }
  const inserted: TableRow = {
    rowIndex: index,
    values,
    treeLevel: isTree ? 1 : undefined
  };
  return [...rows.slice(0, index), inserted, ...rows.slice(index)].map((row, rowIndex) => ({
    ...row,
    rowIndex
  }));
}

export function isAllowedWorkItemType(workItemType: string, allowed?: string[]): boolean {
  if (!allowed || allowed.length === 0) {
    return Boolean(workItemType.trim());
  }
  return allowed.some((name) => name.toLowerCase() === workItemType.trim().toLowerCase());
}

export function statesForWorkItemType(
  workItemType: string,
  statesByType?: Record<string, string[]>
): string[] | undefined {
  if (!statesByType) {
    return undefined;
  }
  if (statesByType[workItemType]) {
    return statesByType[workItemType];
  }
  const match = Object.keys(statesByType).find((name) => name.toLowerCase() === workItemType.trim().toLowerCase());
  return match ? statesByType[match] : undefined;
}

export function isAllowedState(
  state: string,
  workItemType: string,
  statesByType?: Record<string, string[]>
): boolean {
  const trimmed = state.trim();
  if (!trimmed) {
    return true;
  }
  const allowed = statesForWorkItemType(workItemType, statesByType);
  if (!allowed || allowed.length === 0) {
    return true;
  }
  return allowed.some((name) => name.toLowerCase() === trimmed.toLowerCase());
}

export function defaultStateForType(
  workItemType: string,
  statesByType?: Record<string, string[]>
): string {
  return statesForWorkItemType(workItemType, statesByType)?.[0] || "";
}

export function moveWorkbookColumn(
  columns: WorkbookColumn[],
  referenceName: string,
  delta: number
): WorkbookColumn[] {
  const rest = treeLevelCount(columns) > 0 ? columns.filter((column) => !column.titleLevel) : columns;
  const index = rest.findIndex((column) => column.referenceName === referenceName);
  if (index < 0) {
    throw new Error("Choose a column to move.");
  }
  const nextIndex = index + delta;
  if (nextIndex < 0 || nextIndex >= rest.length) {
    return columns;
  }
  const reordered = [...rest];
  const [item] = reordered.splice(index, 1);
  reordered.splice(nextIndex, 0, item);
  const titles = columns.filter((column) => column.titleLevel);
  if (titles.length === 0) {
    return reordered;
  }
  const typeIndex = reordered.findIndex((column) => column.referenceName === "System.WorkItemType");
  if (typeIndex >= 0) {
    return [...reordered.slice(0, typeIndex + 1), ...titles, ...reordered.slice(typeIndex + 1)];
  }
  return [...titles, ...reordered];
}

export function flattenClassificationPaths(node: ClassificationNode | undefined, parentPath = ""): string[] {
  if (!node?.name?.trim()) {
    return [];
  }
  const path = parentPath ? `${parentPath}\\${node.name.trim()}` : node.name.trim();
  const paths = [path];
  for (const child of node.children || []) {
    paths.push(...flattenClassificationPaths(child, path));
  }
  return paths.slice(0, MAX_CLASSIFICATION_PATHS);
}

export function isAllowedPath(value: string, paths?: string[]): boolean {
  const trimmed = value.trim();
  if (!trimmed) {
    return true;
  }
  if (!paths || paths.length === 0) {
    return true;
  }
  return paths.some((path) => path.toLowerCase() === trimmed.toLowerCase());
}

export function isAllowedPicklistValue(
  value: string,
  allowed?: string[],
  suggested = false
): boolean {
  const trimmed = value.trim();
  if (!trimmed || suggested) {
    return true;
  }
  if (!allowed || allowed.length === 0) {
    return true;
  }
  return allowed.some((item) => String(item).toLowerCase() === trimmed.toLowerCase());
}

export function isAllowedTransition(
  fromState: string,
  toState: string,
  workItemType: string,
  transitionsByType?: Record<string, Record<string, string[]>>
): boolean {
  const to = toState.trim();
  if (!to) {
    return true;
  }
  const transitions = transitionsForType(workItemType, transitionsByType);
  if (!transitions || Object.keys(transitions).length === 0) {
    return true;
  }
  const from = fromState.trim();
  if (!from) {
    const initial = listForState(transitions, "") || listForState(transitions, "*");
    if (!initial || initial.length === 0) {
      return true;
    }
    return initial.some((name) => name.toLowerCase() === to.toLowerCase());
  }
  if (from.toLowerCase() === to.toLowerCase()) {
    return true;
  }
  const allowed = listForState(transitions, from);
  if (!allowed || allowed.length === 0) {
    return true;
  }
  return allowed.some((name) => name.toLowerCase() === to.toLowerCase());
}

export function applyCatalogToColumns(columns: WorkbookColumn[], catalog?: ProcessCatalog | null): WorkbookColumn[] {
  if (!catalog) {
    return columns.map((column) => ({
      ...column,
      html: column.html || isHtmlField(column.referenceName),
      identity: column.identity,
      date: column.date || isDateTimeField(column.referenceName)
    }));
  }
  const html = new Set(catalog.htmlFields.map((name) => name.toLowerCase()));
  const identity = new Set(catalog.identityFields.map((name) => name.toLowerCase()));
  const dates = new Set(
    catalog.fields
      .filter((field) => field.isDate || isDateTimeField(field.referenceName, field.type))
      .map((field) => field.referenceName.toLowerCase())
  );
  return columns.map((column) => ({
    ...column,
    html: column.html || html.has(column.referenceName.toLowerCase()) || isHtmlField(column.referenceName),
    identity: column.identity || identity.has(column.referenceName.toLowerCase()),
    date:
      column.date ||
      dates.has(column.referenceName.toLowerCase()) ||
      isDateTimeField(column.referenceName)
  }));
}

export function emptyProcessCatalog(): ProcessCatalog {
  return {
    types: [],
    statesByType: {},
    transitionsByType: {},
    picklists: {},
    suggestedPicklists: [],
    htmlFields: [],
    identityFields: Array.from(IDENTITY_FIELDS),
    identities: [],
    areaPaths: [],
    iterationPaths: [],
    fields: []
  };
}

export function indentPathLabel(path: string): string {
  const parts = path.split("\\");
  if (parts.length <= 1) {
    return path;
  }
  return `${"  ".repeat(parts.length - 1)}${parts[parts.length - 1]}`;
}

export function isTreePathField(referenceName: string): boolean {
  return TREE_PATH_FIELDS.has(referenceName);
}

function transitionsForType(
  workItemType: string,
  transitionsByType?: Record<string, Record<string, string[]>>
): Record<string, string[]> | undefined {
  if (!transitionsByType) {
    return undefined;
  }
  if (transitionsByType[workItemType]) {
    return transitionsByType[workItemType];
  }
  const match = Object.keys(transitionsByType).find(
    (name) => name.toLowerCase() === workItemType.trim().toLowerCase()
  );
  return match ? transitionsByType[match] : undefined;
}

function listForState(transitions: Record<string, string[]>, state: string): string[] | undefined {
  if (Object.prototype.hasOwnProperty.call(transitions, state)) {
    return transitions[state];
  }
  const match = Object.keys(transitions).find((name) => name.toLowerCase() === state.toLowerCase());
  return match ? transitions[match] : undefined;
}
