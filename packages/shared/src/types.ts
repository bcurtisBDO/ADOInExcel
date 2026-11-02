export type OpenAction = "OpenQuery" | "OpenItems";
export type ListType = "flat" | "tree";

export type CellValue = string | number | boolean;

export interface WorkbookColumn {
  referenceName: string;
  name: string;
  titleLevel?: number;
  identity?: boolean;
  html?: boolean;
  date?: boolean;
}

export interface WorkbookMetadata {
  version: number;
  orgUrl: string;
  project: string;
  action: OpenAction;
  listType: ListType;
  treeLevels?: number;
  queryId?: string;
  queryName?: string;
  workItemIds?: number[];
  columns: WorkbookColumn[];
  loadedAt: string;
  revisions: Record<string, number>;
  snapshot: Record<string, Record<string, CellValue>>;
  parents: Record<string, number>;
  workItemTypes?: string[];
  statesByType?: Record<string, string[]>;
  picklists?: Record<string, string[]>;
  suggestedPicklists?: string[];
  htmlFields?: string[];
  identityFields?: string[];
  identities?: string[];
  transitionsByType?: Record<string, Record<string, string[]>>;
  areaPaths?: string[];
  iterationPaths?: string[];
}

export interface JsonPatchOperation {
  op: "add" | "replace" | "remove" | "test";
  path: string;
  value?: unknown;
}

export type PublishKind = "create" | "update" | "skip" | "conflict" | "error";

export interface PublishChange {
  kind: PublishKind;
  rowIndex: number;
  id?: number;
  workItemType?: string;
  operations?: JsonPatchOperation[];
  message: string;
  desiredParentId?: number | null;
  desiredParentRowIndex?: number;
  previousParentId?: number | null;
}

export interface PublishPlan {
  changes: PublishChange[];
  updateCount: number;
  createCount: number;
  skipCount: number;
  errors: string[];
}

export interface WorkItemRelation {
  rel?: string;
  url?: string;
  attributes?: Record<string, unknown>;
}

export interface WorkItemLinkInfo {
  index: number;
  rel: string;
  label: string;
  url: string;
  id?: number;
  name?: string;
}

export interface WorkItemCommentInfo {
  id: number;
  text: string;
  createdBy: string;
  createdDate: string;
}

export interface WorkItemDetails {
  id: number;
  rev?: number;
  title: string;
  webUrl: string;
  parentId?: number;
  parentTitle?: string;
  parentUrl?: string;
  links: WorkItemLinkInfo[];
  attachments: WorkItemLinkInfo[];
  comments: WorkItemCommentInfo[];
}

export interface WorkItemLike {
  id?: number;
  rev?: number;
  fields?: Record<string, unknown>;
  relations?: WorkItemRelation[];
  url?: string;
}

export interface QueryColumn {
  referenceName?: string;
  name?: string;
}

export interface QueryWorkItemRef {
  id: number;
  url?: string;
}

export interface QueryRelation {
  rel?: string;
  source?: { id?: number };
  target?: { id?: number };
}

export interface QueryResultLike {
  queryType?: unknown;
  queryResultType?: unknown;
  columns?: QueryColumn[];
  workItems?: QueryWorkItemRef[];
  workItemRelations?: QueryRelation[];
}

export interface SavedQuery {
  id: string;
  name: string;
  path: string;
  isFolder?: boolean;
}

export interface AdoProject {
  id: string;
  name: string;
}

export interface TableRow {
  rowIndex: number;
  id?: number;
  values: Record<string, CellValue>;
  treeLevel?: number;
  parentId?: number;
  parentRowIndex?: number;
}

export interface TreeLayout {
  order: number[];
  parentById: Record<number, number>;
  depthById: Record<number, number>;
  maxDepth: number;
}

export interface IdentityPick {
  displayName: string;
  uniqueName?: string;
  descriptor?: string;
}

export type AdoCredential = string | { type: "pat" | "bearer"; value: string };

export interface WorkItemTypeInfo {
  name: string;
  referenceName?: string;
  description?: string;
  isDisabled?: boolean;
}

export interface ProcessFieldInfo {
  referenceName: string;
  name: string;
  type?: string;
  readOnly?: boolean;
  isIdentity?: boolean;
  isHtml?: boolean;
  isDate?: boolean;
  isPicklistSuggested?: boolean;
  allowedValues?: string[];
}

export interface ClassificationNode {
  name?: string;
  children?: ClassificationNode[];
}

export interface ProcessCatalog {
  types: string[];
  statesByType: Record<string, string[]>;
  transitionsByType: Record<string, Record<string, string[]>>;
  picklists: Record<string, string[]>;
  suggestedPicklists: string[];
  htmlFields: string[];
  identityFields: string[];
  identities: string[];
  areaPaths: string[];
  iterationPaths: string[];
  fields: ProcessFieldInfo[];
}
