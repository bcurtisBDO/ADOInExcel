import { WORKBOOK_METADATA_VERSION } from "./constants";
import { workItemToSnapshot } from "./fields";
import { parentsFromLayout } from "./tree";
import {
  CellValue,
  ListType,
  OpenAction,
  ProcessCatalog,
  TreeLayout,
  WorkbookColumn,
  WorkbookMetadata,
  WorkItemLike
} from "./types";

export function buildMetadata(options: {
  orgUrl: string;
  project: string;
  action: OpenAction;
  columns: WorkbookColumn[];
  items: WorkItemLike[];
  queryId?: string;
  queryName?: string;
  workItemIds?: number[];
  listType?: ListType;
  layout?: TreeLayout | null;
  workItemTypes?: string[];
  statesByType?: Record<string, string[]>;
  picklists?: Record<string, string[]>;
  suggestedPicklists?: string[];
  htmlFields?: string[];
  identityFields?: string[];
  transitionsByType?: Record<string, Record<string, string[]>>;
  areaPaths?: string[];
  iterationPaths?: string[];
  catalog?: ProcessCatalog | null;
}): WorkbookMetadata {
  const revisions: Record<string, number> = {};
  const snapshot: Record<string, Record<string, CellValue>> = {};

  for (const item of options.items) {
    if (!item.id) {
      continue;
    }
    const key = String(item.id);
    if (typeof item.rev === "number") {
      revisions[key] = item.rev;
    }
    snapshot[key] = workItemToSnapshot(item, options.columns);
  }

  const listType = options.listType || (options.layout ? "tree" : "flat");
  const fromCatalog = options.catalog ? catalogFields(options.catalog) : {};

  return {
    version: WORKBOOK_METADATA_VERSION,
    orgUrl: options.orgUrl,
    project: options.project,
    action: options.action,
    listType,
    treeLevels: options.layout?.maxDepth,
    queryId: options.queryId,
    queryName: options.queryName,
    workItemIds: options.workItemIds,
    columns: options.columns,
    loadedAt: new Date().toISOString(),
    revisions,
    snapshot,
    parents: parentsFromLayout(options.layout),
    workItemTypes: options.workItemTypes ?? fromCatalog.workItemTypes,
    statesByType: options.statesByType ?? fromCatalog.statesByType,
    picklists: options.picklists ?? fromCatalog.picklists,
    suggestedPicklists: options.suggestedPicklists ?? fromCatalog.suggestedPicklists,
    htmlFields: options.htmlFields ?? fromCatalog.htmlFields,
    identityFields: options.identityFields ?? fromCatalog.identityFields,
    transitionsByType: options.transitionsByType ?? fromCatalog.transitionsByType,
    areaPaths: options.areaPaths ?? fromCatalog.areaPaths,
    iterationPaths: options.iterationPaths ?? fromCatalog.iterationPaths
  };
}

export function catalogFields(catalog: ProcessCatalog): Pick<
  WorkbookMetadata,
  | "workItemTypes"
  | "statesByType"
  | "picklists"
  | "suggestedPicklists"
  | "htmlFields"
  | "identityFields"
  | "transitionsByType"
  | "areaPaths"
  | "iterationPaths"
> {
  return {
    workItemTypes: catalog.types,
    statesByType: catalog.statesByType,
    picklists: catalog.picklists,
    suggestedPicklists: catalog.suggestedPicklists,
    htmlFields: catalog.htmlFields,
    identityFields: catalog.identityFields,
    transitionsByType: catalog.transitionsByType,
    areaPaths: catalog.areaPaths,
    iterationPaths: catalog.iterationPaths
  };
}

export const META_AREA_PATH_HEADER = "areaPaths";
export const META_ITERATION_PATH_HEADER = "iterationPaths";

export function packClassificationPathGrid(
  areaPaths: string[] = [],
  iterationPaths: string[] = []
): string[][] {
  const rows = Math.max(areaPaths.length, iterationPaths.length);
  const grid: string[][] = [[META_AREA_PATH_HEADER, META_ITERATION_PATH_HEADER]];
  for (let index = 0; index < rows; index += 1) {
    grid.push([areaPaths[index] || "", iterationPaths[index] || ""]);
  }
  return grid;
}

export function unpackClassificationPathGrid(
  values: unknown[][] | undefined | null
): { areaPaths: string[]; iterationPaths: string[] } {
  const empty = { areaPaths: [] as string[], iterationPaths: [] as string[] };
  if (!values || values.length === 0) {
    return empty;
  }

  const first = String(values[0]?.[0] ?? "").trim();
  if (first.startsWith("{")) {
    return parsePersistedPathJson(first);
  }

  const header = values[0].map((cell) => String(cell ?? "").trim().toLowerCase());
  const hasHeader = header[0] === META_AREA_PATH_HEADER.toLowerCase();
  const start = hasHeader ? 1 : 0;
  const areaPaths: string[] = [];
  const iterationPaths: string[] = [];
  for (let row = start; row < values.length; row += 1) {
    const area = String(values[row]?.[0] ?? "").trim();
    const iteration = String(values[row]?.[1] ?? "").trim();
    if (area) {
      areaPaths.push(area);
    }
    if (iteration) {
      iterationPaths.push(iteration);
    }
  }
  return { areaPaths, iterationPaths };
}

export function parsePersistedPathJson(raw: unknown): { areaPaths: string[]; iterationPaths: string[] } {
  const empty = { areaPaths: [] as string[], iterationPaths: [] as string[] };
  if (raw === null || raw === undefined || raw === "") {
    return empty;
  }
  try {
    const parsed = JSON.parse(String(raw)) as { areaPaths?: unknown; iterationPaths?: unknown };
    return {
      areaPaths: Array.isArray(parsed.areaPaths) ? parsed.areaPaths.map((path) => String(path)).filter(Boolean) : [],
      iterationPaths: Array.isArray(parsed.iterationPaths)
        ? parsed.iterationPaths.map((path) => String(path)).filter(Boolean)
        : []
    };
  } catch {
    return empty;
  }
}

export function catalogFromMetadata(metadata: WorkbookMetadata): ProcessCatalog {
  return {
    types: metadata.workItemTypes || [],
    statesByType: metadata.statesByType || {},
    transitionsByType: metadata.transitionsByType || {},
    picklists: metadata.picklists || {},
    suggestedPicklists: metadata.suggestedPicklists || [],
    htmlFields: metadata.htmlFields || [],
    identityFields: metadata.identityFields || [],
    identities: metadata.identities || [],
    areaPaths: metadata.areaPaths || [],
    iterationPaths: metadata.iterationPaths || [],
    fields: []
  };
}

export function headersFromColumns(columns: WorkbookColumn[]): string[] {
  return columns.map((column) => column.name);
}

const META_JSON_PREFIX = "ADO1:";
const META_CHUNK_SIZE = 24000;

export function packMetadataSheet(metadata: WorkbookMetadata): CellValue[][] {
  const areaPaths = metadata.areaPaths || [];
  const iterationPaths = metadata.iterationPaths || [];
  const body: WorkbookMetadata = { ...metadata, areaPaths: undefined, iterationPaths: undefined };
  const withPaths = { ...body, areaPaths, iterationPaths };
  const json = JSON.stringify(jsonFitsCell(withPaths) ? withPaths : body);
  const chunks: string[] = [];
  for (let index = 0; index < json.length; index += META_CHUNK_SIZE) {
    chunks.push(`${META_JSON_PREFIX}${json.slice(index, index + META_CHUNK_SIZE)}`);
  }
  const pathGrid = packClassificationPathGrid(areaPaths, iterationPaths);
  const rows = Math.max(chunks.length, pathGrid.length, 1);
  const grid: CellValue[][] = [];
  for (let row = 0; row < rows; row += 1) {
    grid.push([chunks[row] || "", pathGrid[row]?.[0] || "", pathGrid[row]?.[1] || ""]);
  }
  return grid;
}

function jsonFitsCell(value: unknown): boolean {
  try {
    return JSON.stringify(value).length <= META_CHUNK_SIZE;
  } catch {
    return false;
  }
}

export function parseMetadataSheet(values: unknown[][] | undefined | null): WorkbookMetadata | null {
  if (!values || values.length === 0) {
    return null;
  }
  const chunks: string[] = [];
  for (const row of values) {
    const cell = row?.[0];
    if (cell === null || cell === undefined || cell === "") {
      continue;
    }
    const raw = String(cell);
    if (/^#(NAME|VALUE|REF|N\/A|NULL|DIV\/0|NUM|GETTING_DATA)[!?]/i.test(raw.trim())) {
      continue;
    }
    let text = raw.replace(/^\uFEFF/, "").replace(/^'/, "");
    if (text.startsWith(META_JSON_PREFIX)) {
      text = text.slice(META_JSON_PREFIX.length);
    }
    chunks.push(text);
  }
  for (let end = chunks.length; end >= 1; end -= 1) {
    const parsed = parseMetadata(chunks.slice(0, end).join(""));
    if (parsed) {
      return parsed;
    }
  }
  for (const chunk of chunks) {
    const trimmed = chunk.trim();
    if (trimmed.startsWith("{")) {
      const single = parseMetadata(trimmed);
      if (single) {
        return single;
      }
    }
  }
  return null;
}

export function parseMetadata(raw: unknown): WorkbookMetadata | null {
  if (raw === null || raw === undefined || raw === "") {
    return null;
  }

  let text = String(raw).trim().replace(/^\uFEFF/, "");
  if (text.startsWith(META_JSON_PREFIX)) {
    text = text.slice(META_JSON_PREFIX.length).trim();
  }
  if (!text) {
    return null;
  }

  try {
    const parsed = JSON.parse(text) as WorkbookMetadata;
    if (!parsed || (parsed.version !== 1 && parsed.version !== 2) || !Array.isArray(parsed.columns)) {
      return null;
    }
    parsed.revisions = parsed.revisions || {};
    parsed.snapshot = parsed.snapshot || {};
    parsed.parents = parsed.parents || {};
    parsed.listType = parsed.listType || (parsed.treeLevels && parsed.treeLevels > 1 ? "tree" : "flat");
    parsed.columns = parsed.columns.map((column) => {
      const titleMatch = /^Title\s+(\d+)$/i.exec(column.name);
      if (titleMatch && !column.titleLevel) {
        return { ...column, referenceName: "System.Title", titleLevel: Number(titleMatch[1]) };
      }
      return column;
    });
    return parsed;
  } catch {
    return null;
  }
}
