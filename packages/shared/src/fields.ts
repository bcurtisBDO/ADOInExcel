import {
  DEFAULT_COLUMNS,
  IDENTITY_FIELDS,
  MAX_IDENTITY_PICKLIST,
  PARENT_TITLE_FIELD,
  READ_ONLY_FIELDS,
  REQUIRED_FIELDS
} from "./constants";
import { htmlToMarkdown, isHtmlField, looksLikeHtml } from "./html";
import {
  CellValue,
  QueryColumn,
  QueryResultLike,
  TableRow,
  TreeLayout,
  WorkbookColumn,
  WorkItemLike
} from "./types";

export function chunk<T>(items: T[], size: number): T[][] {
  const result: T[][] = [];
  for (let i = 0; i < items.length; i += size) {
    result.push(items.slice(i, i + size));
  }
  return result;
}

/** Run async work over items with a fixed concurrency cap (order of results matches input). */
export async function mapConcurrent<T, R>(
  items: T[],
  concurrency: number,
  worker: (item: T, index: number) => Promise<R>
): Promise<R[]> {
  if (items.length === 0) {
    return [];
  }
  const limit = Math.max(1, Math.min(concurrency, items.length));
  const results = new Array<R>(items.length);
  let next = 0;
  async function run(): Promise<void> {
    while (next < items.length) {
      const index = next;
      next += 1;
      results[index] = await worker(items[index], index);
    }
  }
  await Promise.all(Array.from({ length: limit }, () => run()));
  return results;
}

export function stripHtml(value: string): string {
  return value
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .trim();
}

const EXCEL_EPOCH_SERIAL = 25569;
const MS_PER_DAY = 86400000;

export function isDateTimeField(referenceName?: string, type?: string): boolean {
  const typeName = (type || "").toLowerCase();
  if (typeName === "datetime" || typeName === "date") {
    return true;
  }
  if (!referenceName) {
    return false;
  }
  const leaf = referenceName.split(".").pop() || referenceName;
  return /Date$/i.test(leaf) || /DateTime$/i.test(leaf);
}

export function isIsoDateString(value: unknown): boolean {
  return typeof value === "string" && /^\d{4}-\d{2}-\d{2}(?:$|T|\s)/.test(value.trim());
}

export function parseDateOnly(value: unknown): { year: number; month: number; day: number } | undefined {
  if (value === null || value === undefined || value === "") {
    return undefined;
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    const utc = new Date(Math.round((value - EXCEL_EPOCH_SERIAL) * MS_PER_DAY));
    if (Number.isNaN(utc.getTime())) {
      return undefined;
    }
    return { year: utc.getUTCFullYear(), month: utc.getUTCMonth() + 1, day: utc.getUTCDate() };
  }
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return { year: value.getUTCFullYear(), month: value.getUTCMonth() + 1, day: value.getUTCDate() };
  }
  if (typeof value === "string") {
    const trimmed = value.trim();
    const iso = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) {
      return { year: Number(iso[1]), month: Number(iso[2]), day: Number(iso[3]) };
    }
    const parsed = Date.parse(trimmed);
    if (!Number.isNaN(parsed)) {
      const date = new Date(parsed);
      return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate() };
    }
  }
  return undefined;
}

export function dateToExcelSerial(parts: { year: number; month: number; day: number }): number {
  return Date.UTC(parts.year, parts.month - 1, parts.day) / MS_PER_DAY + EXCEL_EPOCH_SERIAL;
}

export function dateToAdoValue(value: unknown): string | null {
  const parts = parseDateOnly(value);
  if (!parts) {
    return value === "" || value === null || value === undefined ? null : String(value);
  }
  const month = String(parts.month).padStart(2, "0");
  const day = String(parts.day).padStart(2, "0");
  return `${parts.year}-${month}-${day}T00:00:00.000Z`;
}

export function dateToYmd(value: unknown): string {
  const parts = parseDateOnly(value);
  if (!parts) {
    return "";
  }
  return `${parts.year}-${String(parts.month).padStart(2, "0")}-${String(parts.day).padStart(2, "0")}`;
}

export function excelShortDateFormat(): string {
  const parts = new Intl.DateTimeFormat(undefined, {
    year: "numeric",
    month: "numeric",
    day: "numeric"
  }).formatToParts(new Date(Date.UTC(2026, 8, 18)));
  const tokens = parts
    .filter((part) => part.type === "year" || part.type === "month" || part.type === "day")
    .map((part) => (part.type === "year" ? "yyyy" : part.type === "month" ? "m" : "d"));
  return tokens.join("/") || "m/d/yyyy";
}

export function fieldToCell(value: unknown, referenceName?: string, htmlFields?: string[]): CellValue {
  if (value === null || value === undefined) {
    return "";
  }
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return dateToExcelSerial(parseDateOnly(value)!);
  }
  if (typeof value === "number" || typeof value === "boolean") {
    return value;
  }
  if (typeof value === "string") {
    if (isDateTimeField(referenceName) || isIsoDateString(value)) {
      const parts = parseDateOnly(value);
      if (parts) {
        return dateToExcelSerial(parts);
      }
    }
    if (isHtmlField(referenceName, htmlFields) || looksLikeHtml(value)) {
      return htmlToMarkdown(value);
    }
    return value;
  }
  if (Array.isArray(value)) {
    return value.map((item) => String(fieldToCell(item, referenceName, htmlFields))).filter(Boolean).join("; ");
  }
  if (typeof value === "object") {
    const record = value as Record<string, unknown>;
    if (typeof record.displayName === "string") {
      const unique = typeof record.uniqueName === "string" ? record.uniqueName : "";
      if (unique && unique !== record.displayName) {
        return `${record.displayName} <${unique}>`;
      }
      return record.displayName;
    }
    if (typeof record.uniqueName === "string") {
      return record.uniqueName;
    }
  }
  return String(value);
}

export function isIdentityValue(value: unknown): boolean {
  if (!value || typeof value !== "object") {
    return false;
  }
  const record = value as Record<string, unknown>;
  return typeof record.displayName === "string" || typeof record.uniqueName === "string";
}

export function normalizeCompare(value: unknown): string {
  if (value === null || value === undefined || value === "") {
    return "";
  }
  if (typeof value === "boolean") {
    return value ? "true" : "false";
  }
  if (typeof value === "number") {
    return String(value);
  }
  return String(value).replace(/\r\n/g, "\n").trim();
}

export function valuesEqual(left: unknown, right: unknown, referenceName?: string): boolean {
  if (isDateTimeField(referenceName) || isIsoDateString(left) || isIsoDateString(right)) {
    return dateToYmd(left) === dateToYmd(right);
  }
  if (typeof left === "number" && typeof right === "number") {
    return left === right;
  }
  return normalizeCompare(left) === normalizeCompare(right);
}

export function parseWorkItemId(value: unknown): number | undefined {
  if (typeof value === "number" && Number.isFinite(value) && value > 0) {
    return Math.trunc(value);
  }
  if (typeof value === "string" && value.trim()) {
    const parsed = Number(value.trim());
    if (Number.isFinite(parsed) && parsed > 0) {
      return Math.trunc(parsed);
    }
  }
  return undefined;
}

export function idsFromQueryResult(result: QueryResultLike): number[] {
  if (result.workItems && result.workItems.length > 0) {
    return result.workItems.map((item) => item.id);
  }

  if (result.workItemRelations && result.workItemRelations.length > 0) {
    const ids: number[] = [];
    const seen = new Set<number>();
    for (const relation of result.workItemRelations) {
      const id = relation.target?.id;
      if (id && !seen.has(id)) {
        seen.add(id);
        ids.push(id);
      }
    }
    return ids;
  }

  return [];
}

export function isTreeQuery(result: QueryResultLike | undefined | null): boolean {
  if (!result) {
    return false;
  }
  const type = String(result.queryType ?? "").toLowerCase();
  if (type === "tree" || type === "2" || type === "onehop" || type === "3") {
    return true;
  }
  return Boolean(result.workItemRelations?.some((relation) => relation.source?.id && relation.target?.id));
}

export function mergeColumns(queryColumns?: QueryColumn[], requested?: string[]): WorkbookColumn[] {
  const byRef = new Map<string, WorkbookColumn>();

  const friendlyName = new Map(DEFAULT_COLUMNS.map((column) => [column.referenceName, column.name]));

  const add = (referenceName: string, name?: string) => {
    if (!referenceName || byRef.has(referenceName)) {
      return;
    }
    byRef.set(referenceName, {
      referenceName,
      name: name || friendlyName.get(referenceName) || referenceName.replace(/^.*\./, ""),
      identity: IDENTITY_FIELDS.has(referenceName)
    });
  };

  for (const field of REQUIRED_FIELDS) {
    add(field);
  }

  if (requested && requested.length > 0) {
    for (const field of requested) {
      add(field);
    }
  } else if (queryColumns) {
    for (const column of queryColumns) {
      if (column.referenceName) {
        add(column.referenceName, column.name);
      }
    }
  }

  return Array.from(byRef.values());
}

export function markIdentityColumns(columns: WorkbookColumn[], items: WorkItemLike[] = []): WorkbookColumn[] {
  return columns.map((column) => {
    if (column.identity || IDENTITY_FIELDS.has(column.referenceName)) {
      return { ...column, identity: true };
    }
    const identity = items.some((item) => isIdentityValue(item.fields?.[column.referenceName]));
    return identity ? { ...column, identity: true } : column;
  });
}

export function markDateColumns(columns: WorkbookColumn[], items: WorkItemLike[] = []): WorkbookColumn[] {
  return columns.map((column) => {
    if (column.date || isDateTimeField(column.referenceName)) {
      return { ...column, date: true };
    }
    const date = items.some((item) => isIsoDateString(item.fields?.[column.referenceName]));
    return date ? { ...column, date: true } : column;
  });
}

export function withTitleLevels(columns: WorkbookColumn[], _levels?: number): WorkbookColumn[] {
  return singleTitleColumns(columns);
}

export function singleTitleColumns(columns: WorkbookColumn[]): WorkbookColumn[] {
  const rest = columns.filter((column) => column.referenceName !== "System.Title" && !column.titleLevel);
  const title: WorkbookColumn = { referenceName: "System.Title", name: "Title" };
  const typeIndex = rest.findIndex((column) => column.referenceName === "System.WorkItemType");
  if (typeIndex >= 0) {
    return [...rest.slice(0, typeIndex + 1), title, ...rest.slice(typeIndex + 1)];
  }
  return [title, ...rest];
}

export function treeLevelCount(columns: WorkbookColumn[]): number {
  const levels = columns.filter((column) => column.titleLevel).map((column) => column.titleLevel || 0);
  return levels.length ? Math.max(...levels) : 0;
}

export function isReadOnlyField(referenceName: string): boolean {
  return READ_ONLY_FIELDS.has(referenceName);
}

export function isSyntheticField(referenceName: string): boolean {
  return referenceName === PARENT_TITLE_FIELD || referenceName.startsWith("Ado.");
}

export function hasParentTitleColumn(columns: WorkbookColumn[]): boolean {
  return columns.some((column) => column.referenceName === PARENT_TITLE_FIELD && !column.titleLevel);
}

export function fieldReferenceNames(columns: WorkbookColumn[]): string[] {
  const names = new Set<string>();
  for (const column of columns) {
    const referenceName = column.titleLevel ? "System.Title" : column.referenceName;
    if (isSyntheticField(referenceName)) {
      continue;
    }
    names.add(referenceName);
  }
  names.add("System.Title");
  names.add("System.Parent");
  return Array.from(names);
}

export function titleMapFromItems(items: WorkItemLike[]): Map<number, string> {
  const titles = new Map<number, string>();
  for (const item of items) {
    if (!item.id) {
      continue;
    }
    titles.set(item.id, String(item.fields?.["System.Title"] || "").trim());
  }
  return titles;
}

export function applyParentTitles(items: WorkItemLike[], titles: Map<number, string>): WorkItemLike[] {
  return items.map((item) => {
    const parentId = parentIdFromRelations(item);
    const parentTitle = parentId ? titles.get(parentId) || "" : "";
    return {
      ...item,
      fields: {
        ...item.fields,
        [PARENT_TITLE_FIELD]: parentTitle
      }
    };
  });
}

export function missingParentIds(items: WorkItemLike[], titles: Map<number, string>): number[] {
  const missing = new Set<number>();
  for (const item of items) {
    const parentId = parentIdFromRelations(item);
    if (parentId && !titles.has(parentId)) {
      missing.add(parentId);
    }
  }
  return Array.from(missing);
}

export function workItemToSnapshot(item: WorkItemLike, columns: WorkbookColumn[]): Record<string, CellValue> {
  const snapshot: Record<string, CellValue> = {};
  const seen = new Set<string>();
  for (const column of columns) {
    const referenceName = column.titleLevel ? "System.Title" : column.referenceName;
    if (seen.has(referenceName)) {
      continue;
    }
    seen.add(referenceName);
    snapshot[referenceName] = fieldToCell(
      item.fields?.[referenceName],
      referenceName,
      column.html ? [referenceName] : undefined
    );
  }
  return snapshot;
}

export function itemsToTable(
  items: WorkItemLike[],
  columns: WorkbookColumn[],
  layout?: TreeLayout | null
): CellValue[][] {
  return items.map((item) =>
    columns.map((column) => {
      if (column.titleLevel) {
        const depth = item.id && layout ? layout.depthById[item.id] || 1 : 1;
        return depth === column.titleLevel
          ? fieldToCell(item.fields?.["System.Title"], "System.Title")
          : "";
      }
      return fieldToCell(
        item.fields?.[column.referenceName],
        column.referenceName,
        column.html ? [column.referenceName] : undefined
      );
    })
  );
}

export function rowsToTable(rows: TableRow[], columns: WorkbookColumn[]): CellValue[][] {
  return rows.map((row) =>
    columns.map((column) => {
      if (column.titleLevel) {
        return (row.treeLevel || 1) === column.titleLevel ? row.values["System.Title"] ?? "" : "";
      }
      return row.values[column.referenceName] ?? "";
    })
  );
}

export function sanitizeFileName(name: string): string {
  const safe = name.replace(/[<>:"/\\|?*\u0000-\u001f]/g, "_").trim();
  return (safe || "work-items").slice(0, 80);
}

export function normalizeOrgUrl(input: string): string {
  return input.trim().replace(/\/+$/, "");
}

/** Lower-cased org identity; `https://org.visualstudio.com` and `https://dev.azure.com/org` are the same organization. */
export function organizationKey(orgUrl: string): string {
  const normalized = normalizeOrgUrl(orgUrl).toLowerCase();
  const legacy = normalized.match(/^https?:\/\/([^./]+)\.visualstudio\.com$/);
  if (legacy) {
    return `https://dev.azure.com/${legacy[1]}`;
  }
  return normalized.replace(/^http:\/\//, "https://");
}

export function sameOrganization(left: string, right: string): boolean {
  return organizationKey(left) === organizationKey(right);
}

export function workItemWebUrl(orgUrl: string, project: string, id: number): string {
  return `${normalizeOrgUrl(orgUrl)}/${encodeURIComponent(project)}/_workitems/edit/${id}`;
}

/** Opens the saved query's results page in Azure DevOps. */
export function queryWebUrl(orgUrl: string, project: string, queryId: string): string {
  return `${normalizeOrgUrl(orgUrl)}/${encodeURIComponent(project)}/_queries/query/${encodeURIComponent(queryId)}/`;
}

export function workItemIdFromRelationUrl(url?: string): number | undefined {
  const match = String(url || "").match(/workItems\/(\d+)/i);
  return match ? Number(match[1]) : undefined;
}

export function encodeBasicPat(pat: string): string {
  const raw = `:${pat}`;
  if (typeof btoa === "function") {
    return btoa(raw);
  }
  return Buffer.from(raw, "utf8").toString("base64");
}

export function vsspsOrigin(orgUrl: string): string {
  const url = normalizeOrgUrl(orgUrl);
  const dev = url.match(/^https:\/\/dev\.azure\.com\/([^/]+)/i);
  if (dev) {
    return `https://vssps.dev.azure.com/${dev[1]}`;
  }
  const visualStudio = url.match(/^https:\/\/([^.]+)\.visualstudio\.com/i);
  if (visualStudio) {
    return `https://${visualStudio[1]}.vssps.visualstudio.com`;
  }
  return url;
}

export function formatIdentityCell(identity: { displayName: string; uniqueName?: string }): string {
  if (identity.uniqueName && identity.uniqueName !== identity.displayName) {
    return `${identity.displayName} <${identity.uniqueName}>`;
  }
  return identity.displayName;
}

export function uniqueSortedLabels(values: string[], limit = MAX_IDENTITY_PICKLIST): string[] {
  const seen = new Set<string>();
  const labels: string[] = [];
  for (const value of values) {
    const trimmed = String(value || "").trim();
    if (!trimmed) {
      continue;
    }
    const key = trimmed.toLowerCase();
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    labels.push(trimmed);
  }
  labels.sort((left, right) => left.localeCompare(right));
  return labels.slice(0, limit);
}

export function identityLabelsFromPicks(picks: Array<{ displayName: string; uniqueName?: string }>): string[] {
  return uniqueSortedLabels(picks.map((pick) => formatIdentityCell(pick)));
}

export function identityLabelsFromSnapshot(
  snapshot: Record<string, Record<string, CellValue>> | undefined,
  columns: WorkbookColumn[]
): string[] {
  const refs = columns.filter((column) => column.identity).map((column) => column.referenceName);
  const values: string[] = [];
  for (const row of Object.values(snapshot || {})) {
    for (const ref of refs) {
      const cell = row[ref];
      if (cell !== undefined && cell !== null && String(cell).trim()) {
        values.push(String(cell));
      }
    }
  }
  return uniqueSortedLabels(values);
}

export function parentIdFromRelations(item: WorkItemLike): number | undefined {
  for (const relation of item.relations || []) {
    if (relation.rel === "System.LinkTypes.Hierarchy-Reverse" && relation.url) {
      const id = workItemIdFromRelationUrl(relation.url);
      if (id) {
        return id;
      }
    }
  }
  return parseWorkItemId(item.fields?.["System.Parent"]);
}
