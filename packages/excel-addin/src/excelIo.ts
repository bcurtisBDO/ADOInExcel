/// <reference types="office-js" />
import {
  META_SHEET,
  PICKLIST_SHEET,
  TABLE_NAME,
  WORK_ITEMS_SHEET,
  CellValue,
  TableRow,
  TreeLayout,
  WorkbookColumn,
  WorkbookMetadata,
  WorkItemLike,
  collapseTitleLevels,
  childOutlineRanges,
  excelShortDateFormat,
  headersFromColumns,
  itemsToTable,
  packMetadataSheet,
  parseMetadata,
  parseMetadataSheet,
  parseWorkItemId,
  rowsToTable,
  treeLevelFromParents,
  unpackClassificationPathGrid,
  uniqueSortedLabels,
  identityLabelsFromSnapshot,
  isReadOnlyField,
  workItemWebUrl
} from "@ado-open-in-excel/shared";

export interface SheetSelection {
  rowIndex: number;
  columnName: string;
  value: string;
}

export interface PathLists {
  areaPaths?: string[];
  iterationPaths?: string[];
}

let pathLists: PathLists = {};

export function setClassificationPaths(lists: PathLists): void {
  pathLists = {
    areaPaths: lists.areaPaths || [],
    iterationPaths: lists.iterationPaths || []
  };
}

let identityLabels: string[] = [];

export function setIdentityLabels(labels: string[]): void {
  identityLabels = labels || [];
}

async function getSheet(context: Excel.RequestContext, name: string, create: boolean): Promise<Excel.Worksheet | null> {
  const sheet = context.workbook.worksheets.getItemOrNullObject(name);
  await context.sync();
  if (sheet.isNullObject) {
    if (!create) {
      return null;
    }
    const created = context.workbook.worksheets.add(name);
    await context.sync();
    return created;
  }
  return sheet;
}

const DOCUMENT_META_COUNT = "adoOpenInExcel.metaCount";
const DOCUMENT_META_CHUNK = "adoOpenInExcel.meta.";
const DOCUMENT_META_CHUNK_SIZE = 8000;
const WORKBOOK_BINDING_SETTING = "adoOpenInExcel.binding";

let sessionMetadata: WorkbookMetadata | null = null;

export async function writeMetadata(metadata: WorkbookMetadata): Promise<void> {
  sessionMetadata = metadata;
  try {
    await Excel.run(async (context) => {
      const sheet = await getSheet(context, META_SHEET, true);
      if (!sheet) {
        return;
      }
      sheet.visibility = Excel.SheetVisibility.visible;
      await context.sync();
      const used = sheet.getUsedRangeOrNullObject();
      used.load(["isNullObject", "rowCount", "columnCount"]);
      await context.sync();
      if (!used.isNullObject) {
        used.clear();
        await context.sync();
      }
      const rows = writeMetaCells(sheet, metadata);
      await context.sync();
      const leftover = sheet.getUsedRangeOrNullObject();
      leftover.load(["isNullObject", "rowCount", "columnCount"]);
      await context.sync();
      if (!leftover.isNullObject && leftover.rowCount > rows) {
        sheet.getRangeByIndexes(rows, 0, leftover.rowCount - rows, leftover.columnCount).clear();
        await context.sync();
      }
      await writeWorkbookBinding(context, metadata);
      sheet.visibility = Excel.SheetVisibility.hidden;
      await context.sync();
    });
  } catch {
    // Sheet write is best-effort; session and document settings still bind the workbook.
  }
  await saveDocumentMetadata(metadata);
}

export async function readMetadata(): Promise<WorkbookMetadata | null> {
  if (sessionMetadata) {
    applyPathLists(sessionMetadata);
    return sessionMetadata;
  }
  const fromSettings = readDocumentMetadata();
  if (fromSettings) {
    sessionMetadata = fromSettings;
    applyPathLists(fromSettings);
    return fromSettings;
  }
  return Excel.run(async (context) => {
    let sheet = await getSheet(context, META_SHEET, false);
    if (!sheet) {
      sheet = await findMetadataSheet(context);
    }
    if (sheet) {
      const values = await readSheetValues(context, sheet);
      if (values?.length) {
        const metadata = parseMetadataSheet(values);
        sheet.visibility = Excel.SheetVisibility.hidden;
        await context.sync();
        if (metadata) {
          const pathGrid = values.map((row: unknown[]) => (row || []).slice(1));
          const persisted = unpackClassificationPathGrid(pathGrid);
          const merged: WorkbookMetadata = {
            ...metadata,
            areaPaths: persisted.areaPaths.length ? persisted.areaPaths : metadata.areaPaths,
            iterationPaths: persisted.iterationPaths.length ? persisted.iterationPaths : metadata.iterationPaths
          };
          sessionMetadata = merged;
          applyPathLists(merged);
          return merged;
        }
      } else {
        sheet.visibility = Excel.SheetVisibility.hidden;
        await context.sync();
      }
    }
    const fromWorkbook = await readWorkbookBinding(context);
    if (fromWorkbook) {
      sessionMetadata = fromWorkbook;
      applyPathLists(fromWorkbook);
    }
    return fromWorkbook;
  });
}

function applyPathLists(metadata: WorkbookMetadata): void {
  if (!pathLists.areaPaths?.length && (metadata.areaPaths?.length || metadata.iterationPaths?.length)) {
    setClassificationPaths({
      areaPaths: metadata.areaPaths,
      iterationPaths: metadata.iterationPaths
    });
  }
}

function readDocumentMetadata(): WorkbookMetadata | null {
  const settings = typeof Office !== "undefined" ? Office.context?.document?.settings : undefined;
  if (!settings) {
    return null;
  }
  const count = Number(settings.get(DOCUMENT_META_COUNT) || 0);
  if (!count) {
    return null;
  }
  const json = Array.from({ length: count }, (_, index) => String(settings.get(`${DOCUMENT_META_CHUNK}${index}`) || "")).join(
    ""
  );
  return parseMetadata(json);
}

async function saveDocumentMetadata(metadata: WorkbookMetadata): Promise<void> {
  const settings = typeof Office !== "undefined" ? Office.context?.document?.settings : undefined;
  if (!settings) {
    return;
  }
  if (await writeDocumentSettings(settings, metadata)) {
    return;
  }
  await writeDocumentSettings(settings, compactMetadata(metadata));
}

async function writeDocumentSettings(
  settings: Office.Settings,
  metadata: WorkbookMetadata
): Promise<boolean> {
  const json = JSON.stringify(metadata);
  const count = Math.max(1, Math.ceil(json.length / DOCUMENT_META_CHUNK_SIZE));
  const previous = Number(settings.get(DOCUMENT_META_COUNT) || 0);
  for (let index = 0; index < count; index += 1) {
    settings.set(
      `${DOCUMENT_META_CHUNK}${index}`,
      json.slice(index * DOCUMENT_META_CHUNK_SIZE, (index + 1) * DOCUMENT_META_CHUNK_SIZE)
    );
  }
  for (let index = count; index < previous; index += 1) {
    settings.remove(`${DOCUMENT_META_CHUNK}${index}`);
  }
  settings.set(DOCUMENT_META_COUNT, String(count));
  return new Promise((resolve) => {
    settings.saveAsync((result) => {
      resolve(result?.status !== Office.AsyncResultStatus.Failed);
    });
  });
}

function compactMetadata(metadata: WorkbookMetadata): WorkbookMetadata {
  return {
    version: metadata.version,
    orgUrl: metadata.orgUrl,
    project: metadata.project,
    action: metadata.action,
    listType: metadata.listType,
    treeLevels: metadata.treeLevels,
    queryId: metadata.queryId,
    queryName: metadata.queryName,
    workItemIds: metadata.workItemIds,
    columns: metadata.columns,
    loadedAt: metadata.loadedAt,
    revisions: metadata.revisions,
    snapshot: {},
    parents: metadata.parents,
    workItemTypes: metadata.workItemTypes,
    statesByType: metadata.statesByType,
    htmlFields: metadata.htmlFields,
    identityFields: metadata.identityFields,
    transitionsByType: metadata.transitionsByType
  };
}

async function writeWorkbookBinding(context: Excel.RequestContext, metadata: WorkbookMetadata): Promise<void> {
  try {
    const json = JSON.stringify(compactMetadata(metadata));
    const setting = context.workbook.settings.getItemOrNullObject(WORKBOOK_BINDING_SETTING);
    setting.load("isNullObject");
    await context.sync();
    if (setting.isNullObject) {
      context.workbook.settings.add(WORKBOOK_BINDING_SETTING, json);
    } else {
      setting.value = json;
    }
    await context.sync();
  } catch {
    // Workbook settings are a fallback when the hidden sheet cannot be read.
  }
}

async function readWorkbookBinding(context: Excel.RequestContext): Promise<WorkbookMetadata | null> {
  try {
    const setting = context.workbook.settings.getItemOrNullObject(WORKBOOK_BINDING_SETTING);
    setting.load(["isNullObject", "value"]);
    await context.sync();
    if (setting.isNullObject) {
      return null;
    }
    return parseMetadata(setting.value);
  } catch {
    return null;
  }
}

async function findMetadataSheet(context: Excel.RequestContext): Promise<Excel.Worksheet | null> {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();
  const match = sheets.items.find((item) => item.name.replace(/^_/, "").toLowerCase() === "adoopeninexcel");
  return match || null;
}

async function readSheetValues(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet
): Promise<unknown[][] | null> {
  sheet.visibility = Excel.SheetVisibility.visible;
  await context.sync();
  const used = sheet.getUsedRangeOrNullObject();
  used.load(["values", "isNullObject"]);
  await context.sync();
  if (!used.isNullObject && used.values?.length) {
    return used.values as unknown[][];
  }
  const fallback = sheet.getRange("A1:C4000");
  fallback.load("values");
  await context.sync();
  const rows = (fallback.values || []).filter((row) =>
    (row || []).some((cell) => cell !== null && cell !== undefined && String(cell).trim() !== "")
  );
  return rows.length ? (rows as unknown[][]) : null;
}

export interface WriteGridOptions {
  skipAutofit?: boolean;
  skipValidation?: boolean;
  skipHyperlinks?: boolean;
}

export async function writeWorkItems(
  items: WorkItemLike[],
  columns: WorkbookColumn[],
  metadata: WorkbookMetadata,
  layout?: TreeLayout | null,
  options?: WriteGridOptions
): Promise<void> {
  const treeLevels =
    metadata.listType === "tree"
      ? items.map((item) =>
          item.id && layout?.depthById[item.id]
            ? layout.depthById[item.id]
            : treeLevelFromParents(item.id, metadata.parents) || 1
        )
      : undefined;
  await writeGrid(
    [headersFromColumns(columns), ...itemsToTable(items, columns, layout)],
    metadata,
    columns,
    treeLevels,
    options
  );
}

export async function writeRows(
  rows: TableRow[],
  columns: WorkbookColumn[],
  metadata: WorkbookMetadata,
  options?: WriteGridOptions
): Promise<void> {
  const treeLevels =
    metadata.listType === "tree" ? rows.map((row) => row.treeLevel || 1) : undefined;
  await writeGrid(
    [headersFromColumns(columns), ...rowsToTable(rows, columns)],
    metadata,
    columns,
    treeLevels,
    options
  );
}

async function writeGrid(
  values: CellValue[][],
  metadata: WorkbookMetadata,
  columns: WorkbookColumn[],
  treeLevels?: number[],
  options?: WriteGridOptions
): Promise<void> {
  let committed = false;
  try {
    await Excel.run(async (context) => {
      const sheet = await getSheet(context, WORK_ITEMS_SHEET, true);
      if (!sheet) {
        return;
      }

      const selection = context.workbook.getSelectedRange();
      selection.load(["rowIndex", "columnIndex"]);
      const selectionSheet = selection.worksheet;
      selectionSheet.load("name");
      const existing = sheet.tables.getItemOrNullObject(TABLE_NAME);
      existing.load("name");
      await context.sync();

      const headers = values[0] || [];
      const restore =
        selectionSheet.name === WORK_ITEMS_SHEET
          ? { row: selection.rowIndex, col: selection.columnIndex }
          : undefined;

      let table: Excel.Table;
      if (!existing.isNullObject && (await tableHeadersMatch(context, existing, headers))) {
        try {
          await replaceTableValues(context, sheet, existing, values, columns, options);
          table = existing;
        } catch {
          table = await recreateTable(context, sheet, values, columns, options);
        }
      } else {
        table = await recreateTable(context, sheet, values, columns, options);
      }
      table.getRange().format.verticalAlignment = Excel.VerticalAlignment.top;
      await context.sync();
      committed = true;

      if (!options?.skipHyperlinks) {
        try {
          await applyIdHyperlinks(context, table, metadata);
        } catch {
          // Hyperlinks are best-effort.
        }
      }
      try {
        await applyTreePresentation(context, sheet, table, treeLevels || [], metadata.listType === "tree");
      } catch {
        // Indent and outline are best-effort.
      }

      try {
        sheet.activate();
        await restoreGridSelection(context, sheet, values.length, headers.length, restore);
      } catch {
        // Selection restore is best-effort.
      }
    });
  } finally {
    if (committed) {
      await writeMetadata(metadata);
      if (!options?.skipValidation) {
        try {
          await reapplyProcessValidation(metadata);
        } catch {
          // Dropdowns are best-effort; the sheet data is already written.
        }
      }
    }
  }
}

async function tableHeadersMatch(
  context: Excel.RequestContext,
  table: Excel.Table,
  headers: CellValue[]
): Promise<boolean> {
  const headerRange = table.getHeaderRowRange();
  headerRange.load("values");
  await context.sync();
  const current = (headerRange.values?.[0] || []).map((value) => String(value));
  return current.length === headers.length && current.every((name, index) => name === String(headers[index]));
}

async function replaceTableValues(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  table: Excel.Table,
  values: CellValue[][],
  columns: WorkbookColumn[],
  options?: WriteGridOptions
): Promise<void> {
  const rows = Math.max(values.length, 1);
  const cols = Math.max((values[0] || []).length, 1);
  const used = sheet.getUsedRangeOrNullObject();
  used.load(["rowCount", "columnCount", "isNullObject"]);
  await context.sync();
  const target = sheet.getRangeByIndexes(0, 0, rows, cols);
  table.resize(target);
  target.numberFormat = gridNumberFormats(values, columns);
  target.values = values;
  if (!used.isNullObject) {
    if (used.rowCount > rows) {
      sheet.getRangeByIndexes(rows, 0, used.rowCount - rows, used.columnCount).clear();
    }
    if (used.columnCount > cols) {
      sheet.getRangeByIndexes(0, cols, Math.max(used.rowCount, rows), used.columnCount - cols).clear();
    }
  }
  if (!options?.skipAutofit) {
    target.format.autofitColumns();
  }
}

async function recreateTable(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  values: CellValue[][],
  columns: WorkbookColumn[],
  options?: WriteGridOptions
): Promise<Excel.Table> {
  await clearRowOutline(context, sheet);
  const tables = sheet.tables;
  tables.load("items/name");
  await context.sync();
  for (const item of tables.items) {
    if (item.name === TABLE_NAME) {
      item.delete();
    }
  }
  const used = sheet.getUsedRangeOrNullObject();
  used.clear();
  await context.sync();

  const headers = values[0] || [];
  const range = sheet.getRangeByIndexes(0, 0, values.length, headers.length);
  range.numberFormat = gridNumberFormats(values, columns);
  range.values = values;

  const table = sheet.tables.add(range, true);
  table.name = TABLE_NAME;
  table.style = "TableStyleMedium7";
  const headerRange = table.getHeaderRowRange();
  headerRange.format.fill.color = "#217346";
  headerRange.format.font.color = "#FFFFFF";
  headerRange.format.font.bold = true;
  if (!options?.skipAutofit) {
    range.format.autofitColumns();
  }
  return table;
}

async function restoreGridSelection(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  rowCount: number,
  columnCount: number,
  preferred?: { row: number; col: number }
): Promise<void> {
  // tables.add() selects the table and Excel switches to Table Design. Select a
  // cell outside the table first so that tab closes, then a single cell inside.
  sheet.getRangeByIndexes(0, Math.max(columnCount, 1), 1, 1).select();
  await context.sync();
  const lastRow = Math.max(rowCount - 1, 0);
  const lastCol = Math.max(columnCount - 1, 0);
  const row = Math.min(Math.max(preferred?.row ?? Math.min(1, lastRow), 0), lastRow);
  const col = Math.min(Math.max(preferred?.col ?? 0, 0), lastCol);
  sheet.getRangeByIndexes(row, col, 1, 1).select();
  await context.sync();
}

async function clearRowOutline(context: Excel.RequestContext, sheet: Excel.Worksheet): Promise<void> {
  for (let attempt = 0; attempt < 8; attempt += 1) {
    try {
      const used = sheet.getUsedRangeOrNullObject();
      used.load("isNullObject");
      await context.sync();
      if (used.isNullObject) {
        break;
      }
      used.ungroup(Excel.GroupOption.byRows);
      await context.sync();
    } catch {
      break;
    }
  }
}

async function applyTreePresentation(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  table: Excel.Table,
  treeLevels: number[],
  isTree: boolean
): Promise<void> {
  await clearRowOutline(context, sheet);
  table.columns.load("items/name, items/index");
  const dataBody = table.getDataBodyRange();
  dataBody.load("rowCount");
  await context.sync();
  const count = dataBody.rowCount;
  if (!count) {
    return;
  }

  if (!isTree) {
    const titleColumn = findTableColumn(table, ["Title", "System.Title"]);
    if (titleColumn) {
      const titleBody = titleColumn.getDataBodyRange();
      for (let index = 0; index < count; index += 1) {
        titleBody.getCell(index, 0).format.indentLevel = 0;
      }
    }
    dataBody.format.fill.clear();
    await context.sync();
    return;
  }

  const titleColumn = findTableColumn(table, ["Title", "System.Title"]);
  if (!titleColumn) {
    return;
  }
  const titleBody = titleColumn.getDataBodyRange();
  const levels = Array.from({ length: count }, (_, index) => Math.max(1, treeLevels[index] || 1));
  table.showBandedRows = false;
  for (let index = 0; index < count; index += 1) {
    titleBody.getCell(index, 0).format.indentLevel = Math.min(15, levels[index] - 1);
    dataBody.getRow(index).format.fill.color = treeLevelFill(levels[index]);
  }
  await context.sync();
  for (const range of childOutlineRanges(levels)) {
    const rowCount = range.end - range.start + 1;
    if (rowCount < 1) {
      continue;
    }
    sheet.getRangeByIndexes(range.start + 1, 0, rowCount, 1).group(Excel.GroupOption.byRows);
  }
  try {
    await context.sync();
    sheet.showOutlineLevels(8, 1);
    await context.sync();
  } catch {
    // Excel allows a limited number of outline levels.
  }
}

/** Soft green fills by nest depth; deeper levels get lighter so roots stay easiest to spot. */
const TREE_LEVEL_FILLS = ["#C6D9C6", "#D9E8D9", "#E6F0E6", "#EFF5EF", "#F5F8F5", "#FAFCFA"];

function treeLevelFill(level: number): string {
  const index = Math.min(Math.max(level, 1), TREE_LEVEL_FILLS.length) - 1;
  return TREE_LEVEL_FILLS[index];
}

export async function restoreTreePresentation(metadata: WorkbookMetadata): Promise<void> {
  if (metadata.listType !== "tree") {
    return;
  }
  const rows = await readTableRows(metadata.columns);
  await Excel.run(async (context) => {
    const sheet = await getSheet(context, WORK_ITEMS_SHEET, false);
    if (!sheet) {
      return;
    }
    const table = sheet.tables.getItemOrNullObject(TABLE_NAME);
    await context.sync();
    if (table.isNullObject) {
      return;
    }
    await applyTreePresentation(
      context,
      sheet,
      table,
      rows.map((row) => row.treeLevel || 1),
      true
    );
    await context.sync();
  });
}

async function applyIdHyperlinks(
  context: Excel.RequestContext,
  table: Excel.Table,
  metadata: WorkbookMetadata
): Promise<void> {
  table.columns.load("items/name, items/index");
  await context.sync();
  const column = findTableColumn(table, ["ID", "System.Id"]);
  if (!column || !metadata.orgUrl || !metadata.project) {
    return;
  }
  const body = column.getDataBodyRange();
  body.load("values, rowCount");
  await context.sync();
  const max = Math.min(body.rowCount, 2000);
  const chunkSize = 250;
  for (let start = 0; start < max; start += chunkSize) {
    const end = Math.min(start + chunkSize, max);
    for (let index = start; index < end; index += 1) {
      const id = Number(body.values?.[index]?.[0]);
      if (!Number.isFinite(id) || id <= 0) {
        continue;
      }
      body.getCell(index, 0).hyperlink = {
        address: workItemWebUrl(metadata.orgUrl, metadata.project, id),
        screenTip: `Open work item ${id} in Azure DevOps`,
        textToDisplay: String(id)
      };
    }
    await context.sync();
  }
}

function writeMetaCells(sheet: Excel.Worksheet, metadata: WorkbookMetadata): number {
  const grid = packMetadataSheet(metadata);
  const rows = Math.max(grid.length, 1);
  const cols = Math.max((grid[0] || []).length, 1);
  const range = sheet.getRangeByIndexes(0, 0, rows, cols);
  range.numberFormat = Array.from({ length: rows }, () => Array.from({ length: cols }, () => "@"));
  range.values = grid;
  return rows;
}

function gridNumberFormats(values: CellValue[][], columns: WorkbookColumn[]): string[][] {
  const dateFormat = excelShortDateFormat();
  return values.map((row, rowIndex) =>
    row.map((_, colIndex) => (rowIndex === 0 || !columns[colIndex]?.date ? "@" : dateFormat))
  );
}

function columnLetter(index: number): string {
  let position = index + 1;
  let letter = "";
  while (position > 0) {
    const remainder = (position - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    position = Math.floor((position - 1) / 26);
  }
  return letter;
}

async function applyProcessValidation(
  context: Excel.RequestContext,
  table: Excel.Table,
  metadata: WorkbookMetadata
): Promise<void> {
  const types = metadata.workItemTypes || [];
  const statesByType = metadata.statesByType || {};
  const extraLists: Array<{ header: string; values: string[]; stop: boolean }> = [];
  const areaPaths = pathLists.areaPaths?.length ? pathLists.areaPaths : metadata.areaPaths || [];
  const iterationPaths = pathLists.iterationPaths?.length
    ? pathLists.iterationPaths
    : metadata.iterationPaths || [];
  if (areaPaths.length > 0) {
    extraLists.push({ header: "Area Path", values: areaPaths, stop: true });
  }
  if (iterationPaths.length > 0) {
    extraLists.push({ header: "Iteration Path", values: iterationPaths, stop: true });
  }
  const suggested = new Set((metadata.suggestedPicklists || []).map((name) => name.toLowerCase()));
  for (const [referenceName, values] of Object.entries(metadata.picklists || {})) {
    if (!values.length) {
      continue;
    }
    extraLists.push({
      header: referenceName,
      values,
      stop: !suggested.has(referenceName.toLowerCase())
    });
  }
  const people = uniqueSortedLabels([
    ...identityLabels,
    ...(metadata.identities || []),
    ...identityLabelsFromSnapshot(metadata.snapshot, metadata.columns)
  ]);
  if (people.length > 0) {
    for (const column of metadata.columns) {
      if (!column.identity || column.titleLevel || isReadOnlyField(column.referenceName)) {
        continue;
      }
      extraLists.push({
        header: column.referenceName,
        values: people,
        stop: false
      });
    }
  }

  if (types.length === 0 && extraLists.length === 0) {
    return;
  }

  const grid = buildValidationGrid(types, statesByType, extraLists);
  const pick = await writePicklistSheet(context, grid);
  try {
    table.rows.load("count");
    table.columns.load("items/name, items/index");
    await context.sync();
    if (table.rows.count < 1) {
      return;
    }

    if (types.length > 0) {
      const lastType = columnLetter(types.length - 1);
      const typeHeaders = picklistAddress(`$A$1:$${lastType}$1`);
      const typeColumn = findTableColumn(table, ["Work Item Type", "System.WorkItemType"]);
      const stateColumn = findTableColumn(table, ["State", "System.State"]);

      if (typeColumn) {
        applyListValidation(
          typeColumn.getDataBodyRange(),
          `=${typeHeaders}`,
          "Work Item Type",
          "Select a work item type from this project's process template.",
          true
        );
      }

      if (stateColumn && typeColumn && Object.values(statesByType).some((states) => states.length > 0)) {
        const body = stateColumn.getDataBodyRange();
        body.load("rowIndex");
        await context.sync();
        const typeCell = `${columnLetter(typeColumn.index)}${body.rowIndex + 1}`;
        const stateList = picklistAddress("$A$2");
        applyListValidation(
          body,
          `=OFFSET(${stateList},0,MATCH(${typeCell},${typeHeaders},0)-1,COUNTA(OFFSET(${stateList},0,MATCH(${typeCell},${typeHeaders},0)-1,100,1)),1)`,
          "State",
          "Select a state that is valid for this work item type.",
          true
        );
      }
    }

    const listStart = types.length === 0 ? 0 : types.length + 1;
    extraLists.forEach((list, offset) => {
      const column = findListColumn(table, metadata, list.header);
      if (!column) {
        return;
      }
      const col = columnLetter(listStart + offset);
      const lastRow = 1 + list.values.length;
      const identity = Boolean(metadata.columns.find((item) => item.referenceName === list.header)?.identity);
      applyListValidation(
        column.getDataBodyRange(),
        `=${picklistAddress(`$${col}$2:$${col}$${lastRow}`)}`,
        identity ? "Assigned To" : list.header,
        identity
          ? "Choose a person from this project. You can also type a name that is not in the list."
          : "Select a value from this project's process template.",
        list.stop
      );
    });

    for (const column of metadata.columns) {
      if (!column.html) {
        continue;
      }
      const tableColumn = findTableColumn(table, [column.name, column.referenceName]);
      if (tableColumn) {
        tableColumn.getDataBodyRange().format.wrapText = true;
      }
    }

    await context.sync();
  } finally {
    pick.visibility = Excel.SheetVisibility.hidden;
    await context.sync();
  }
}

async function writePicklistSheet(context: Excel.RequestContext, grid: string[][]): Promise<Excel.Worksheet> {
  let pick = context.workbook.worksheets.getItemOrNullObject(PICKLIST_SHEET);
  await context.sync();
  if (pick.isNullObject) {
    pick = context.workbook.worksheets.add(PICKLIST_SHEET);
    await context.sync();
  }
  pick.visibility = Excel.SheetVisibility.visible;
  await context.sync();
  const used = pick.getUsedRangeOrNullObject();
  used.load("isNullObject");
  await context.sync();
  if (!used.isNullObject) {
    used.clear();
    await context.sync();
  }
  const rows = Math.max(grid.length, 1);
  const cols = Math.max((grid[0] || []).length, 1);
  pick.getRangeByIndexes(0, 0, rows, cols).values = grid;
  await context.sync();
  return pick;
}

function picklistAddress(a1: string): string {
  return `'${PICKLIST_SHEET}'!${a1}`;
}

export async function reapplyProcessValidation(metadata: WorkbookMetadata): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = await getSheet(context, WORK_ITEMS_SHEET, false);
    if (!sheet) {
      return;
    }
    const table = sheet.tables.getItemOrNullObject(TABLE_NAME);
    await context.sync();
    if (table.isNullObject) {
      return;
    }
    try {
      await applyProcessValidation(context, table, metadata);
    } catch {
      // Dropdowns are best-effort.
    }
    const header = table.getHeaderRowRange();
    header.load("columnCount");
    table.rows.load("count");
    await context.sync();
    sheet.activate();
    await restoreGridSelection(context, sheet, table.rows.count + 1, header.columnCount || 1);
  });
}

function buildValidationGrid(
  types: string[],
  statesByType: Record<string, string[]>,
  lists: Array<{ header: string; values: string[] }>
): string[][] {
  const maxStates = types.length ? Math.max(1, ...types.map((type) => (statesByType[type] || []).length)) : 0;
  const maxList = Math.max(0, ...lists.map((list) => list.values.length));
  const rows = 1 + Math.max(maxStates, maxList, 1);
  const listStart = types.length === 0 ? 0 : types.length + 1;
  const cols = Math.max(types.length, listStart + lists.length, 1);
  const grid: string[][] = Array.from({ length: rows }, () => Array.from({ length: cols }, () => ""));
  types.forEach((type, col) => {
    grid[0][col] = type;
    (statesByType[type] || []).forEach((state, row) => {
      grid[row + 1][col] = state;
    });
  });
  lists.forEach((list, offset) => {
    const col = listStart + offset;
    grid[0][col] = list.header;
    list.values.forEach((value, row) => {
      grid[row + 1][col] = value;
    });
  });
  return grid;
}

function findTableColumn(table: Excel.Table, names: string[]): Excel.TableColumn | undefined {
  const wanted = new Set(names.map((name) => name.toLowerCase()));
  return table.columns.items.find((column) => wanted.has(column.name.toLowerCase()));
}

function findListColumn(table: Excel.Table, metadata: WorkbookMetadata, header: string): Excel.TableColumn | undefined {
  const column = metadata.columns.find(
    (item) => item.referenceName === header || item.name === header
  );
  const names = column ? [column.name, column.referenceName, header] : [header];
  if (header === "Area Path") {
    names.push("System.AreaPath");
  }
  if (header === "Iteration Path") {
    names.push("System.IterationPath");
  }
  if (header === "System.AssignedTo" || header === "Assigned To") {
    names.push("Assigned To", "System.AssignedTo");
  }
  return findTableColumn(table, names);
}

function applyListValidation(
  range: Excel.Range,
  source: string,
  title: string,
  message: string,
  stop: boolean
): void {
  range.dataValidation.clear();
  range.dataValidation.rule = {
    list: {
      inCellDropDown: true,
      source
    }
  };
  range.dataValidation.errorAlert = {
    showAlert: true,
    style: stop ? Excel.DataValidationAlertStyle.stop : Excel.DataValidationAlertStyle.warning,
    title,
    message
  };
}

export async function readTableRows(columns: WorkbookColumn[]): Promise<TableRow[]> {
  const metadata = await readMetadata();
  return Excel.run(async (context) => {
    const sheet = await getSheet(context, WORK_ITEMS_SHEET, false);
    if (!sheet) {
      return [];
    }

    const table = sheet.tables.getItemOrNullObject(TABLE_NAME);
    await context.sync();

    let headerValues: string[] = [];
    let bodyValues: (string | number | boolean)[][] = [];
    let indentByRow: number[] = [];

    if (!table.isNullObject) {
      table.rows.load("count");
      const header = table.getHeaderRowRange();
      header.load("values");
      await context.sync();
      headerValues = (header.values?.[0] || []).map((value) => String(value));
      if (table.rows.count > 0) {
        const body = table.getDataBodyRange();
        body.load("values");
        const titleIndex = headerValues.findIndex((name) => name === "Title" || name === "System.Title");
        const indentCells: Excel.Range[] = [];
        if (titleIndex >= 0) {
          for (let index = 0; index < table.rows.count; index += 1) {
            const cell = body.getCell(index, titleIndex);
            cell.format.load("indentLevel");
            indentCells.push(cell);
          }
        }
        await context.sync();
        bodyValues = (body.values || []) as (string | number | boolean)[][];
        indentByRow = indentCells.map((cell) => Number(cell.format.indentLevel) || 0);
      }
    } else {
      const used = sheet.getUsedRangeOrNullObject();
      used.load("values");
      await context.sync();
      if (used.isNullObject || !used.values || used.values.length === 0) {
        return [];
      }
      headerValues = (used.values[0] || []).map((value) => String(value));
      bodyValues = used.values.slice(1) as (string | number | boolean)[][];
    }

    const indexByName = new Map<string, number>();
    headerValues.forEach((name, index) => indexByName.set(name, index));
    const hasIndent = indentByRow.some((value) => value > 0);

    return bodyValues.map((line, rowIndex) => {
      const values: Record<string, CellValue> = {};
      const collapsed = collapseTitleLevels(headerValues, line, columns);
      for (const column of columns) {
        if (column.titleLevel) {
          continue;
        }
        const index = indexByName.has(column.name)
          ? indexByName.get(column.name)!
          : indexByName.get(column.referenceName);
        const raw = index === undefined ? "" : line[index];
        values[column.referenceName] = (raw ?? "") as CellValue;
      }
      if (collapsed.title) {
        values["System.Title"] = collapsed.title;
      }
      let treeLevel = collapsed.treeLevel;
      if (treeLevel === undefined && hasIndent) {
        treeLevel = (indentByRow[rowIndex] || 0) + 1;
      }
      if (treeLevel === undefined && metadata?.listType === "tree") {
        treeLevel = treeLevelFromParents(parseWorkItemId(values["System.Id"]), metadata.parents) || 1;
      }
      return {
        rowIndex,
        id: parseWorkItemId(values["System.Id"]),
        values,
        treeLevel
      };
    });
  });
}

export async function replaceCreatedIds(created: Array<{ rowIndex: number; id: number }>): Promise<void> {
  if (created.length === 0) {
    return;
  }

  const metadata = await readMetadata();
  await Excel.run(async (context) => {
    const sheet = await getSheet(context, WORK_ITEMS_SHEET, false);
    if (!sheet) {
      return;
    }
    const table = sheet.tables.getItemOrNullObject(TABLE_NAME);
    await context.sync();
    if (table.isNullObject) {
      return;
    }
    const header = table.getHeaderRowRange();
    header.load("values");
    await context.sync();
    const headers = (header.values?.[0] || []).map((value) => String(value));
    const idIndex = headers.findIndex((name) => name === "ID" || name === "System.Id");
    if (idIndex < 0) {
      return;
    }
    const body = table.getDataBodyRange();
    for (const item of created) {
      const cell = body.getCell(item.rowIndex, idIndex);
      cell.values = [[item.id]];
      if (metadata?.orgUrl && metadata?.project) {
        cell.hyperlink = {
          address: workItemWebUrl(metadata.orgUrl, metadata.project, item.id),
          screenTip: `Open work item ${item.id} in Azure DevOps`,
          textToDisplay: String(item.id)
        };
      }
    }
    await context.sync();
  });
}

export async function getSheetSelection(): Promise<SheetSelection | null> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["rowIndex", "columnIndex", "values"]);
    const sheet = range.worksheet;
    sheet.load("name");
    await context.sync();
    if (sheet.name !== WORK_ITEMS_SHEET || range.rowIndex < 1) {
      return null;
    }
    const header = sheet.getRangeByIndexes(0, range.columnIndex, 1, 1);
    header.load("values");
    await context.sync();
    return {
      rowIndex: range.rowIndex - 1,
      columnName: String(header.values?.[0]?.[0] || ""),
      value: String(range.values?.[0]?.[0] ?? "")
    };
  });
}

export async function setSelectedCell(value: string): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.values = [[value]];
    await context.sync();
  });
}
