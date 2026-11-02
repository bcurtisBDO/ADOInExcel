import * as XLSX from "xlsx";
import {
  META_SHEET,
  WORK_ITEMS_SHEET,
  CellValue,
  TreeLayout,
  WorkbookColumn,
  WorkbookMetadata,
  WorkItemLike,
  dateToExcelSerial,
  excelShortDateFormat,
  headersFromColumns,
  isIsoDateString,
  itemsToTable,
  packMetadataSheet,
  parseDateOnly,
  sanitizeFileName,
  topAlignStylesXml
} from "@ado-open-in-excel/shared";

export interface WorkbookFile {
  blob: Blob;
  fileName: string;
}

export function buildWorkbookBuffer(
  items: WorkItemLike[],
  columns: WorkbookColumn[],
  metadata: WorkbookMetadata,
  layout?: TreeLayout | null
): ArrayBuffer {
  const headers = headersFromColumns(columns);
  const data = itemsToTable(items, columns, layout);
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([headers, ...data] as CellValue[][]);
  applyDateFormats(sheet, columns, data.length);
  XLSX.utils.book_append_sheet(workbook, sheet, WORK_ITEMS_SHEET);

  const metaSheet = XLSX.utils.aoa_to_sheet(packMetadataSheet(metadata) as CellValue[][]);
  XLSX.utils.book_append_sheet(workbook, metaSheet, META_SHEET);
  workbook.Workbook = workbook.Workbook || {};
  workbook.Workbook.Sheets = (workbook.SheetNames || []).map((name) => ({
    name,
    Hidden: name === META_SHEET ? 1 : 0
  }));

  return topAlignCells(XLSX.write(workbook, { bookType: "xlsx", type: "array" }) as ArrayBuffer);
}

function topAlignCells(buffer: ArrayBuffer): ArrayBuffer {
  try {
    const zip = XLSX.CFB.read(new Uint8Array(buffer), { type: "array" });
    const index = (zip.FullPaths as string[]).findIndex((path) => path.endsWith("/xl/styles.xml"));
    if (index < 0) {
      return buffer;
    }
    const xml = new TextDecoder().decode(zip.FileIndex[index].content);
    XLSX.CFB.utils.cfb_add(zip, "/xl/styles.xml", new TextEncoder().encode(topAlignStylesXml(xml)));
    const patched = XLSX.CFB.write(zip, { fileType: "zip", type: "array" }) as ArrayLike<number>;
    return Uint8Array.from(patched).buffer;
  } catch {
    // Alignment is cosmetic; keep the unpatched workbook rather than fail the download.
    return buffer;
  }
}

function applyDateFormats(sheet: XLSX.WorkSheet, columns: WorkbookColumn[], dataRows: number): void {
  const format = excelShortDateFormat();
  columns.forEach((column, colIndex) => {
    if (!column.date) {
      return;
    }
    for (let row = 1; row <= dataRows; row += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: colIndex });
      const cell = sheet[address] as XLSX.CellObject | undefined;
      if (!cell) {
        continue;
      }
      if (typeof cell.v === "string" && (isIsoDateString(cell.v) || parseDateOnly(cell.v))) {
        const parts = parseDateOnly(cell.v);
        if (parts) {
          cell.v = dateToExcelSerial(parts);
        }
      }
      if (typeof cell.v === "number") {
        cell.t = "n";
        cell.z = format;
      }
    }
  });
}

export function workbookFile(buffer: ArrayBuffer, name: string): WorkbookFile {
  return {
    blob: new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }),
    fileName: `${sanitizeFileName(name)}.xlsx`
  };
}

export function triggerDownload(blob: Blob, fileName: string): void {
  const nav = window.navigator as Navigator & {
    msSaveOrOpenBlob?: (file: Blob, fileName: string) => boolean;
  };
  if (typeof nav.msSaveOrOpenBlob === "function") {
    nav.msSaveOrOpenBlob(blob, fileName);
    return;
  }
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.style.display = "none";
  document.body.appendChild(link);
  link.click();
  window.setTimeout(() => {
    link.remove();
    URL.revokeObjectURL(url);
  }, 60_000);
}
