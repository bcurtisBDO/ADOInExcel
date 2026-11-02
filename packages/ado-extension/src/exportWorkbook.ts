import { getClient } from "azure-devops-extension-api";
import {
  WorkItemErrorPolicy,
  WorkItemTrackingRestClient
} from "azure-devops-extension-api/WorkItemTracking";
import {
  DEFAULT_COLUMNS,
  MAX_WORK_ITEMS,
  WORK_ITEM_BATCH_SIZE,
  WorkbookColumn,
  WorkItemLike,
  buildMetadata,
  buildTreeLayout,
  chunk,
  fieldReferenceNames,
  idsFromQueryResult,
  markIdentityColumns,
  markDateColumns,
  mergeColumns,
  orderItems,
  QueryResultLike,
  singleTitleColumns
} from "@ado-open-in-excel/shared";
import { DownloadDialogConfig, getProjectContext } from "./host";
import { buildWorkbookBuffer, workbookFile } from "./workbook";

async function fetchWorkItems(project: string, ids: number[], fields: string[]): Promise<WorkItemLike[]> {
  const client = getClient(WorkItemTrackingRestClient);
  const items: WorkItemLike[] = [];
  for (const group of chunk(ids, WORK_ITEM_BATCH_SIZE)) {
    const batch = await client.getWorkItems(
      group,
      project,
      fields,
      undefined,
      undefined,
      WorkItemErrorPolicy.Omit
    );
    items.push(...(batch as WorkItemLike[]));
  }
  return items;
}

export async function exportWorkbook(options: {
  project: string;
  orgUrl: string;
  action: "OpenQuery" | "OpenItems";
  ids: number[];
  columns: WorkbookColumn[];
  fileName: string;
  queryId?: string;
  queryName?: string;
  queryResult?: QueryResultLike;
}): Promise<{ blob: Blob; fileName: string; count: number }> {
  if (options.ids.length === 0) {
    throw new Error("No work items were found to open in Excel.");
  }
  if (options.ids.length > MAX_WORK_ITEMS) {
    throw new Error(`This list has ${options.ids.length} work items. Reduce it below ${MAX_WORK_ITEMS} and try again.`);
  }

  const layout = buildTreeLayout(options.queryResult);
  const baseColumns = markIdentityColumns(layout ? singleTitleColumns(options.columns) : options.columns);
  const fields = fieldReferenceNames(baseColumns);
  const fetched = await fetchWorkItems(options.project, options.ids, fields);
  const columns = markDateColumns(baseColumns, fetched);
  const items = layout ? orderItems(fetched, layout) : fetched;
  const metadata = buildMetadata({
    orgUrl: options.orgUrl,
    project: options.project,
    action: options.action,
    columns,
    items,
    queryId: options.queryId,
    queryName: options.queryName,
    workItemIds: options.action === "OpenItems" ? options.ids : undefined,
    listType: layout ? "tree" : "flat",
    layout
  });

  const buffer = buildWorkbookBuffer(items, columns, metadata, layout);
  return { ...workbookFile(buffer, options.fileName), count: items.length };
}

export async function runDownload(
  config: DownloadDialogConfig
): Promise<{ blob: Blob; fileName: string; count: number }> {
  const { name: project, orgUrl } = await getProjectContext();
  const client = getClient(WorkItemTrackingRestClient);

  if (config.kind === "query") {
    if (!config.queryId) {
      throw new Error("Unable to perform operation. To use this extension, queries must be saved in My Queries or Shared Queries.");
    }
    const result = await client.queryById(config.queryId, project);
    const ids = idsFromQueryResult(result);
    const columns = mergeColumns(result.columns, undefined);
    return exportWorkbook({
      project,
      orgUrl,
      action: "OpenQuery",
      ids,
      columns,
      fileName: config.queryName || "query",
      queryId: config.queryId,
      queryName: config.queryName,
      queryResult: result
    });
  }

  const ids = config.ids || [];
  if (ids.length === 0) {
    throw new Error("Select one or more work items first.");
  }
  const columns = mergeColumns(DEFAULT_COLUMNS, config.columns);
  return exportWorkbook({
    project,
    orgUrl,
    action: "OpenItems",
    ids,
    columns,
    fileName: ids.length === 1 ? `work-item-${ids[0]}` : "work-items"
  });
}
