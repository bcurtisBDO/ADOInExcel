import { isSupportedQueryId } from "@ado-open-in-excel/shared";
import {
  DownloadDialogConfig,
  IActionContext,
  idsFromContext,
  notify,
  openDownloadDialog,
  queryFromContext
} from "./host";

async function startDownload(config: DownloadDialogConfig): Promise<void> {
  await openDownloadDialog(config);
}

export type { IActionContext };

export const openQueryAction = {
  execute: async (actionContext: IActionContext) => {
    try {
      const query = queryFromContext(actionContext);
      if (!query?.id || !isSupportedQueryId(query.id)) {
        throw new Error("Unable to perform operation. To use this extension, queries must be saved in My Queries or Shared Queries.");
      }
      await startDownload({
        kind: "query",
        queryId: query.id,
        queryName: query.name
      });
    } catch (error) {
      await notify("ADO in Excel", error instanceof Error ? error.message : String(error));
    }
  }
};

export const openQueryOnToolbarAction = {
  execute: async (actionContext: IActionContext) => {
    await openQueryAction.execute(actionContext);
  }
};

export const openWorkItemsAction = {
  execute: async (actionContext: IActionContext) => {
    try {
      const ids = idsFromContext(actionContext);
      if (!ids || ids.length === 0) {
        throw new Error("Select one or more work items first.");
      }
      await startDownload({
        kind: "items",
        ids,
        columns: actionContext.columns
      });
    } catch (error) {
      await notify("ADO in Excel", error instanceof Error ? error.message : String(error));
    }
  }
};
