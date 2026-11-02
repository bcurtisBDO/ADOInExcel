import {
  CommonServiceIds,
  IGlobalMessagesService,
  IHostPageLayoutService,
  ILocationService,
  IProjectPageService
} from "azure-devops-extension-api";
import * as SDK from "azure-devops-extension-sdk";
import { normalizeOrgUrl } from "@ado-open-in-excel/shared";

export interface IQueryObject {
  id: string;
  isPublic: boolean;
  name: string;
  path: string;
  wiql: string;
}

export interface IActionContext {
  id?: number;
  workItemId?: number;
  query?: IQueryObject;
  queryId?: string;
  queryText?: string;
  ids?: number[];
  workItemIds?: number[];
  columns?: string[];
}

export interface DownloadDialogConfig {
  kind: "query" | "items";
  queryId?: string;
  queryName?: string;
  ids?: number[];
  columns?: string[];
}

export async function getProjectContext(): Promise<{ name: string; orgUrl: string }> {
  const projectService = await SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService);
  const project = await projectService.getProject();
  if (!project) {
    throw new Error("Unable to perform operation. Not a valid project context.");
  }

  const locationService = await SDK.getService<ILocationService>(CommonServiceIds.LocationService);
  const collectionUri = await locationService.getServiceLocation();
  return {
    name: project.name,
    orgUrl: normalizeOrgUrl(collectionUri)
  };
}

export async function notify(title: string, message: string): Promise<void> {
  try {
    const layout = await SDK.getService<IHostPageLayoutService>(CommonServiceIds.HostPageLayoutService);
    layout.openMessageDialog(message, {
      title,
      okText: "Close",
      showCancel: false
    });
  } catch {
    try {
      const messages = await SDK.getService<IGlobalMessagesService>(CommonServiceIds.GlobalMessagesService);
      messages.addToast({ duration: 5000, message: `${title}: ${message}` });
    } catch {
      alert(`${title}\n\n${message}`);
    }
  }
}

export async function toast(message: string): Promise<void> {
  try {
    const messages = await SDK.getService<IGlobalMessagesService>(CommonServiceIds.GlobalMessagesService);
    messages.addToast({ duration: 4000, message });
  } catch {
    // Action iframes can still proceed without a toast.
  }
}

export function queryFromContext(actionContext: IActionContext): IQueryObject | undefined {
  if (actionContext?.query?.id) {
    return actionContext.query;
  }
  if (actionContext?.queryId) {
    return {
      id: actionContext.queryId,
      isPublic: true,
      name: "query",
      path: "",
      wiql: actionContext.queryText || ""
    };
  }
  return undefined;
}

export function idsFromContext(actionContext: IActionContext): number[] | undefined {
  return (
    actionContext.ids ||
    actionContext.workItemIds ||
    (actionContext.workItemId && actionContext.workItemId > 0 ? [actionContext.workItemId] : undefined) ||
    (actionContext.id && actionContext.id > 0 ? [actionContext.id] : undefined)
  );
}

function downloadDialogId(): string {
  const extension = SDK.getExtensionContext();
  return `${extension.publisherId}.${extension.extensionId}.download-dialog`;
}

export async function openDownloadDialog(config: DownloadDialogConfig): Promise<void> {
  await toast("Opening ADO in Excel…");
  const layout = await SDK.getService<IHostPageLayoutService>(CommonServiceIds.HostPageLayoutService);
  const contributionId = downloadDialogId();
  await new Promise<void>((resolve, reject) => {
    try {
      layout.openCustomDialog(contributionId, {
        title: "ADO in Excel",
        configuration: config,
        lightDismiss: false,
        onClose: () => resolve()
      } as Parameters<IHostPageLayoutService["openCustomDialog"]>[1]);
    } catch (error) {
      reject(error instanceof Error ? error : new Error(String(error)));
    }
  });
}
