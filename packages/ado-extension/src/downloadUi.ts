import * as SDK from "azure-devops-extension-sdk";
import { runDownload } from "./exportWorkbook";
import { DownloadDialogConfig } from "./host";
import { triggerDownload } from "./workbook";

function ensureUi(): { status: HTMLElement; actions: HTMLElement } {
  let status = document.getElementById("status");
  let actions = document.getElementById("actions");
  if (!status || !actions) {
    document.body.innerHTML = "";
    status = document.createElement("p");
    status.id = "status";
    actions = document.createElement("p");
    actions.id = "actions";
    document.body.appendChild(status);
    document.body.appendChild(actions);
  }
  return { status, actions };
}

function showError(message: string): void {
  const { status } = ensureUi();
  status.className = "error";
  status.textContent = message;
}

function closeDialog(result = false): void {
  const configuration = SDK.getConfiguration() as {
    dialog?: { close: (value?: boolean) => void };
  };
  try {
    configuration.dialog?.close(result);
  } catch {
    // The host chrome can still close the dialog.
  }
}

function renderActions(buttons: Array<{ label: string; primary?: boolean; onClick: () => void }>): void {
  const { actions } = ensureUi();
  actions.innerHTML = "";
  for (const item of buttons) {
    const button = document.createElement("button");
    button.type = "button";
    button.textContent = item.label;
    if (!item.primary) {
      button.style.background = "#605e5c";
      button.style.marginLeft = "8px";
    }
    button.addEventListener("click", item.onClick);
    actions.appendChild(button);
  }
}

export function resolveDownloadConfig(raw: Record<string, unknown> | undefined): DownloadDialogConfig {
  if (!raw) {
    return {} as DownloadDialogConfig;
  }
  if (raw.kind === "query" || raw.kind === "items") {
    return raw as unknown as DownloadDialogConfig;
  }
  const nested = raw.configuration;
  if (nested && typeof nested === "object") {
    const config = nested as Record<string, unknown>;
    if (config.kind === "query" || config.kind === "items") {
      return config as unknown as DownloadDialogConfig;
    }
  }
  return {} as DownloadDialogConfig;
}

export async function runDownloadDialog(config: DownloadDialogConfig): Promise<void> {
  const { status } = ensureUi();
  status.className = "";
  status.textContent = "Preparing workbook…";
  renderActions([{ label: "Close", onClick: () => closeDialog(false) }]);
  if (!config.kind) {
    const message = "The download dialog did not receive a query or work item selection.";
    showError(message);
    throw new Error(message);
  }

  try {
    status.textContent = "Loading work items and building the workbook…";
    const prepared = await runDownload(config);
    status.textContent =
      `Ready: ${prepared.count} work item(s) in ${prepared.fileName}. Azure DevOps cannot start Excel from the browser — click Download, then open the file. If Excel shows Protected View, click Enable Editing before Publish.`;
    renderActions([
      {
        label: "Download workbook",
        primary: true,
        onClick: () => triggerDownload(prepared.blob, prepared.fileName)
      },
      { label: "Close", onClick: () => closeDialog(true) }
    ]);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    showError(message);
    renderActions([{ label: "Close", onClick: () => closeDialog(false) }]);
    throw error instanceof Error ? error : new Error(message);
  }
}
