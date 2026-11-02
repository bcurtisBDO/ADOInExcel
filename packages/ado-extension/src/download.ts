import * as SDK from "azure-devops-extension-sdk";
import { resolveDownloadConfig, runDownloadDialog } from "./downloadUi";

async function boot(): Promise<void> {
  await SDK.init({ loaded: false });
  await SDK.ready();
  try {
    const config = resolveDownloadConfig(SDK.getConfiguration());
    await runDownloadDialog(config);
    SDK.notifyLoadSucceeded();
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    try {
      await SDK.notifyLoadFailed(message);
    } catch {
      document.body.textContent = message;
    }
  }
}

void boot();
