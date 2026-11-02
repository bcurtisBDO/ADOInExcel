import { openSelectedWorkItem, publish, refresh } from "../operations";

async function runCommand(
  event: Office.AddinCommands.Event,
  work: () => Promise<{ message: string }>
): Promise<void> {
  try {
    const result = await work();
    console.log(result.message);
  } catch (error) {
    console.error(error);
  } finally {
    event.completed();
  }
}

function refreshWorkItems(event: Office.AddinCommands.Event): void {
  void runCommand(event, refresh);
}

function publishWorkItems(event: Office.AddinCommands.Event): void {
  void runCommand(event, publish);
}

function openWorkItemInAdo(event: Office.AddinCommands.Event): void {
  void runCommand(event, openSelectedWorkItem);
}

Office.onReady(() => {
  Office.actions.associate("refreshWorkItems", refreshWorkItems);
  Office.actions.associate("publishWorkItems", publishWorkItems);
  Office.actions.associate("openWorkItemInAdo", openWorkItemInAdo);
});
