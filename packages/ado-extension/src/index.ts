import * as SDK from "azure-devops-extension-sdk";
import * as App from "./app";

SDK.register("openQueryOnToolbarAction", () => {
  return {
    execute: async (context: unknown) => {
      await App.openQueryOnToolbarAction.execute(context as App.IActionContext);
    }
  };
});

SDK.register("openQueryAction", () => {
  return {
    execute: async (context: unknown) => {
      await App.openQueryAction.execute(context as App.IActionContext);
    }
  };
});

SDK.register("openWorkItemsAction", () => {
  return {
    execute: async (context: unknown) => {
      await App.openWorkItemsAction.execute(context as App.IActionContext);
    }
  };
});

SDK.init();
