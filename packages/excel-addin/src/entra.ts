import { AccountInfo, PublicClientApplication } from "@azure/msal-browser";
import { ADO_SCOPES, ENTRA_AUTHORITY, authRedirectUri } from "./entraConfig";

const pcaByClient = new Map<string, PublicClientApplication>();

async function getPca(clientId: string): Promise<PublicClientApplication> {
  const existing = pcaByClient.get(clientId);
  if (existing) {
    return existing;
  }
  const pca = new PublicClientApplication({
    auth: {
      clientId,
      authority: ENTRA_AUTHORITY,
      redirectUri: authRedirectUri()
    },
    cache: {
      cacheLocation: "localStorage"
    }
  });
  await pca.initialize();
  pcaByClient.set(clientId, pca);
  return pca;
}

export async function signInWithEntra(clientId: string): Promise<{ username: string }> {
  const redirectUri = authRedirectUri();
  const url = `${redirectUri}?clientId=${encodeURIComponent(clientId)}`;

  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      url,
      { height: 65, width: 35, promptBeforeOpen: false },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(result.error?.message || "Could not open the Microsoft sign-in dialog."));
          return;
        }
        const dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          dialog.close();
          try {
            const payload = JSON.parse((arg as { message: string }).message) as {
              ok?: boolean;
              username?: string;
              error?: string;
            };
            if (payload.ok) {
              resolve({ username: payload.username || "Signed in" });
            } else {
              reject(new Error(payload.error || "Microsoft sign-in failed."));
            }
          } catch {
            reject(new Error("Microsoft sign-in returned an unexpected response."));
          }
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          reject(new Error("Microsoft sign-in was cancelled."));
        });
      }
    );
  });
}

export async function acquireAdoToken(clientId: string): Promise<string> {
  const pca = await getPca(clientId);
  const account = currentAccount(pca);
  if (!account) {
    throw new Error("Sign in with Microsoft first.");
  }
  try {
    const silent = await pca.acquireTokenSilent({
      account,
      scopes: ADO_SCOPES
    });
    return silent.accessToken;
  } catch {
    throw new Error("Your Microsoft session expired. Sign in with Microsoft again.");
  }
}

export async function currentEntraUsername(clientId: string): Promise<string | undefined> {
  try {
    const pca = await getPca(clientId);
    return currentAccount(pca)?.username;
  } catch {
    return undefined;
  }
}

export async function signOutEntra(clientId: string): Promise<void> {
  const pca = await getPca(clientId);
  const account = currentAccount(pca);
  if (account) {
    await pca.clearCache();
  }
}

function currentAccount(pca: PublicClientApplication): AccountInfo | undefined {
  const accounts = pca.getAllAccounts();
  return accounts[0];
}
