import { PublicClientApplication } from "@azure/msal-browser";
import { ADO_SCOPES, ENTRA_AUTHORITY, authRedirectUri } from "../entraConfig";

Office.onReady(async () => {
  const params = new URLSearchParams(window.location.search);
  const clientId = params.get("clientId") || "";
  if (!clientId) {
    send({ ok: false, error: "Missing Entra application (client) ID." });
    return;
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

  try {
    const redirected = await pca.handleRedirectPromise();
    if (redirected) {
      send({ ok: true, username: redirected.account?.username });
      return;
    }
    await pca.loginRedirect({
      scopes: ADO_SCOPES,
      prompt: "select_account"
    });
  } catch (error) {
    send({
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    });
  }
});

function send(payload: { ok: boolean; username?: string; error?: string }): void {
  try {
    Office.context.ui.messageParent(JSON.stringify(payload));
  } catch {
    document.body.innerHTML = `<p>${payload.error || "Signed in. You can close this window."}</p>`;
  }
}
