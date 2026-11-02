export const ADO_SCOPES = ["499b84ac-1321-427f-aa17-267ca6975798/.default"];
export const ENTRA_AUTHORITY = "https://login.microsoftonline.com/organizations";

/** Resolved next to the current page so hosting under a sub-path (e.g. GitHub Pages) works. */
export function authRedirectUri(): string {
  return new URL("auth.html", window.location.href).href;
}
