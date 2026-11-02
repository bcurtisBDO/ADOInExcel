const SETTINGS_KEY = "ado-open-in-excel-settings";

export type AuthKind = "pat" | "entra";

export interface ConnectionSettings {
  orgUrl: string;
  authKind: AuthKind;
  pat?: string;
  entraClientId?: string;
  entraUsername?: string;
}

async function storage(): Promise<OfficeRuntime.Storage | Storage> {
  if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
    return OfficeRuntime.storage;
  }
  return window.localStorage;
}

export async function loadSettings(): Promise<ConnectionSettings | null> {
  const store = await storage();
  const raw = await Promise.resolve(store.getItem(SETTINGS_KEY));
  if (!raw) {
    return null;
  }
  try {
    const parsed = JSON.parse(raw) as ConnectionSettings;
    if (!parsed.orgUrl) {
      return null;
    }
    if (!parsed.authKind) {
      parsed.authKind = parsed.pat ? "pat" : "entra";
    }
    return parsed;
  } catch {
    return null;
  }
}

export async function saveSettings(settings: ConnectionSettings): Promise<void> {
  const store = await storage();
  await Promise.resolve(store.setItem(SETTINGS_KEY, JSON.stringify(settings)));
}

export async function clearSettings(): Promise<void> {
  const store = await storage();
  await Promise.resolve(store.removeItem(SETTINGS_KEY));
}

export function isGuid(value: string): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value.trim());
}
