export type CloudSyncEntity = 'users' | 'clients' | 'orders' | 'documents';

type ViteImportMeta = ImportMeta & {
  env?: Record<string, unknown>;
};

type CloudSyncRow<T> = {
  id: string;
  payload: T;
  updated_at?: string;
};

const tableByEntity: Record<CloudSyncEntity, string> = {
  users: 'crm_users',
  clients: 'crm_clients',
  orders: 'crm_orders',
  documents: 'crm_documents',
};

function envValue(key: string) {
  return String(((import.meta as ViteImportMeta).env ?? {})[key] ?? '').trim();
}

export function cloudSyncConfig() {
  const url = envValue('VITE_SUPABASE_URL').replace(/\/$/, '');
  const anonKey = envValue('VITE_SUPABASE_ANON_KEY');
  const enabled = envValue('VITE_CRM_CLOUD_SYNC') === 'true' && Boolean(url && anonKey);
  const intervalMs = Math.max(Number(envValue('VITE_CRM_SYNC_INTERVAL_MS')) || 5000, 2000);
  return { enabled, url, anonKey, intervalMs };
}

function supabaseHeaders(anonKey: string) {
  return {
    apikey: anonKey,
    Authorization: `Bearer ${anonKey}`,
    'Content-Type': 'application/json',
  };
}

export async function fetchCloudEntity<T>(entity: CloudSyncEntity): Promise<T[]> {
  const config = cloudSyncConfig();
  if (!config.enabled) return [];
  const table = tableByEntity[entity];
  const response = await fetch(`${config.url}/rest/v1/${table}?select=id,payload,updated_at&order=updated_at.desc`, {
    headers: supabaseHeaders(config.anonKey),
  });
  if (!response.ok) throw new Error(`Cloud sync read failed: ${entity} ${response.status}`);
  const rows = await response.json() as Array<CloudSyncRow<T>>;
  return rows.map((row) => row.payload).filter(Boolean);
}

export async function pushCloudEntity<T>(entity: CloudSyncEntity, items: T[], getId: (item: T) => string) {
  const config = cloudSyncConfig();
  if (!config.enabled) return;
  const table = tableByEntity[entity];
  const updatedAt = new Date().toISOString();
  const rows = items
    .map((item) => ({ id: getId(item), payload: item, updated_at: updatedAt }))
    .filter((row) => row.id);
  if (rows.length === 0) return;
  const response = await fetch(`${config.url}/rest/v1/${table}`, {
    method: 'POST',
    headers: {
      ...supabaseHeaders(config.anonKey),
      Prefer: 'resolution=merge-duplicates',
    },
    body: JSON.stringify(rows),
  });
  if (!response.ok) throw new Error(`Cloud sync write failed: ${entity} ${response.status}`);
}

export function stableCloudId(value: string) {
  const source = value.trim().toLowerCase();
  let hash = 0;
  for (let index = 0; index < source.length; index += 1) {
    hash = ((hash << 5) - hash) + source.charCodeAt(index);
    hash |= 0;
  }
  const slug = source
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
  return `${slug || 'record'}-${Math.abs(hash).toString(36)}`;
}
