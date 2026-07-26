import { describe, expect, it } from 'vitest';
import { getDb, getRow } from '../src/db.js';
import { app, authHeaders, jsonHeaders, jsonOf } from './helpers/http.js';

type CreatedKey = { key: string; keyHash: string };
type KeyRow = { keyHash: string; name: string; active: boolean; dailyLimit: number | null };

async function createKey(name = 'quota-key'): Promise<CreatedKey> {
  const res = await app.request('/api/keys', {
    method: 'POST',
    headers: jsonHeaders(),
    body: JSON.stringify({ name }),
  });
  expect(res.status).toBe(201);
  return jsonOf<CreatedKey>(res);
}

async function patchKey(keyHash: string, body: Record<string, unknown>) {
  return app.request(`/api/keys/${encodeURIComponent(keyHash)}`, {
    method: 'PATCH',
    headers: jsonHeaders(),
    body: JSON.stringify(body),
  });
}

async function listKeys(): Promise<KeyRow[]> {
  const res = await app.request('/api/keys', { headers: authHeaders() });
  return (await jsonOf<{ keys: KeyRow[] }>(res)).keys;
}

describe('PATCH /api/keys/:key', () => {
  it('persists dailyLimit instead of 500ing on the camelCase field (regression)', async () => {
    const { keyHash } = await createKey();

    const res = await patchKey(keyHash, { dailyLimit: 100 });
    expect(res.status).toBe(200);

    const stored = getRow<{ daily_limit: number | null }>(getDb(), `SELECT daily_limit FROM api_keys WHERE key = ?`, keyHash);
    expect(stored?.daily_limit).toBe(100);
    expect((await listKeys()).find((k) => k.keyHash === keyHash)?.dailyLimit).toBe(100);
  });

  it('clears dailyLimit when set to null', async () => {
    const { keyHash } = await createKey('clear-key');
    await patchKey(keyHash, { dailyLimit: 50 });

    const res = await patchKey(keyHash, { dailyLimit: null });
    expect(res.status).toBe(200);
    const stored = getRow<{ daily_limit: number | null }>(getDb(), `SELECT daily_limit FROM api_keys WHERE key = ?`, keyHash);
    expect(stored?.daily_limit).toBeNull();
  });

  it('still updates name and active', async () => {
    const { keyHash } = await createKey('old-name');

    expect((await patchKey(keyHash, { name: '  new-name  ', active: false })).status).toBe(200);
    const row = (await listKeys()).find((k) => k.keyHash === keyHash);
    expect(row?.name).toBe('new-name');
    expect(row?.active).toBe(false);
  });

  it('enforces a persisted dailyLimit on the auth path', async () => {
    const { key, keyHash } = await createKey('enforced-key');
    await patchKey(keyHash, { dailyLimit: 1 });

    const first = await app.request('/api/providers', { headers: authHeaders(key) });
    expect(first.status).toBe(200);
    const second = await app.request('/api/providers', { headers: authHeaders(key) });
    expect(second.status).toBe(429);
  });
});
