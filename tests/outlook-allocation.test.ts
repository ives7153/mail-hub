import { describe, expect, it } from 'vitest';
import { dispatch } from '../src/dispatcher.js';
import { allRows, getDb, getRow } from '../src/db.js';
import { OutlookProvider } from '../src/providers/outlook.js';
import { app, authHeaders, jsonHeaders } from './helpers/http.js';

describe('Outlook account allocation', () => {
  it('reserves the selected account with the final inbox id during creation', async () => {
    const db = getDb();
    db.prepare(
      `INSERT INTO outlook_accounts (email, password, client_id, refresh_token, token_status)
       VALUES (?, ?, ?, ?, ?)`,
    ).run('reserve@outlook.test', 'pw', 'client', 'refresh', 'valid');

    const result = await dispatch({ provider: 'outlook', ownerKey: 'owner-a' });

    const assigned = getRow<{ assigned_inbox_id: string }>(
      db,
      `SELECT assigned_inbox_id FROM outlook_accounts WHERE email = ?`,
      'reserve@outlook.test',
    );
    expect(assigned?.assigned_inbox_id).toBe(result.id);

    const inbox = getRow<{ id: string; owner_key: string }>(db, `SELECT id, owner_key FROM inboxes WHERE id = ?`, result.id);
    expect(inbox).toEqual({ id: result.id, owner_key: 'owner-a' });
  });

  it('does not hand out an already reserved Outlook account', async () => {
    const db = getDb();
    db.prepare(
      `INSERT INTO outlook_accounts (email, password, client_id, refresh_token, token_status, assigned_inbox_id)
       VALUES (?, ?, ?, ?, ?, ?)`,
    ).run('busy@outlook.test', 'pw', 'client', 'refresh', 'valid', 'existing-inbox');

    await expect(dispatch({ provider: 'outlook', domain: 'outlook.test' })).rejects.toThrow(/无可用账号/);
  });

  it('does not hand out a pending OAuth account without token credentials', async () => {
    const db = getDb();
    db.prepare(
      `INSERT INTO outlook_accounts (email, password, token_status)
       VALUES (?, ?, ?)`,
    ).run('pending@outlook.test', 'pw', 'pending_oauth');

    const provider = new OutlookProvider();

    await expect(provider.getDomains()).resolves.toEqual([]);
    await expect(dispatch({ provider: 'outlook', domain: 'outlook.test' })).rejects.toThrow(/无可用账号/);
    const assigned = getRow<{ assigned_inbox_id: string | null }>(
      db,
      `SELECT assigned_inbox_id FROM outlook_accounts WHERE email = ?`,
      'pending@outlook.test',
    );
    expect(assigned?.assigned_inbox_id).toBeNull();
  });

  it('filters available Outlook domains by target service reuse', async () => {
    const db = getDb();
    db.prepare(
      `INSERT INTO outlook_accounts (email, password, client_id, refresh_token, token_status, used_services)
       VALUES (?, ?, ?, ?, ?, ?)`,
    ).run('used@one.test', 'pw', 'client', 'refresh', 'valid', JSON.stringify(['service.test']));
    db.prepare(
      `INSERT INTO outlook_accounts (email, password, client_id, refresh_token, token_status)
       VALUES (?, ?, ?, ?, ?)`,
    ).run('free@two.test', 'pw', 'client', 'refresh', 'valid');

    const provider = new OutlookProvider();

    await expect(provider.getDomains({ for: 'service.test' })).resolves.toEqual(['two.test']);
    await expect(provider.getDomains()).resolves.toEqual(['one.test', 'two.test']);
  });

  it('allocates distinct accounts for concurrent Outlook requests', async () => {
    const db = getDb();
    for (const email of ['a@outlook.test', 'b@outlook.test']) {
      db.prepare(
        `INSERT INTO outlook_accounts (email, password, client_id, refresh_token, token_status)
         VALUES (?, ?, ?, ?, ?)`,
      ).run(email, 'pw', 'client', 'refresh', 'valid');
    }

    const results = await Promise.all([
      dispatch({ provider: 'outlook', domain: 'outlook.test' }),
      dispatch({ provider: 'outlook', domain: 'outlook.test' }),
    ]);
    const addresses = results.map((result) => result.address).sort();
    const assigned = allRows<{ email: string; assigned_inbox_id: string }>(
      db,
      `SELECT email, assigned_inbox_id FROM outlook_accounts ORDER BY email`,
    );

    expect(addresses).toEqual(['a@outlook.test', 'b@outlook.test']);
    expect(new Set(assigned.map((row) => row.assigned_inbox_id)).size).toBe(2);
    expect(assigned.every((row) => row.assigned_inbox_id)).toBe(true);
  });

  it('does not alternate failures when one domain is exhausted for a target service', async () => {
    const db = getDb();
    db.prepare(
      `INSERT INTO outlook_accounts (email, password, client_id, refresh_token, token_status, used_services)
       VALUES (?, ?, ?, ?, ?, ?)`,
    ).run('used@blocked.test', 'pw', 'client', 'refresh', 'valid', JSON.stringify(['netlify.com']));

    for (const email of ['one@open.test', 'two@open.test', 'three@open.test']) {
      db.prepare(
        `INSERT INTO outlook_accounts (email, password, client_id, refresh_token, token_status)
         VALUES (?, ?, ?, ?, ?)`,
      ).run(email, 'pw', 'client', 'refresh', 'valid');
    }

    const results = [];
    for (let i = 0; i < 3; i++) {
      results.push(await dispatch({ provider: 'outlook', for: 'netlify.com' }));
    }

    expect(results.map((result) => result.address).sort()).toEqual([
      'one@open.test',
      'three@open.test',
      'two@open.test',
    ]);
  });

  it('does not record the service on a failure report by default, and the settings endpoint agrees (regression)', async () => {
    const db = getDb();
    db.prepare(
      `INSERT INTO outlook_accounts (email, password, client_id, refresh_token, token_status)
       VALUES (?, ?, ?, ?, ?)`,
    ).run('unset@outlook.test', 'pw', 'client', 'refresh', 'valid');
    const created = await dispatch({ provider: 'outlook', for: 'svc-unset.com' });

    // No settings row exists. used_services is an irreversible blacklist, so a
    // failure must not burn the account for that service unless the operator
    // opted in — and the settings endpoint must report the same default, not
    // the opposite of what the report path does.
    const res = await app.request(`/api/inbox/${created.id}/report`, {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({ success: false }),
    });
    expect(res.status).toBe(200);

    const settingsRes = await app.request('/api/outlook/settings', { headers: authHeaders() });
    const settings = await settingsRes.json() as { recordFailService: boolean };
    expect(settings.recordFailService).toBe(false);

    const account = getRow<{ used_services: string }>(
      db,
      `SELECT used_services FROM outlook_accounts WHERE email = ?`,
      'unset@outlook.test',
    );
    expect(JSON.parse(account!.used_services)).toEqual([]);
  });

  it('records the service on a failure report once the operator opts in', async () => {
    const db = getDb();
    db.prepare(
      `INSERT INTO outlook_accounts (email, password, client_id, refresh_token, token_status)
       VALUES (?, ?, ?, ?, ?)`,
    ).run('optin@outlook.test', 'pw', 'client', 'refresh', 'valid');

    const patched = await app.request('/api/outlook/settings', {
      method: 'PATCH',
      headers: jsonHeaders(),
      body: JSON.stringify({ recordFailService: true }),
    });
    expect(patched.status).toBe(200);

    const created = await dispatch({ provider: 'outlook', for: 'svc-optin.com' });
    await app.request(`/api/inbox/${created.id}/report`, {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({ success: false }),
    });

    const settingsRes = await app.request('/api/outlook/settings', { headers: authHeaders() });
    expect(((await settingsRes.json()) as { recordFailService: boolean }).recordFailService).toBe(true);

    const account = getRow<{ used_services: string }>(
      db,
      `SELECT used_services FROM outlook_accounts WHERE email = ?`,
      'optin@outlook.test',
    );
    expect(JSON.parse(account!.used_services)).toEqual(['svc-optin.com']);
  });
});
