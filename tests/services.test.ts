import { describe, expect, it } from 'vitest';
import { dispatch } from '../src/dispatcher.js';
import { getDb, getRow, seedServiceStats } from '../src/db.js';
import { registry } from '../src/providers/registry.js';
import { FakeProvider } from './helpers/fake-provider.js';
import { app, jsonHeaders, jsonOf } from './helpers/http.js';

type StatsRow = {
  name: string;
  inbox_count: number;
  success_count: number;
  fail_count: number;
  first_used_at: string | null;
  last_used_at: string | null;
};

type ServicesResponse = {
  summary: { totalServices: number; totalInboxes: number; totalFailures: number; totalBlocks: number };
  services: {
    name: string; totalInboxes: number; activeInboxes: number;
    successCount: number; failCount: number; blockCount: number;
    lastUsed: string | null;
  }[];
};

function statsRow(name: string): StatsRow | undefined {
  return getRow<StatsRow>(getDb(), `SELECT * FROM service_stats WHERE name = ?`, name);
}

function insertInbox(id: string, service: string | null, status = 'active') {
  getDb().prepare(
    `INSERT INTO inboxes (id, provider, address, auth_data, target_service, status) VALUES (?, 'fake', ?, '{}', ?, ?)`
  ).run(id, `${id}@example.test`, service, status);
}

async function report(inboxId: string, success: boolean, service?: string) {
  return app.request(`/api/inbox/${inboxId}/report`, {
    method: 'POST',
    headers: jsonHeaders(),
    body: JSON.stringify({ success, ...(service ? { service } : {}) }),
  });
}

describe('durable service stats', () => {
  it('dispatch bumps inbox_count and stores a trimmed target_service', async () => {
    registry.register(new FakeProvider({ rateLimit: { createPerMinute: 10, pollPerMinute: 10 } }));
    await dispatch({ for: '  svc-a.com  ', provider: 'fake' });

    const row = statsRow('svc-a.com');
    expect(row).toBeDefined();
    expect(row?.inbox_count).toBe(1);
    expect(row?.first_used_at).toBeTruthy();
    expect(row?.last_used_at).toBeTruthy();

    const inbox = getRow<{ target_service: string }>(getDb(), `SELECT target_service FROM inboxes WHERE target_service = 'svc-a.com'`);
    expect(inbox?.target_service).toBe('svc-a.com');

    await dispatch({ for: 'svc-a.com', provider: 'fake' });
    expect(statsRow('svc-a.com')?.inbox_count).toBe(2);
  });

  it('report success/failure bumps durable counters', async () => {
    insertInbox('rep-1', 'svc-b.com');

    await report('rep-1', true);
    expect(statsRow('svc-b.com')?.success_count).toBe(1);
    expect(statsRow('svc-b.com')?.fail_count).toBe(0);

    await report('rep-1', false);
    const row = statsRow('svc-b.com');
    expect(row?.success_count).toBe(1);
    expect(row?.fail_count).toBe(1);
  });

  it('report with explicit service creates a stats row even without prior tracking', async () => {
    insertInbox('rep-2', null);
    await report('rep-2', true, 'svc-override.com');
    expect(statsRow('svc-override.com')?.success_count).toBe(1);
  });

  it('report tolerates a non-string service instead of 500ing (regression)', async () => {
    insertInbox('rep-nonstring', 'svc-typed.com');
    const res = await app.request('/api/inbox/rep-nonstring/report', {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({ success: true, service: 12345 }),
    });
    expect(res.status).toBe(200);
    // Falls back to the inbox's own target_service, and provider_stats still records.
    expect(statsRow('svc-typed.com')?.success_count).toBe(1);
    const ps = getRow<{ success_count: number }>(getDb(), `SELECT success_count FROM provider_stats WHERE provider = 'fake'`);
    expect(ps?.success_count).toBe(1);
  });

  it('report trims service so fail_log/blocks share one canonical name (regression)', async () => {
    const db = getDb();
    db.prepare(`INSERT INTO block_rules (service, provider, threshold, window_hours, scope, domain_level, enabled) VALUES ('*', '*', 1, 24, 'per_service', 2, 1)`).run();
    insertInbox('rep-pad', null);

    const res = await app.request('/api/inbox/rep-pad/report', {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({ success: false, service: '  padded.com  ' }),
    });
    expect(res.status).toBe(200);

    // All four keyed off the trimmed name — a padded blocks row would be
    // invisible to the dispatcher's `WHERE service = ?` lookup.
    expect(statsRow('padded.com')?.fail_count).toBe(1);
    expect(getRow<{ service: string }>(db, `SELECT service FROM fail_log`)?.service).toBe('padded.com');
    expect(getRow<{ service: string }>(db, `SELECT service FROM blocks`)?.service).toBe('padded.com');
  });

  it('service stays listed after its inboxes and fail_log are purged (regression)', async () => {
    registry.register(new FakeProvider());
    await dispatch({ for: 'svc-c.com', provider: 'fake' });
    const inbox = getRow<{ id: string }>(getDb(), `SELECT id FROM inboxes WHERE target_service = 'svc-c.com'`);
    await report(inbox!.id, false);

    // Simulate the retention purge wiping the raw history.
    getDb().prepare(`DELETE FROM inboxes`).run();
    getDb().prepare(`DELETE FROM fail_log`).run();

    const res = await app.request('/api/services', { headers: jsonHeaders() });
    const data = await jsonOf<ServicesResponse>(res);
    const svc = data.services.find((s) => s.name === 'svc-c.com');
    expect(svc).toBeDefined();
    expect(svc?.totalInboxes).toBe(1);
    expect(svc?.failCount).toBe(1);
    expect(svc?.activeInboxes).toBe(0);

    const detailRes = await app.request(`/api/services/${encodeURIComponent('svc-c.com')}`, { headers: jsonHeaders() });
    const detail = await jsonOf<{ stats: { totalInboxes: number; failCount: number } | null; inboxes: unknown[] }>(detailRes);
    expect(detail.stats?.totalInboxes).toBe(1);
    expect(detail.stats?.failCount).toBe(1);
    expect(detail.inboxes).toHaveLength(0);
  });

  it('GET /api/services merges live active counts with durable totals', async () => {
    registry.register(new FakeProvider());
    await dispatch({ for: 'svc-d.com', provider: 'fake' });

    const res = await app.request('/api/services', { headers: jsonHeaders() });
    const data = await jsonOf<ServicesResponse>(res);
    const svc = data.services.find((s) => s.name === 'svc-d.com');
    expect(svc?.totalInboxes).toBe(1);
    expect(svc?.activeInboxes).toBe(1);
    expect(data.summary.totalServices).toBeGreaterThanOrEqual(1);
  });

  it('seedServiceStats backfills from retained history and is idempotent', () => {
    const db = getDb();
    insertInbox('seed-1', 'seeded.com', 'closed');
    insertInbox('seed-2', 'seeded.com', 'active');
    db.prepare(`INSERT INTO fail_log (service, provider, domain) VALUES ('seeded.com', 'fake', 'example.test')`).run();
    db.prepare(`INSERT INTO fail_log (service, provider, domain) VALUES ('faillog-only.com', 'fake', 'example.test')`).run();

    seedServiceStats(db);
    expect(statsRow('seeded.com')?.inbox_count).toBe(2);
    expect(statsRow('seeded.com')?.fail_count).toBe(1);
    expect(statsRow('faillog-only.com')?.fail_count).toBe(1);

    // Re-running (every boot) must not double-count.
    seedServiceStats(db);
    expect(statsRow('seeded.com')?.inbox_count).toBe(2);
    expect(statsRow('seeded.com')?.fail_count).toBe(1);
    expect(statsRow('faillog-only.com')?.fail_count).toBe(1);
  });
});
