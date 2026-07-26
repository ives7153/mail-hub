import { describe, expect, it } from 'vitest';
import { getDb } from '../src/db.js';
import { app, authHeaders, jsonOf } from './helpers/http.js';

type ListResponse = {
  inboxes: { id: string }[];
  page: number;
  pageSize: number;
  total: number;
};

function seedInboxes(count: number, status = 'active') {
  const db = getDb();
  const ins = db.prepare(
    `INSERT INTO inboxes (id, provider, address, auth_data, target_service, status, created_at)
     VALUES (?, 'mailtm', ?, '{}', 'paging.test', ?, datetime('now', ?))`,
  );
  db.transaction(() => {
    for (let i = 0; i < count; i++) {
      ins.run(`page-${String(i).padStart(3, '0')}`, `p${i}@example.test`, status, `-${i} minutes`);
    }
  })();
}

async function list(query = ''): Promise<ListResponse> {
  const res = await app.request(`/api/inboxes${query}`, { headers: authHeaders() });
  expect(res.status).toBe(200);
  return jsonOf<ListResponse>(res);
}

describe('GET /api/inboxes pagination contract', () => {
  it('reports the unfiltered total even when the page is truncated', async () => {
    seedInboxes(137);

    const first = await list();
    // The frontend walks pages using `total`; if this ever reported the page
    // length instead, the dashboard would silently show a partial list again.
    expect(first.inboxes).toHaveLength(50);
    expect(first.total).toBe(137);
    expect(first.pageSize).toBe(50);
    expect(first.page).toBe(1);
  });

  it('honors pageSize up to the 100 cap and walks pages without gaps or repeats', async () => {
    seedInboxes(137);

    const p1 = await list('?page=1&pageSize=100');
    expect(p1.inboxes).toHaveLength(100);
    expect(p1.pageSize).toBe(100);

    const p2 = await list('?page=2&pageSize=100');
    expect(p2.inboxes).toHaveLength(37);

    const ids = new Set([...p1.inboxes, ...p2.inboxes].map((i) => i.id));
    expect(ids.size).toBe(137);
  });

  it('caps pageSize at 100 so a walk always terminates', async () => {
    seedInboxes(137);

    const res = await list('?page=1&pageSize=1000');
    expect(res.pageSize).toBe(100);
    expect(res.inboxes).toHaveLength(100);
  });

  it('keeps total consistent with the active filter', async () => {
    seedInboxes(30, 'active');
    getDb().prepare(`UPDATE inboxes SET status = 'closed' WHERE id LIKE 'page-00%'`).run();

    const active = await list('?status=active');
    const closed = await list('?status=closed');
    expect(active.total + closed.total).toBe(30);
    expect(closed.total).toBe(10);
  });
});
