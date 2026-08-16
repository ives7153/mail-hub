import { Hono } from 'hono';
import { allRows, buildSetClause, getDb, getRow, logActivity } from '../db.js';
import { requireAdmin, type AdminEnv } from './admin.js';

export const accountsRoutes = new Hono<AdminEnv>();

accountsRoutes.use('/admin/accounts*', requireAdmin);

type MailboxCard = {
  email: string;
  source: string;
  remark: string;
  tags: string;
  createdAt: string;
  registrationCount: number;
};

type RegistrationRow = {
  id: number;
  email: string;
  app_name: string;
  username: string;
  password: string;
  memo: string;
  created_at: string;
  updated_at: string;
};

// Union of pool mailboxes + manual registry. Same address across sources is
// deduplicated with manual entries winning (rn=1 picks source 'manual' first).
// better-sqlite3 ships SQLite >=3.25, so window functions are available.
const MAILBOXES_SQL = `
  WITH all_emails AS (
    SELECT email, 'outlook' AS source, '' AS remark, '' AS tags, created_at FROM outlook_accounts
    UNION ALL
    SELECT LOWER(user), 'imap', '', '', created_at FROM imap_accounts
    UNION ALL
    SELECT hme, 'icloud', '', '', created_at FROM icloud_addresses
    UNION ALL
    SELECT email, 'manual', remark, tags, created_at FROM mailbox_registry
  ),
  ranked AS (
    SELECT *, ROW_NUMBER() OVER (
      PARTITION BY LOWER(email)
      ORDER BY CASE source WHEN 'manual' THEN 0 ELSE 1 END
    ) AS rn
    FROM all_emails
  )
  SELECT r.email, r.source, r.remark, r.tags, r.created_at AS createdAt,
    (SELECT COUNT(*) FROM mailbox_registrations x WHERE LOWER(x.email) = r.email) AS registrationCount
  FROM ranked r WHERE r.rn = 1
  ORDER BY r.email
`;

function normalizeEmail(value: unknown): string {
  return String(value ?? '').trim().toLowerCase();
}

// Confirms the email exists in the UNION (pool or manual) so registrations can
// only attach to a real mailbox.
function emailExists(db: ReturnType<typeof getDb>, email: string): boolean {
  const row = getRow<{ n: number }>(
    db,
    `SELECT 1 AS n FROM (${MAILBOXES_SQL}) WHERE LOWER(email) = ? LIMIT 1`,
    email,
  );
  return !!row;
}

accountsRoutes.get('/admin/accounts', (c) => {
  const db = getDb();
  const emails = allRows<MailboxCard>(db, MAILBOXES_SQL);
  return c.json({ emails });
});

accountsRoutes.post('/admin/accounts/emails', async (c) => {
  const body = await c.req.json().catch(() => ({}));
  const email = normalizeEmail(body.email);
  if (!email.includes('@') || !email.includes('.')) {
    return c.json({ error: 'Invalid email address' }, 400);
  }
  const remark = String(body.remark ?? '').trim();
  const tags = String(body.tags ?? '').trim();

  const db = getDb();
  const existing = getRow<{ email: string }>(db, `SELECT email FROM mailbox_registry WHERE email = ?`, email);
  if (existing) return c.json({ error: 'Email already exists' }, 409);

  db.prepare(`INSERT INTO mailbox_registry (email, remark, tags) VALUES (?, ?, ?)`).run(email, remark, tags);
  logActivity('green', `Added mailbox: ${email}`);
  return c.json({
    email,
    source: 'manual',
    remark,
    tags,
    createdAt: new Date().toISOString(),
    registrationCount: 0,
  }, 201);
});

accountsRoutes.patch('/admin/accounts/emails/:email', async (c) => {
  const email = normalizeEmail(c.req.param('email'));
  const body = await c.req.json().catch(() => ({}));
  const db = getDb();

  if (!emailExists(db, email)) return c.json({ error: 'Not found' }, 404);
  const row = getRow<{ source: string }>(db, `SELECT source FROM mailbox_registry WHERE email = ?`, email);
  if (!row || row.source !== 'manual') return c.json({ error: 'Pool emails cannot be edited' }, 400);

  const setClause = buildSetClause(body, {
    remark: (v) => String(v ?? '').trim(),
    tags: (v) => String(v ?? '').trim(),
  });
  if (!setClause) return c.json({ error: 'No fields to update' }, 400);

  db.prepare(`UPDATE mailbox_registry SET ${setClause.setClause}, updated_at = datetime('now') WHERE email = ?`)
    .run(...setClause.params, email);
  logActivity('blue', `Updated mailbox: ${email}`);
  return c.json({ ok: true });
});

accountsRoutes.delete('/admin/accounts/emails/:email', (c) => {
  const email = normalizeEmail(c.req.param('email'));
  const db = getDb();

  if (!emailExists(db, email)) return c.json({ error: 'Not found' }, 404);
  const row = getRow<{ source: string }>(db, `SELECT source FROM mailbox_registry WHERE email = ?`, email);
  if (!row || row.source !== 'manual') return c.json({ error: 'Pool emails cannot be deleted' }, 400);

  const del = db.transaction(() => {
    db.prepare(`DELETE FROM mailbox_registrations WHERE LOWER(email) = ?`).run(email);
    db.prepare(`DELETE FROM mailbox_registry WHERE email = ?`).run(email);
  });
  del();
  logActivity('amber', `Deleted mailbox: ${email}`);
  return c.json({ ok: true });
});

// Registrations for one mailbox. Supports ?email= filter and ?keyword= search.
accountsRoutes.get('/admin/accounts/registrations', (c) => {
  const db = getDb();
  const email = normalizeEmail(c.req.query('email') ?? '');
  const keyword = String(c.req.query('keyword') ?? '').trim();

  let sql = `SELECT * FROM mailbox_registrations`;
  const where: string[] = [];
  const params: string[] = [];
  if (email) { where.push('LOWER(email) = ?'); params.push(email); }
  if (keyword) {
    where.push('(app_name LIKE ? OR username LIKE ? OR memo LIKE ?)');
    const like = `%${keyword}%`;
    params.push(like, like, like);
  }
  if (where.length) sql += ` WHERE ${where.join(' AND ')}`;
  sql += ` ORDER BY app_name`;

  const rows = allRows<RegistrationRow>(db, sql, ...params);
  return c.json({ registrations: rows });
});

accountsRoutes.post('/admin/accounts/emails/:email/registrations', async (c) => {
  const email = normalizeEmail(c.req.param('email'));
  const body = await c.req.json().catch(() => ({}));
  const appName = String(body.appName ?? '').trim();
  if (!appName) return c.json({ error: 'appName is required' }, 400);

  const db = getDb();
  if (!emailExists(db, email)) return c.json({ error: 'Mailbox not found' }, 404);

  const result = db.prepare(`
    INSERT INTO mailbox_registrations (email, app_name, username, password, memo)
    VALUES (?, ?, ?, ?, ?)
  `).run(
    email,
    appName,
    String(body.username ?? '').trim(),
    String(body.password ?? ''),
    String(body.memo ?? '').trim(),
  );
  logActivity('green', `Added registration for ${appName} (${email})`);
  return c.json({ id: Number(result.lastInsertRowid), email, app_name: appName }, 201);
});

accountsRoutes.patch('/admin/accounts/registrations/:id', async (c) => {
  const id = c.req.param('id');
  const body = await c.req.json().catch(() => ({}));
  const db = getDb();

  const existing = getRow<{ id: number }>(db, `SELECT id FROM mailbox_registrations WHERE id = ?`, id);
  if (!existing) return c.json({ error: 'Not found' }, 404);

  const setClause = buildSetClause(body, {
    app_name: (v) => String(v ?? '').trim(),
    username: (v) => String(v ?? '').trim(),
    password: (v) => String(v ?? ''),
    memo: (v) => String(v ?? '').trim(),
  });
  if (!setClause) return c.json({ error: 'No fields to update' }, 400);

  db.prepare(`UPDATE mailbox_registrations SET ${setClause.setClause}, updated_at = datetime('now') WHERE id = ?`)
    .run(...setClause.params, id);
  logActivity('blue', `Updated registration #${id}`);
  return c.json({ ok: true });
});

accountsRoutes.delete('/admin/accounts/registrations/:id', (c) => {
  const id = c.req.param('id');
  const db = getDb();
  const row = getRow<{ app_name: string; email: string }>(
    db,
    `SELECT app_name, email FROM mailbox_registrations WHERE id = ?`,
    id,
  );
  if (!row) return c.json({ error: 'Not found' }, 404);

  db.prepare(`DELETE FROM mailbox_registrations WHERE id = ?`).run(id);
  logActivity('amber', `Deleted registration: ${row.app_name} (${row.email})`);
  return c.json({ ok: true });
});
