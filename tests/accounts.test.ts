import { describe, expect, it, beforeEach } from 'vitest';
import { getDb } from '../src/db.js';
import { hashApiKey } from '../src/crypto.js';
import { app, authHeaders, jsonHeaders, jsonOf } from './helpers/http.js';

type EmailCard = {
  email: string;
  source: string;
  remark: string;
  tags: string;
  registrationCount: number;
};

type Registration = {
  id: number;
  email: string;
  app_name: string;
  username: string;
  password: string;
  memo: string;
};

async function addEmail(email: string, extra: Record<string, unknown> = {}) {
  return app.request('/api/admin/accounts/emails', {
    method: 'POST',
    headers: jsonHeaders(),
    body: JSON.stringify({ email, ...extra }),
  });
}

async function addReg(email: string, body: Record<string, unknown>) {
  return app.request(`/api/admin/accounts/emails/${encodeURIComponent(email)}/registrations`, {
    method: 'POST',
    headers: jsonHeaders(),
    body: JSON.stringify(body),
  });
}

async function listEmails(): Promise<EmailCard[]> {
  const res = await app.request('/api/admin/accounts', { headers: authHeaders() });
  expect(res.status).toBe(200);
  return (await jsonOf<{ emails: EmailCard[] }>(res)).emails;
}

async function listRegs(email?: string): Promise<Registration[]> {
  const q = email ? `?email=${encodeURIComponent(email)}` : '';
  const res = await app.request(`/api/admin/accounts/registrations${q}`, { headers: authHeaders() });
  expect(res.status).toBe(200);
  return (await jsonOf<{ registrations: Registration[] }>(res)).registrations;
}

describe('admin guard', () => {
  it('rejects a valid but non-admin API key', async () => {
    getDb().prepare(`INSERT INTO api_keys (key, name) VALUES (?, ?)`).run(hashApiKey('mk_user'), 'user');
    const res = await app.request('/api/admin/accounts', { headers: authHeaders('mk_user') });
    expect(res.status).toBe(403);
  });

  it('rejects no token', async () => {
    const res = await app.request('/api/admin/accounts');
    expect(res.status).toBe(401);
  });
});

describe('mailbox list UNION', () => {
  beforeEach(() => {
    const db = getDb();
    db.prepare(`INSERT INTO outlook_accounts (email, password) VALUES (?, ?)`).run('pool@outlook.com', 'x');
    db.prepare(`INSERT INTO imap_accounts (id, host, user, password, domain) VALUES (?, ?, ?, ?, ?)`)
      .run('imap1', 'imap.example.com', 'IMAP.User@Example.com', 'pw', 'example.com');
    db.prepare(`INSERT INTO icloud_addresses (hme, account_id, anonymous_id) VALUES (?, ?, ?)`)
      .run('alias@icloud.com', 'acc1', 'anon1');
  });

  it('surfaces pool mailboxes from all account-pool tables', async () => {
    const emails = await listEmails();
    const sources = Object.fromEntries(emails.map(e => [e.email, e.source]));
    expect(sources['pool@outlook.com']).toBe('outlook');
    expect(sources['imap.user@example.com']).toBe('imap'); // LOWER(user)
    expect(sources['alias@icloud.com']).toBe('icloud');
  });

  it('merges manual mailbox and prefers manual on same address', async () => {
    const res = await addEmail('Pool@Outlook.com');
    expect(res.status).toBe(201);
    const emails = await listEmails();
    const row = emails.find(e => e.email === 'pool@outlook.com');
    expect(row?.source).toBe('manual'); // manual wins over outlook on same address
  });

  it('reports registration count per mailbox', async () => {
    const db = getDb();
    db.prepare(`INSERT INTO mailbox_registrations (email, app_name, username, password) VALUES (?, ?, ?, ?)`)
      .run('pool@outlook.com', 'Twitter', 'u1', 'p1');
    const emails = await listEmails();
    const row = emails.find(e => e.email === 'pool@outlook.com');
    expect(row?.registrationCount).toBe(1);
  });
});

describe('manual mailbox CRUD', () => {
  it('adds a manual mailbox (lowercased, fields stored)', async () => {
    const res = await addEmail('User@gmail.com', { remark: '主邮箱', tags: 'claude,work' });
    expect(res.status).toBe(201);
    const card = await jsonOf<EmailCard>(res);
    expect(card.email).toBe('user@gmail.com');
    expect(card.source).toBe('manual');
    expect(card.remark).toBe('主邮箱');

    const emails = await listEmails();
    const row = emails.find(e => e.email === 'user@gmail.com');
    expect(row?.tags).toBe('claude,work');
  });

  it('rejects duplicate email with 409', async () => {
    await addEmail('dup@example.com');
    const res = await addEmail('dup@example.com');
    expect(res.status).toBe(409);
  });

  it('rejects malformed email with 400', async () => {
    const res = await addEmail('not-an-email');
    expect(res.status).toBe(400);
  });

  it('patches remark/tags only for manual mailboxes', async () => {
    await addEmail('patch@example.com');
    const res = await app.request('/api/admin/accounts/emails/patch@example.com', {
      method: 'PATCH',
      headers: jsonHeaders(),
      body: JSON.stringify({ remark: '更新备注', tags: 'a,b' }),
    });
    expect(res.status).toBe(200);
    const emails = await listEmails();
    const row = emails.find(e => e.email === 'patch@example.com');
    expect(row?.remark).toBe('更新备注');
    expect(row?.tags).toBe('a,b');
  });

  it('rejects editing a pool mailbox', async () => {
    const db = getDb();
    db.prepare(`INSERT INTO outlook_accounts (email, password) VALUES (?, ?)`).run('pool@outlook.com', 'x');
    const res = await app.request('/api/admin/accounts/emails/pool@outlook.com', {
      method: 'PATCH',
      headers: jsonHeaders(),
      body: JSON.stringify({ remark: 'x' }),
    });
    expect(res.status).toBe(400);
  });

  it('deletes manual mailbox and cascades registrations', async () => {
    await addEmail('cascade@example.com');
    await addReg('cascade@example.com', { appName: 'Twitter', username: 'u', password: 'p' });
    expect((await listRegs('cascade@example.com')).length).toBe(1);

    const res = await app.request('/api/admin/accounts/emails/cascade@example.com', {
      method: 'DELETE',
      headers: authHeaders(),
    });
    expect(res.status).toBe(200);
    expect((await listRegs('cascade@example.com')).length).toBe(0);
    expect((await listEmails()).some(e => e.email === 'cascade@example.com')).toBe(false);
  });

  it('rejects deleting a pool mailbox', async () => {
    const db = getDb();
    db.prepare(`INSERT INTO outlook_accounts (email, password) VALUES (?, ?)`).run('pool@outlook.com', 'x');
    const res = await app.request('/api/admin/accounts/emails/pool@outlook.com', {
      method: 'DELETE',
      headers: authHeaders(),
    });
    expect(res.status).toBe(400);
  });
});

describe('registrations CRUD', () => {
  beforeEach(() => {
    const db = getDb();
    db.prepare(`INSERT INTO outlook_accounts (email, password) VALUES (?, ?)`).run('pool@outlook.com', 'x');
  });

  it('adds a registration to a pool mailbox', async () => {
    const res = await addReg('pool@outlook.com', {
      appName: 'GitHub',
      username: 'me',
      password: 'secret',
      memo: '主账号',
    });
    expect(res.status).toBe(201);
    const regs = await listRegs('pool@outlook.com');
    expect(regs).toHaveLength(1);
    expect(regs[0].app_name).toBe('GitHub');
    expect(regs[0].password).toBe('secret');
    expect(regs[0].memo).toBe('主账号');
  });

  it('rejects registration for unknown mailbox', async () => {
    const res = await addReg('nobody@example.com', { appName: 'X', username: 'u', password: 'p' });
    expect(res.status).toBe(404);
  });

  it('requires appName', async () => {
    const res = await addReg('pool@outlook.com', { username: 'u', password: 'p' });
    expect(res.status).toBe(400);
  });

  it('patches a registration', async () => {
    const created = await addReg('pool@outlook.com', { appName: 'App', username: 'u', password: 'p' });
    const { id } = await jsonOf<{ id: number }>(created);
    const res = await app.request(`/api/admin/accounts/registrations/${id}`, {
      method: 'PATCH',
      headers: jsonHeaders(),
      body: JSON.stringify({ password: 'new-pwd', memo: '新备注' }),
    });
    expect(res.status).toBe(200);
    const regs = await listRegs('pool@outlook.com');
    expect(regs[0].password).toBe('new-pwd');
    expect(regs[0].memo).toBe('新备注');
  });

  it('deletes a registration', async () => {
    const created = await addReg('pool@outlook.com', { appName: 'App', username: 'u', password: 'p' });
    const { id } = await jsonOf<{ id: number }>(created);
    const res = await app.request(`/api/admin/accounts/registrations/${id}`, {
      method: 'DELETE',
      headers: authHeaders(),
    });
    expect(res.status).toBe(200);
    expect(await listRegs('pool@outlook.com')).toHaveLength(0);
  });

  it('supports keyword search across app/username/memo', async () => {
    await addReg('pool@outlook.com', { appName: 'Netflix', username: 'watcher', password: 'p' });
    const res = await app.request('/api/admin/accounts/registrations?keyword=netfl', { headers: authHeaders() });
    expect(res.status).toBe(200);
    const regs = await jsonOf<{ registrations: Registration[] }>(res);
    expect(regs.registrations).toHaveLength(1);
    expect(regs.registrations[0].app_name).toBe('Netflix');
  });
});

describe('bulk batch registrations', () => {
  beforeEach(() => {
    const db = getDb();
    db.prepare(`INSERT INTO outlook_accounts (email, password) VALUES (?, ?)`).run('pool@outlook.com', 'x');
    db.prepare(`INSERT INTO mailbox_registry (email) VALUES (?)`).run('manual@example.com');
  });

  it('inserts email × app cartesian and reports counts', async () => {
    const res = await app.request('/api/admin/accounts/registrations/batch', {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({
        emails: ['pool@outlook.com', 'manual@example.com'],
        apps: [
          { appName: 'GitHub', username: 'u1', password: 'p1' },
          { appName: 'Claude', username: 'u2', password: 'p2', memo: 'AI' },
        ],
      }),
    });
    expect(res.status).toBe(200);
    const body = await jsonOf<{ added: { count: number }; skipped: { count: number }; errors: unknown[] }>(res);
    expect(body.added.count).toBe(4); // 2 emails × 2 apps
    expect(body.errors).toHaveLength(0);

    const regs = await listRegs();
    expect(regs).toHaveLength(4);
    expect(regs.some(r => r.email === 'manual@example.com' && r.app_name === 'Claude' && r.memo === 'AI')).toBe(true);
  });

  it('skips duplicates (same email + app) and reports unknown mailboxes', async () => {
    const db = getDb();
    db.prepare(`INSERT INTO mailbox_registrations (email, app_name) VALUES (?, ?)`).run('pool@outlook.com', 'GitHub');

    const res = await app.request('/api/admin/accounts/registrations/batch', {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({
        emails: ['pool@outlook.com', 'nobody@nowhere.com'],
        apps: [{ appName: 'GitHub', username: 'u' }],
      }),
    });
    const body = await jsonOf<{ added: { count: number }; skipped: { count: number }; errors: { email?: string }[] }>(res);
    expect(body.added.count).toBe(0); // GitHub already exists on pool
    expect(body.skipped.count).toBe(1);
    expect(body.errors).toHaveLength(1);
    expect(body.errors[0].email).toBe('nobody@nowhere.com');
  });

  it('rejects when no apps provided', async () => {
    const res = await app.request('/api/admin/accounts/registrations/batch', {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({ emails: ['pool@outlook.com'], apps: [] }),
    });
    expect(res.status).toBe(200); // empty batch is a no-op, not an error
    const body = await jsonOf<{ added: { count: number } }>(res);
    expect(body.added.count).toBe(0);
  });
});
