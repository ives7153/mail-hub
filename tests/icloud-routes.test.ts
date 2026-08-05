import { describe, it, expect, vi, afterEach } from 'vitest';
import { existsSync, mkdirSync, rmSync, writeFileSync } from 'fs';
import { dirname } from 'path';
import { app, authHeaders, jsonHeaders } from './helpers/http.js';
import { getDb } from '../src/db.js';
import { hashApiKey } from '../src/crypto.js';
import { trustTokenPath } from '../src/providers/icloud-auth.js';

const srpMockState = vi.hoisted(() => ({ rejectPassword: false }));

vi.mock('icloudjs', () => {
  class FakeService {
    authStore = { trustToken: '', getHeaders: () => ({ Cookie: '' }) };
    status = 'MfaRequested';
    awaitReady = Promise.resolve();
    async authenticate(): Promise<void> {
      if (srpMockState.rejectPassword) throw new Error('Apple rejected the password');
    }
  }
  return { default: FakeService };
});

afterEach(() => {
  srpMockState.rejectPassword = false;
  vi.unstubAllGlobals();
});

describe('iCloud admin routes', () => {
  it('refuses a non-admin caller', async () => {
    // A request with no token at all is answered 401 by the global /api/*
    // middleware and never reaches this router, so the caller has to be a real
    // regular API key for the admin boundary itself to be under test.
    getDb().prepare(`INSERT INTO api_keys (key, name) VALUES (?, ?)`).run(hashApiKey('mk_user'), 'user');

    const res = await app.request('/api/icloud/accounts', { headers: authHeaders('mk_user') });
    expect(res.status).toBe(403);
  });

  it('creates an account without echoing its secrets back', async () => {
    const res = await app.request('/api/icloud/accounts', {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({
        appleId: 'me@icloud.com',
        cookies: 'X-APPLE-WEBAUTH-TOKEN=secret; X-APPLE-WEBAUTH-USER=whoami',
        imapUser: 'me@icloud.com',
        imapPassword: 'abcd-efgh-ijkl-mnop',
      }),
    });

    expect(res.status).toBe(200);
    const body = await res.json() as { account: Record<string, unknown> };
    expect(body.account.apple_id).toBe('me@icloud.com');
    expect(JSON.stringify(body)).not.toContain('secret');
    expect(JSON.stringify(body)).not.toContain('abcd-efgh');
  });

  it('starts a new account as pending, so nothing serves it before /test', async () => {
    await app.request('/api/icloud/accounts', {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({
        appleId: 'untested@icloud.com',
        cookies: 'X-APPLE-WEBAUTH-TOKEN=t; X-APPLE-WEBAUTH-USER=u',
        // A wrong/blank IMAP password is invisible until the first poll fails;
        // the account must not mint or dispatch on the strength of the cookie
        // alone. getAccountById and refill both take only active/degraded.
        imapPassword: '',
      }),
    });

    const row = getDb().prepare(
      `SELECT status FROM icloud_accounts WHERE apple_id = 'untested@icloud.com'`,
    ).get() as { status: string };
    expect(row.status).toBe('pending');
  });

  it('accepts a whole Copy-as-cURL paste rather than only the Cookie header', async () => {
    const res = await app.request('/api/icloud/accounts', {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({
        appleId: 'curl@icloud.com',
        // Digging the Cookie line out of this by hand is the chore that
        // produces truncated pastes and a confusing failure an hour later.
        cookies: `curl 'https://p68-maildomainws.icloud.com/v1/hme/list' \\\n  -H 'accept: */*' \\\n  -H 'cookie: X-APPLE-WEBAUTH-TOKEN=tok; X-APPLE-WEBAUTH-USER=usr' \\\n  --compressed`,
        imapPassword: 'pw',
      }),
    });

    expect(res.status).toBe(200);
    const stored = getDb().prepare(
      `SELECT cookies FROM icloud_accounts WHERE apple_id = 'curl@icloud.com'`,
    ).get() as { cookies: string };
    expect(stored.cookies).toBe('X-APPLE-WEBAUTH-TOKEN=tok; X-APPLE-WEBAUTH-USER=usr');
  });

  it('names the missing cookie instead of storing a half-copied session', async () => {
    const res = await app.request('/api/icloud/accounts', {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({
        appleId: 'partial@icloud.com',
        // The /reportStats request carries the token but not the user, and it
        // is the one people happen to right-click.
        cookies: 'X-APPLE-WEBAUTH-TOKEN=tok; other=1',
        imapPassword: 'pw',
      }),
    });

    expect(res.status).toBe(400);
    const body = await res.json() as { missing: string[]; error: string };
    expect(body.missing).toEqual(['X-APPLE-WEBAUTH-USER']);
    expect(body.error).toContain('X-APPLE-WEBAUTH-USER');
    // Rejected means not stored — a half-copied session must not sit there
    // looking configured until the first poll fails.
    const row = getDb().prepare(
      `SELECT COUNT(*) AS c FROM icloud_accounts WHERE apple_id = 'partial@icloud.com'`,
    ).get() as { c: number };
    expect(row.c).toBe(0);
  });

  it('swaps an expired cookie in place without losing the address pool', async () => {
    getDb().prepare(
      `INSERT INTO icloud_accounts (id, apple_id, cookies, hme_service_url)
       VALUES ('acc-ck', 'ck@icloud.com', 'X-APPLE-WEBAUTH-TOKEN=old; X-APPLE-WEBAUTH-USER=old', 'https://stale.test')`,
    ).run();
    getDb().prepare(
      `INSERT INTO icloud_addresses (hme, account_id, anonymous_id) VALUES ('keep@icloud.com', 'acc-ck', 'anon-keep')`,
    ).run();

    const res = await app.request('/api/icloud/accounts/acc-ck/cookies', {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({ cookies: 'X-APPLE-WEBAUTH-TOKEN=new; X-APPLE-WEBAUTH-USER=new' }),
    });

    expect(res.status).toBe(200);
    const row = getDb().prepare(
      `SELECT cookies, hme_service_url FROM icloud_accounts WHERE id = 'acc-ck'`,
    ).get() as { cookies: string; hme_service_url: string };
    expect(row.cookies).toContain('=new');
    // The partition host belongs to the old session; keeping it would send the
    // new cookie to a host that no longer answers for it.
    expect(row.hme_service_url).toBe('');

    // Deleting and recreating the account would have burned this alias: every
    // address is one of the 750 an Apple ID ever gets.
    const pool = getDb().prepare(
      `SELECT COUNT(*) AS c FROM icloud_addresses WHERE account_id = 'acc-ck'`,
    ).get() as { c: number };
    expect(pool.c).toBe(1);
  });

  it('never returns cookies or the app-specific password when listing accounts', async () => {
    // The list route carries its own column list, so the create route's
    // omission proves nothing about it.
    getDb().prepare(
      `INSERT INTO icloud_accounts (id, apple_id, cookies, imap_password)
       VALUES ('acc-list', 'me@icloud.com', 'X-APPLE-WEBAUTH-TOKEN=cookie-secret', 'wxyz-app-password')`,
    ).run();

    const res = await app.request('/api/icloud/accounts', { headers: authHeaders() });

    expect(res.status).toBe(200);
    const raw = await res.text();
    expect(raw).toContain('acc-list');
    expect(raw).not.toContain('cookie-secret');
    expect(raw).not.toContain('wxyz-app-password');
  });

  it('mints addresses into the pool and marks them as Mail Hub owned', async () => {
    getDb().prepare(
      `INSERT INTO icloud_accounts (id, apple_id, cookies, hme_service_url)
       VALUES ('acc-1', 'me@icloud.com', 'c', 'https://svc.test')`,
    ).run();

    const bodies: Record<string, unknown>[] = [];
    vi.stubGlobal('fetch', vi.fn(async (url: string, init?: RequestInit) => {
      if (init?.body) bodies.push(JSON.parse(String(init.body)));
      const isReserve = String(url).includes('/reserve');
      return {
        ok: true,
        status: 200,
        headers: { get: () => null },
        text: async () => JSON.stringify(isReserve
          ? { success: true, result: { hme: { hme: 'fresh@icloud.com', anonymousId: 'anon-1', label: 'mail-hub', note: '', isActive: true } } }
          : { success: true, result: { hme: 'fresh@icloud.com' } }),
      };
    }));

    const res = await app.request('/api/icloud/accounts/acc-1/generate', {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({ count: 1 }),
    });

    expect(res.status).toBe(200);
    const row = getDb().prepare(
      `SELECT state, anonymous_id FROM icloud_addresses WHERE hme = 'fresh@icloud.com'`,
    ).get() as { state: string; anonymous_id: string };
    expect(row.state).toBe('free');
    expect(row.anonymous_id).toBe('anon-1');
    expect(bodies.some((b) => b.label === 'mail-hub')).toBe(true);
  });

  it('refuses an SRP sign-in for an account with no password saved', async () => {
    getDb().prepare(
      `INSERT INTO icloud_accounts (id, apple_id) VALUES ('acc-srp', 'me@icloud.com')`,
    ).run();

    const res = await app.request('/api/icloud/accounts/acc-srp/srp/begin', {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({}),
    });

    // Saying which credential is missing beats a stack trace out of the SRP
    // library, which is what a bare handshake attempt would produce.
    expect(res.status).toBe(400);
    expect((await res.json() as { error: string }).error).toContain('password');
  });

  it('keeps the working credential state when Apple rejects a replacement password', async () => {
    getDb().prepare(
      `INSERT INTO icloud_accounts (id, apple_id, auth_mode, password)
       VALUES ('acc-srp-reject', 'working@icloud.com', 'cookie', 'working-password')`,
    ).run();
    srpMockState.rejectPassword = true;

    const res = await app.request('/api/icloud/accounts/acc-srp-reject/srp/begin', {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({ password: 'mistyped-password' }),
    });

    expect(res.status).toBe(400);
    const row = getDb().prepare(
      `SELECT password, auth_mode FROM icloud_accounts WHERE id = 'acc-srp-reject'`,
    ).get() as { password: string; auth_mode: string };
    expect(row).toEqual({ password: 'working-password', auth_mode: 'cookie' });

    srpMockState.rejectPassword = false;
    const accepted = await app.request('/api/icloud/accounts/acc-srp-reject/srp/begin', {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({ password: 'accepted-password' }),
    });
    expect(accepted.status).toBe(200);
    const acceptedRow = getDb().prepare(
      `SELECT password, auth_mode FROM icloud_accounts WHERE id = 'acc-srp-reject'`,
    ).get() as { password: string; auth_mode: string };
    expect(acceptedRow).toEqual({ password: 'accepted-password', auth_mode: 'srp' });
  });

  it('refuses a 2FA code for a session it never issued', async () => {
    const res = await app.request('/api/icloud/accounts/acc-srp/srp/complete', {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({ sessionId: 'never-issued', code: '123456' }),
    });

    expect(res.status).toBe(400);
  });

  it('reports Apple’s own message when generation is refused', async () => {
    getDb().prepare(
      `INSERT INTO icloud_accounts (id, apple_id, cookies, hme_service_url)
       VALUES ('acc-2', 'me@icloud.com', 'c', 'https://svc.test')`,
    ).run();
    vi.stubGlobal('fetch', vi.fn(async () => ({
      ok: true,
      status: 200,
      headers: { get: () => null },
      text: async () => JSON.stringify({ success: false, error: { errorMessage: 'Rate limited, try later' } }),
    })));

    const res = await app.request('/api/icloud/accounts/acc-2/generate', {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({ count: 1 }),
    });

    const body = await res.json() as { error?: string; created?: number };
    expect(body.created).toBe(0);
    expect(body.error).toContain('Rate limited, try later');
  });

  it('deactivates a burned address at Apple and never hands it out again', async () => {
    getDb().prepare(
      `INSERT INTO icloud_accounts (id, apple_id, cookies, hme_service_url)
       VALUES ('acc-rt', 'me@icloud.com', 'c', 'https://svc.test')`,
    ).run();
    getDb().prepare(
      `INSERT INTO icloud_addresses (hme, account_id, anonymous_id) VALUES ('burned@icloud.com', 'acc-rt', 'anon-burn')`,
    ).run();

    const bodies: unknown[] = [];
    vi.stubGlobal('fetch', vi.fn(async (_u: string, init?: RequestInit) => {
      if (init?.body) bodies.push(JSON.parse(String(init.body)));
      return { ok: true, status: 200, headers: { get: () => null }, text: async () => JSON.stringify({ success: true, result: {} }) };
    }));

    const res = await app.request('/api/icloud/addresses/burned@icloud.com/retire', {
      method: 'POST', headers: jsonHeaders(),
    });

    expect(res.status).toBe(200);
    // This is the one place Apple-side deactivation belongs: the address is
    // burned, so the slot is spent either way and recycling it would hand a
    // tenant an alias a target already blocks.
    expect(bodies).toContainEqual({ anonymousId: 'anon-burn' });
    const row = getDb().prepare(
      `SELECT state FROM icloud_addresses WHERE hme = 'burned@icloud.com'`,
    ).get() as { state: string };
    expect(row.state).toBe('retired');
  });

  it('claims the address before Apple is called, so dispatch cannot grab it mid-retire', async () => {
    getDb().prepare(
      `INSERT INTO icloud_accounts (id, apple_id, cookies, hme_service_url)
       VALUES ('acc-order', 'me@icloud.com', 'c', 'https://svc.test')`,
    ).run();
    getDb().prepare(
      `INSERT INTO icloud_addresses (hme, account_id, anonymous_id) VALUES ('mid@icloud.com', 'acc-order', 'anon-mid')`,
    ).run();

    const statesSeenByApple: string[] = [];
    vi.stubGlobal('fetch', vi.fn(async () => {
      const row = getDb().prepare(
        `SELECT state FROM icloud_addresses WHERE hme = 'mid@icloud.com'`,
      ).get() as { state: string };
      statesSeenByApple.push(row.state);
      return { ok: true, status: 200, headers: { get: () => null }, text: async () => JSON.stringify({ success: true, result: {} }) };
    }));

    const res = await app.request('/api/icloud/addresses/mid@icloud.com/retire', {
      method: 'POST', headers: jsonHeaders(),
    });

    expect(res.status).toBe(200);
    // The row must already be off the market while Apple deactivates it. When
    // it was still 'free' here, a dispatch could claim it in that window and
    // end up holding an alias Apple had just killed.
    expect(statesSeenByApple).toEqual(['retiring']);
  });

  it('puts the address back as it was only when Apple explicitly rejects deactivation', async () => {
    getDb().prepare(
      `INSERT INTO icloud_accounts (id, apple_id, cookies, hme_service_url)
       VALUES ('acc-fail', 'me@icloud.com', 'c', 'https://svc.test')`,
    ).run();
    getDb().prepare(
      `INSERT INTO icloud_addresses (hme, account_id, anonymous_id) VALUES ('stay@icloud.com', 'acc-fail', 'anon-stay')`,
    ).run();
    vi.stubGlobal('fetch', vi.fn(async () => ({
      ok: true, status: 200, headers: { get: () => null },
      text: async () => JSON.stringify({ success: false, error: { errorMessage: 'deactivation is not allowed for this alias' } }),
    })));

    const res = await app.request('/api/icloud/addresses/stay@icloud.com/retire', {
      method: 'POST', headers: jsonHeaders(),
    });

    expect(res.status).toBe(502);
    // A completed response that explicitly rejected the op proves the alias
    // still forwards, so it must not be stranded in 'retiring' — that would
    // quietly shrink the pool by one for ever.
    const row = getDb().prepare(
      `SELECT state FROM icloud_addresses WHERE hme = 'stay@icloud.com'`,
    ).get() as { state: string };
    expect(row.state).toBe('free');
  });

  it('keeps an address retiring when the deactivation result is ambiguous', async () => {
    getDb().prepare(
      `INSERT INTO icloud_accounts (id, apple_id, cookies, hme_service_url)
       VALUES ('acc-timeout', 'me@icloud.com', 'c', 'https://svc.test')`,
    ).run();
    getDb().prepare(
      `INSERT INTO icloud_addresses (hme, account_id, anonymous_id) VALUES ('uncertain@icloud.com', 'acc-timeout', 'anon-uncertain')`,
    ).run();
    vi.stubGlobal('fetch', vi.fn(async () => { throw new Error('network timeout'); }));
    const output: string[] = [];
    const logSpy = vi.spyOn(process.stdout, 'write').mockImplementation((chunk) => {
      output.push(String(chunk));
      return true;
    });

    let res: Response;
    try {
      res = await app.request('/api/icloud/addresses/uncertain@icloud.com/retire', {
        method: 'POST', headers: jsonHeaders(),
      });
    } finally {
      logSpy.mockRestore();
    }

    // A timeout can land after Apple already deactivated, so returning the
    // alias to the pool would hand a dead address to the next tenant. Stay
    // fail-closed in 'retiring'.
    expect(res.status).toBe(502);
    expect(getDb().prepare(`SELECT state FROM icloud_addresses WHERE hme = 'uncertain@icloud.com'`).get()).toEqual({ state: 'retiring' });
    const warnings = output
      .flatMap((chunk) => chunk.trim().split('\n'))
      .filter(Boolean)
      .map((line) => JSON.parse(line) as Record<string, unknown>)
      .filter((entry) => entry.level === 'warn' && entry.module === 'icloud-route');
    expect(warnings).toHaveLength(1);
    expect(warnings[0]).toMatchObject({ address: 'uncertain@icloud.com' });
    expect(warnings[0]?.operatorAction).toMatch(/inspect.*retry/i);
    expect(warnings[0]?.error).toContain('network timeout');
  });

  it.each([
    ['returns HTTP 5xx', async () => ({ ok: false, status: 503, headers: { get: () => null }, text: async () => 'upstream unavailable' })],
    ['returns malformed JSON', async () => ({ ok: true, status: 200, headers: { get: () => null }, text: async () => '<html>gateway</html>' })],
  ])('keeps an address retiring when Apple %s', async (_caseName, response) => {
    const suffix = _caseName.replace(/\W+/g, '-');
    getDb().prepare(
      `INSERT INTO icloud_accounts (id, apple_id, cookies, hme_service_url)
       VALUES (?, ?, 'c', 'https://svc.test')`,
    ).run(`acc-${suffix}`, `${suffix}@icloud.com`);
    getDb().prepare(
      `INSERT INTO icloud_addresses (hme, account_id, anonymous_id) VALUES (?, ?, ?)`,
    ).run(`${suffix}@icloud.com`, `acc-${suffix}`, `anon-${suffix}`);
    vi.stubGlobal('fetch', vi.fn(response));

    const res = await app.request(`/api/icloud/addresses/${encodeURIComponent(`${suffix}@icloud.com`)}/retire`, {
      method: 'POST', headers: jsonHeaders(),
    });

    expect(res.status).toBe(502);
    expect(getDb().prepare(`SELECT state FROM icloud_addresses WHERE hme = ?`).get(`${suffix}@icloud.com`)).toEqual({ state: 'retiring' });
  });

  it('restores the old state for Apple’s structured 409 operation rejection', async () => {
    getDb().prepare(
      `INSERT INTO icloud_accounts (id, apple_id, cookies, hme_service_url)
       VALUES ('acc-409', '409@icloud.com', 'c', 'https://svc.test')`,
    ).run();
    getDb().prepare(
      `INSERT INTO icloud_addresses (hme, account_id, anonymous_id) VALUES ('reject409@icloud.com', 'acc-409', 'anon-409')`,
    ).run();
    vi.stubGlobal('fetch', vi.fn(async () => ({
      ok: false, status: 409, headers: { get: () => null },
      text: async () => JSON.stringify({ success: false, error: { errorMessage: 'operation rejected' } }),
    })));

    const res = await app.request('/api/icloud/addresses/reject409@icloud.com/retire', {
      method: 'POST', headers: jsonHeaders(),
    });

    // A 409 whose body says success:false is a definitive business rejection,
    // so the alias is still live and returns to 'free'.
    expect(res.status).toBe(502);
    expect(getDb().prepare(`SELECT state FROM icloud_addresses WHERE hme = 'reject409@icloud.com'`).get()).toEqual({ state: 'free' });
  });

  it('refuses to retire an address already mid-retirement, so two calls cannot both claim it', async () => {
    getDb().prepare(
      `INSERT INTO icloud_accounts (id, apple_id, cookies, hme_service_url)
       VALUES ('acc-conc', 'me@icloud.com', 'c', 'https://svc.test')`,
    ).run();
    // The row a first request has already claimed. The claim matches only
    // free/retired, never 'retiring' — because SQLite counts a no-op
    // retiring→retiring UPDATE as changes===1, so matching it would let a
    // second caller "win" and, if the first then rolled back to free, its
    // success write would target a no-longer-'retiring' row and vanish,
    // leaving a deactivated alias dispatchable.
    getDb().prepare(
      `INSERT INTO icloud_addresses (hme, account_id, anonymous_id, state)
       VALUES ('mid2@icloud.com', 'acc-conc', 'anon-mid2', 'retiring')`,
    ).run();
    const fetchSpy = vi.fn();
    vi.stubGlobal('fetch', fetchSpy);

    const res = await app.request('/api/icloud/addresses/mid2@icloud.com/retire', {
      method: 'POST', headers: jsonHeaders(),
    });

    expect(res.status).toBe(409);
    // The second caller must not have called Apple at all.
    expect(fetchSpy).not.toHaveBeenCalled();
    const row = getDb().prepare(
      `SELECT state FROM icloud_addresses WHERE hme = 'mid2@icloud.com'`,
    ).get() as { state: string };
    expect(row.state).toBe('retiring');
  });

  it('erases the on-disk trust token when the account is deleted', async () => {
    getDb().prepare(
      `INSERT INTO icloud_accounts (id, apple_id, auth_mode, trust_token)
       VALUES ('acc-del', 'del@icloud.com', 'srp', 'tok-xyz')`,
    ).run();
    // The token file is what the SRP flow writes; it is a 2FA-bypass credential
    // and must not survive the row it belongs to.
    const path = trustTokenPath('del@icloud.com');
    mkdirSync(dirname(path), { recursive: true });
    writeFileSync(path, 'tok-xyz', 'utf8');
    expect(existsSync(path)).toBe(true);

    const res = await app.request('/api/icloud/accounts/acc-del', {
      method: 'DELETE', headers: authHeaders('admin-secret'),
    });

    expect(res.status).toBe(200);
    expect(existsSync(path)).toBe(false);
  });

  it('keeps account deletion retryable when the trust token cannot be erased', async () => {
    getDb().prepare(
      `INSERT INTO icloud_accounts (id, apple_id, auth_mode, trust_token)
       VALUES ('acc-stuck-token', 'stuck@icloud.com', 'srp', 'tok-stuck')`,
    ).run();
    getDb().prepare(
      `INSERT INTO icloud_addresses (hme, account_id, anonymous_id)
       VALUES ('keep-stuck@icloud.com', 'acc-stuck-token', 'anon-stuck')`,
    ).run();
    const path = trustTokenPath('stuck@icloud.com');
    mkdirSync(path, { recursive: true });
    writeFileSync(`${path}/credential`, 'tok-stuck', 'utf8');

    try {
      const res = await app.request('/api/icloud/accounts/acc-stuck-token', {
        method: 'DELETE', headers: authHeaders('admin-secret'),
      });

      expect(res.status).toBe(500);
      expect((await res.json() as { error: string }).error).toContain('trust token');
      expect(existsSync(`${path}/credential`)).toBe(true);
      const account = getDb().prepare(
        `SELECT id FROM icloud_accounts WHERE id = 'acc-stuck-token'`,
      ).get();
      const address = getDb().prepare(
        `SELECT hme FROM icloud_addresses WHERE account_id = 'acc-stuck-token'`,
      ).get();
      expect(account).toBeDefined();
      expect(address).toBeDefined();
    } finally {
      rmSync(path, { recursive: true, force: true });
    }
  });

  it('refuses to retire an address a live inbox still holds', async () => {
    getDb().prepare(
      `INSERT INTO icloud_accounts (id, apple_id, cookies, hme_service_url)
       VALUES ('acc-rt2', 'me@icloud.com', 'c', 'https://svc.test')`,
    ).run();
    getDb().prepare(
      `INSERT INTO icloud_addresses (hme, account_id, anonymous_id, state, assigned_inbox_id)
       VALUES ('busy@icloud.com', 'acc-rt2', 'anon-busy', 'assigned', 'ib-live')`,
    ).run();

    const res = await app.request('/api/icloud/addresses/busy@icloud.com/retire', {
      method: 'POST', headers: jsonHeaders(),
    });

    expect(res.status).toBe(409);
  });
});
