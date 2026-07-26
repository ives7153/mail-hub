import { describe, expect, it } from 'vitest';
import { getDb } from '../src/db.js';
import { OutlookProvider, buildAliasAddress, stripPlusTag } from '../src/providers/outlook.js';
import { app, authHeaders, jsonHeaders } from './helpers/http.js';

/**
 * Plus addressing is an opt-in per request: the account still serves exactly one
 * inbox, so the tag only buys a distinct address at the target service. The
 * account row stays keyed on the bare email, which is what every lookup
 * (refresh token, api_type, used_services, admin list) must resolve to.
 *
 * Delivery was verified live against the production pool on 2026-07-26: a
 * message sent to `<account>+probe@outlook.com` reached the base mailbox in ~6s
 * with the tag preserved in toRecipients.
 */

function insertAccount(email: string): void {
  getDb().prepare(
    `INSERT INTO outlook_accounts (email, password, client_id, refresh_token, token_status)
     VALUES (?, 'pw', 'cid', 'rt-' || ?, 'valid')`,
  ).run(email, email);
}

describe('stripPlusTag', () => {
  it('removes the tag and keeps the domain', () => {
    expect(stripPlusTag('user+github@outlook.com')).toBe('user@outlook.com');
  });

  it('leaves an untagged address untouched', () => {
    expect(stripPlusTag('user@outlook.com')).toBe('user@outlook.com');
  });

  it('strips only the first plus and ignores one in the domain part', () => {
    expect(stripPlusTag('user+a+b@outlook.com')).toBe('user@outlook.com');
    expect(stripPlusTag('nope')).toBe('nope');
    expect(stripPlusTag('@outlook.com')).toBe('@outlook.com');
  });
});

describe('buildAliasAddress', () => {
  it('appends a sanitized tag', () => {
    expect(buildAliasAddress('user@outlook.com', 'GitHub')).toBe('user+github@outlook.com');
    expect(buildAliasAddress('user@outlook.com', 'a b/c')).toBe('user+abc@outlook.com');
  });

  it('keeps the local part within the RFC 5321 64-octet limit', () => {
    const base = 'a'.repeat(60);
    const addr = buildAliasAddress(`${base}@outlook.com`, 'toolongtag')!;
    expect(addr.split('@')[0].length).toBeLessThanOrEqual(64);
    expect(addr.startsWith(`${base}+`)).toBe(true);
  });

  it('returns null when no usable tag survives, so the caller falls back', () => {
    expect(buildAliasAddress('user@outlook.com', '///')).toBeNull();
    expect(buildAliasAddress('user@outlook.com', '')).toBeNull();
    expect(buildAliasAddress('a'.repeat(64) + '@outlook.com', 'tag')).toBeNull();
    expect(buildAliasAddress('notanemail', 'tag')).toBeNull();
  });
});

describe('Outlook alias allocation', () => {
  it('uses the plain account address unless an alias is requested', async () => {
    insertAccount('plain@outlook.com');
    const inbox = await new OutlookProvider().createInbox({ inboxId: 'i-plain' });

    expect(inbox.address).toBe('plain@outlook.com');
    expect(inbox.authData.email).toBe('plain@outlook.com');
  });

  it('generates the tag server-side when alias is requested', async () => {
    insertAccount('aliased@outlook.com');
    const inbox = await new OutlookProvider().createInbox({ inboxId: 'i-alias', alias: true });

    // The caller never supplies a tag, so assert the shape, not a literal.
    expect(inbox.address).toMatch(/^aliased\+[a-z0-9]{8}@outlook\.com$/);
    // The credential path (refresh token, api_type, release) resolves via this.
    expect(inbox.authData.email).toBe('aliased@outlook.com');
  });

  it('gives two inboxes on the same account distinct tags', async () => {
    insertAccount('reuse@outlook.com');
    const provider = new OutlookProvider();
    const first = await provider.createInbox({ inboxId: 'i-r1', alias: true });
    await provider.releaseInbox(first, 'i-r1');
    const second = await provider.createInbox({ inboxId: 'i-r2', alias: true });

    expect(first.address).not.toBe(second.address);
    expect(stripPlusTag(first.address)).toBe('reuse@outlook.com');
    expect(stripPlusTag(second.address)).toBe('reuse@outlook.com');
  });

  it('still consumes exactly one account — aliasing is not extra capacity', async () => {
    insertAccount('solo@outlook.com');
    const provider = new OutlookProvider();
    await provider.createInbox({ inboxId: 'i-first', alias: true });

    await expect(provider.createInbox({ inboxId: 'i-second', alias: true }))
      .rejects.toThrow(/无可用账号/);
  });

  it('releases the account when an alias inbox is released', async () => {
    insertAccount('rel@outlook.com');
    const provider = new OutlookProvider();
    const inbox = await provider.createInbox({ inboxId: 'i-rel', alias: true });

    await provider.releaseInbox(inbox, 'i-rel');

    const row = getDb().prepare(`SELECT assigned_inbox_id FROM outlook_accounts WHERE email = 'rel@outlook.com'`).get() as { assigned_inbox_id: string | null };
    expect(row.assigned_inbox_id).toBeNull();
  });

  it('falls back to the plain address when no tag can fit the local part', async () => {
    // 60-char base leaves room for '+' plus 3 chars, so an 8-char tag is cut
    // down rather than overflowing; a 64-char base leaves no room at all.
    const full = 'a'.repeat(64);
    insertAccount(`${full}@outlook.com`);
    const inbox = await new OutlookProvider().createInbox({ inboxId: 'i-fb', alias: true });

    expect(inbox.address).toBe(`${full}@outlook.com`);
  });
});

/**
 * used_services is a permanent anti-reuse blacklist keyed on the ACCOUNT email.
 * Looking it up by the inbox address would silently match no row for an alias
 * inbox, leaving the blacklist unwritten — the account would be handed to the
 * same service again.
 */
describe('report route resolves the account behind an alias inbox', () => {
  async function createInboxFor(email: string, tag: string | undefined, id: string): Promise<void> {
    insertAccount(email);
    const inbox = await new OutlookProvider().createInbox({ inboxId: id, ...(tag ? { alias: true } : {}) });
    getDb().prepare(
      `INSERT INTO inboxes (id, provider, address, auth_data, api_base, target_service, status)
       VALUES (?, 'outlook', ?, ?, '', 'netlify.com', 'active')`,
    ).run(id, inbox.address, JSON.stringify(inbox.authData));
  }

  it('records used_services against the base account for an alias inbox', async () => {
    await createInboxFor('report-alias@outlook.com', 'netlify', 'r-alias');

    const res = await app.request('/api/inbox/r-alias/report', {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({ success: true, service: 'netlify.com' }),
    });
    expect(res.status).toBe(200);

    const row = getDb().prepare(`SELECT used_services FROM outlook_accounts WHERE email = 'report-alias@outlook.com'`).get() as { used_services: string };
    expect(JSON.parse(row.used_services)).toEqual(['netlify.com']);
  });

  it('still records for a plain (non-alias) inbox', async () => {
    await createInboxFor('report-plain@outlook.com', undefined, 'r-plain');

    await app.request('/api/inbox/r-plain/report', {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({ success: true, service: 'netlify.com' }),
    });

    const row = getDb().prepare(`SELECT used_services FROM outlook_accounts WHERE email = 'report-plain@outlook.com'`).get() as { used_services: string };
    expect(JSON.parse(row.used_services)).toEqual(['netlify.com']);
  });

  /**
   * The point of the alias: one account may register at the SAME service more
   * than once, because each registration gets a fresh address. A plain request
   * for that service still finds the account excluded, so turning the flag off
   * restores the original guarantee. Exercised through the real API, since the
   * used_services filter sits in the account-selection SQL and is also applied
   * by getDomains — miss either and dispatch fails before createInbox runs.
   */
  it('lets an aliased request reuse an account already used for that service', async () => {
    insertAccount('reuse-ok@outlook.com');

    async function createAndBurn(alias: boolean): Promise<{ status: number; address?: string }> {
      const res = await app.request('/api/inbox', {
        method: 'POST', headers: jsonHeaders(),
        body: JSON.stringify({ for: 'guarded.com', provider: 'outlook', alias }),
      });
      const body = await res.json() as { id?: string; address?: string };
      if (res.status === 201 && body.id) {
        await app.request(`/api/inbox/${body.id}/report`, {
          method: 'POST', headers: jsonHeaders(),
          body: JSON.stringify({ success: true, service: 'guarded.com' }),
        });
        // Close so only used_services — not the 1:1 assignment — can exclude it.
        await app.request(`/api/inbox/${body.id}`, { method: 'DELETE', headers: authHeaders() });
      }
      return { status: res.status, address: body.address };
    }

    const first = await createAndBurn(true);
    expect(first.status).toBe(201);

    const used = getDb().prepare(`SELECT used_services FROM outlook_accounts WHERE email = 'reuse-ok@outlook.com'`).get() as { used_services: string };
    expect(JSON.parse(used.used_services)).toEqual(['guarded.com']);

    // Same account, same service, second registration — allowed, new address.
    const second = await createAndBurn(true);
    expect(second.status).toBe(201);
    expect(second.address).toMatch(/^reuse-ok\+[a-z0-9]{8}@outlook\.com$/);
    expect(second.address).not.toBe(first.address);

    // …and a third, to show it is not a one-off exemption.
    const third = await createAndBurn(true);
    expect(third.status).toBe(201);
    expect(new Set([first.address, second.address, third.address]).size).toBe(3);

    // Without the flag the original protection still holds.
    const plain = await app.request('/api/inbox', {
      method: 'POST', headers: jsonHeaders(),
      body: JSON.stringify({ for: 'guarded.com', provider: 'outlook' }),
    });
    expect(plain.status).not.toBe(201);
  });

  it('excludes the account from later dispatch for that service', async () => {
    await createInboxFor('report-burn@outlook.com', 'svc', 'r-burn');
    await app.request('/api/inbox/r-burn/report', {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({ success: true, service: 'burned.com' }),
    });
    // Free the account so only the used_services filter can exclude it.
    getDb().prepare(`UPDATE outlook_accounts SET assigned_inbox_id = NULL WHERE email = 'report-burn@outlook.com'`).run();

    await expect(new OutlookProvider().createInbox({ inboxId: 'i-again', for: 'burned.com' }))
      .rejects.toThrow(/无可用账号/);
  });
});

/**
 * The feature exists so an API caller can turn it on per request — the caller
 * never invents a tag. These lock the wire contract of POST /api/inbox.
 */
describe('POST /api/inbox honours the alias flag', () => {
  async function create(body: Record<string, unknown>): Promise<{ status: number; address?: string }> {
    const res = await app.request('/api/inbox', {
      method: 'POST', headers: jsonHeaders(), body: JSON.stringify(body),
    });
    const json = await res.json() as { address?: string };
    return { status: res.status, address: json.address };
  }

  it('returns a tagged address when alias is true', async () => {
    insertAccount('apion@outlook.com');
    const r = await create({ for: 'alias-on.test', provider: 'outlook', alias: true });

    expect(r.status).toBe(201);
    expect(r.address).toMatch(/^apion\+[a-z0-9]{8}@outlook\.com$/);
  });

  it('returns the plain address when alias is absent or false', async () => {
    insertAccount('apioff@outlook.com');
    const r = await create({ for: 'alias-off.test', provider: 'outlook' });
    expect(r.address).toBe('apioff@outlook.com');

    insertAccount('apifalse@outlook.com');
    const r2 = await create({ for: 'alias-false.test', provider: 'outlook', alias: false });
    expect(r2.address).toBe('apifalse@outlook.com');
  });

  it('ignores a non-boolean alias rather than treating it as opt-in', async () => {
    insertAccount('apistr@outlook.com');
    const r = await create({ for: 'alias-str.test', provider: 'outlook', alias: 'yes' });

    expect(r.address).toBe('apistr@outlook.com');
  });
});

describe('admin account list links alias inboxes back to the account', () => {
  it('reports last_inbox_id for an alias inbox', async () => {
    insertAccount('lastinbox@outlook.com');
    const inbox = await new OutlookProvider().createInbox({ inboxId: 'i-last', alias: true });
    getDb().prepare(
      `INSERT INTO inboxes (id, provider, address, auth_data, api_base, status)
       VALUES ('i-last', 'outlook', ?, ?, '', 'active')`,
    ).run(inbox.address, JSON.stringify(inbox.authData));

    const res = await app.request('/api/outlook/accounts', { headers: authHeaders() });
    const body = await res.json() as { accounts: { email: string; last_inbox_id: string | null }[] };
    const account = body.accounts.find((a) => a.email === 'lastinbox@outlook.com');

    expect(account?.last_inbox_id).toBe('i-last');
  });
});
