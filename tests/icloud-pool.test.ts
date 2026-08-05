import { describe, it, expect } from 'vitest';
import { getDb } from '../src/db.js';
import { inCooldown, setCooldown, clearCooldown, reconcileAccount, reapOrphanedAddresses } from '../src/providers/icloud.js';
import { MAILHUB_HME_LABEL } from '../src/providers/icloud-client.js';

function seed(id = 'acc-1'): void {
  getDb().prepare(`INSERT INTO icloud_accounts (id, apple_id) VALUES (?, 'me@icloud.com')`).run(id);
}

describe('refill cooldown', () => {
  it('reports no cooldown for a fresh account', () => {
    seed();
    expect(inCooldown('acc-1')).toBe(false);
  });

  it('holds an account off for the requested window', () => {
    seed();
    setCooldown('acc-1', 45 * 60 * 1000, 'Rate limited by Apple');
    expect(inCooldown('acc-1')).toBe(true);

    const row = getDb().prepare(
      `SELECT pool_cooldown_until, last_error FROM icloud_accounts WHERE id = 'acc-1'`,
    ).get() as { pool_cooldown_until: string; last_error: string };
    expect(row.pool_cooldown_until).toBeTruthy();
    // Apple's failure taxonomy is unknown, so the raw text is what we keep.
    expect(row.last_error).toBe('Rate limited by Apple');
  });

  it('lets an expired cooldown through', () => {
    seed();
    // Written directly: setCooldown refuses non-positive windows, and what is
    // under test here is reading an already-elapsed timestamp, not setting one.
    getDb().prepare(
      `UPDATE icloud_accounts SET pool_cooldown_until = datetime('now', '-1 second') WHERE id = 'acc-1'`,
    ).run();
    // Surviving a restart is the whole reason this lives in the database
    // rather than the in-memory rateLimiter, so it is compared in SQL.
    expect(inCooldown('acc-1')).toBe(false);
  });

  it('refuses a cooldown window that is not a positive number', () => {
    seed();
    // datetime('now', 'NaN seconds') yields NULL — "no cooldown at all", the
    // exact opposite of every caller's intent. That must be a loud error, not
    // a silent green light to hammer Apple.
    expect(() => setCooldown('acc-1', Number.NaN, 'x')).toThrow(/positive number/);
    expect(() => setCooldown('acc-1', -1000, 'x')).toThrow(/positive number/);
    expect(inCooldown('acc-1')).toBe(false);
  });

  it('clears a cooldown on success', () => {
    seed();
    setCooldown('acc-1', 60_000, 'oops');
    clearCooldown('acc-1');
    expect(inCooldown('acc-1')).toBe(false);
  });

  it('honours a cooldown this process never set', () => {
    // Stands in for the restart the column exists to survive: the timestamp is
    // written directly, so no in-memory state in this process has ever seen
    // this account. An implementation that answered from a Map — even one that
    // mirrored every write to the database — reads free here and hammers Apple
    // on the first tick after every deploy. Only a SQL comparison against the
    // stored value can pass.
    getDb().prepare(
      `INSERT INTO icloud_accounts (id, apple_id, pool_cooldown_until)
       VALUES ('acc-restart', 'me@icloud.com', datetime('now', '+45 minutes'))`,
    ).run();
    expect(inCooldown('acc-restart')).toBe(true);
  });
});

function fakeClient(entries: Array<{ hme: string; anonymousId: string; label: string }>) {
  return {
    listWithForwarding: async () => ({
      hmeEmails: entries.map((e) => ({ ...e, note: '', isActive: true })),
      selectedForwardTo: 'owner@example.com',
      forwardToEmails: [],
    }),
  } as never;
}

describe('pool reconciliation', () => {
  it('adopts a marked address that never made it into the database', async () => {
    seed('acc-r1');
    // reserve succeeding and the insert failing leaves an address that
    // permanently consumes one of the 750 slots while being invisible locally.
    const result = await reconcileAccount('acc-r1', fakeClient([
      { hme: 'lost@icloud.com', anonymousId: 'anon-lost', label: MAILHUB_HME_LABEL },
    ]));

    expect(result.adopted).toEqual(['lost@icloud.com']);
    const row = getDb().prepare(
      `SELECT state FROM icloud_addresses WHERE hme = 'lost@icloud.com'`,
    ).get() as { state: string };
    expect(row.state).toBe('free');
  });

  it('refuses an unmarked address and reports it instead', async () => {
    seed('acc-r2');
    // /v2/hme/list returns every alias on the Apple ID, including ones the
    // owner made by hand for private use. Adopting one would hand their mail
    // — password resets included — to an unrelated tenant.
    const result = await reconcileAccount('acc-r2', fakeClient([
      { hme: 'personal@icloud.com', anonymousId: 'anon-personal', label: 'my bank' },
    ]));

    expect(result.adopted).toEqual([]);
    expect(result.unowned).toEqual(['personal@icloud.com']);
    const row = getDb().prepare(
      `SELECT COUNT(*) AS c FROM icloud_addresses WHERE hme = 'personal@icloud.com'`,
    ).get() as { c: number };
    expect(row.c).toBe(0);
  });

  it('does not duplicate an address already known locally', async () => {
    seed('acc-r3');
    getDb().prepare(
      `INSERT INTO icloud_addresses (hme, account_id, anonymous_id, state) VALUES ('known@icloud.com', 'acc-r3', 'anon-known', 'assigned')`,
    ).run();

    const result = await reconcileAccount('acc-r3', fakeClient([
      { hme: 'known@icloud.com', anonymousId: 'anon-known', label: MAILHUB_HME_LABEL },
    ]));

    expect(result.adopted).toEqual([]);
    // Re-adopting would reset a live assignment back to free and hand the
    // address to a second inbox.
    const row = getDb().prepare(
      `SELECT state FROM icloud_addresses WHERE hme = 'known@icloud.com'`,
    ).get() as { state: string };
    expect(row.state).toBe('assigned');
  });
});

describe('orphan reaping', () => {
  it('returns an address whose inbox no longer exists', () => {
    seed('acc-o1');
    getDb().prepare(
      `INSERT INTO icloud_addresses (hme, account_id, anonymous_id, state, assigned_inbox_id)
       VALUES ('orphan@icloud.com', 'acc-o1', 'anon-o', 'assigned', 'gone-forever')`,
    ).run();

    // A crash between claiming an address and persisting the inbox would
    // otherwise strand it as assigned for good, draining the pool one
    // interrupted request at a time.
    expect(reapOrphanedAddresses()).toBe(1);
    const row = getDb().prepare(
      `SELECT state, assigned_inbox_id FROM icloud_addresses WHERE hme = 'orphan@icloud.com'`,
    ).get() as { state: string; assigned_inbox_id: string | null };
    expect(row.state).toBe('free');
    expect(row.assigned_inbox_id).toBeNull();
  });

  it('leaves an address whose inbox is alive', () => {
    seed('acc-o2');
    getDb().prepare(
      `INSERT INTO inboxes (id, provider, address, auth_data, api_base, expires_at)
       VALUES ('live-inbox', 'icloud', 'held@icloud.com', '{}', '', datetime('now', '+1 day'))`,
    ).run();
    getDb().prepare(
      `INSERT INTO icloud_addresses (hme, account_id, anonymous_id, state, assigned_inbox_id)
       VALUES ('held@icloud.com', 'acc-o2', 'anon-h', 'assigned', 'live-inbox')`,
    ).run();

    expect(reapOrphanedAddresses()).toBe(0);
  });

  it('spares an address claimed moments ago whose inbox row is still coming', () => {
    seed('acc-fresh');
    getDb().prepare(
      `INSERT INTO icloud_addresses (hme, account_id, anonymous_id, state, assigned_inbox_id, assigned_at)
       VALUES ('fresh@icloud.com', 'acc-fresh', 'anon-f', 'assigned', 'being-written', datetime('now'))`,
    ).run();

    // The dispatcher claims first and writes the inbox row moments later.
    // Reaping inside that window returns a just-claimed address to the pool
    // and two live inboxes end up reading one alias.
    expect(reapOrphanedAddresses()).toBe(0);
    const row = getDb().prepare(
      `SELECT state FROM icloud_addresses WHERE hme = 'fresh@icloud.com'`,
    ).get() as { state: string };
    expect(row.state).toBe('assigned');
  });

  it('still reaps an orphan once it is old enough to be real', () => {
    seed('acc-stale');
    getDb().prepare(
      `INSERT INTO icloud_addresses (hme, account_id, anonymous_id, state, assigned_inbox_id, assigned_at)
       VALUES ('stale@icloud.com', 'acc-stale', 'anon-s', 'assigned', 'gone-forever', datetime('now', '-20 minutes'))`,
    ).run();

    expect(reapOrphanedAddresses()).toBe(1);
  });

  it('never touches a retired address', () => {
    seed('acc-o3');
    getDb().prepare(
      `INSERT INTO icloud_addresses (hme, account_id, anonymous_id, state, assigned_inbox_id)
       VALUES ('burned@icloud.com', 'acc-o3', 'anon-b', 'retired', 'gone')`,
    ).run();

    // Retired means the address was burned and deactivated at Apple. Recycling
    // it would hand a tenant an alias that no longer receives.
    expect(reapOrphanedAddresses()).toBe(0);
  });
});

describe('reconciliation and deactivated aliases', () => {
  it('takes a deactivated alias back as retired, not as a usable address', async () => {
    seed('acc-inactive');

    // A deactivated alias silently drops everything sent to it. Adopting one
    // into the free pool hands a tenant an address that can never receive, and
    // they sit waiting for a code that will not come.
    const result = await reconcileAccount('acc-inactive', {
      listWithForwarding: async () => ({
        hmeEmails: [{ hme: 'dead@icloud.com', anonymousId: 'anon-dead', label: MAILHUB_HME_LABEL, note: '', isActive: false }],
        selectedForwardTo: 'owner@example.com',
        forwardToEmails: [],
      }),
    } as never);

    expect(result.adopted).toEqual([]);
    const row = getDb().prepare(
      `SELECT state FROM icloud_addresses WHERE hme = 'dead@icloud.com'`,
    ).get() as { state: string };
    expect(row.state).toBe('retired');
  });

  it('does not split one Apple slot into two local addresses', async () => {
    seed('acc-dup');
    getDb().prepare(
      `INSERT INTO icloud_addresses (hme, account_id, anonymous_id) VALUES ('old-name@icloud.com', 'acc-dup', 'anon-same')`,
    ).run();

    // anonymousId is Apple's identity for the alias; hme is only our primary
    // key. Matching on the primary key alone would insert this as a brand-new
    // free address, so one Apple slot would show up as two local ones.
    const result = await reconcileAccount('acc-dup', {
      listWithForwarding: async () => ({
        hmeEmails: [{ hme: 'new-name@icloud.com', anonymousId: 'anon-same', label: MAILHUB_HME_LABEL, note: '', isActive: true }],
        selectedForwardTo: 'owner@example.com',
        forwardToEmails: [],
      }),
    } as never);

    expect(result.adopted).toEqual([]);
    const count = getDb().prepare(
      `SELECT COUNT(*) AS c FROM icloud_addresses WHERE account_id = 'acc-dup'`,
    ).get() as { c: number };
    expect(count.c).toBe(1);
  });
});
