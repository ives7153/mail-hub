import { describe, it, expect } from 'vitest';
import { getDb } from '../src/db.js';

describe('iCloud schema', () => {
  it('stores an account with its IMAP read-side columns', () => {
    const db = getDb();
    db.prepare(
      `INSERT INTO icloud_accounts (id, apple_id, region, auth_mode, cookies, imap_user, imap_password)
       VALUES ('acc-1', 'me@icloud.com', 'global', 'cookie', 'X-APPLE=1', 'me@icloud.com', 'app-specific')`,
    ).run();

    const row = db.prepare(
      `SELECT apple_id, region, imap_host, imap_port, imap_tls, status FROM icloud_accounts WHERE id = 'acc-1'`,
    ).get() as { apple_id: string; region: string; imap_host: string; imap_port: number; imap_tls: number; status: string };

    expect(row.apple_id).toBe('me@icloud.com');
    expect(row.region).toBe('global');
    // Defaults exist so an operator only supplies the app-specific password.
    expect(row.imap_host).toBe('imap.mail.me.com');
    expect(row.imap_port).toBe(993);
    expect(row.imap_tls).toBe(1);
    expect(row.status).toBe('active');
  });

  it('rejects two addresses sharing one anonymous_id', () => {
    const db = getDb();
    db.prepare(`INSERT INTO icloud_accounts (id, apple_id) VALUES ('acc-2', 'me@icloud.com')`).run();
    db.prepare(
      `INSERT INTO icloud_addresses (hme, account_id, anonymous_id) VALUES ('a@icloud.com', 'acc-2', 'anon-1')`,
    ).run();

    // Reconciliation adopts by anonymous_id; without this constraint a lost
    // reserve response and a later adoption would create two rows for one
    // Apple-side address.
    expect(() =>
      db.prepare(
        `INSERT INTO icloud_addresses (hme, account_id, anonymous_id) VALUES ('b@icloud.com', 'acc-2', 'anon-1')`,
      ).run(),
    ).toThrow(/UNIQUE/);
  });

  it('defaults a new address to the free pool', () => {
    const db = getDb();
    db.prepare(`INSERT INTO icloud_accounts (id, apple_id) VALUES ('acc-3', 'me@icloud.com')`).run();
    db.prepare(
      `INSERT INTO icloud_addresses (hme, account_id, anonymous_id) VALUES ('c@icloud.com', 'acc-3', 'anon-2')`,
    ).run();

    const row = db.prepare(
      `SELECT state, use_count, assigned_inbox_id FROM icloud_addresses WHERE hme = 'c@icloud.com'`,
    ).get() as { state: string; use_count: number; assigned_inbox_id: string | null };

    expect(row.state).toBe('free');
    expect(row.use_count).toBe(0);
    expect(row.assigned_inbox_id).toBeNull();
  });
});
