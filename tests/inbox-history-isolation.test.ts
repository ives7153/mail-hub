import { describe, expect, it } from 'vitest';
import { getDb } from '../src/db.js';
import { registry } from '../src/providers/registry.js';
import { app, authHeaders } from './helpers/http.js';
import { FakeProvider } from './helpers/fake-provider.js';
import { BaseProvider, type InboxData, type Message, type MessageDetail, type ProviderMeta } from '../src/providers/base.js';
import { parseInboxTimestamp } from '../src/inbox-lifecycle.js';

/**
 * Pool providers (Outlook 1:1-but-reused, YYDS, IMAP catch-all) hand out
 * mailboxes that already contain someone else's mail. The inbox row's
 * created_at is the only boundary between "this tenant's mail" and history, so
 * every user-facing message surface must apply it. /code already did;
 * /messages did not, and leaked the previous owner's mail into the UI.
 */

class HistoryProvider extends BaseProvider {
  meta: ProviderMeta = {
    name: 'history',
    displayName: 'History Mail',
    type: 'api',
    tier: 'free',
    trustLevel: 5,
    rateLimit: { createPerMinute: 60, pollPerMinute: 60 },
    retention: 'test',
    features: { customUsername: true, pollInbox: true, realtime: false, attachments: false },
  };

  constructor(private readonly messages: Message[]) {
    super();
  }

  async getDomains(): Promise<string[]> { return ['example.test']; }

  async createInbox(): Promise<InboxData> {
    return { address: 'pooled@example.test', authData: {}, provider: this.meta.name, apiBase: '' };
  }

  async getMessages(): Promise<Message[]> { return this.messages; }

  async getMessage(_inbox: InboxData, messageId: string): Promise<MessageDetail> {
    const found = this.messages.find((m) => m.id === messageId);
    if (!found) throw new Error(`no such message ${messageId}`);
    return { ...found, text: 'body' };
  }
}

function insertInbox(id: string, provider: string, createdAt: string): void {
  getDb().prepare(
    `INSERT INTO inboxes (id, provider, address, auth_data, api_base, created_at, status)
     VALUES (?, ?, 'pooled@example.test', '{}', '', ?, 'active')`,
  ).run(id, provider, createdAt);
}

const PREVIOUS_TENANT: Message = {
  id: 'old-1',
  from: 'newsletter@spam.test',
  subject: "theo, here's how teams use Quotient",
  excerpt: '',
  receivedAt: '2026-07-23T03:10:13Z',
};
const CURRENT_TENANT: Message = {
  id: 'new-1',
  from: 'Anthropic <no-reply@mail.anthropic.com>',
  subject: 'Your login code is 123456',
  excerpt: '',
  receivedAt: '2026-07-26T14:02:30Z',
};

describe('pooled inbox history isolation', () => {
  it('GET /messages hides mail that predates the inbox', async () => {
    registry.register(new HistoryProvider([CURRENT_TENANT, PREVIOUS_TENANT]));
    // created_at is stored by SQLite's datetime('now') — UTC, space-separated.
    insertInbox('iso-messages', 'history', '2026-07-26 14:02:18');

    const res = await app.request('/api/inbox/iso-messages/messages', { headers: authHeaders() });
    expect(res.status).toBe(200);
    const body = await res.json() as { messages: Message[] };

    expect(body.messages.map((m) => m.id)).toEqual(['new-1']);
    registry.unregister('history');
  });

  it('GET /messages/:mid refuses to open a pre-inbox message', async () => {
    registry.register(new HistoryProvider([CURRENT_TENANT, PREVIOUS_TENANT]));
    insertInbox('iso-detail', 'history', '2026-07-26 14:02:18');

    const leaked = await app.request('/api/inbox/iso-detail/messages/old-1', { headers: authHeaders() });
    expect(leaked.status).toBe(404);

    const own = await app.request('/api/inbox/iso-detail/messages/new-1', { headers: authHeaders() });
    expect(own.status).toBe(200);
    registry.unregister('history');
  });

  it('keeps messages whose timestamp is missing or unparseable', async () => {
    registry.register(new HistoryProvider([
      { id: 'no-date', from: 'a@b.test', subject: 'undated', excerpt: '', receivedAt: '' },
    ]));
    insertInbox('iso-undated', 'history', '2026-07-26 14:02:18');

    const res = await app.request('/api/inbox/iso-undated/messages', { headers: authHeaders() });
    const body = await res.json() as { messages: Message[] };

    expect(body.messages.map((m) => m.id)).toEqual(['no-date']);
    registry.unregister('history');
  });

  it('does not filter when created_at is unparseable', async () => {
    registry.register(new HistoryProvider([PREVIOUS_TENANT]));
    // created_at is NOT NULL, but a legacy/corrupt row can still hold a value
    // SQLite never produced. No boundary is knowable, so nothing is dropped.
    insertInbox('iso-nocreated', 'history', 'not-a-timestamp');

    const res = await app.request('/api/inbox/iso-nocreated/messages', { headers: authHeaders() });
    const body = await res.json() as { messages: Message[] };

    expect(body.messages.map((m) => m.id)).toEqual(['old-1']);
    registry.unregister('history');
  });

  it('still returns the inbox metadata envelope alongside filtered messages', async () => {
    const fake = new FakeProvider();
    registry.register(fake);
    insertInbox('iso-envelope', 'fake', '2020-01-01 00:00:00');

    const res = await app.request('/api/inbox/iso-envelope/messages', { headers: authHeaders() });
    const body = await res.json() as { messages: Message[]; status: string; address: string; provider: string };

    expect(body.status).toBe('active');
    expect(body.address).toBe('pooled@example.test');
    expect(body.provider).toBe('fake');
    expect(body.messages).toHaveLength(1);
    registry.unregister('fake');
  });
});

/**
 * created_at is written by SQLite as UTC 'YYYY-MM-DD HH:MM:SS'. `new Date()` on
 * that shape applies the LOCAL zone, so on any non-UTC host the boundary shifts
 * by the offset — east of UTC it lands early and history leaks through, west of
 * it the tenant's own mail is hidden. Production runs UTC, which is why /code's
 * existing filter appeared to work.
 */
describe('parseInboxTimestamp', () => {
  it('reads SQLite datetime() output as UTC regardless of host timezone', () => {
    expect(parseInboxTimestamp('2026-07-26 14:02:18')).toBe(Date.UTC(2026, 6, 26, 14, 2, 18));
  });

  it('still honours an explicit timezone offset', () => {
    expect(parseInboxTimestamp('2026-07-26T14:02:18Z')).toBe(Date.UTC(2026, 6, 26, 14, 2, 18));
    expect(parseInboxTimestamp('2026-07-26T23:02:18+09:00')).toBe(Date.UTC(2026, 6, 26, 14, 2, 18));
  });

  it('returns 0 for missing or unparseable input', () => {
    expect(parseInboxTimestamp(undefined)).toBe(0);
    expect(parseInboxTimestamp('')).toBe(0);
    expect(parseInboxTimestamp('not a date')).toBe(0);
  });
});
