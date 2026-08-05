import { describe, expect, it } from 'vitest';
import { getDb } from '../src/db.js';
import { registry } from '../src/providers/registry.js';
import { app, authHeaders } from './helpers/http.js';
import { FakeProvider } from './helpers/fake-provider.js';
import { BaseProvider, type InboxData, type Message, type MessageDetail, type ProviderMeta } from '../src/providers/base.js';
import { isMessageWithinInboxLifetime, parseInboxTimestamp } from '../src/inbox-lifecycle.js';

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

  constructor(
    private readonly messages: Message[],
    private readonly details: Record<string, Partial<MessageDetail>> = {},
  ) {
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
    return { ...found, text: 'body', ...this.details[messageId] };
  }
}

function insertInbox(
  id: string,
  provider: string,
  createdAt: string,
  opts: { closedAt?: string | null; expiresAt?: string | null; status?: string } = {},
): void {
  getDb().prepare(
    `INSERT INTO inboxes (id, provider, address, auth_data, api_base, created_at, closed_at, expires_at, status)
     VALUES (?, ?, 'pooled@example.test', '{}', '', ?, ?, ?, ?)`,
  ).run(id, provider, createdAt, opts.closedAt ?? null, opts.expiresAt ?? null, opts.status ?? 'active');
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

type ReadSurface = 'messages' | 'detail' | 'code';

async function readIds(surface: ReadSurface, inboxId: string, messageId: string): Promise<string[]> {
  const suffix = surface === 'messages' ? 'messages' : surface === 'detail' ? `messages/${messageId}` : 'code';
  const res = await app.request(`/api/inbox/${inboxId}/${suffix}`, { headers: authHeaders() });
  if (surface === 'messages') {
    expect(res.status).toBe(200);
    return ((await res.json()) as { messages: Message[] }).messages.map((message) => message.id);
  }
  if (surface === 'detail') return res.status === 200 ? [messageId] : [];
  expect(res.status).toBe(200);
  const body = await res.json() as { messageId: string | null };
  return body.messageId ? [body.messageId] : [];
}

describe.each<ReadSurface>(['messages', 'detail', 'code'])('%s lease end isolation', (surface) => {
  it('fails closed for the whole imprecise created_at second', async () => {
    const ambiguous: Message = {
      id: 'ambiguous-start', from: 'sender@test', subject: 'Code 101010', excerpt: '', receivedAt: '2026-07-26T14:02:00.100Z',
    };
    const safelyAfter: Message = {
      id: 'safe-start', from: 'sender@test', subject: 'Code 202020', excerpt: '', receivedAt: '2026-07-26T14:02:01.000Z',
    };
    registry.register(new HistoryProvider([ambiguous]));
    insertInbox(`strict-start-${surface}`, 'history', '2026-07-26 14:02:00', {
      expiresAt: '2099-01-01 00:00:00', status: 'active',
    });

    expect(await readIds(surface, `strict-start-${surface}`, 'ambiguous-start')).not.toContain('ambiguous-start');
    registry.register(new HistoryProvider([safelyAfter]));
    expect(await readIds(surface, `strict-start-${surface}`, 'safe-start')).toContain('safe-start');
    registry.unregister('history');
  });

  it.each([
    {
      label: 'closed_at',
      inbox: { closedAt: '2026-07-26 14:03:00', expiresAt: '2026-07-26 15:00:00', status: 'closed' },
    },
    {
      label: 'expires_at',
      inbox: { expiresAt: '2026-07-26 14:03:00', status: 'active' },
    },
  ])('rejects mail at or after $label without upper-bound slack', async ({ inbox }) => {
    const beforeEnd: Message = {
      id: 'before-end', from: 'sender@test', subject: 'Code 111111', excerpt: '', receivedAt: '2026-07-26T14:02:59Z',
    };
    const atEnd: Message = {
      id: 'at-end', from: 'sender@test', subject: 'Code 222222', excerpt: '', receivedAt: '2026-07-26T14:03:00Z',
    };
    registry.register(new HistoryProvider([beforeEnd]));
    insertInbox(`end-${surface}-${inbox.status}`, 'history', '2026-07-26 14:02:00', inbox);

    const before = await readIds(surface, `end-${surface}-${inbox.status}`, 'before-end');
    registry.register(new HistoryProvider([atEnd]));
    const after = await readIds(surface, `end-${surface}-${inbox.status}`, 'at-end');

    expect(before).toContain('before-end');
    expect(after).not.toContain('at-end');
    registry.unregister('history');
  });

  it('keeps adjacent lease windows non-overlapping at the handoff instant', async () => {
    const handoff: Message = {
      id: 'handoff', from: 'sender@test', subject: 'Code 333333', excerpt: '', receivedAt: '2026-07-26T14:03:00Z',
    };
    const safelyNext: Message = {
      id: 'safely-next', from: 'sender@test', subject: 'Code 343434', excerpt: '', receivedAt: '2026-07-26T14:03:01Z',
    };
    registry.register(new HistoryProvider([handoff]));
    insertInbox(`previous-${surface}`, 'history', '2026-07-26 14:02:00', {
      closedAt: '2026-07-26 14:03:00', status: 'closed',
    });
    insertInbox(`next-${surface}`, 'history', '2026-07-26 14:03:00', {
      expiresAt: '2099-01-01 00:00:00', status: 'active',
    });

    expect(await readIds(surface, `previous-${surface}`, 'handoff')).not.toContain('handoff');
    expect(await readIds(surface, `next-${surface}`, 'handoff')).not.toContain('handoff');
    registry.register(new HistoryProvider([safelyNext]));
    expect(await readIds(surface, `next-${surface}`, 'safely-next')).toContain('safely-next');
    registry.unregister('history');
  });

  it.each([
    { label: 'missing', receivedAt: '' },
    { label: 'unparseable', receivedAt: 'not-a-date' },
  ])('fails closed for $label message dates after a lease has ended', async ({ receivedAt }) => {
    const message: Message = {
      id: 'unknown-date', from: 'sender@test', subject: 'Code 444444', excerpt: '', receivedAt,
    };
    registry.register(new HistoryProvider([message]));
    insertInbox(`unknown-${surface}-${receivedAt || 'missing'}`, 'history', '2026-07-26 14:02:00', {
      closedAt: '2026-07-26 14:03:00', status: 'closed',
    });

    expect(await readIds(surface, `unknown-${surface}-${receivedAt || 'missing'}`, 'unknown-date'))
      .not.toContain('unknown-date');
    registry.unregister('history');
  });

  it('fails closed for legacy closed rows with no usable end timestamp', async () => {
    const message: Message = {
      id: 'legacy', from: 'sender@test', subject: 'Code 555555', excerpt: '', receivedAt: '2026-07-26T14:02:30Z',
    };
    registry.register(new HistoryProvider([message]));
    insertInbox(`legacy-${surface}`, 'history', '2026-07-26 14:02:00', { status: 'closed' });

    expect(await readIds(surface, `legacy-${surface}`, 'legacy')).not.toContain('legacy');
    registry.unregister('history');
  });
});

it('/code revalidates the detailed message timestamp before extracting a code', async () => {
  const summary: Message = {
    id: 'changed-date', from: 'sender@test', subject: 'Code 666666', excerpt: '', receivedAt: '2026-07-26T14:02:30Z',
  };
  registry.register(new HistoryProvider([summary], {
    'changed-date': { receivedAt: '2026-07-26T14:03:01Z', text: 'Code 666666' },
  }));
  insertInbox('detail-recheck', 'history', '2026-07-26 14:02:00', {
    closedAt: '2026-07-26 14:03:00', status: 'closed',
  });

  const res = await app.request('/api/inbox/detail-recheck/code', { headers: authHeaders() });
  const body = await res.json() as { codes: unknown[]; messageId: string | null };
  expect(body.codes).toEqual([]);
  expect(body.messageId).toBeNull();
  registry.unregister('history');
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

  it('keeps the legacy numeric helper contract strictly lower-bounded', () => {
    const start = Date.UTC(2026, 6, 26, 14, 2, 0);
    expect(isMessageWithinInboxLifetime('2026-07-26T14:01:59Z', start)).toBe(false);
    expect(isMessageWithinInboxLifetime('2026-07-26T14:02:00Z', start)).toBe(true);
  });
});
