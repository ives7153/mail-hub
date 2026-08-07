import { beforeEach, describe, expect, it, vi } from 'vitest';

const imapState = vi.hoisted(() => ({
  connectFailure: '' as '' | 'auth' | 'network',
  rejectedAccessTokens: [] as string[],
  constructorOptions: [] as Array<Record<string, any>>,
  tokenCalls: 0,
  httpMailCalls: 0,
  listCalls: 0,
  locks: [] as Array<{ path: string; options: Record<string, unknown> | undefined; released: boolean }>,
  logoutCount: 0,
  closeCount: 0,
  mutationCalls: 0,
  folders: [
    { path: 'INBOX', specialUse: '\\Inbox', listed: true, subscribed: true, flags: new Set<string>(), delimiter: '/', name: 'INBOX', parent: [], parentPath: '' },
    { path: '垃圾邮件', specialUse: '\\Junk', listed: true, subscribed: true, flags: new Set<string>(), delimiter: '/', name: '垃圾邮件', parent: [], parentPath: '' },
  ],
  messages: {
    INBOX: [
      { uid: 7, subject: 'Inbox mail', receivedAt: '2026-08-07T10:00:00.000Z', body: 'inbox body' },
    ],
    '垃圾邮件': [
      { uid: 7, subject: 'Junk mail', receivedAt: '2026-08-07T11:00:00.000Z', body: 'junk body' },
    ],
  } as Record<string, Array<{ uid: number; subject: string; receivedAt: string; body: string }>>,
}));

vi.mock('imapflow', () => {
  class FakeImapFlow {
    private selected = 'INBOX';
    private readonly options: Record<string, any>;

    constructor(options: Record<string, any>) {
      this.options = options;
      imapState.constructorOptions.push(options);
    }

    async connect(): Promise<void> {
      if (imapState.connectFailure === 'auth' || imapState.rejectedAccessTokens.includes(this.options.auth?.accessToken)) {
        throw Object.assign(new Error('Command failed'), {
          authenticationFailed: true,
          responseText: 'AUTHENTICATE failed',
        });
      }
      if (imapState.connectFailure === 'network') {
        throw Object.assign(new Error('connect timed out'), { code: 'ETIMEDOUT' });
      }
    }

    once(): void {}

    async logout(): Promise<void> {
      imapState.logoutCount++;
    }

    close(): void {
      imapState.closeCount++;
    }

    async list(): Promise<typeof imapState.folders> {
      imapState.listCalls++;
      return imapState.folders;
    }

    async mailboxOpen(path: string): Promise<void> {
      this.selected = path;
    }

    async getMailboxLock(path: string, options?: Record<string, unknown>): Promise<{ release(): void }> {
      this.selected = path;
      const lock = { path, options, released: false };
      imapState.locks.push(lock);
      return { release: () => { lock.released = true; } };
    }

    async search(): Promise<number[]> {
      return (imapState.messages[this.selected] || []).map((message) => message.uid);
    }

    async *fetch(range: number[]): AsyncGenerator<unknown> {
      for (const uid of range) {
        const message = (imapState.messages[this.selected] || []).find((candidate) => candidate.uid === uid);
        if (!message) continue;
        yield {
          uid: message.uid,
          internalDate: new Date(message.receivedAt),
          envelope: {
            from: [{ name: 'Sender', address: 'sender@example.test' }],
            subject: message.subject,
            date: new Date(message.receivedAt),
          },
        };
      }
    }

    async fetchOne(uid: string): Promise<unknown> {
      const message = (imapState.messages[this.selected] || []).find((candidate) => candidate.uid === Number(uid));
      if (!message) return false;
      return {
        uid: message.uid,
        internalDate: new Date(message.receivedAt),
        envelope: {
          from: [{ name: 'Sender', address: 'sender@example.test' }],
          subject: message.subject,
          date: new Date(message.receivedAt),
        },
        bodyStructure: { type: 'text/plain' },
      };
    }

    async download(uid: string): Promise<{ meta: { charset: string }; content: AsyncIterable<Buffer> }> {
      const message = (imapState.messages[this.selected] || []).find((candidate) => candidate.uid === Number(uid));
      if (!message) throw new Error('message not found');
      return {
        meta: { charset: 'utf-8' },
        content: (async function* () { yield Buffer.from(message.body); })(),
      };
    }

    async messageFlagsAdd(): Promise<void> { imapState.mutationCalls++; }
    async messageMove(): Promise<void> { imapState.mutationCalls++; }
    async messageDelete(): Promise<void> { imapState.mutationCalls++; }
  }

  return { ImapFlow: FakeImapFlow };
});

import { getDb } from '../src/db.js';
import { checkToken, OutlookProvider } from '../src/providers/outlook.js';
import { app, jsonHeaders } from './helpers/http.js';

function insertAccount(email: string, apiType = 'imap'): void {
  getDb().prepare(
    `INSERT INTO outlook_accounts (email, password, client_id, refresh_token, token_status, api_type)
     VALUES (?, 'pw', 'client-id', ?, 'valid', ?)`,
  ).run(email, `refresh-${email}`, apiType);
}

function stubMicrosoftHttp(opts: { rotatedRefreshToken?: string; accessTokens?: string[] } = {}): void {
  vi.stubGlobal('fetch', vi.fn(async (url: string) => {
    const target = String(url);
    if (target.includes('/oauth2/v2.0/token')) {
      const accessToken = opts.accessTokens?.[imapState.tokenCalls] || opts.accessTokens?.at(-1) || 'opaque-access-token';
      imapState.tokenCalls++;
      return new Response(JSON.stringify({
        access_token: accessToken,
        expires_in: 3600,
        ...(opts.rotatedRefreshToken ? { refresh_token: opts.rotatedRefreshToken } : {}),
      }), {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
      });
    }
    if (target.includes('graph.microsoft.com') || target.includes('outlook.office.com/api/')) {
      imapState.httpMailCalls++;
      return new Response('{}', { status: 401 });
    }
    return new Response('{}', { status: 404 });
  }));
}

function inbox(email: string) {
  return {
    address: email,
    authData: { email, password: 'pw', clientId: 'client-id', refreshToken: `refresh-${email}` },
    provider: 'outlook',
    apiBase: '',
  };
}

beforeEach(() => {
  vi.unstubAllGlobals();
  imapState.connectFailure = '';
  imapState.rejectedAccessTokens = [];
  imapState.constructorOptions = [];
  imapState.tokenCalls = 0;
  imapState.httpMailCalls = 0;
  imapState.listCalls = 0;
  imapState.locks = [];
  imapState.logoutCount = 0;
  imapState.closeCount = 0;
  imapState.mutationCalls = 0;
});

describe('Outlook IMAP XOAUTH2 mailbox transport', () => {
  it('reads Inbox and special-use Junk directly with collision-free ids', async () => {
    const email = 'known-imap@outlook.com';
    insertAccount(email);
    stubMicrosoftHttp();

    const provider = new OutlookProvider();
    const messages = await provider.getMessages(inbox(email));

    expect(messages.map((message) => message.subject)).toEqual(['Junk mail', 'Inbox mail']);
    expect(new Set(messages.map((message) => message.id)).size).toBe(2);
    expect(imapState.httpMailCalls).toBe(0);
    expect(imapState.listCalls).toBe(1);
    expect(imapState.constructorOptions[0]?.auth).toEqual({
      user: email,
      accessToken: 'opaque-access-token',
    });
    expect(imapState.constructorOptions[0]).toMatchObject({
      connectionTimeout: 10000,
      socketTimeout: 30000,
    });
    expect(imapState.locks.map((lock) => [lock.path, lock.options])).toEqual([
      ['INBOX', { readOnly: true }],
      ['垃圾邮件', { readOnly: true }],
    ]);
    expect(imapState.locks.every((lock) => lock.released)).toBe(true);
    expect(imapState.mutationCalls).toBe(0);
  });

  it('uses the folder encoded in the message id when reading detail', async () => {
    const email = 'detail-imap@outlook.com';
    insertAccount(email);
    stubMicrosoftHttp();
    const provider = new OutlookProvider();
    const messages = await provider.getMessages(inbox(email));
    const junk = messages.find((message) => message.subject === 'Junk mail');

    const detail = await provider.getMessage(inbox(email), junk!.id);

    expect(detail.subject).toBe('Junk mail');
    expect(detail.text).toBe('junk body');
    expect(imapState.locks.at(-1)).toMatchObject({ path: '垃圾邮件', released: true });
    expect(imapState.mutationCalls).toBe(0);
  });

  it.each([
    ['a mailbox outside Inbox and special-use Junk', 'Sent Items', '7'],
    ['an IMAP UID range', 'INBOX', '1:*'],
  ])('rejects a forged message id containing %s', async (_description, mailbox, uid) => {
    const email = 'folder-boundary@outlook.com';
    insertAccount(email);
    stubMicrosoftHttp();
    const forgedId = `imap:${Buffer.from(JSON.stringify([mailbox, uid])).toString('base64url')}`;

    await expect(new OutlookProvider().getMessage(inbox(email), forgedId)).rejects.toThrow('Invalid Outlook IMAP message id');

    expect(imapState.locks).toHaveLength(0);
    expect(imapState.mutationCalls).toBe(0);
  });

  it.each([
    ['empty', ''],
    ['unrecognized', 'legacy-api'],
  ])('discovers IMAP while polling an account whose api type is %s', async (label, apiType) => {
    const email = `unknown-${label}-imap@outlook.com`;
    insertAccount(email, apiType);
    stubMicrosoftHttp();

    const messages = await new OutlookProvider().getMessages(inbox(email));

    expect(messages.map((message) => message.subject)).toEqual(['Junk mail', 'Inbox mail']);
    // Each HTTP transport reads Inbox and Junk as one capability attempt.
    expect(imapState.httpMailCalls).toBe(4);
    expect(getDb().prepare(`SELECT api_type FROM outlook_accounts WHERE email = ?`).get(email)).toEqual({ api_type: 'imap' });
  });

  it('persists a refresh token rotated during ordinary mailbox polling', async () => {
    const email = 'poll-rotation@outlook.com';
    insertAccount(email);
    stubMicrosoftHttp({ rotatedRefreshToken: 'rotated-during-poll' });

    await new OutlookProvider().getMessages(inbox(email));

    expect(getDb().prepare(
      `SELECT refresh_token, token_renewed_at FROM outlook_accounts WHERE email = ?`,
    ).get(email)).toMatchObject({
      refresh_token: 'rotated-during-poll',
      token_renewed_at: expect.any(String),
    });
  });

  it('evicts a rejected cached access token and retries IMAP once with a fresh token', async () => {
    const email = 'retry-imap@outlook.com';
    insertAccount(email);
    imapState.rejectedAccessTokens = ['stale-access-token'];
    stubMicrosoftHttp({ accessTokens: ['stale-access-token', 'fresh-access-token'] });

    const messages = await new OutlookProvider().getMessages(inbox(email));

    expect(messages).toHaveLength(2);
    expect(imapState.tokenCalls).toBe(2);
    expect(imapState.constructorOptions.map((options) => options.auth.accessToken)).toEqual([
      'stale-access-token',
      'fresh-access-token',
    ]);
    expect(imapState.closeCount).toBe(1);
  });

  it('evicts the rotated credential cache key before retrying a rejected access token', async () => {
    const email = 'retry-rotated-imap@outlook.com';
    insertAccount(email);
    imapState.rejectedAccessTokens = ['stale-rotated-access-token'];
    stubMicrosoftHttp({
      rotatedRefreshToken: 'rotated-before-retry',
      accessTokens: ['stale-rotated-access-token', 'fresh-rotated-access-token'],
    });

    const messages = await new OutlookProvider().getMessages(inbox(email));

    expect(messages).toHaveLength(2);
    expect(imapState.tokenCalls).toBe(2);
    expect(imapState.constructorOptions.map((options) => options.auth.accessToken)).toEqual([
      'stale-rotated-access-token',
      'fresh-rotated-access-token',
    ]);
    expect(getDb().prepare(`SELECT refresh_token FROM outlook_accounts WHERE email = ?`).get(email)).toEqual({
      refresh_token: 'rotated-before-retry',
    });
  });
});

describe('Outlook capability detection', () => {
  it('classifies an opaque Thunderbird token as valid IMAP after both HTTP APIs reject it', async () => {
    const email = 'detect-imap@outlook.com';
    insertAccount(email, '');
    stubMicrosoftHttp();

    await expect(checkToken(email, 'client-id', `refresh-${email}`)).resolves.toEqual({
      status: 'valid',
      apiType: 'imap',
    });
  });

  it('persists valid plus imap through the /outlook/check route', async () => {
    const email = 'route-check-imap@outlook.com';
    insertAccount(email, '');
    getDb().prepare(`UPDATE outlook_accounts SET token_status = '' WHERE email = ?`).run(email);
    stubMicrosoftHttp();

    const response = await app.request('/api/outlook/check', {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({ emails: [email] }),
    });

    expect(response.status).toBe(200);
    await expect(response.json()).resolves.toMatchObject({
      valid: 1,
      invalid: 0,
      unknown: 0,
      results: [{ email, valid: true, status: 'valid', apiType: 'imap' }],
    });
    expect(getDb().prepare(
      `SELECT token_status, api_type FROM outlook_accounts WHERE email = ?`,
    ).get(email)).toEqual({ token_status: 'valid', api_type: 'imap' });
  });

  it('refreshes and rechecks a cached access token before /outlook/check can mark the account invalid', async () => {
    const email = 'route-check-stale-cache@outlook.com';
    insertAccount(email, 'imap');
    stubMicrosoftHttp({ accessTokens: ['stale-check-access-token', 'fresh-check-access-token'] });
    await new OutlookProvider().getMessages(inbox(email));
    imapState.rejectedAccessTokens = ['stale-check-access-token'];

    const response = await app.request('/api/outlook/check', {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({ emails: [email] }),
    });

    expect(response.status).toBe(200);
    await expect(response.json()).resolves.toMatchObject({
      valid: 1,
      invalid: 0,
      unknown: 0,
      results: [{ email, valid: true, status: 'valid', apiType: 'imap' }],
    });
    expect(imapState.tokenCalls).toBe(2);
    expect(getDb().prepare(
      `SELECT token_status, api_type FROM outlook_accounts WHERE email = ?`,
    ).get(email)).toEqual({ token_status: 'valid', api_type: 'imap' });
  });

  it.each([
    ['temporary IMAP connection failure is inconclusive', 'network', 'unknown'],
    ['explicit rejection by every transport is invalid', 'auth', 'invalid'],
  ] as const)('%s', async (_name, connectFailure, expected) => {
    const email = `${connectFailure}@outlook.com`;
    insertAccount(email, '');
    imapState.connectFailure = connectFailure;
    stubMicrosoftHttp();

    const result = await checkToken(email, 'client-id', `refresh-${email}`);

    expect(result).toEqual({ status: expected, apiType: '' });
    expect(imapState.constructorOptions).toHaveLength(1);
  });

  it('treats a malformed successful token response as unknown, not invalid', async () => {
    const email = 'malformed-token@outlook.com';
    insertAccount(email, '');
    vi.stubGlobal('fetch', vi.fn(async (url: string) => {
      if (String(url).includes('/oauth2/v2.0/token')) {
        return new Response('{}', { status: 200, headers: { 'Content-Type': 'application/json' } });
      }
      throw new Error('mail APIs must not be called without an access token');
    }));

    await expect(checkToken(email, 'client-id', `refresh-${email}`)).resolves.toEqual({
      status: 'unknown',
      apiType: '',
    });
    expect(imapState.constructorOptions).toHaveLength(0);
  });
});

describe('Outlook token renewal capability validation', () => {
  it('records valid IMAP only after the rotated token can read the mailbox', async () => {
    const email = 'renew-imap@outlook.com';
    insertAccount(email, '');
    stubMicrosoftHttp({ rotatedRefreshToken: 'rotated-refresh-token' });

    const response = await app.request('/api/outlook/renew', {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({ emails: [email] }),
    });

    expect(response.status).toBe(200);
    await expect(response.json()).resolves.toMatchObject({
      results: [{ email, renewed: true, status: 'renewed', apiType: 'imap' }],
    });
    expect(getDb().prepare(
      `SELECT refresh_token, token_status, api_type FROM outlook_accounts WHERE email = ?`,
    ).get(email)).toEqual({
      refresh_token: 'rotated-refresh-token',
      token_status: 'valid',
      api_type: 'imap',
    });
  });

  it.each([
    ['temporary IMAP failure keeps the prior status', 'network', 'unknown', 'valid'],
    ['explicit rejection by all transports marks the account invalid', 'auth', 'invalid', 'invalid'],
  ] as const)('%s', async (_name, failure, expectedResult, expectedStoredStatus) => {
    const email = `renew-${failure}@outlook.com`;
    insertAccount(email, '');
    imapState.connectFailure = failure;
    stubMicrosoftHttp({ rotatedRefreshToken: `rotated-${failure}` });

    const response = await app.request('/api/outlook/renew', {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({ emails: [email] }),
    });
    const body = await response.json() as { results: Array<{ status: string }> };

    expect(body.results[0].status).toBe(expectedResult);
    expect(getDb().prepare(
      `SELECT refresh_token, token_status FROM outlook_accounts WHERE email = ?`,
    ).get(email)).toEqual({
      refresh_token: `rotated-${failure}`,
      token_status: expectedStoredStatus,
    });
  });
});
