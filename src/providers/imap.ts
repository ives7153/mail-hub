import { ImapFlow } from 'imapflow';
import { BaseProvider, PROVIDER, type InboxData, type Message, type MessageDetail } from './base.js';
import { allRows, getDb, getRow } from '../db.js';
import { randomString } from '../utils.js';
import { createLogger } from '../logger.js';
import { errorMessage, logIgnoredError } from '../errors.js';

const log = createLogger('imap');

interface ImapAccount {
  id: string;
  host: string;
  port: number;
  user: string;
  password: string;
  domain: string;
  tls: number;
  status: string;
}

function getActiveAccounts(): ImapAccount[] {
  return allRows<ImapAccount>(
    getDb(),
    `SELECT id, host, port, user, password, domain, tls, status FROM imap_accounts WHERE status = 'active'`,
  );
}

function getAccountById(id: string): ImapAccount | undefined {
  return getRow<ImapAccount>(
    getDb(),
    `SELECT id, host, port, user, password, domain, tls, status FROM imap_accounts WHERE id = ? AND status = 'active'`,
    id,
  );
}

function getAccountByDomain(domain: string): ImapAccount | undefined {
  return getRow<ImapAccount>(
    getDb(),
    `SELECT id, host, port, user, password, domain, tls, status FROM imap_accounts WHERE domain = ? AND status = 'active' LIMIT 1`,
    domain,
  );
}

async function connectImap(account: ImapAccount): Promise<ImapFlow> {
  const client = new ImapFlow({
    host: account.host,
    port: account.port,
    secure: account.tls === 1,
    auth: { user: account.user, pass: account.password },
    logger: false,
  });
  await client.connect();
  return client;
}

/** The subset of imapflow's BODYSTRUCTURE tree this module needs. */
interface BodyNode {
  part?: string;
  type?: string;
  disposition?: string;
  parameters?: { charset?: string };
  childNodes?: BodyNode[];
}

/**
 * Resolve which body parts actually hold the displayable text/html.
 *
 * Part numbers cannot be assumed: '1'/'2' only line up for a flat
 * multipart/alternative. A single-part message has no numbered children,
 * and under multipart/mixed the text lives at '1.1'/'1.2' while '2' is an
 * attachment. Worse, asking for a part that does not exist fails the whole
 * FETCH rather than just that part, so the structure must be read first.
 */
export function selectBodyParts(root: BodyNode | undefined): { text?: string; html?: string } {
  if (!root) return {};

  // Non-multipart message: RFC 3501 numbers the whole body as part 1.
  if (!root.childNodes?.length) {
    const type = root.type ?? '';
    if (type === 'text/html') return { html: '1' };
    if (type.startsWith('text/')) return { text: '1' };
    return {};
  }

  let text: string | undefined;
  let html: string | undefined;
  const walk = (node: BodyNode): void => {
    for (const child of node.childNodes ?? []) {
      if (child.childNodes?.length) {
        walk(child);
        continue;
      }
      // An attachment is not the message body even when it is text/*.
      if (child.disposition === 'attachment') continue;
      if (!text && child.type === 'text/plain') text = child.part;
      if (!html && child.type === 'text/html') html = child.part;
    }
  };
  walk(root);
  return { text, html };
}

/** Decode a body buffer using the part's declared charset, not a blind utf8 cast. */
export function decodeBody(buf: Buffer, charset?: string): string {
  if (!buf.length) return '';
  const label = (charset || 'utf-8').trim();
  try {
    return new TextDecoder(label).decode(buf);
  } catch {
    // Unknown/unsupported label — utf8 is the least-bad fallback.
    return buf.toString('utf8');
  }
}

interface PoolEntry { clientPromise: Promise<ImapFlow>; timer: ReturnType<typeof setTimeout>; }
const pool = new Map<string, PoolEntry>();
const IDLE_MS = 5 * 60 * 1000;
// A busy catch-all mailbox can match hundreds of UIDs; poll only the newest.
const POLL_FETCH_LIMIT = 20;

function evictClient(id: string, entry?: PoolEntry): void {
  const current = pool.get(id);
  if (!current) return;
  // Entry-matched eviction: an async error callback must not kill a newer
  // client that has since replaced the failed one.
  if (entry && current !== entry) return;
  clearTimeout(current.timer);
  pool.delete(id);
  current.clientPromise
    .then((client) => client.logout())
    .catch((error: unknown) => {
      logIgnoredError(log, 'IMAP pooled client logout failed', error, { accountId: id });
    });
}

async function getPooledClient(account: ImapAccount): Promise<ImapFlow> {
  const existing = pool.get(account.id);
  if (existing) {
    clearTimeout(existing.timer);
    existing.timer = setTimeout(() => evictClient(account.id, existing), IDLE_MS);
    return existing.clientPromise;
  }
  // The entry is registered synchronously (holding a promise) so concurrent
  // callers share one connection instead of racing to open duplicates.
  const entry: PoolEntry = {
    clientPromise: connectImap(account).then((client) => {
      client.once('error', () => evictClient(account.id, entry));
      return client;
    }),
    timer: setTimeout(() => evictClient(account.id, entry), IDLE_MS),
  };
  pool.set(account.id, entry);
  try {
    return await entry.clientPromise;
  } catch (e) {
    if (pool.get(account.id) === entry) {
      clearTimeout(entry.timer);
      pool.delete(account.id);
    }
    throw e;
  }
}

export class ImapProvider extends BaseProvider {
  meta = {
    name: PROVIDER.IMAP,
    displayName: 'IMAP / 域名邮箱',
    type: 'api' as const,
    tier: 'free' as const,
    trustLevel: 10,
    rateLimit: { createPerMinute: 60, pollPerMinute: 10 },
    retention: '24h',
    features: {
      customUsername: true,
      pollInbox: true,
      realtime: false,
      attachments: true,
    },
  };

  async getDomains(): Promise<string[]> {
    const accounts = getActiveAccounts();
    return [...new Set(accounts.map((a) => a.domain))];
  }

  async createInbox(opts?: { domain?: string; username?: string }): Promise<InboxData> {
    let account: ImapAccount | undefined;

    if (opts?.domain) {
      account = getAccountByDomain(opts.domain);
    }

    if (!account) {
      const accounts = getActiveAccounts();
      if (accounts.length === 0) throw new Error('No active IMAP accounts configured');
      account = accounts[Math.floor(Math.random() * accounts.length)];
    }

    const username = opts?.username || randomString(12);

    return {
      address: `${username}@${account.domain}`,
      authData: {
        imapAccountId: account.id,
        username,
        domain: account.domain,
      },
      provider: this.meta.name,
      apiBase: `imap://${account.host}`,
    };
  }

  async getMessages(inbox: InboxData): Promise<Message[]> {
    const account = getAccountById(inbox.authData.imapAccountId);
    if (!account) throw new Error(`IMAP account ${inbox.authData.imapAccountId} not found`);

    const client = await getPooledClient(account);
    try {
      const lock = await client.getMailboxLock('INBOX');
      try {
        const toAddr = inbox.address;
        const uids = await client.search({ to: toAddr }, { uid: true });
        if (!uids || uids.length === 0) return [];
        const recent = uids.slice(-POLL_FETCH_LIMIT);
        const messages: Message[] = [];
        for await (const fetched of client.fetch(recent, { envelope: true }, { uid: true })) {
          messages.push({
            id: String(fetched.uid),
            from: fetched.envelope?.from?.[0]?.address ?? '',
            subject: fetched.envelope?.subject ?? '',
            excerpt: '',
            receivedAt: fetched.envelope?.date?.toISOString() ?? '',
          });
        }
        return messages;
      } finally {
        lock.release();
      }
    } catch (e) {
      evictClient(account.id);
      throw e;
    }
  }

  async getMessage(inbox: InboxData, messageId: string): Promise<MessageDetail> {
    const account = getAccountById(inbox.authData.imapAccountId);
    if (!account) throw new Error(`IMAP account ${inbox.authData.imapAccountId} not found`);

    const client = await getPooledClient(account);
    try {
      const lock = await client.getMailboxLock('INBOX');
      try {
        const fetched = await client.fetchOne(messageId, {
          uid: true,
          envelope: true,
          bodyStructure: true,
        }, { uid: true });

        if (!fetched) throw new Error(`Message ${messageId} not found`);

        const parts = selectBodyParts(fetched.bodyStructure as BodyNode | undefined);

        // download() applies the Content-Transfer-Encoding decoder, so
        // quoted-printable soft breaks cannot split a verification code the
        // way a raw toString() left them.
        const readPart = async (part: string): Promise<string> => {
          const { meta, content } = await client.download(messageId, part, { uid: true });
          const chunks: Buffer[] = [];
          for await (const chunk of content) chunks.push(chunk as Buffer);
          return decodeBody(Buffer.concat(chunks), meta?.charset);
        };

        let text = '';
        let html = '';
        if (parts.text) {
          try { text = await readPart(parts.text); } catch (error) {
            log.warn('failed to read IMAP text body part', { accountId: account.id, messageId, part: parts.text, error: errorMessage(error) });
          }
        }
        if (parts.html) {
          try { html = await readPart(parts.html); } catch (error) {
            log.warn('failed to read IMAP html body part', { accountId: account.id, messageId, part: parts.html, error: errorMessage(error) });
          }
        }

        return {
          id: messageId,
          from: fetched.envelope?.from?.[0]?.address ?? '',
          subject: fetched.envelope?.subject ?? '',
          excerpt: text.slice(0, 200),
          receivedAt: fetched.envelope?.date?.toISOString() ?? '',
          text: text || undefined,
          html: html || undefined,
        };
      } finally {
        lock.release();
      }
    } catch (e) {
      evictClient(account.id);
      throw e;
    }
  }
}

export async function testImapConnection(account: ImapAccount): Promise<{ ok: boolean; error?: string }> {
  let client: ImapFlow;
  try {
    client = await connectImap(account);
  } catch (e) {
    return { ok: false, error: errorMessage(e) };
  }
  try {
    await client.mailboxOpen('INBOX');
    return { ok: true };
  } catch (e) {
    return { ok: false, error: errorMessage(e) };
  } finally {
    await client.logout().catch((error: unknown) => {
      logIgnoredError(log, 'IMAP test logout failed', error, { accountId: account.id });
    });
  }
}
