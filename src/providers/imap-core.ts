import { ImapFlow, type SearchObject } from 'imapflow';
import { createHash } from 'crypto';
import type { Message, MessageDetail } from './base.js';
import { createLogger } from '../logger.js';
import { errorMessage, logIgnoredError } from '../errors.js';

const log = createLogger('imap-core');

/**
 * Everything a pooled IMAP connection needs, decoupled from any one table.
 *
 * `poolKey` is supplied by the caller rather than derived from an id: two
 * providers share this module's pool, and their account identifiers would
 * otherwise occupy the same key space. Callers namespace it — for example
 * `imap:${id}`, `icloud:${id}`, or `outlook:${email}`.
 */
export interface ImapCreds {
  poolKey: string;
  host: string;
  port: number;
  user: string;
  password?: string;
  accessToken?: string;
  tls: boolean;
  proxy?: string;
  connectionTimeout?: number;
  socketTimeout?: number;
}

// A busy catch-all mailbox can match hundreds of UIDs; poll only the newest.
export const POLL_FETCH_LIMIT = 20;

async function connect(creds: ImapCreds): Promise<ImapFlow> {
  const client = new ImapFlow({
    host: creds.host,
    port: creds.port,
    secure: creds.tls,
    auth: creds.accessToken
      ? { user: creds.user, accessToken: creds.accessToken }
      : { user: creds.user, pass: creds.password },
    ...(creds.proxy ? { proxy: creds.proxy } : {}),
    ...(creds.connectionTimeout ? { connectionTimeout: creds.connectionTimeout } : {}),
    ...(creds.socketTimeout ? { socketTimeout: creds.socketTimeout } : {}),
    logger: false,
  });
  try {
    await client.connect();
    return client;
  } catch (e) {
    client.close();
    throw e;
  }
}

/** The subset of imapflow's BODYSTRUCTURE tree this module needs. */
export interface BodyNode {
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

interface PoolEntry {
  clientPromise: Promise<ImapFlow>;
  timer: ReturnType<typeof setTimeout>;
  credentialFingerprint: string;
}
const pool = new Map<string, PoolEntry>();
const IDLE_MS = 5 * 60 * 1000;

function credentialFingerprint(creds: ImapCreds): string {
  return createHash('sha256')
    .update(JSON.stringify([
      creds.host,
      creds.port,
      creds.user,
      creds.password,
      creds.accessToken,
      creds.tls,
      creds.proxy,
      creds.connectionTimeout,
      creds.socketTimeout,
    ]))
    .digest('hex');
}

export function evictClient(poolKey: string, entry?: PoolEntry): void {
  const current = pool.get(poolKey);
  if (!current) return;
  // Entry-matched eviction: an async error callback must not kill a newer
  // client that has since replaced the failed one.
  if (entry && current !== entry) return;
  clearTimeout(current.timer);
  pool.delete(poolKey);
  current.clientPromise
    .then((client) => client.logout().catch(() => client.close()))
    .catch((error: unknown) => {
      logIgnoredError(log, 'IMAP pooled client logout failed', error, { poolKey });
    });
}

/**
 * Hand back the entry alongside the client.
 *
 * Callers need it to evict the connection they actually used. Evicting by key
 * alone races: between a request failing and its catch block running, another
 * caller can have replaced a dead entry with a fresh client, and the bare
 * eviction would log that newcomer out from under every request now sharing it.
 * A local failure (a malformed body, a missing part) would take down a healthy
 * connection the same way.
 */
async function acquire(creds: ImapCreds): Promise<{ client: ImapFlow; entry: PoolEntry }> {
  const fingerprint = credentialFingerprint(creds);
  const existing = pool.get(creds.poolKey);
  if (existing?.credentialFingerprint === fingerprint) {
    clearTimeout(existing.timer);
    existing.timer = setTimeout(() => evictClient(creds.poolKey, existing), IDLE_MS);
    return { client: await existing.clientPromise, entry: existing };
  }
  if (existing) evictClient(creds.poolKey, existing);
  // The entry is registered synchronously (holding a promise) so concurrent
  // callers share one connection instead of racing to open duplicates.
  const entry: PoolEntry = {
    clientPromise: connect(creds).then((client) => {
      client.once('error', () => evictClient(creds.poolKey, entry));
      return client;
    }),
    timer: setTimeout(() => evictClient(creds.poolKey, entry), IDLE_MS),
    credentialFingerprint: fingerprint,
  };
  pool.set(creds.poolKey, entry);
  try {
    return { client: await entry.clientPromise, entry };
  } catch (e) {
    if (pool.get(creds.poolKey) === entry) {
      clearTimeout(entry.timer);
      pool.delete(creds.poolKey);
    }
    throw e;
  }
}

/** Case-insensitive exact address match against the envelope recipients. */
function addressedTo(
  envelope: { to?: { address?: string }[]; cc?: { address?: string }[] } | undefined,
  recipient: string,
  strictRecipient = false,
): boolean {
  const all = [...(envelope?.to ?? []), ...(envelope?.cc ?? [])];
  // A catch-all mailbox fails open for Bcc-only mail. Personal forwarding
  // mailboxes fail closed because an unverifiable message may be private.
  if (all.length === 0) return !strictRecipient;
  const want = recipient.toLowerCase();
  return all.some((a) => (a.address ?? '').toLowerCase() === want);
}

/**
 * List the newest messages matching `criteria`.
 *
 * The criteria is a parameter rather than a fixed `{ to }` because two
 * providers sort the same shared mailbox by different recipient evidence:
 * a catch-all domain answers on To, and which header an iCloud alias
 * survives in is settled empirically (see scripts/verify-icloud-imap.ts).
 */
export async function fetchMessagesBySearch(
  creds: ImapCreds,
  criteria: SearchObject,
  opts: { limit?: number; recipient?: string; strictRecipient?: boolean; mailbox?: string } = {},
): Promise<Message[]> {
  const limit = opts.limit ?? POLL_FETCH_LIMIT;
  const { client, entry } = await acquire(creds);
  try {
    const lock = await client.getMailboxLock(opts.mailbox ?? 'INBOX', { readOnly: true });
    try {
      const uids = await client.search(criteria, { uid: true });
      if (!uids || uids.length === 0) return [];
      const recent = uids.slice(-limit);
      const messages: Message[] = [];
      for await (const fetched of client.fetch(recent, { envelope: true, internalDate: true }, { uid: true })) {
        // SEARCH narrows; this decides. TO is a substring match on the
        // envelope field, so the search alone would hand one tenant another
        // tenant's mail whenever one address is a prefix of the other.
        if (opts.recipient && !addressedTo(fetched.envelope, opts.recipient, opts.strictRecipient)) continue;
        messages.push({
          id: String(fetched.uid),
          from: fetched.envelope?.from?.[0]?.address ?? '',
          subject: fetched.envelope?.subject ?? '',
          excerpt: '',
          receivedAt: fetched.internalDate
            ? (fetched.internalDate instanceof Date ? fetched.internalDate.toISOString() : fetched.internalDate)
            : fetched.envelope?.date?.toISOString() ?? '',
        });
      }
      return messages;
    } finally {
      lock.release();
    }
  } catch (e) {
    evictClient(creds.poolKey, entry);
    throw e;
  }
}

/**
 * Read one message by UID.
 *
 * `opts.recipient` is not optional in spirit: a UID names a message in the
 * whole shared mailbox, and callers reach this with an id straight off the
 * request path. Without the check, a tenant who guesses a UID — they are small
 * sequential integers — reads a neighbour's body, verification code included,
 * even though the listing correctly hid it.
 */
export async function fetchMessageDetail(
  creds: ImapCreds,
  uid: string,
  opts: { recipient?: string; strictRecipient?: boolean; mailbox?: string } = {},
): Promise<MessageDetail> {
  const { client, entry } = await acquire(creds);
  try {
    const lock = await client.getMailboxLock(opts.mailbox ?? 'INBOX', { readOnly: true });
    try {
      const fetched = await client.fetchOne(uid, {
        uid: true,
        envelope: true,
        internalDate: true,
        bodyStructure: true,
      }, { uid: true });

      if (!fetched) throw new Error(`Message ${uid} not found`);

      // Same error as a genuinely absent UID, deliberately: distinguishing the
      // two would turn this endpoint into an oracle for which UIDs are live.
      //
      // `strictRecipient` decides what a message with no envelope recipients at
      // all means, and the right answer differs by whose mailbox is being read.
      // A catch-all domain mailbox exists only to serve these addresses, so
      // failing open there costs at most a stray while failing closed would
      // silently lose Bcc-only mail. An iCloud forwarding mailbox is the
      // operator's own personal inbox, where the same leniency hands a tenant
      // their private mail — so that caller asks to fail closed.
      if (opts.recipient) {
        if (!addressedTo(fetched.envelope, opts.recipient, opts.strictRecipient)) {
          throw new Error(`Message ${uid} not found`);
        }
      }

      const parts = selectBodyParts(fetched.bodyStructure as BodyNode | undefined);

      // download() applies the Content-Transfer-Encoding decoder, so
      // quoted-printable soft breaks cannot split a verification code the
      // way a raw toString() left them.
      const readPart = async (part: string): Promise<string> => {
        const { meta, content } = await client.download(uid, part, { uid: true });
        const chunks: Buffer[] = [];
        for await (const chunk of content) chunks.push(chunk as Buffer);
        return decodeBody(Buffer.concat(chunks), meta?.charset);
      };

      let text = '';
      let html = '';
      if (parts.text) {
        try { text = await readPart(parts.text); } catch (error) {
          log.warn('failed to read IMAP text body part', { poolKey: creds.poolKey, uid, part: parts.text, error: errorMessage(error) });
        }
      }
      if (parts.html) {
        try { html = await readPart(parts.html); } catch (error) {
          log.warn('failed to read IMAP html body part', { poolKey: creds.poolKey, uid, part: parts.html, error: errorMessage(error) });
        }
      }

      return {
        id: uid,
        from: fetched.envelope?.from?.[0]?.address ?? '',
        subject: fetched.envelope?.subject ?? '',
        excerpt: text.slice(0, 200),
        receivedAt: fetched.internalDate
          ? (fetched.internalDate instanceof Date ? fetched.internalDate.toISOString() : fetched.internalDate)
          : fetched.envelope?.date?.toISOString() ?? '',
        text: text || undefined,
        html: html || undefined,
      };
    } finally {
      lock.release();
    }
  } catch (e) {
    evictClient(creds.poolKey, entry);
    throw e;
  }
}

/**
 * Say what the server actually refused.
 *
 * imapflow's `message` is the bare word "Command failed" for every rejection,
 * which tells an operator nothing about whether their password is wrong, the
 * mailbox is missing, or the host refused the connection. The useful text is
 * on the error object next to it, so read it.
 */
export function describeImapError(error: unknown): string {
  const base = errorMessage(error);
  if (!error || typeof error !== 'object') return base;

  const e = error as {
    responseText?: string;
    responseStatus?: string;
    serverResponseCode?: string;
    authenticationFailed?: boolean;
    code?: string;
  };

  const detail = e.responseText?.trim();
  const parts = [detail && detail !== base ? detail : undefined];

  if (e.serverResponseCode) parts.push(`[${e.serverResponseCode}]`);
  if (e.code && e.code !== e.serverResponseCode) parts.push(`(${e.code})`);

  const described = parts.filter(Boolean).join(' ');
  if (!described) return base;

  // The server's own words first — "Command failed" adds nothing in front of
  // them, but the auth verdict is worth stating because it decides which
  // credential to go fix.
  return e.authenticationFailed ? `authentication rejected: ${described}` : described;
}

export function isImapAuthenticationError(error: unknown): boolean {
  return Boolean(error && typeof error === 'object' && (error as { authenticationFailed?: unknown }).authenticationFailed === true);
}

export async function findMailboxBySpecialUse(creds: ImapCreds, specialUse: string): Promise<string | undefined> {
  const { client, entry } = await acquire(creds);
  try {
    const mailboxes = await client.list();
    return mailboxes.find((mailbox) => mailbox.specialUse?.toLowerCase() === specialUse.toLowerCase())?.path;
  } catch (e) {
    evictClient(creds.poolKey, entry);
    throw e;
  }
}

export async function assertMailboxReadable(creds: ImapCreds, mailbox = 'INBOX'): Promise<void> {
  const client = await connect(creds);
  try {
    const lock = await client.getMailboxLock(mailbox, { readOnly: true });
    try {
      await client.search({ all: true }, { uid: true });
    } finally {
      lock.release();
    }
  } finally {
    await client.logout().catch(() => client.close());
  }
}

export async function testConnection(creds: ImapCreds): Promise<{ ok: boolean; error?: string }> {
  try {
    await assertMailboxReadable(creds);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: describeImapError(e) };
  }
}
