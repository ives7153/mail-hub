import { createHash, randomUUID } from 'crypto';
import { BaseProvider, PROVIDER, type InboxData, type Message, type MessageDetail, type ProviderMeta } from './base.js';
import { allRows, getDb, getRow } from '../db.js';
import { createConnection } from 'net';
import { fetchWithTimeout, formatSender, randomString } from '../utils.js';
import { errorMessage, UpstreamHttpError } from '../errors.js';

const OAUTH2_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
const GRAPH_INBOX_URL = 'https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages';
const GRAPH_JUNK_URL = 'https://graph.microsoft.com/v1.0/me/mailFolders/junkemail/messages';
const OUTLOOK_INBOX_URL = 'https://outlook.office.com/api/v2.0/me/mailfolders/inbox/messages';
const OUTLOOK_JUNK_URL = 'https://outlook.office.com/api/v2.0/me/mailfolders/junkemail/messages';
const IMAP_HOST = 'outlook.office365.com';
const IMAP_PORT = 993;

const TOKEN_TTL = 55 * 60 * 1000;
const tokenCache = new Map<string, { token: string; expiresAt: number }>();

interface GraphMessage {
  id: string;
  subject?: string;
  from?: { emailAddress?: { name?: string; address?: string } };
  receivedDateTime?: string;
  body?: { content?: string; contentType?: string };
  bodyPreview?: string;
}

interface OAuthResponse {
  access_token?: string;
  refresh_token?: string;
  error?: string;
  error_description?: string;
}

/**
 * The OAuth endpoint deterministically rejected the credentials (bad/expired
 * refresh token, revoked consent). Distinct from network errors, throttling
 * and 5xx, which say nothing about token validity.
 */
export class OAuthRejectedError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'OAuthRejectedError';
  }
}

export type TokenCheckStatus = 'valid' | 'invalid' | 'unknown';

interface CountRow { c: number }

const ALLOCABLE_ACCOUNT_WHERE = `assigned_inbox_id IS NULL
  AND client_id != ''
  AND refresh_token != ''
  AND COALESCE(token_status, '') NOT IN ('invalid', 'no_token', 'pending_oauth')`;

/**
 * Two accounts in the pool routinely share a client_id, so the refresh token is
 * the only thing telling their cached access tokens apart — and an access token
 * IS the mailbox. This used to key on `refreshToken.slice(-8)`, which throws
 * away all but 8 characters: two accounts whose tokens end alike collide, and
 * the second one is served the first one's mail. Real Microsoft tokens make that
 * astronomically unlikely, but the truncation bought nothing (this is an
 * in-memory map key, not storage), and "unlikely cross-account mail mixing" is
 * not a property worth keeping. Hash the whole token.
 */
function cacheKey(clientId: string, refreshToken: string): string {
  return `${clientId}:${createHash('sha256').update(refreshToken).digest('hex')}`;
}

/**
 * `local+tag@domain` → `local@domain`. Plus-addressed mail is delivered to the
 * base mailbox, so the account row is always keyed on the stripped form. Any
 * lookup that treats an inbox address as an account email must go through this
 * (or, better, read `authData.email`, which holds the account identity
 * verbatim). Returns the input unchanged when there is no tag.
 */
export function stripPlusTag(address: string): string {
  const at = address.lastIndexOf('@');
  if (at <= 0) return address;
  const local = address.slice(0, at);
  const plus = local.indexOf('+');
  if (plus < 0) return address;
  return local.slice(0, plus) + address.slice(at);
}

// RFC 5321 caps the local part at 64 octets; base + '+' + tag must fit.
const MAX_LOCAL_PART = 64;

/**
 * Microsoft accepts any legal SMTP local-part characters after the '+', but a
 * tag that reaches a signup form should be boring: letters, digits, dash,
 * underscore and dot only. Anything else is dropped rather than escaped, so a
 * caller-supplied tag can never produce an unroutable address.
 */
function sanitizeTag(raw: string): string {
  return raw.toLowerCase().replace(/[^a-z0-9._-]/g, '').replace(/^[.\-_]+/, '').slice(0, 24);
}

/**
 * Builds `local+tag@domain`, truncating the tag (never the base) if the local
 * part would exceed the RFC limit. Returns null when no usable tag survives, so
 * the caller falls back to the plain account address instead of shipping a
 * malformed one.
 */
export function buildAliasAddress(accountEmail: string, tag: string): string | null {
  const at = accountEmail.lastIndexOf('@');
  if (at <= 0) return null;
  const base = accountEmail.slice(0, at);
  const domain = accountEmail.slice(at);
  const budget = MAX_LOCAL_PART - base.length - 1;
  if (budget < 1) return null;
  const clean = sanitizeTag(tag).slice(0, budget);
  if (!clean) return null;
  return `${base}+${clean}${domain}`;
}

function getCachedToken(clientId: string, refreshToken: string): string | null {
  const entry = tokenCache.get(cacheKey(clientId, refreshToken));
  if (entry && Date.now() < entry.expiresAt) return entry.token;
  return null;
}

function setCachedToken(clientId: string, refreshToken: string, token: string): void {
  tokenCache.set(cacheKey(clientId, refreshToken), { token, expiresAt: Date.now() + TOKEN_TTL });
}

export function evictCachedToken(clientId: string, refreshToken: string): void {
  tokenCache.delete(cacheKey(clientId, refreshToken));
}

/** Test hook: module-level cache must not leak between test cases. */
export function resetTokenCache(): void {
  tokenCache.clear();
}

async function fetchOAuthToken(clientId: string, refreshToken: string): Promise<{ accessToken: string; newRefreshToken?: string }> {
  const body = new URLSearchParams({
    client_id: clientId,
    refresh_token: refreshToken,
    grant_type: 'refresh_token',
  });
  const res = await fetchWithTimeout(OAUTH2_URL, {
    timeout: 10000,
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString(),
  });
  const data = await res.json().catch(() => ({})) as OAuthResponse;
  if (!res.ok) {
    const detail = [data.error, data.error_description].filter(Boolean).join(': ') || `HTTP ${res.status}`;
    // 400/401/403 carry an OAuth error verdict; anything else (429/5xx) is infrastructure.
    if (res.status === 400 || res.status === 401 || res.status === 403) {
      throw new OAuthRejectedError(`OAuth refresh rejected: ${detail}`);
    }
    throw new UpstreamHttpError(`OAuth token endpoint error: ${detail}`, res.status, res.headers.get('Retry-After'));
  }
  if (!data.access_token) throw new OAuthRejectedError('OAuth response missing access_token');
  return { accessToken: data.access_token, newRefreshToken: data.refresh_token };
}

async function obtainAccessToken(clientId: string, refreshToken: string): Promise<string> {
  const cached = getCachedToken(clientId, refreshToken);
  if (cached) return cached;
  const result = await fetchOAuthToken(clientId, refreshToken);
  setCachedToken(clientId, refreshToken, result.accessToken);
  return result.accessToken;
}

async function fetchMailsGraph(accessToken: string, folderUrl: string, count = 20): Promise<GraphMessage[]> {
  const params = new URLSearchParams({
    $top: String(count),
    $orderby: 'receivedDateTime desc',
    $select: 'id,subject,from,receivedDateTime,body,bodyPreview',
  });
  const res = await fetchWithTimeout(`${folderUrl}?${params}`, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (res.status === 401) throw new Error('API 401');
  if (!res.ok) return [];
  const data = await res.json() as { value?: GraphMessage[] };
  // Normalize immediately: the Outlook REST API returns PascalCase fields, and
  // downstream dedup/sort must never see un-normalized ids/timestamps.
  return (data.value || []).map(normalizeMessage);
}

async function fetchMailsBothApis(accessToken: string, apiType: string, count = 20): Promise<{ messages: GraphMessage[]; apiType: string }> {
  if (apiType === 'outlook') {
    const [inboxMsgs, junkMsgs] = await Promise.all([
      fetchMailsGraph(accessToken, OUTLOOK_INBOX_URL, count),
      fetchMailsGraph(accessToken, OUTLOOK_JUNK_URL, count),
    ]);
    return { messages: mergeMessages(inboxMsgs, junkMsgs, count), apiType: 'outlook' };
  }
  try {
    const [inboxMsgs, junkMsgs] = await Promise.all([
      fetchMailsGraph(accessToken, GRAPH_INBOX_URL, count),
      fetchMailsGraph(accessToken, GRAPH_JUNK_URL, count),
    ]);
    return { messages: mergeMessages(inboxMsgs, junkMsgs, count), apiType: 'graph' };
  } catch (e) {
    if (errorMessage(e).includes('401')) {
      const [inboxMsgs, junkMsgs] = await Promise.all([
        fetchMailsGraph(accessToken, OUTLOOK_INBOX_URL, count),
        fetchMailsGraph(accessToken, OUTLOOK_JUNK_URL, count),
      ]);
      return { messages: mergeMessages(inboxMsgs, junkMsgs, count), apiType: 'outlook' };
    }
    throw e;
  }
}

function mergeMessages(inboxMsgs: GraphMessage[], junkMsgs: GraphMessage[], limit = 20): GraphMessage[] {
  const merged = new Map<string, GraphMessage>();
  for (const m of [...inboxMsgs, ...junkMsgs]) {
    if (!merged.has(m.id)) merged.set(m.id, m);
  }
  return [...merged.values()]
    .sort((a, b) => (b.receivedDateTime || '').localeCompare(a.receivedDateTime || ''))
    .slice(0, limit);
}

async function fetchSingleMessage(accessToken: string, messageId: string, apiType: string): Promise<GraphMessage> {
  const urls = apiType === 'outlook'
    ? [`https://outlook.office.com/api/v2.0/me/messages/${messageId}?$select=id,subject,from,receivedDateTime,body,bodyPreview`]
    : [
        `https://graph.microsoft.com/v1.0/me/messages/${messageId}?$select=id,subject,from,receivedDateTime,body,bodyPreview`,
        `https://outlook.office.com/api/v2.0/me/messages/${messageId}?$select=id,subject,from,receivedDateTime,body,bodyPreview`,
      ];
  for (const url of urls) {
    const res = await fetchWithTimeout(url, { headers: { Authorization: `Bearer ${accessToken}` } });
    if (res.status === 401) throw new Error('API 401');
    if (res.ok) return normalizeMessage(await res.json());
  }
  throw new Error('无法获取邮件详情');
}

function normalizeMessage(msg: any): GraphMessage {
  return {
    id: msg.id || msg.Id || '',
    subject: msg.subject || msg.Subject || '',
    from: msg.from || msg.From ? {
      emailAddress: {
        name: (msg.from?.emailAddress || msg.From?.EmailAddress)?.name || (msg.from?.emailAddress || msg.From?.EmailAddress)?.Name || '',
        address: (msg.from?.emailAddress || msg.From?.EmailAddress)?.address || (msg.from?.emailAddress || msg.From?.EmailAddress)?.Address || '',
      }
    } : undefined,
    receivedDateTime: msg.receivedDateTime || msg.ReceivedDateTime || '',
    body: msg.body || msg.Body ? {
      content: (msg.body || msg.Body)?.content || (msg.body || msg.Body)?.Content || '',
      contentType: ((msg.body || msg.Body)?.contentType || (msg.body || msg.Body)?.ContentType || '').toLowerCase(),
    } : undefined,
    bodyPreview: msg.bodyPreview || msg.BodyPreview || '',
  };
}

function graphMsgToMessage(normalized: GraphMessage): Message {
  return {
    id: normalized.id,
    from: formatSender(normalized.from?.emailAddress || {}),
    subject: normalized.subject || '',
    excerpt: normalized.bodyPreview || '',
    receivedAt: normalized.receivedDateTime || '',
  };
}

function graphMsgToDetail(normalized: GraphMessage): MessageDetail {
  const bodyObj = normalized.body || {};
  const content = bodyObj.content || '';
  const isHtml = bodyObj.contentType === 'html';
  return {
    ...graphMsgToMessage(normalized),
    text: isHtml ? '' : content,
    html: isHtml ? content : '',
  };
}

/**
 * The account identity behind an inbox. `authData.email` holds it verbatim;
 * the fallback exists for rows written before that field, where `address` was
 * always the bare account email — strip any tag so an alias inbox can never
 * resolve to a non-existent account row.
 */
function accountEmailOf(inbox: InboxData): string {
  return inbox.authData.email || stripPlusTag(inbox.address);
}

/**
 * One access-token attempt with a single retry after evicting the cache on 401,
 * which is the only failure a fresh token can fix. Both the inbox path and the
 * account-mailbox path funnel through here so the retry, the api_type probe and
 * its write-back cannot drift apart.
 */
async function withAccessToken<T>(clientId: string, refreshToken: string, run: (token: string) => Promise<T>): Promise<T> {
  const accessToken = await obtainAccessToken(clientId, refreshToken);
  try {
    return await run(accessToken);
  } catch (e) {
    if (!errorMessage(e).includes('401')) throw e;
    evictCachedToken(clientId, refreshToken);
    return run(await obtainAccessToken(clientId, refreshToken));
  }
}

async function pollMailbox(email: string, clientId: string, refreshToken: string, limit: number): Promise<Message[]> {
  const db = getDb();
  const apiType = getRow<{ api_type: string }>(db, `SELECT api_type FROM outlook_accounts WHERE email = ?`, email)?.api_type || '';
  return withAccessToken(clientId, refreshToken, async (token) => {
    const result = await fetchMailsBothApis(token, apiType, limit);
    if (result.apiType && result.apiType !== apiType) {
      db.prepare(`UPDATE outlook_accounts SET api_type = ? WHERE email = ?`).run(result.apiType, email);
    }
    return result.messages.map(graphMsgToMessage);
  });
}

async function readMailboxMessage(email: string, clientId: string, refreshToken: string, messageId: string): Promise<MessageDetail> {
  const apiType = getRow<{ api_type: string }>(getDb(), `SELECT api_type FROM outlook_accounts WHERE email = ?`, email)?.api_type || '';
  return withAccessToken(clientId, refreshToken, async (token) => graphMsgToDetail(await fetchSingleMessage(token, messageId, apiType)));
}

function accountCredentials(email: string): { clientId: string; refreshToken: string } {
  const row = getRow<{ client_id: string; refresh_token: string }>(
    getDb(),
    `SELECT client_id, refresh_token FROM outlook_accounts WHERE email = ?`,
    email,
  );
  if (!row) throw new Error(`Outlook 账号不存在: ${email}`);
  if (!row.client_id || !row.refresh_token) throw new Error(`Outlook 账号 ${email} 缺少令牌凭据`);
  return { clientId: row.client_id, refreshToken: row.refresh_token };
}

/**
 * The whole mailbox behind an account, with no inbox lease in the picture.
 * An inbox is a lease over this mailbox, so its message list is deliberately
 * clipped to the lease window (isMessageWithinInboxLifetime) — that boundary is
 * a tenant boundary and must never be widened. Seeing what else is in the
 * mailbox is a different question, asked of the account, and answerable only to
 * an admin. Bounded by `limit` because Graph pages and we do not.
 */
export async function fetchAccountMailbox(email: string, limit = 50): Promise<Message[]> {
  const { clientId, refreshToken } = accountCredentials(email);
  return pollMailbox(email, clientId, refreshToken, limit);
}

export async function fetchAccountMessage(email: string, messageId: string): Promise<MessageDetail> {
  const { clientId, refreshToken } = accountCredentials(email);
  return readMailboxMessage(email, clientId, refreshToken, messageId);
}

export class OutlookProvider extends BaseProvider {
  meta: ProviderMeta = {
    name: PROVIDER.OUTLOOK,
    displayName: 'Outlook',
    type: 'api',
    tier: 'paid',
    trustLevel: 4,
    rateLimit: { createPerMinute: 60, pollPerMinute: 30 },
    retention: 'Permanent',
    features: {
      // The account's local part is fixed, so an arbitrary username is not
      // possible — only a `+tag` suffix on that fixed base. Advertising
      // customUsername would promise something this provider cannot do.
      customUsername: false,
      pollInbox: true,
      realtime: false,
      attachments: true,
      alias: true,
    },
  };

  private getFreshRefreshToken(email: string): string | null {
    const row = getRow<{ refresh_token: string }>(getDb(), `SELECT refresh_token FROM outlook_accounts WHERE email = ?`, email);
    return row?.refresh_token || null;
  }

  async getDomains(opts?: { for?: string; alias?: boolean }): Promise<string[]> {
    const db = getDb();
    let whereClauses = ALLOCABLE_ACCOUNT_WHERE;
    const params: unknown[] = [];
    if (opts?.for && !opts.alias) {
      whereClauses += ` AND (used_services IS NULL OR used_services NOT LIKE ?)`;
      params.push(`%"${opts.for.replace(/"/g, '\\"')}"%`);
    }
    const rows = allRows<{ domain: string }>(db,
      `SELECT DISTINCT SUBSTR(email, INSTR(email, '@') + 1) as domain
       FROM outlook_accounts WHERE ${whereClauses}`,
      ...params,
    );
    return rows.map((r) => r.domain);
  }

  async createInbox(opts?: { domain?: string; for?: string; inboxId?: string; alias?: boolean }): Promise<InboxData> {
    const db = getDb();
    const inboxId = opts?.inboxId ?? `pending-${randomUUID()}`;

    let whereClauses = ALLOCABLE_ACCOUNT_WHERE;
    const selectParams: unknown[] = [];
    if (opts?.domain) {
      whereClauses += ` AND email LIKE ?`;
      selectParams.push(`%@${opts.domain}`);
    }
    // used_services is the anti-reuse blacklist for the ACCOUNT's own address.
    // A fresh alias is a new address at the target service, which is the point
    // of asking for one, so an aliased request may reuse a burned account. The
    // record is still written on report, so a later PLAIN request for that
    // service still finds the account excluded.
    if (opts?.for && !opts.alias) {
      whereClauses += ` AND (used_services IS NULL OR used_services NOT LIKE ?)`;
      selectParams.push(`%"${opts.for.replace(/"/g, '\\"')}"%`);
    }
    const params: unknown[] = [inboxId, ...selectParams];

    const sql = `UPDATE outlook_accounts SET assigned_inbox_id = ?
      WHERE email = (
        SELECT email FROM outlook_accounts
        WHERE ${whereClauses}
        ORDER BY CASE WHEN token_status = 'valid' THEN 0 ELSE 1 END, created_at ASC
        LIMIT 1
      ) AND assigned_inbox_id IS NULL
      RETURNING email, password, client_id, refresh_token`;

    const allocate = db.transaction(() => {
      const row = db.prepare(sql).get(...params) as { email: string; password: string; client_id: string; refresh_token: string } | undefined;

      if (!row) {
        const total = getRow<CountRow>(db, `SELECT COUNT(*) AS c FROM outlook_accounts`)?.c ?? 0;
        const invalid = getRow<CountRow>(db, `SELECT COUNT(*) AS c FROM outlook_accounts WHERE token_status = 'invalid'`)?.c ?? 0;
        const pending = getRow<CountRow>(db, `SELECT COUNT(*) AS c FROM outlook_accounts WHERE token_status IN ('pending_oauth', 'no_token') OR client_id = '' OR refresh_token = ''`)?.c ?? 0;
        const assigned = getRow<CountRow>(db, `SELECT COUNT(*) AS c FROM outlook_accounts WHERE assigned_inbox_id IS NOT NULL AND COALESCE(token_status, '') NOT IN ('invalid', 'no_token', 'pending_oauth')`)?.c ?? 0;
        const available = getRow<CountRow>(db, `SELECT COUNT(*) AS c FROM outlook_accounts WHERE ${whereClauses}`, ...selectParams)?.c ?? 0;
        const valid = available;
        const parts: string[] = [`共${total}个账号`];
        if (invalid > 0) parts.push(`${invalid}个无效`);
        if (pending > 0) parts.push(`${pending}个待补全`);
        if (assigned > 0) parts.push(`${assigned}个已分配`);
        if (valid > 0 && opts?.for) parts.push(`剩余${valid}个均已用于 ${opts.for}`);
        if (valid === 0 && !opts?.for) parts.push(`无空闲账号`);
        throw new Error(`Outlook 账号池中无可用账号 (${parts.join(', ')})`);
      }

      const { email, password, client_id: clientId, refresh_token: refreshToken } = row;
      if (!clientId || !refreshToken) {
        throw new Error(`Outlook 账号 ${email} 缺少令牌凭据`);
      }

      // Plus addressing is opt-in per request: the caller asks for an alias and
      // the tag is generated here, so no caller has to invent one or can probe
      // for which tags exist. The account still serves exactly one inbox, so the
      // tag buys a distinct address at the target service, not extra pool
      // capacity — and nothing has to sort shared mail by recipient.
      // `authData.email` stays the ACCOUNT address: every credential lookup
      // (refresh token, api_type, used_services) keys off it, and only
      // `address` carries the tag.
      const aliasAddress = opts?.alias ? buildAliasAddress(email, randomString(8)) : null;

      return {
        address: aliasAddress ?? email,
        authData: { email, password, clientId, refreshToken },
        provider: this.meta.name,
        apiBase: '',
      };
    });

    return allocate();
  }

  async getMessages(inbox: InboxData): Promise<Message[]> {
    const email = accountEmailOf(inbox);
    const freshToken = this.getFreshRefreshToken(email) || inbox.authData.refreshToken;
    return pollMailbox(email, inbox.authData.clientId, freshToken, 20);
  }

  async getMessage(inbox: InboxData, messageId: string): Promise<MessageDetail> {
    const email = accountEmailOf(inbox);
    const freshToken = this.getFreshRefreshToken(email) || inbox.authData.refreshToken;
    return readMailboxMessage(email, inbox.authData.clientId, freshToken, messageId);
  }

  async deleteInbox(inbox: InboxData): Promise<void> {
    const db = getDb();
    db.prepare(`UPDATE outlook_accounts SET assigned_inbox_id = NULL WHERE email = ?`).run(inbox.authData.email);
  }

  async releaseInbox(inbox: InboxData, inboxId: string): Promise<void> {
    const email = accountEmailOf(inbox);
    getDb().prepare(
      `UPDATE outlook_accounts SET assigned_inbox_id = NULL WHERE assigned_inbox_id = ? OR email = ?`
    ).run(inboxId, email);
  }
}

/**
 * Never throws. 'invalid' only on a deterministic OAuth rejection or definitive
 * 401/403 from both mail APIs; network errors, throttling and 5xx yield
 * 'unknown' so callers never destroy accounts over an infrastructure blip.
 */
export async function checkToken(_email: string, clientId: string, refreshToken: string): Promise<{ status: TokenCheckStatus; apiType: string }> {
  let token: string;
  try {
    token = await obtainAccessToken(clientId, refreshToken);
  } catch (e) {
    return { status: e instanceof OAuthRejectedError ? 'invalid' : 'unknown', apiType: '' };
  }

  let inconclusive = false;
  for (const [url, apiType] of [[GRAPH_INBOX_URL, 'graph'], [OUTLOOK_INBOX_URL, 'outlook']] as const) {
    try {
      const res = await fetchWithTimeout(`${url}?$top=1`, {
        headers: { Authorization: `Bearer ${token}` },
      });
      if (res.ok) return { status: 'valid', apiType };
      if (res.status !== 401 && res.status !== 403) inconclusive = true;
    } catch {
      inconclusive = true;
    }
  }
  return { status: inconclusive ? 'unknown' : 'invalid', apiType: '' };
}

/**
 * Returns the rotated credentials, or null when the endpoint accepted the token
 * but did not rotate it. Throws OAuthRejectedError on deterministic rejection
 * and UpstreamHttpError/network errors on infrastructure failure — callers
 * must only mark accounts invalid on OAuthRejectedError.
 */
export async function renewToken(clientId: string, refreshToken: string): Promise<{ newRefreshToken: string; accessToken: string } | null> {
  const result = await fetchOAuthToken(clientId, refreshToken);
  if (!result.newRefreshToken) return null;
  return { newRefreshToken: result.newRefreshToken, accessToken: result.accessToken };
}
