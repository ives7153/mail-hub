import { registry } from './providers/registry.js';
import type { InboxData } from './providers/base.js';
import { createLogger } from './logger.js';
import { logIgnoredError } from './errors.js';

const log = createLogger('inbox-lifecycle');

export interface StoredInbox extends InboxData {
  id: string;
}

/**
 * SQLite writes created_at via datetime('now'), which is UTC in the shape
 * 'YYYY-MM-DD HH:MM:SS' — no zone marker. `new Date()` reads that as LOCAL
 * time, so on a host east of UTC the boundary lands earlier than the real
 * creation instant and history leaks through; west of UTC it lands later and
 * hides the tenant's own mail. Normalize to UTC before parsing. Returns 0 when
 * there is no usable timestamp, which callers treat as "no boundary known, do
 * not filter".
 */
export function parseInboxTimestamp(value: string | null | undefined): number {
  if (!value) return 0;
  const normalized = /^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}(\.\d+)?$/.test(value)
    ? `${value.replace(' ', 'T')}Z`
    : value;
  const ms = Date.parse(normalized);
  return Number.isFinite(ms) ? ms : 0;
}

/**
 * Pool providers reuse mailboxes that already hold the previous tenant's mail
 * (Outlook accounts are recycled, YYDS keys and IMAP catch-alls are shared), so
 * the inbox's own creation time is the only boundary between "mine" and
 * "history". The 60s slack absorbs clock skew between us and the mail host.
 *
 * Messages with a missing or unparseable timestamp are KEPT: dropping them
 * would silently lose real mail from providers with sloppy date fields, and a
 * visible stray beats a swallowed verification code.
 */
export function isMessageWithinInboxLifetime(receivedAt: string | undefined, inboxCreatedAtMs: number): boolean {
  if (!inboxCreatedAtMs) return true;
  if (!receivedAt) return true;
  const received = Date.parse(receivedAt);
  if (!Number.isFinite(received)) return true;
  return received >= inboxCreatedAtMs - 60000;
}

export function rowToInboxData(row: { address: string; auth_data: string; provider: string; api_base: string | null }): InboxData {
  return {
    address: row.address,
    authData: JSON.parse(row.auth_data),
    provider: row.provider,
    apiBase: row.api_base || '',
  };
}

export function parseStoredInbox(row: {
  id: string;
  provider: string;
  address: string;
  auth_data: string;
  api_base: string | null;
}): StoredInbox {
  return { id: row.id, ...rowToInboxData(row) };
}

export async function releaseInboxResources(
  inbox: StoredInbox,
  opts: { deleteExternal?: boolean } = {}
): Promise<void> {
  const provider = registry.get(inbox.provider);

  if (opts.deleteExternal) {
    await provider?.deleteInbox(inbox).catch((error: unknown) => {
      logIgnoredError(log, 'provider inbox deletion failed', error, { inboxId: inbox.id, provider: inbox.provider });
    });
  }

  await provider?.releaseInbox(inbox, inbox.id).catch((error: unknown) => {
    logIgnoredError(log, 'provider inbox release failed', error, { inboxId: inbox.id, provider: inbox.provider });
  });
}
