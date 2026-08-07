import type { Message, MessageDetail } from './base.js';
import {
  assertMailboxReadable,
  fetchMessageDetail,
  fetchMessagesBySearch,
  findMailboxBySpecialUse,
  type ImapCreds,
} from './imap-core.js';

const IMAP_HOST = 'outlook.office365.com';
const IMAP_PORT = 993;
const IMAP_ID_PREFIX = 'imap:';

export function outlookImapCreds(email: string, accessToken: string, proxy?: string): ImapCreds {
  return {
    poolKey: `outlook:${email}`,
    host: IMAP_HOST,
    port: IMAP_PORT,
    user: email,
    accessToken,
    tls: true,
    ...(proxy ? { proxy } : {}),
    connectionTimeout: 10000,
    socketTimeout: 30000,
  };
}

function encodeMessageId(mailbox: string, uid: string): string {
  return `${IMAP_ID_PREFIX}${Buffer.from(JSON.stringify([mailbox, uid])).toString('base64url')}`;
}

function decodeMessageId(messageId: string): { mailbox: string; uid: string } {
  if (!messageId.startsWith(IMAP_ID_PREFIX)) throw new Error('Invalid Outlook IMAP message id');
  try {
    const value = JSON.parse(Buffer.from(messageId.slice(IMAP_ID_PREFIX.length), 'base64url').toString('utf8')) as unknown;
    if (
      !Array.isArray(value)
      || value.length !== 2
      || typeof value[0] !== 'string'
      || typeof value[1] !== 'string'
      || !/^[1-9]\d{0,9}$/.test(value[1])
      || Number(value[1]) > 0xffffffff
    ) {
      throw new Error('invalid shape');
    }
    return { mailbox: value[0], uid: value[1] };
  } catch {
    throw new Error('Invalid Outlook IMAP message id');
  }
}

export async function fetchOutlookImapMessages(creds: ImapCreds, limit: number): Promise<Message[]> {
  const junk = await findMailboxBySpecialUse(creds, '\\Junk');
  const mailboxes = ['INBOX', ...(junk && junk.toUpperCase() !== 'INBOX' ? [junk] : [])];
  const merged: Message[] = [];
  for (const mailbox of mailboxes) {
    const messages = await fetchMessagesBySearch(creds, { all: true }, { mailbox, limit });
    merged.push(...messages.map((message) => ({
      ...message,
      id: encodeMessageId(mailbox, message.id),
    })));
  }
  return merged
    .sort((a, b) => (b.receivedAt || '').localeCompare(a.receivedAt || ''))
    .slice(0, limit);
}

export async function fetchOutlookImapMessage(creds: ImapCreds, messageId: string): Promise<MessageDetail> {
  const { mailbox, uid } = decodeMessageId(messageId);
  if (mailbox !== 'INBOX') {
    const junk = await findMailboxBySpecialUse(creds, '\\Junk');
    if (!junk || mailbox !== junk) throw new Error('Invalid Outlook IMAP message id');
  }
  const message = await fetchMessageDetail(creds, uid, { mailbox });
  return { ...message, id: messageId };
}

export async function checkOutlookImap(creds: ImapCreds): Promise<void> {
  await assertMailboxReadable(creds);
}
