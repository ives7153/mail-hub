export const PROVIDER = {
  OUTLOOK: 'outlook',
  YYDS: 'yyds',
  IMAP: 'imap',
  MAILTM: 'mailtm',
  MAILGW: 'mailgw',
  TEMPMAIL_LOL: 'tempmail-lol',
  TEMPMAIL_ING: 'tempmail-ing',
} as const;

export type ProviderName = typeof PROVIDER[keyof typeof PROVIDER];

export interface ProviderMeta {
  name: string;
  displayName: string;
  type: 'api' | 'alias';
  tier: 'free' | 'paid';
  trustLevel: number;
  rateLimit: {
    createPerMinute: number;
    pollPerMinute: number;
  };
  retention: string;
  features: {
    customUsername: boolean;
    pollInbox: boolean;
    realtime: boolean;
    attachments: boolean;
    /**
     * Provider can hand out a sub-address (e.g. Outlook plus addressing) on
     * request. Optional so providers that cannot do it need no change.
     */
    alias?: boolean;
  };
}

export interface InboxData {
  address: string;
  authData: Record<string, string>;
  provider: string;
  apiBase: string;
  expiresAt?: string;
}

export interface Message {
  id: string;
  from: string;
  subject: string;
  excerpt: string;
  receivedAt: string;
}

export interface MessageDetail extends Message {
  text?: string;
  html?: string;
}

export type ProviderDomainMode = 'endpoint' | 'static' | 'from_create';

export abstract class BaseProvider {
  abstract meta: ProviderMeta;

  getDomainMode(): ProviderDomainMode {
    return 'endpoint';
  }

  /**
   * `alias` mirrors the create request: a provider that scopes its candidate
   * accounts by prior use must not exclude them when the caller will be handed
   * a fresh sub-address, or dispatch would reject the request before
   * createInbox ever runs.
   */
  abstract getDomains(opts?: { for?: string; alias?: boolean }): Promise<string[]>;
  abstract createInbox(opts?: { domain?: string; username?: string; for?: string; subdomain?: string; inboxId?: string; alias?: boolean }): Promise<InboxData>;
  abstract getMessages(inbox: InboxData): Promise<Message[]>;
  abstract getMessage(inbox: InboxData, messageId: string): Promise<MessageDetail>;

  async deleteInbox(_inbox: InboxData): Promise<void> {}
  async releaseInbox(_inbox: InboxData, _inboxId?: string): Promise<void> {}
}
