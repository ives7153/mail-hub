/**
 * Manual verification: does an iCloud Hide My Email alias survive into a header
 * IMAP can search?
 *
 * The entire read path assumes it does. Nothing else in the design can be
 * trusted until this runs green against a real account, and this script is
 * excluded from `npm test` because it needs live credentials.
 *
 * The host is an argument because Hide My Email forwards to whatever address
 * the account owner verified — frequently not an Apple mailbox at all, in which
 * case the alias has to be found in that provider's IMAP instead.
 *
 * Usage:
 *   $secure = Read-Host 'IMAP app password' -AsSecureString
 *   $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
 *   try {
 *     $env:MAILHUB_IMAP_PASSWORD = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
 *     npx tsx scripts/verify-icloud-imap.ts --user <imap-user> --hme <hme-address> [--host <host>]
 *   } finally {
 *     [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
 *     Remove-Item Env:MAILHUB_IMAP_PASSWORD -ErrorAction SilentlyContinue
 *   }
 *
 * Send a message to <hme-address> first, then run it.
 */
import { ImapFlow, type SearchObject } from 'imapflow';

function usage(): never {
  console.error('usage: npx tsx scripts/verify-icloud-imap.ts --user <imap-user> --hme <hme-address> [--host <host>]');
  console.error('provide MAILHUB_IMAP_PASSWORD or pipe the password on standard input; passwords are never accepted as arguments');
  process.exit(1);
}

function parseOptions(args: string[]): { user: string; hme: string; host?: string } {
  const parsed = new Map<string, string>();
  const allowed = new Set(['--user', '--hme', '--host']);
  for (let i = 0; i < args.length; i += 2) {
    const name = args[i];
    const value = args[i + 1];
    if (!allowed.has(name) || parsed.has(name) || !value || value.startsWith('--')) usage();
    parsed.set(name, value);
  }
  const user = parsed.get('--user');
  const hme = parsed.get('--hme');
  if (!user || !hme) usage();
  return { user, hme, host: parsed.get('--host') };
}

const args = process.argv.slice(2);
if (args.length % 2 !== 0) usage();
const { user, hme, host: hostArg } = parseOptions(args);

async function readPassword(): Promise<string> {
  const fromEnvironment = process.env.MAILHUB_IMAP_PASSWORD;
  if (fromEnvironment) return fromEnvironment;
  if (process.stdin.isTTY) {
    console.error('provide MAILHUB_IMAP_PASSWORD or pipe the password on standard input; passwords are never accepted as arguments');
    process.exit(1);
  }
  let text = '';
  for await (const chunk of process.stdin) text += String(chunk);
  const password = text.replace(/[\r\n]+$/, '');
  if (!password) {
    console.error('provide MAILHUB_IMAP_PASSWORD or pipe the password on standard input; passwords are never accepted as arguments');
    process.exit(1);
  }
  return password;
}

/** Same table the admin UI suggests from, so both agree on where mail lands. */
const IMAP_HOSTS: Record<string, string> = {
  'icloud.com': 'imap.mail.me.com', 'me.com': 'imap.mail.me.com', 'mac.com': 'imap.mail.me.com',
  'gmail.com': 'imap.gmail.com', 'googlemail.com': 'imap.gmail.com',
  'outlook.com': 'outlook.office365.com', 'hotmail.com': 'outlook.office365.com',
  'qq.com': 'imap.qq.com', '163.com': 'imap.163.com', '126.com': 'imap.126.com',
};

// Body matching is absent by design. That a message *contains* an address does
// not make that address its recipient, and the body is attacker controlled — a
// message to A quoting B would be delivered into B's inbox.
const CANDIDATES: Array<{ name: string; criteria: SearchObject; diagnosticOnly?: boolean }> = [
  { name: "{ to }", criteria: { to: hme } },
  { name: "header delivered-to", criteria: { header: { 'delivered-to': hme } } },
  { name: "header x-original-to", criteria: { header: { 'x-original-to': hme } } },
  { name: "header envelope-to", criteria: { header: { 'envelope-to': hme } } },
  // Diagnostic only. That a message *contains* an address does not make that
  // address its recipient, and the body is attacker-controlled — this line
  // exists to prove the mail arrived at all, never to route it.
  { name: "body (diagnostic only)", criteria: { body: hme }, diagnosticOnly: true },
];

async function main(): Promise<void> {
  const password = await readPassword();
  const host = hostArg || IMAP_HOSTS[user.split('@')[1]?.toLowerCase() ?? ''] || 'imap.mail.me.com';
  const client = new ImapFlow({
    host,
    port: 993,
    secure: true,
    auth: { user, pass: password },
    logger: false,
  });

  await client.connect();
  const lock = await client.getMailboxLock('INBOX');
  try {
    console.log(`\nSearching INBOX of ${user} for ${hme}\n`);

    for (const candidate of CANDIDATES) {
      try {
        const uids = await client.search(candidate.criteria, { uid: true });
        const hits = uids ? uids.length : 0;
        const tag = candidate.diagnosticOnly ? (hits > 0 ? 'seen' : 'miss') : (hits > 0 ? 'HIT ' : 'miss');
        console.log(`${tag}  ${candidate.name.padEnd(26)} ${hits} message(s)`);
      } catch (e) {
        console.log(`ERR   ${candidate.name.padEnd(24)} ${(e as Error).message}`);
      }
    }

    // Whatever the matrix says, read the real thing. Guessing which header
    // Apple kept is exactly the mistake this script exists to prevent.
    console.log('\n--- raw headers of the newest message in the mailbox ---\n');
    const newest = await client.search({ all: true }, { uid: true });
    const lastUid = newest && newest.length ? newest[newest.length - 1] : undefined;
    if (lastUid === undefined) {
      console.log('(mailbox is empty — send a message to the alias first)');
      return;
    }
    const { content } = await client.download(String(lastUid), '', { uid: true });
    const chunks: Buffer[] = [];
    for await (const chunk of content) chunks.push(chunk as Buffer);
    const raw = Buffer.concat(chunks).toString('utf8');
    console.log(raw.split(/\r?\n\r?\n/)[0]);
  } finally {
    lock.release();
    await client.logout();
  }
}

main().catch(() => {
  // IMAP libraries may include the attempted AUTH command in their error;
  // never print it because it can contain the supplied password.
  console.error('verification failed');
  process.exit(1);
});
