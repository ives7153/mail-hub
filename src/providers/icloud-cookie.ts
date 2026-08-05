/**
 * Turn whatever the operator pasted into a Cookie header.
 *
 * Getting an iCloud session out of a browser produces wildly different text
 * depending on which menu item was clicked — "Copy as cURL" in three shell
 * dialects, "Copy as PowerShell", a cookie-extension JSON export, or the raw
 * header. Asking someone to find the Cookie line inside a 4KB cURL command and
 * paste only that part is how you get a truncated cookie and a confusing
 * "session not valid" an hour later, so every one of those shapes is accepted.
 */

/** Without these two the session is not usable, whatever else came along. */
export const REQUIRED_ICLOUD_COOKIES = ['X-APPLE-WEBAUTH-TOKEN', 'X-APPLE-WEBAUTH-USER'];

export interface ParsedCookies {
  /** Normalised `name=value; name=value` ready for a Cookie header. */
  cookies: string;
  /** Cookie names found, in the order they will be sent. */
  names: string[];
  /** Required names that were not present. */
  missing: string[];
}

/** `[{ name, value }, …]` — what cookie manager extensions export. */
function fromJsonExport(input: string): string | undefined {
  const trimmed = input.trim();
  if (!trimmed.startsWith('[') && !trimmed.startsWith('{')) return undefined;
  try {
    const parsed = JSON.parse(trimmed) as unknown;
    const list = Array.isArray(parsed) ? parsed : [parsed];
    const pairs = list
      .filter((e): e is { name: string; value: string } =>
        !!e && typeof e === 'object' && typeof (e as { name?: unknown }).name === 'string')
      .map((e) => `${e.name}=${String(e.value ?? '')}`);
    return pairs.length ? pairs.join('; ') : undefined;
  } catch {
    return undefined;
  }
}

/**
 * `New-Object System.Net.Cookie("name", "value", "/", "domain")` — one line per
 * cookie, which is what "Copy as PowerShell" emits.
 */
function fromPowerShell(input: string): string | undefined {
  const re = /New-Object\s+System\.Net\.Cookie\s*\(\s*"([^"]+)"\s*,\s*"([^"]*)"/gi;
  const pairs: string[] = [];
  for (const m of input.matchAll(re)) pairs.push(`${m[1]}=${m[2]}`);
  return pairs.length ? pairs.join('; ') : undefined;
}

/**
 * A cURL command: the cookies ride in `-H 'cookie: …'` or `-b '…'`.
 *
 * The quote character varies by shell (bash uses single, cmd uses double), and
 * the header name's case varies by browser, so both are matched loosely.
 */
function fromCurl(input: string): string | undefined {
  const header = input.match(/-H\s+(['"])\s*cookie\s*:\s*([\s\S]*?)\1/i);
  if (header) return header[2];
  const bFlag = input.match(/(?:^|\s)-b\s+(['"])([\s\S]*?)\1/);
  if (bFlag) return bFlag[2];
  const bBare = input.match(/(?:^|\s)-b\s+([^\s'"]+)/);
  if (bBare) return bBare[1];
  return undefined;
}

/** Strip a leading `Cookie:` if the header name came along for the ride. */
function stripHeaderName(value: string): string {
  return value.replace(/^\s*cookie\s*:\s*/i, '');
}

export function parseCookieBlob(input: string): ParsedCookies {
  const raw = fromJsonExport(input)
    ?? fromPowerShell(input)
    ?? fromCurl(input)
    ?? stripHeaderName(input);

  // Later wins: a paste that repeats a cookie is almost always a refreshed
  // value appended after a stale one.
  const seen = new Map<string, string>();
  for (const part of raw.split(';')) {
    const eq = part.indexOf('=');
    if (eq <= 0) continue;
    const name = part.slice(0, eq).trim();
    // A name with whitespace or quotes in it is shell noise, not a cookie.
    if (!name || !/^[!#$%&'*+\-.^_`|~0-9A-Za-z]+$/.test(name)) continue;
    const value = part.slice(eq + 1).trim().replace(/^"(.*)"$/s, '$1');
    seen.set(name, value);
  }

  const names = [...seen.keys()];
  const lower = new Set(names.map((n) => n.toLowerCase()));
  return {
    cookies: names.map((n) => `${n}=${seen.get(n)}`).join('; '),
    names,
    missing: REQUIRED_ICLOUD_COOKIES.filter((r) => !lower.has(r.toLowerCase())),
  };
}
