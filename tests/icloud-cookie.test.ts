import { describe, it, expect } from 'vitest';
import { parseCookieBlob, REQUIRED_ICLOUD_COOKIES } from '../src/providers/icloud-cookie.js';

// Whatever DevTools or a cookie extension puts on the clipboard should work.
// Making the operator dig the Cookie header out of a 4KB cURL command by hand
// is the kind of chore that guarantees a truncated paste and a confusing
// "session not valid" an hour later.

const TOKEN = 'X-APPLE-WEBAUTH-TOKEN';
const USER = 'X-APPLE-WEBAUTH-USER';

describe('parseCookieBlob', () => {
  it('takes a bare Cookie header value', () => {
    const r = parseCookieBlob(`${TOKEN}=abc; ${USER}=def`);
    expect(r.cookies).toBe(`${TOKEN}=abc; ${USER}=def`);
    expect(r.missing).toEqual([]);
  });

  it('takes the header with its name still attached', () => {
    const r = parseCookieBlob(`Cookie: ${TOKEN}=abc; ${USER}=def`);
    expect(r.cookies).toContain(`${TOKEN}=abc`);
    expect(r.names).toContain(USER);
  });

  it('takes a Copy as cURL (bash) paste', () => {
    const blob = [
      `curl 'https://p68-maildomainws.icloud.com/v1/hme/list' \\`,
      `  -H 'accept: */*' \\`,
      `  -H 'cookie: ${TOKEN}=abc; ${USER}=def; X-APPLE-DS-WEB-SESSION-TOKEN=ghi' \\`,
      `  -H 'origin: https://www.icloud.com' \\`,
      `  --compressed`,
    ].join('\n');

    const r = parseCookieBlob(blob);
    expect(r.names).toEqual(expect.arrayContaining([TOKEN, USER, 'X-APPLE-DS-WEB-SESSION-TOKEN']));
    expect(r.cookies).not.toContain('curl');
    expect(r.cookies).not.toContain('origin');
    expect(r.missing).toEqual([]);
  });

  it('takes a Copy as cURL (cmd) paste, which quotes with double quotes', () => {
    const blob = `curl "https://p68.icloud.com/v1/hme/list" -H "cookie: ${TOKEN}=abc; ${USER}=def" --compressed`;
    const r = parseCookieBlob(blob);
    expect(r.names).toEqual(expect.arrayContaining([TOKEN, USER]));
  });

  it('takes curl -b as well as -H cookie', () => {
    const r = parseCookieBlob(`curl https://x -b '${TOKEN}=abc; ${USER}=def'`);
    expect(r.names).toEqual(expect.arrayContaining([TOKEN, USER]));
  });

  it('takes a Copy as PowerShell paste', () => {
    const blob = [
      '$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession',
      `$session.Cookies.Add((New-Object System.Net.Cookie("${TOKEN}", "abc", "/", "icloud.com")))`,
      `$session.Cookies.Add((New-Object System.Net.Cookie("${USER}", "def", "/", "icloud.com")))`,
      'Invoke-WebRequest -UseBasicParsing -Uri "https://p68.icloud.com/v1/hme/list" -WebSession $session',
    ].join('\n');

    const r = parseCookieBlob(blob);
    expect(r.cookies).toBe(`${TOKEN}=abc; ${USER}=def`);
  });

  it('takes a Cookie-Editor JSON export', () => {
    const blob = JSON.stringify([
      { domain: '.icloud.com', name: TOKEN, value: 'abc', httpOnly: true },
      { domain: '.icloud.com', name: USER, value: 'def' },
    ]);
    const r = parseCookieBlob(blob);
    expect(r.cookies).toBe(`${TOKEN}=abc; ${USER}=def`);
  });

  it('names exactly which required cookie is absent', () => {
    // The usual cause is copying a feedbackws/reportStats request, which
    // carries the token but not the user. Saying which one is missing beats a
    // generic rejection an hour before the first poll fails.
    const r = parseCookieBlob(`${TOKEN}=abc; some-other=1`);
    expect(r.missing).toEqual([USER]);
    expect(REQUIRED_ICLOUD_COOKIES).toContain(USER);
  });

  it('keeps cookies it does not recognise rather than allowlisting', () => {
    // An allowlist that misses one Apple session cookie breaks auth silently
    // and looks exactly like an expired session.
    const r = parseCookieBlob(`${TOKEN}=a; ${USER}=b; X-APPLE-WEB-SOMETHING-NEW=c`);
    expect(r.cookies).toContain('X-APPLE-WEB-SOMETHING-NEW=c');
  });

  it('survives a value containing = and quotes', () => {
    const r = parseCookieBlob(`${TOKEN}="v1:base64==padding"; ${USER}=def`);
    expect(r.cookies).toContain('v1:base64==padding');
  });

  it('drops duplicates, keeping the last occurrence', () => {
    const r = parseCookieBlob(`${TOKEN}=old; ${USER}=b; ${TOKEN}=new`);
    expect(r.cookies).toContain(`${TOKEN}=new`);
    expect(r.cookies).not.toContain('old');
  });

  it('reports nothing found for an empty or junk paste', () => {
    expect(parseCookieBlob('').names).toEqual([]);
    expect(parseCookieBlob('hello world').missing).toEqual(REQUIRED_ICLOUD_COOKIES);
  });
});
