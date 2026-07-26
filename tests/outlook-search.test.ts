import { describe, expect, it } from 'vitest';
import { getDb } from '../src/db.js';
import { app, authHeaders, jsonOf } from './helpers/http.js';

type AccountsResponse = { accounts: { email: string; group_name: string | null }[] };

function insertAccount(email: string, group = 'Ungrouped') {
  getDb().prepare(
    `INSERT INTO outlook_accounts (email, password, group_name) VALUES (?, 'pw', ?)`
  ).run(email, group);
}

async function search(q: string): Promise<AccountsResponse> {
  const res = await app.request(`/api/outlook/accounts?q=${encodeURIComponent(q)}`, { headers: authHeaders() });
  expect(res.status).toBe(200);
  return jsonOf<AccountsResponse>(res);
}

describe('outlook account pool search', () => {
  it('matches email substrings case-preserving via LIKE', async () => {
    insertAccount('alice.work@outlook.com');
    insertAccount('bob@hotmail.com');

    const data = await search('alice');
    expect(data.accounts.map((a) => a.email)).toEqual(['alice.work@outlook.com']);
  });

  it('matches group name', async () => {
    insertAccount('one@outlook.com', 'batch-2024');
    insertAccount('two@outlook.com', 'other');

    const data = await search('batch-2024');
    expect(data.accounts.map((a) => a.email)).toEqual(['one@outlook.com']);
  });

  it('escapes LIKE wildcards so % and _ are literal', async () => {
    insertAccount('percent%weird@outlook.com');
    insertAccount('under_score@outlook.com');
    insertAccount('plain@outlook.com');

    // Interior wildcards are what discriminates: without ESCAPE, 'p%t' would
    // also match plain@outlook.com and 'plai_' would match plain@outlook.com.
    expect((await search('p%t')).accounts).toEqual([]);
    expect((await search('plai_')).accounts).toEqual([]);

    // The literal characters still match the rows that really contain them.
    expect((await search('percent%w')).accounts.map((a) => a.email)).toEqual(['percent%weird@outlook.com']);
    expect((await search('under_s')).accounts.map((a) => a.email)).toEqual(['under_score@outlook.com']);
  });

  it('combines with other filters', async () => {
    insertAccount('mix-long@outlook.com');
    getDb().prepare(`UPDATE outlook_accounts SET account_type = 'long' WHERE email = 'mix-long@outlook.com'`).run();
    insertAccount('mix-short@outlook.com');

    const res = await app.request('/api/outlook/accounts?q=mix&type=long', { headers: authHeaders() });
    const data = await jsonOf<AccountsResponse>(res);
    expect(data.accounts.map((a) => a.email)).toEqual(['mix-long@outlook.com']);
  });
});
