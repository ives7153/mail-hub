import { describe, expect, it } from 'vitest';
import { spawnSync } from 'child_process';
import { readFileSync } from 'fs';
import { resolve } from 'path';

const script = resolve(process.cwd(), 'scripts/verify-icloud-imap.ts');
const secret = 'argv-secret-must-not-appear';

function run(args: string[], opts: { env?: NodeJS.ProcessEnv; input?: string } = {}): ReturnType<typeof spawnSync> {
  const env = { ...process.env, MAILHUB_IMAP_PASSWORD: undefined, ...opts.env };
  return spawnSync(process.execPath, ['--import', 'tsx', script, ...args], {
    cwd: process.cwd(),
    encoding: 'utf8',
    env,
    input: opts.input,
    timeout: 2_000,
  });
}

describe('verify-icloud-imap credential input', () => {
  it('shows secure-string conversion and always clears native and environment secrets', () => {
    const source = readFileSync(script, 'utf8');

    expect(source).toContain('[Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)');
    expect(source).toContain('try {');
    expect(source).toContain('[Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)');
    expect(source).toContain('Remove-Item Env:MAILHUB_IMAP_PASSWORD -ErrorAction SilentlyContinue');
  });

  it('rejects the legacy positional password shape without echoing the supplied secret', () => {
    const result = run(['user@example.com', secret, 'alias@icloud.com']);
    const output = `${result.stdout}${result.stderr}`;

    expect(result.status).not.toBe(0);
    expect(output).toMatch(/usage|password.*stdin|password.*environment/i);
    expect(output).not.toContain(secret);
  });

  it('fails safely when neither stdin nor the explicit password environment variable is present', () => {
    const result = run(['--user', 'user@example.com', '--hme', 'alias@icloud.com']);
    const output = `${result.stdout}${result.stderr}`;

    expect(result.status).not.toBe(0);
    expect(output).toMatch(/MAILHUB_IMAP_PASSWORD|standard input/i);
    expect(output).not.toContain(secret);
  });

  it.each([
    ['unknown option', ['--user', 'user@example.com', '--wat', 'x', '--hme', 'alias@icloud.com']],
    ['duplicate option', ['--user', 'user@example.com', '--user', 'other@example.com', '--hme', 'alias@icloud.com']],
    ['missing option value', ['--user', '--hme', 'alias@icloud.com']],
    ['positional argument', ['--user', 'user@example.com', '--hme', 'alias@icloud.com', 'extra']],
  ])('rejects %s without echoing credentials', (_caseName, args) => {
    const result = run(args, { env: { MAILHUB_IMAP_PASSWORD: secret } });
    const output = `${result.stdout}${result.stderr}`;
    expect(result.status).not.toBe(0);
    expect(output).toMatch(/usage/i);
    expect(output).not.toContain(secret);
  });

  it('accepts valid flags with an environment password and reaches a local refused port', () => {
    const args = ['--user', 'user@example.com', '--hme', 'alias@icloud.com', '--host', '127.0.0.1'];
    expect(args).not.toContain(secret);

    const result = run(args, { env: { MAILHUB_IMAP_PASSWORD: secret } });
    const output = `${result.stdout}${result.stderr}`;

    expect(result.status).not.toBe(0);
    expect(output).toContain('verification failed');
    expect(output).not.toMatch(/usage/i);
    expect(output).not.toContain(secret);
  });

  it('accepts a password from standard input and reaches a local refused port', () => {
    const args = ['--user', 'user@example.com', '--hme', 'alias@icloud.com', '--host', '127.0.0.1'];
    expect(args).not.toContain(secret);

    const result = run(args, { input: `${secret}\n` });
    const output = `${result.stdout}${result.stderr}`;

    expect(result.status).not.toBe(0);
    expect(output).toContain('verification failed');
    expect(output).not.toMatch(/usage/i);
    expect(output).not.toContain(secret);
  });

  it('prefers the environment password without waiting for standard input', () => {
    const args = ['--user', 'user@example.com', '--hme', 'alias@icloud.com', '--host', '127.0.0.1'];
    const result = run(args, { env: { MAILHUB_IMAP_PASSWORD: secret } });
    const output = `${result.stdout}${result.stderr}`;

    expect(result.status).not.toBe(0);
    expect(output).toContain('verification failed');
    expect(output).not.toContain(secret);
  });
});
