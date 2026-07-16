import { APP_VERSION } from './version.js';
import { fetchWithTimeout } from './utils.js';

const GITHUB_TAGS_URL = 'https://api.github.com/repos/ydddp/mail-hub/tags?per_page=100';
const GITHUB_TAGS_FEED_URL = 'https://github.com/ydddp/mail-hub/tags.atom';
const DEFAULT_CACHE_TTL_MS = 5 * 60 * 1000;

type VersionTuple = readonly [number, number, number];
type UpdateFetch = (
  url: string,
  options?: RequestInit & { timeout?: number; retries?: number },
) => Promise<Response>;

export interface UpdateCheckResult {
  currentVersion: string;
  latestVersion: string;
  updateAvailable: boolean;
  checkedAt: string;
}

export interface UpdateCheckerOptions {
  currentVersion?: string;
  fetcher?: UpdateFetch;
  now?: () => Date;
  cacheTtlMs?: number;
}

function parseStableVersion(value: unknown): { normalized: string; tuple: VersionTuple } | null {
  if (typeof value !== 'string') return null;
  const match = /^v?(\d+)\.(\d+)\.(\d+)$/.exec(value.trim());
  if (!match) return null;
  const tuple = match.slice(1).map(Number) as unknown as VersionTuple;
  if (tuple.some(part => !Number.isSafeInteger(part))) return null;
  return { normalized: tuple.join('.'), tuple };
}

export function normalizeStableVersion(value: unknown): string | null {
  return parseStableVersion(value)?.normalized ?? null;
}

export function compareStableVersions(left: string, right: string): number {
  const a = parseStableVersion(left);
  const b = parseStableVersion(right);
  if (!a || !b) throw new Error('invalid stable version');

  for (let i = 0; i < 3; i++) {
    if (a.tuple[i] > b.tuple[i]) return 1;
    if (a.tuple[i] < b.tuple[i]) return -1;
  }
  return 0;
}

export function findLatestStableVersion(payload: unknown): string {
  if (!Array.isArray(payload)) throw new Error('invalid GitHub tags response');

  const versions = payload
    .map(tag => tag && typeof tag === 'object'
      ? normalizeStableVersion((tag as { name?: unknown }).name)
      : null)
    .filter((value): value is string => value !== null);

  if (versions.length === 0) throw new Error('no stable version tag found');
  return versions.reduce((latest, candidate) => (
    compareStableVersions(candidate, latest) > 0 ? candidate : latest
  ));
}

function findLatestStableVersionFromAtom(xml: string): string {
  const versions = Array.from(
    xml.matchAll(/<id>[^<]*\/(v?\d+\.\d+\.\d+)<\/id>/g),
    match => normalizeStableVersion(match[1]),
  ).filter((value): value is string => value !== null);

  if (versions.length === 0) throw new Error('invalid GitHub tag feed response');
  return versions.reduce((latest, candidate) => (
    compareStableVersions(candidate, latest) > 0 ? candidate : latest
  ));
}

export function createUpdateChecker(options: UpdateCheckerOptions = {}): () => Promise<UpdateCheckResult> {
  const currentVersion = normalizeStableVersion(options.currentVersion ?? APP_VERSION);
  const fetcher = options.fetcher ?? fetchWithTimeout;
  const now = options.now ?? (() => new Date());
  const cacheTtlMs = options.cacheTtlMs ?? DEFAULT_CACHE_TTL_MS;
  let cache: { result: UpdateCheckResult; expiresAt: number } | undefined;

  return async () => {
    if (!currentVersion) throw new Error('invalid current application version');

    const currentTime = now();
    if (cache && currentTime.getTime() < cache.expiresAt) return cache.result;

    const response = await fetcher(GITHUB_TAGS_URL, {
      headers: {
        Accept: 'application/vnd.github+json',
        'User-Agent': 'Mail-Hub',
      },
      timeout: 10_000,
      retries: 1,
    });

    let latestVersion: string;
    if (response.status === 403 || response.status === 429) {
      const feedResponse = await fetcher(GITHUB_TAGS_FEED_URL, {
        headers: {
          Accept: 'application/atom+xml',
          'User-Agent': 'Mail-Hub',
        },
        timeout: 10_000,
        retries: 1,
      });
      if (!feedResponse.ok) {
        throw new Error(`GitHub tag feed returned status ${feedResponse.status}`);
      }
      latestVersion = findLatestStableVersionFromAtom(await feedResponse.text());
    } else {
      if (!response.ok) throw new Error(`GitHub returned status ${response.status}`);

      let payload: unknown;
      try {
        payload = await response.json();
      } catch {
        throw new Error('invalid GitHub tags response');
      }

      latestVersion = findLatestStableVersion(payload);
    }

    const result: UpdateCheckResult = {
      currentVersion,
      latestVersion,
      updateAvailable: compareStableVersions(latestVersion, currentVersion) > 0,
      checkedAt: currentTime.toISOString(),
    };
    cache = { result, expiresAt: currentTime.getTime() + cacheTtlMs };
    return result;
  };
}

export const checkForUpdates = createUpdateChecker();
