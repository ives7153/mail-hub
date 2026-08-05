import { describe, expect, it } from 'vitest';
import { readFileSync } from 'fs';
import { resolve } from 'path';

const source = readFileSync(resolve(process.cwd(), 'src/public/index.html'), 'utf8');

function addActivityCalls(): string[][] {
  const calls: string[][] = [];
  const needle = 'addActivity(';
  let start = source.indexOf(needle);

  while (start !== -1) {
    if (!source.slice(Math.max(0, start - 9), start).endsWith('function ')) {
      const args: string[] = [];
      let current = '';
      let depth = 1;
      let quote = '';

      for (let i = start + needle.length; i < source.length; i++) {
        const char = source[i];
        if (quote) {
          current += char;
          if (char === '\\') current += source[++i] ?? '';
          else if (char === quote) quote = '';
          continue;
        }
        if (char === '"' || char === "'" || char === '`') {
          quote = char;
          current += char;
        } else if (char === '(') {
          depth++;
          current += char;
        } else if (char === ')') {
          depth--;
          if (depth === 0) {
            args.push(current.trim());
            break;
          }
          current += char;
        } else if (char === ',' && depth === 1) {
          args.push(current.trim());
          current = '';
        } else {
          current += char;
        }
      }
      calls.push(args);
    }
    start = source.indexOf(needle, start + needle.length);
  }

  return calls;
}

describe('admin SPA HTML injection boundaries', () => {
  it('does not use the text-only escaper in an HTML attribute interpolation', () => {
    const unsafeAttributes = [
      ...source.matchAll(/\b[\w:-]+\s*=\s*(["'])\$\{esc\(/g),
      ...source.matchAll(/\b[\w:-]+\s*=\s*(["'])[^"']*\1\s*\+\s*esc\(/g),
    ].map(match => source.slice(0, match.index).split('\n').length);

    expect(unsafeAttributes).toEqual([]);
  });

  it('keeps activity text escaped by default and escapes every raw-HTML interpolation', () => {
    const implementation = source.slice(
      source.indexOf('function addActivity('),
      source.indexOf('function renderActivity('),
    );
    const rawCalls = addActivityCalls().filter(args => args[2] === 'true');
    const unsafeRawCalls = rawCalls.filter(args => {
      const text = args[1] || '';
      if (!text.startsWith('`')) return !/^(['"])[^+]*\1$/.test(text);
      return [...text.matchAll(/\$\{([^}]*)\}/g)]
        .some(match => !match[1].trim().startsWith('esc('));
    });

    expect(implementation).not.toMatch(/\b(?:safe|html|rawHtml)\s*:\s*true\b/);
    expect(implementation).not.toMatch(/function addActivity\([^)]*=\s*true/);
    expect(unsafeRawCalls).toEqual([]);
  });
});
