/**
 * Human-shaped mailbox usernames.
 *
 * randomString(12) produces addresses like `spdm2ole9l80@…`, which read as
 * obviously machine-generated. This mixes a few real-world shapes instead:
 *   markreyes52  lisa.chen  d_watson91  juliahoffman  k.nakamura7
 *
 * Output is always [a-z0-9._], starts with a letter, ends with a letter or
 * digit, and carries at most one separator — the conservative subset that
 * mail systems accept in a local part.
 *
 * Note the space is far smaller than randomString(12): roughly 7M variants,
 * so callers that need a distinct address must check for collisions rather
 * than trust the draw. See generateUniqueUsername in providers/imap.ts.
 */

const FIRST_NAMES = [
  'aaron', 'adam', 'adrian', 'alan', 'albert', 'alex', 'alice', 'amber', 'amelia', 'andre',
  'andrea', 'angela', 'anna', 'anthony', 'april', 'arthur', 'ava', 'barbara', 'ben', 'bianca',
  'brandon', 'brian', 'bruce', 'caleb', 'carla', 'carlos', 'carmen', 'caroline', 'cecilia', 'chloe',
  'chris', 'claire', 'clara', 'colin', 'connor', 'daniel', 'daphne', 'david', 'dean', 'diana',
  'diego', 'dylan', 'edward', 'elena', 'eliza', 'emma', 'eric', 'erika', 'ethan', 'evan',
  'felix', 'fiona', 'frank', 'gabriel', 'grace', 'grant', 'hannah', 'harry', 'heather', 'helen',
  'henry', 'hugo', 'ian', 'irene', 'isaac', 'ivan', 'jacob', 'james', 'jasmine', 'jason',
  'javier', 'jenna', 'jessica', 'joel', 'john', 'jonas', 'jordan', 'jose', 'julia', 'julian',
  'karen', 'kate', 'keith', 'kevin', 'laura', 'leo', 'lily', 'linda', 'lisa', 'lucas',
  'lucy', 'marco', 'maria', 'mark', 'martin', 'mason', 'maya', 'mia', 'micah', 'miguel',
  'nadia', 'nathan', 'nina', 'noah', 'olivia', 'oscar', 'owen', 'paul', 'paula', 'peter',
  'rachel', 'ralph', 'rebecca', 'rita', 'robert', 'ruby', 'ryan', 'samuel', 'sara', 'simon',
  'sofia', 'stella', 'tessa', 'thomas', 'tomas', 'tyler', 'vera', 'victor', 'wendy', 'zoe',
] as const;

const LAST_NAMES = [
  'abbott', 'adams', 'alvarez', 'anderson', 'bailey', 'baker', 'barnes', 'bauer', 'bennett', 'berg',
  'bishop', 'blake', 'boyd', 'brennan', 'brooks', 'bryant', 'burke', 'cameron', 'campbell', 'carter',
  'chen', 'clark', 'cohen', 'coleman', 'collins', 'conrad', 'cooper', 'cortez', 'cross', 'dalton',
  'davies', 'dawson', 'delgado', 'dixon', 'doyle', 'duarte', 'dunn', 'ellis', 'farrell', 'fischer',
  'fleming', 'fletcher', 'flores', 'foster', 'fowler', 'freeman', 'gallagher', 'garcia', 'gibson', 'gomez',
  'graham', 'grant', 'greene', 'hale', 'hansen', 'harper', 'hayes', 'hoffman', 'holt', 'hughes',
  'ibrahim', 'ingram', 'jenkins', 'jensen', 'kaufman', 'keller', 'kim', 'lambert', 'larsen', 'lawson',
  'leblanc', 'lindqvist', 'lopez', 'lynch', 'maddox', 'marsh', 'mendez', 'mercer', 'moreau', 'morrison',
  'nakamura', 'navarro', 'nelson', 'newton', 'nguyen', 'nolan', 'norris', 'oconnell', 'ortiz', 'osborne',
  'palmer', 'parker', 'patel', 'pearson', 'perry', 'pierce', 'quinn', 'ramirez', 'reyes', 'reynolds',
  'richter', 'rivera', 'rowe', 'sandoval', 'santos', 'schmidt', 'sharma', 'shaw', 'silva', 'sinclair',
  'stanton', 'stein', 'sullivan', 'tanaka', 'thornton', 'vaughn', 'vega', 'walsh', 'watson', 'weber',
  'whitaker', 'wilder', 'wong', 'yamada', 'yates', 'zhang',
] as const;

function pick<T>(arr: readonly T[]): T {
  return arr[Math.floor(Math.random() * arr.length)];
}

/** Generate one human-shaped username. */
export function randomUsername(): string {
  const first = pick(FIRST_NAMES);
  const last = pick(LAST_NAMES);

  // No separator is both the most common real-world shape and the safest for
  // upstream validators, so it is weighted highest.
  const r = Math.random();
  const separator = r < 0.5 ? '' : r < 0.8 ? '.' : '_';

  const given = Math.random() < 0.25 ? first[0] : first;

  const d = Math.random();
  const suffix = d < 0.45
    ? ''
    : d < 0.6
      ? String(Math.floor(Math.random() * 10))
      : String(Math.floor(Math.random() * 90) + 10);

  return `${given}${separator}${last}${suffix}`;
}
