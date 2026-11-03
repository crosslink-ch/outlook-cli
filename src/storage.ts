import { randomUUID } from 'node:crypto';
import { constants, type Stats } from 'node:fs';
import { chmod, lstat, mkdir, open, readFile, rename, unlink } from 'node:fs/promises';
import { homedir } from 'node:os';
import { join } from 'node:path';
import { CliError } from './errors.js';

export function configRoot(): string {
  return join(process.env.XDG_CONFIG_HOME || join(homedir(), '.config'), 'outlook-cli');
}

function filePath(name: string, root: string): string {
  if (name !== 'config.json' && name !== 'cache.json' && !/^cache-[0-9a-f-]{36}\.json$/.test(name)) throw new Error('Invalid credential filename');
  return join(root, name);
}

function safe(stat: Stats): boolean {
  return stat.isFile() && stat.uid === process.getuid?.() && (stat.mode & 0o077) === 0;
}

async function ensureRoot(root: string): Promise<void> {
  await mkdir(root, { recursive: true, mode: 0o700 });
  const stat = await lstat(root);
  if (!stat.isDirectory() || stat.uid !== process.getuid?.()) throw new Error('Unsafe credential directory');
  await chmod(root, 0o700);
}

export async function readPrivate(name: string, root = configRoot()): Promise<string | undefined> {
  const path = filePath(name, root);
  try {
    const stat = await lstat(path);
    if (!safe(stat)) throw new Error('Unsafe credential file: expected owner-only regular file');
    const handle = await open(path, constants.O_RDONLY | constants.O_NOFOLLOW);
    try { return await handle.readFile('utf8'); } finally { await handle.close(); }
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'ENOENT') return undefined;
    throw error;
  }
}

export async function writePrivate(name: string, data: string, root = configRoot()): Promise<void> {
  const path = filePath(name, root);
  await ensureRoot(root);
  const existing = await lstat(path).catch((error: NodeJS.ErrnoException) => {
    if (error.code === 'ENOENT') return undefined;
    throw error;
  });
  if (existing && !safe(existing)) throw new Error('Unsafe credential file: refusing overwrite');
  const tmp = join(root, `.tmp-${randomUUID()}`);
  const handle = await open(tmp, constants.O_WRONLY | constants.O_CREAT | constants.O_EXCL, 0o600);
  try {
    await handle.writeFile(data, 'utf8');
    await handle.sync();
  } finally { await handle.close(); }
  try { await rename(tmp, path); } finally { await unlink(tmp).catch(() => {}); }
}

export async function removePrivate(name: string, root = configRoot()): Promise<void> {
  const path = filePath(name, root);
  const stat = await lstat(path).catch((error: NodeJS.ErrnoException) => {
    if (error.code === 'ENOENT') return undefined;
    throw error;
  });
  if (!stat) return;
  if (!safe(stat)) throw new Error('Unsafe credential file: refusing removal');
  await unlink(path);
}

// Short lock only around compare-and-swap, never across authentication/network I/O.
// An interrupted process may leave .config.lock: fail closed rather than guessing it is stale.
export async function replaceConfig(expected: string | undefined, data: string | undefined, root: string): Promise<void> {
  await ensureRoot(root);
  const lock = join(root, '.config.lock');
  let handle;
  try { handle = await open(lock, constants.O_WRONLY | constants.O_CREAT | constants.O_EXCL | constants.O_NOFOLLOW, 0o600); }
  catch { throw new CliError('CONFIG_BUSY', 'Another credential update holds the config lock; retry after it finishes'); }
  try {
    if (await readPrivate('config.json', root) !== expected) throw new CliError('CONCURRENT_AUTH_CHANGE', 'Authentication configuration changed during verification; retry without switching modes');
    if (data === undefined) await removePrivate('config.json', root);
    else await writePrivate('config.json', data, root);
  } finally { await handle.close(); await unlink(lock); }
}
