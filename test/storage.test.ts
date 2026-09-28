import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, chmod, lstat, readFile, symlink } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { readPrivate, writePrivate } from '../src/storage.js';

test('token storage is atomic and owner-only; unsafe files are rejected', async () => {
  const root = await mkdtemp(join(tmpdir(), 'outlook-store-'));
  await writePrivate('cache.json', 'sensitive', root);
  assert.equal((await lstat(root)).mode & 0o777, 0o700);
  assert.equal((await lstat(join(root, 'cache.json'))).mode & 0o777, 0o600);
  assert.equal(await readPrivate('cache.json', root), 'sensitive');
  await writePrivate('cache.json', 'replacement', root);
  assert.equal(await readFile(join(root, 'cache.json'), 'utf8'), 'replacement');
  await chmod(join(root, 'cache.json'), 0o644);
  await assert.rejects(readPrivate('cache.json', root), /Unsafe credential file/);
  await symlink(join(root, 'cache.json'), join(root, 'config.json'));
  await assert.rejects(readPrivate('config.json', root), /Unsafe credential file/);
});
