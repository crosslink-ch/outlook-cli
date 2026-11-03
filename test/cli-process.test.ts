import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, lstat, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';

test('real CLI entrypoint: private offline configuration, contextual help, guarded mutations and JSON exit statuses', async () => {
  const root = await mkdtemp(join(tmpdir(), 'outlook-process-'));
  const invoke = (...args: string[]) => spawnSync(process.execPath, ['--import', 'tsx', 'src/cli.ts', ...args], { encoding: 'utf8', env: { ...process.env, XDG_CONFIG_HOME: root }, timeout: 10000 });
  const tenant = '00000000-0000-0000-0000-000000000001';
  for (const [mode, client] of [['read-only', '10000000-0000-0000-0000-000000000001'], ['read-write', '20000000-0000-0000-0000-000000000001']]) {
    const r = invoke('auth', 'configure', '--mode', mode, '--client-id', client, '--tenant-id', tenant, '--json');
    assert.equal(r.status, 0, r.stderr);
    assert.equal(JSON.parse(r.stdout).data.mode, mode);
    assert.equal(r.stderr, '');
  }
  const dir = join(root, 'outlook-cli');
  assert.equal((await lstat(dir)).mode & 0o777, 0o700);
  assert.equal((await lstat(join(dir, 'config.json'))).mode & 0o777, 0o600);
  const before = await readFile(join(dir, 'config.json'), 'utf8');
  assert.equal(JSON.parse(before).active, undefined, 'configuration must not itself authenticate');
  const rejected = invoke('calendar', 'delete', '--event', 'unused', '--json');
  assert.equal(rejected.status, 1);
  assert.equal(JSON.parse(rejected.stdout).error.code, 'AUTH_REQUIRED');
  const status = invoke('auth', 'status', '--json');
  assert.equal(status.status, 1);
  assert.deepEqual(JSON.parse(status.stdout).operations, []);
  assert.equal(await readFile(join(dir, 'config.json'), 'utf8'), before);
  const help = invoke('calendar', '--help');
  assert.equal(help.status, 0);
  assert.doesNotMatch(help.stdout, /create|update|delete/);
  const invalid = invoke('--json', 'auth', 'configure', '--mode', 'admin', '--client-id', tenant, '--tenant-id', tenant);
  assert.equal(invalid.status, 1);
  assert.equal(JSON.parse(invalid.stdout).error.code, 'INVALID_INPUT');
});
