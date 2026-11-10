import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, mkdir } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';
import * as cli from '../src/program.js';
import { GraphClient } from '../src/graph.js';
import { AuthManager } from '../src/auth.js';
import type { Grant } from '../src/capabilities.js';

const base = ['User.Read', 'Mail.Read'];
const grants = {
  readOnly: { mode: 'read-only' as const, scopes: [...base, 'Mail.ReadWrite', 'Mail.Send'], expiresAt: Date.now() + 3600000 },
  calendarOnly: { mode: 'read-write' as const, scopes: [...base, 'Calendars.ReadWrite'], expiresAt: Date.now() + 3600000 },
  draftsOnly: { mode: 'read-write' as const, scopes: [...base, 'Calendars.ReadWrite', 'Mail.ReadWrite'], expiresAt: Date.now() + 3600000 },
  full: { mode: 'read-write' as const, scopes: [...base, 'Calendars.ReadWrite', 'Mail.ReadWrite', 'Mail.Send'], expiresAt: Date.now() + 3600000 },
};
function session(grant: Grant, network = { calls: 0 }) {
  return { grant, graph: new GraphClient('opaque', async () => { network.calls++; throw new Error('network must not happen'); }, grant), identity: { clientId: 'client', tenantId: 'tenant', homeAccountId: 'home', username: 'test', meId: 'user', cacheFile: 'unused' }, me: { id: 'user' } };
}
function harness(current?: ReturnType<typeof session>) {
  const out: string[] = [];
  const manager = { connection: async () => { if (!current) throw new Error('not signed in'); return current; } } as AuthManager;
  const run = async (...args: string[]) => {
    out.length = 0;
    const code = await cli.run(args, { manager, io: { out: s => out.push(s), err: () => {} } });
    return { code, text: out.join(''), json: () => JSON.parse(out.join('')) };
  };
  return { run };
}
const VERBS = ['create', 'reply', 'update', 'send'];
const listed = (help: string) => VERBS.filter(verb => new RegExp(`^  ${verb}\\b`, 'm').test(help));

test('draft help lists exactly the granted verbs and hidden verbs fail deterministically before any network call', async () => {
  const cases: [ReturnType<typeof session> | undefined, string[], string][] = [
    [undefined, [], 'AUTH_REQUIRED'],
    [session(grants.readOnly), [], 'READ_ONLY_MODE'],
    [session(grants.calendarOnly), [], 'INSUFFICIENT_SCOPE'],
    [session(grants.draftsOnly), ['create', 'reply', 'update'], 'INSUFFICIENT_SCOPE'],
    [session(grants.full), VERBS, ''],
  ];
  for (const [current, visible, code] of cases) {
    const { run } = harness(current);
    const help = await run('draft', '--help');
    assert.equal(help.code, 0);
    assert.deepEqual(listed(help.text), visible);
    assert.match(help.text, current ? new RegExp(`Active mode: ${current.grant.mode}`) : /No verified connection/);
    for (const verb of VERBS.filter(v => !visible.includes(v))) {
      for (const args of [['draft', verb, '--json'], ['draft', verb, '--help', '--json'], ['draft', 'help', verb, '--json'], ['--json', 'draft', verb, '--draft', 'id', '--to', 'a@example.test']]) {
        const result = await run(...args);
        assert.equal(result.code, 1, args.join(' '));
        assert.equal(result.json().error.code, code, args.join(' '));
      }
    }
    for (const verb of visible) assert.equal((await run('draft', verb, '--help')).code, 0);
  }
});

test('top-level help and capabilities advertise drafts only through the verified grant', async () => {
  const unauthenticated = await harness().run('--help');
  assert.match(unauthenticated.text, /draft/);
  assert.doesNotMatch(unauthenticated.text, /Mail is read-only|calendar-only/i);
  assert.equal((await harness().run('--version')).text.trim(), '0.3.0');
  const capabilities = await harness(session(grants.full)).run('capabilities', '--json');
  assert.deepEqual(capabilities.json().data.operations.slice(-4), ['draft.create', 'draft.reply', 'draft.update', 'draft.send']);
  assert.deepEqual((await harness(session(grants.calendarOnly)).run('capabilities', '--json')).json().data.operations.filter((o: string) => o.startsWith('draft')), []);
});

test('draft CLI routes recipients, bodies and IDs to the guarded Graph client and wraps results in the ok/data envelope', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'outlook-draft-body-'));
  const file = join(dir, 'body.txt');
  await writeFile(file, 'Grüezi\nfrom a file\n', 'utf8');
  const current = session(grants.full);
  const calls: any[] = [];
  for (const method of ['createDraft', 'createReplyDraft', 'updateDraft', 'sendDraft']) (current.graph as any)[method] = async (...args: any[]) => { calls.push([method, ...args]); return { outcome: 'stub', id: 'draft-id' }; };
  const { run } = harness(current);
  const commands = [
    ['draft', 'create', '--to', 'a@example.test', '--to', 'b@example.test, c@example.test', '--cc', 'd@example.test,e@example.test', '--bcc', 'f@example.test', '--subject', 'Hello there', '--body', 'Line one', '--html'],
    ['draft', 'create', '--to', 'a@example.test', '--subject', 'From file', '--body-file', file],
    ['draft', 'reply', '--message', 'msg/1', '--all', '--body-file', file],
    ['draft', 'reply', '--message', 'msg/1', '--body', 'Thanks'],
    ['draft', 'update', '--draft', 'draft/1', '--subject', 'Changed'],
    ['draft', 'update', '--draft', 'draft/1', '--to', 'x@example.test,y@example.test', '--cc', '', '--body', '<p>x</p>', '--html'],
    ['draft', 'send', '--draft', 'draft/1'],
  ];
  for (const args of commands) {
    const result = await run(...args, '--json');
    assert.equal(result.code, 0, args.join(' '));
    assert.deepEqual(result.json(), { ok: true, data: { outcome: 'stub', id: 'draft-id' } });
  }
  assert.deepEqual(calls, [
    ['createDraft', { to: ['a@example.test', 'b@example.test', 'c@example.test'], cc: ['d@example.test', 'e@example.test'], bcc: ['f@example.test'], subject: 'Hello there', body: 'Line one', html: true }],
    ['createDraft', { to: ['a@example.test'], subject: 'From file', body: 'Grüezi\nfrom a file\n' }],
    ['createReplyDraft', 'msg/1', { all: true, body: 'Grüezi\nfrom a file\n' }],
    ['createReplyDraft', 'msg/1', { all: false, body: 'Thanks' }],
    ['updateDraft', 'draft/1', { subject: 'Changed' }],
    ['updateDraft', 'draft/1', { to: ['x@example.test', 'y@example.test'], cc: [], body: '<p>x</p>', html: true }],
    ['sendDraft', 'draft/1'],
  ]);
});

test('invalid draft arguments are INVALID_INPUT and make no network call', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'outlook-draft-invalid-'));
  const good = join(dir, 'good.txt');
  await writeFile(good, 'fine');
  await writeFile(join(dir, 'big.txt'), 'x'.repeat(1_048_577));
  await writeFile(join(dir, 'binary.txt'), Buffer.from([0x48, 0xff, 0xfe, 0x00]));
  await mkdir(join(dir, 'folder'));
  const network = { calls: 0 };
  const { run } = harness(session(grants.full, network));
  const create = ['draft', 'create', '--to', 'a@example.test', '--subject', 'S'];
  const invalid = [
    [...create], [...create, '--body', 'x', '--body-file', good],
    [...create, '--body-file', join(dir, 'missing.txt')], [...create, '--body-file', join(dir, 'folder')],
    [...create, '--body-file', join(dir, 'big.txt')], [...create, '--body-file', join(dir, 'binary.txt')],
    ['draft', 'create', '--subject', 'S', '--body', 'x'], ['draft', 'create', '--to', 'a@example.test', '--body', 'x'],
    ['draft', 'create', '--to', 'Anna <a@example.test>', '--subject', 'S', '--body', 'x'], ['draft', 'create', '--to', 'a@example.test b@example.test', '--subject', 'S', '--body', 'x'],
    ['draft', 'create', '--to', 'a@example.test', '--subject', 'S\nBcc: evil@example.test', '--body', 'x'], [...create, '--body', 'x', '--attach', good],
    ['draft', 'reply', '--message', 'm'], ['draft', 'reply', '--body', 'x'], ['draft', 'reply', '--message', 'm', '--body', 'x', '--html'], ['draft', 'reply', '--message', 'm', '--body', 'x', '--to', 'a@example.test'],
    ['draft', 'reply', '--message', 'm', '--body', 'x', '--body-file', good],
    ['draft', 'update', '--draft', 'd'], ['draft', 'update', '--subject', 'S'], ['draft', 'update', '--draft', 'd', '--html'], ['draft', 'update', '--draft', 'd', '--body', 'x', '--body-file', good],
    ['draft', 'update', '--draft', 'd', '--to', 'not-an-address'],
    ['draft', 'send'], ['draft', 'send', '--draft', '..'], ['draft', 'send', '--draft', 'd', '--to', 'a@example.test'], ['draft', 'send', 'd'],
    ['draft', 'delete', '--draft', 'd'], ['draft', 'move', '--draft', 'd'], ['send', '--to', 'a@example.test'],
  ];
  for (const args of invalid) {
    const result = await run(...args, '--json');
    assert.equal(result.code, 1, args.join(' '));
    assert.equal(result.json().error.code, 'INVALID_INPUT', args.join(' '));
  }
  assert.equal(network.calls, 0);
});

test('real CLI entrypoint reports the new version and guards draft verbs without a session', async () => {
  const root = await mkdtemp(join(tmpdir(), 'outlook-draft-process-'));
  const invoke = (...args: string[]) => spawnSync(process.execPath, ['--import', 'tsx', 'src/cli.ts', ...args], { encoding: 'utf8', env: { ...process.env, XDG_CONFIG_HOME: root }, timeout: 10000 });
  assert.equal(invoke('--version').stdout.trim(), '0.3.0');
  for (const args of [['draft', 'send', '--draft', 'unused', '--json'], ['--json', 'draft', 'create', '--to', 'a@example.test', '--subject', 'S', '--body', 'x']]) {
    const rejected = invoke(...args);
    assert.equal(rejected.status, 1);
    assert.equal(JSON.parse(rejected.stdout).error.code, 'AUTH_REQUIRED');
  }
  const help = invoke('draft', '--help');
  assert.equal(help.status, 0);
  assert.deepEqual(listed(help.stdout), []);
});
