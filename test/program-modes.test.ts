import { test } from 'node:test';
import assert from 'node:assert/strict';
import * as cli from '../src/program.js';
import { GraphClient } from '../src/graph.js';
import { AuthManager } from '../src/auth.js';
import type { Grant } from '../src/capabilities.js';

const ro = { mode: 'read-only' as const, scopes: ['User.Read', 'Mail.Read'], expiresAt: Date.now() + 3600000 };
const rw = { ...ro, mode: 'read-write' as const, scopes: [...ro.scopes, 'Calendars.ReadWrite'] };
function session(grant: Grant = ro) {
  return { grant, graph: new GraphClient('opaque', async () => { throw new Error('network must not happen'); }, grant), identity: { clientId: 'client', tenantId: 'tenant', homeAccountId: 'home', username: 'test', meId: 'user', cacheFile: 'unused' }, me: { id: 'user' } };
}

test('calendar discovery and hidden-command guards use the same authenticated operation registry', async () => {
  assert.equal(typeof cli.run, 'function', 'CLI execution must return stable JSON errors');
  for (const current of [undefined, session(ro), session(rw as any)]) {
    const manager = { connection: async () => { if (!current) throw new Error('not signed in'); return current; } } as AuthManager;
    const out: string[] = [];
    const err: string[] = [];
    const options = { manager, io: { out: (s: string) => out.push(s), err: (s: string) => err.push(s) } };
    assert.equal(await cli.run(['calendar', '--help'], options), 0);
    const help = out.join('');
    if (current?.grant.mode === 'read-write') {
      for (const word of ['create', 'update', 'delete', 'events', 'get', 'list']) assert.match(help, new RegExp(word));
    } else {
      assert.doesNotMatch(help, /create|update|delete/);
      for (const verb of ['create', 'update', 'delete']) {
        out.length = 0;
        assert.equal(await cli.run(['calendar', verb, '--json'], options), 1);
        assert.equal(JSON.parse(out.join('')).error.code, current ? 'READ_ONLY_MODE' : 'AUTH_REQUIRED');
        out.length = 0;
        assert.equal(await cli.run(['calendar', verb, '--help', '--json'], options), 1);
        assert.equal(JSON.parse(out.join('')).error.code, current ? 'READ_ONLY_MODE' : 'AUTH_REQUIRED');
      }
    }
  }
});

test('capabilities JSON reports only actual active grants and no unverified identity on auth failure', async () => {
  const out: string[] = [];
  const current = session(ro);
  const manager = { connection: async () => current } as AuthManager;
  assert.equal(await cli.run(['capabilities', '--json'], { manager, io: { out: s => out.push(s), err: () => {} } }), 0);
  const result = JSON.parse(out.join(''));
  assert.equal(result.ok, true);
  assert.equal(result.data.mode, 'read-only');
  assert.deepEqual(result.data.operations, ['me', 'list', 'search', 'read']);
});

test('auth setup and mode-aware import route nonsecret IDs without requiring another mode login', async () => {
  const calls: any[] = [];
  const manager = {
    configure: async (...args: any[]) => { calls.push(['configure', ...args]); },
    login: async (mode: string, _notify: unknown, identity: unknown) => { calls.push(['login', mode, identity]); return { id: 'user' }; },
    importCache: async (...args: any[]) => { calls.push(['import', ...args]); return { id: 'user' }; },
    use: async (mode: string) => { calls.push(['use', mode]); return { id: 'user' }; },
  } as unknown as AuthManager;
  const id = '00000000-0000-0000-0000-000000000001';
  const options = { manager, io: { out: () => {}, err: () => {} } };
  for (const args of [
    ['auth', 'configure', '--mode', 'read-write', '--client-id', id, '--tenant-id', id],
    ['auth', 'login', '--mode', 'read-write'],
    ['auth', 'import-cache', '--mode', 'read-write', '--from', '/private/cache.json'],
    ['auth', 'use', '--mode', 'read-only'],
  ]) assert.equal(await cli.run(args, options), 0);
  assert.deepEqual(calls, [['configure', 'read-write', { clientId: id, tenantId: id }], ['login', 'read-write', undefined], ['import', '/private/cache.json', 'read-write', undefined], ['use', 'read-only']]);
});

test('calendar CLI routes named fields, IDs and explicit attendee approval to the guarded Graph client', async () => {
  const current = session(rw as any);
  const calls: any[] = [];
  for (const method of ['createEvent', 'updateEvent', 'deleteEvent', 'event', 'events', 'calendars']) (current.graph as any)[method] = async (...args: any[]) => { calls.push([method, ...args]); return { id: 'event' }; };
  const options = { manager: { connection: async () => current } as AuthManager, io: { out: () => {}, err: () => {} } };
  for (const args of [
    ['calendar', 'create', '--title', 'Test', '--start', '2027-01-01T10:00:00', '--end', '2027-01-01T11:00:00', '--timezone', 'UTC', '--calendar', 'cal'],
    ['calendar', 'update', '--event', 'id', '--title', 'Updated', '--notify-attendees'],
    ['calendar', 'delete', '--event', 'id', '--calendar', 'cal'],
    ['calendar', 'get', '--event', 'id'],
    ['calendar', 'events', '--start', '2027-01-01T00:00:00Z', '--end', '2027-01-02T00:00:00Z', '--limit', '5'],
    ['calendar', 'list', '--limit', '5'],
  ]) assert.equal(await cli.run(args, options), 0);
  assert.deepEqual(calls, [
    ['createEvent', { title: 'Test', start: '2027-01-01T10:00:00', end: '2027-01-01T11:00:00', timezone: 'UTC' }, 'cal'],
    ['updateEvent', 'id', { title: 'Updated' }, { calendar: undefined, notifyAttendees: true }],
    ['deleteEvent', 'id', { calendar: 'cal', notifyAttendees: undefined }],
    ['event', 'id', undefined],
    ['events', '2027-01-01T00:00:00Z', '2027-01-02T00:00:00Z', 5, undefined],
    ['calendars', 5],
  ]);
});

test('status failures and parser errors produce stable JSON without leaking upstream messages', async () => {
  const out: string[] = [];
  const options = { manager: { connection: async () => { throw new Error('PRIVATE-TOKEN'); } } as unknown as AuthManager, io: { out: (s: string) => out.push(s), err: () => {} } };
  assert.equal(await cli.run(['auth', 'status', '--json'], options), 1);
  assert.equal(JSON.parse(out.join('')).signedIn, false);
  assert.deepEqual(JSON.parse(out.join('')).operations, []);
  assert.doesNotMatch(out.join(''), /PRIVATE-TOKEN/);
  out.length = 0;
  assert.equal(await cli.run(['--json', 'unknown'], options), 1);
  assert.equal(JSON.parse(out.join('')).error.code, 'INVALID_INPUT');
  out.length = 0;
  assert.equal(await cli.run(['calendar', 'help', 'create', '--json'], options), 1);
  assert.equal(JSON.parse(out.join('')).error.code, 'AUTH_REQUIRED');
  out.length = 0;
  assert.equal(await cli.run(['auth', 'configure', '--mode', 'read-write', '--client-id', 'bad', '--tenant-id', 'bad', '--json'], options), 1);
  assert.equal(JSON.parse(out.join('')).error.code, 'INVALID_INPUT');
});
