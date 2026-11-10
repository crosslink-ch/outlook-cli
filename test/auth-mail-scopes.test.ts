import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import * as auth from '../src/auth.js';
import { readPrivate, writePrivate } from '../src/storage.js';
import type { AuthenticationResult, AccountInfo } from '@azure/msal-node';

const tenantId = '00000000-0000-0000-0000-000000000001';
const roId = { clientId: '10000000-0000-0000-0000-000000000001', tenantId };
const rwId = { clientId: '20000000-0000-0000-0000-000000000001', tenantId };
const account: AccountInfo = { homeAccountId: 'user.tenant', localAccountId: 'user', username: 'person@example.test', tenantId, environment: 'login.microsoftonline.com' };
const MINIMUM = ['User.Read', 'Mail.Read', 'Calendars.ReadWrite'];
const FULL = [...MINIMUM, 'Mail.ReadWrite', 'Mail.Send'];
const OLD_OPERATIONS = ['me', 'list', 'search', 'read', 'calendar.list', 'calendar.events', 'calendar.get', 'calendar.create', 'calendar.update', 'calendar.delete'];
const DRAFTS = ['draft.create', 'draft.reply', 'draft.update', 'draft.send'];

// `grant` decides what Microsoft would return for the requested scopes: an array of granted scopes, or undefined to fail.
function fixture(grant: (requested: string[], kind: 'device' | 'silent') => string[] | undefined) {
  const requests: { kind: 'device' | 'silent'; scopes: string[] }[] = [];
  const control = { grant, refresh: false };
  const dependencies: auth.AuthDependencies = {
    createApp(identity) {
      let cache = '';
      const acquire = (kind: 'device' | 'silent') => async (req: any) => {
        requests.push({ kind, scopes: req.scopes });
        const scopes = control.grant(req.scopes, kind);
        if (!scopes) throw new Error('SECRET consent_required must never escape');
        if (kind === 'device' || control.refresh) cache = JSON.stringify({ client: identity.clientId, scopes });
        return { accessToken: 'opaque-token', account, scopes, tenantId, uniqueId: 'user', expiresOn: new Date(Date.now() + 3600000) } as AuthenticationResult;
      };
      return {
        getTokenCache: () => ({ deserialize: (data: string) => { cache = data; }, serialize: () => cache, getAllAccounts: async () => cache ? [account] : [] }),
        acquireTokenByDeviceCode: acquire('device'),
        acquireTokenSilent: acquire('silent'),
      };
    },
    transport: (async () => new Response(JSON.stringify({ id: 'user', mail: account.username }))) as typeof fetch,
  };
  return { dependencies, requests, control };
}
const onlyConsented = (consented: string[]) => (requested: string[]) => requested.every(scope => consented.includes(scope)) ? requested : undefined;

test('read-write login requests the mail-write scopes but accepts a grant holding only the required minimum', async () => {
  const root = await mkdtemp(join(tmpdir(), 'outlook-mail-scope-'));
  const f = fixture(() => MINIMUM);
  const manager = new auth.AuthManager(root, f.dependencies);
  await manager.login('read-write', () => {}, rwId);
  assert.deepEqual(f.requests[0], { kind: 'device', scopes: FULL });
  const status = await manager.status();
  assert.equal(status.mode, 'read-write');
  assert.deepEqual(status.operations, OLD_OPERATIONS);
  f.control.grant = () => ['User.Read', 'Mail.Read', 'Mail.ReadWrite', 'Mail.Send'];
  await assert.rejects(manager.login('read-write', () => {}, rwId), { code: 'INSUFFICIENT_SCOPE' });
});

test('read-write login granted every requested scope lists the draft operations', async () => {
  const root = await mkdtemp(join(tmpdir(), 'outlook-mail-scope-'));
  const f = fixture(requested => requested);
  const manager = new auth.AuthManager(root, f.dependencies);
  await manager.login('read-write', () => {}, rwId);
  assert.deepEqual((await manager.status()).operations, [...OLD_OPERATIONS, ...DRAFTS]);
});

test('an existing three-scope read-write connection keeps working: silent acquisition falls back to the required minimum', async () => {
  const root = await mkdtemp(join(tmpdir(), 'outlook-mail-scope-'));
  const f = fixture(() => MINIMUM);
  const manager = new auth.AuthManager(root, f.dependencies);
  await manager.login('read-write', () => {}, rwId);
  const before = await readPrivate('config.json', root);
  f.control.grant = onlyConsented(MINIMUM);
  f.requests.length = 0;
  const session = await manager.connection();
  assert.deepEqual(f.requests, [{ kind: 'silent', scopes: FULL }, { kind: 'silent', scopes: MINIMUM }]);
  assert.equal(session.grant.mode, 'read-write');
  assert.deepEqual(session.grant.scopes, MINIMUM);
  assert.deepEqual((await manager.status()).operations, OLD_OPERATIONS);
  assert.equal(await readPrivate('config.json', root), before, 'an unchanged cache must not be republished');
  assert.equal((await auth.graphForSavedLogin(root, f.dependencies).then(graph => graph.me())).id, 'user');
  assert.equal((await manager.use('read-write')).id, 'user');
});

test('silent acquisition is AUTH_REQUIRED only when the required minimum also fails, without leaking MSAL errors', async () => {
  const root = await mkdtemp(join(tmpdir(), 'outlook-mail-scope-'));
  const f = fixture(() => MINIMUM);
  const manager = new auth.AuthManager(root, f.dependencies);
  await manager.login('read-write', () => {}, rwId);
  f.control.grant = () => undefined;
  f.requests.length = 0;
  await assert.rejects(manager.connection(), (error: any) => {
    assert.equal(error.code, 'AUTH_REQUIRED');
    assert.doesNotMatch(JSON.stringify(error) + String(error), /SECRET|consent_required/);
    return true;
  });
  assert.deepEqual(f.requests.map(r => r.scopes), [FULL, MINIMUM]);
});

test('a consented tenant silently upgrades the token and the refreshed cache is persisted', async () => {
  const root = await mkdtemp(join(tmpdir(), 'outlook-mail-scope-'));
  const f = fixture(() => MINIMUM);
  const manager = new auth.AuthManager(root, f.dependencies);
  await manager.login('read-write', () => {}, rwId);
  const before = JSON.parse((await readPrivate('config.json', root))!).connections['read-write'].cacheFile;
  assert.deepEqual(JSON.parse((await readPrivate(before, root))!).scopes, MINIMUM);
  f.control.grant = requested => requested;
  f.control.refresh = true;
  f.requests.length = 0;
  const session = await manager.connection();
  assert.deepEqual(f.requests, [{ kind: 'silent', scopes: FULL }]);
  assert.deepEqual(session.grant.scopes, FULL);
  const after = JSON.parse((await readPrivate('config.json', root))!).connections['read-write'].cacheFile;
  assert.notEqual(after, before);
  assert.deepEqual(JSON.parse((await readPrivate(after, root))!).scopes, FULL);
  assert.equal(await readPrivate(before, root), undefined);
  f.control.refresh = false;
  assert.deepEqual((await manager.status()).operations, [...OLD_OPERATIONS, ...DRAFTS]);
});

test('import-cache accepts a three-scope read-write cache by falling back to the required minimum', async () => {
  const root = await mkdtemp(join(tmpdir(), 'outlook-mail-scope-'));
  const source = await mkdtemp(join(tmpdir(), 'outlook-mail-scope-source-'));
  await writePrivate('cache.json', 'import-private-cache', source);
  const f = fixture(onlyConsented(MINIMUM));
  const manager = new auth.AuthManager(root, f.dependencies);
  assert.equal((await manager.importCache(join(source, 'cache.json'), 'read-write', rwId)).id, 'user');
  assert.deepEqual(f.requests, [{ kind: 'silent', scopes: FULL }, { kind: 'silent', scopes: MINIMUM }]);
  const state = JSON.parse((await readPrivate('config.json', root))!);
  assert.equal(state.active, 'read-write');
  assert.equal(await readPrivate(state.connections['read-write'].cacheFile, root), 'import-private-cache');
});

test('read-only mode still requests and requires only User.Read and Mail.Read, with no fallback attempt', async () => {
  const root = await mkdtemp(join(tmpdir(), 'outlook-mail-scope-'));
  const f = fixture(requested => requested);
  const manager = new auth.AuthManager(root, f.dependencies);
  await manager.login('read-only', () => {}, roId);
  await manager.connection();
  assert.deepEqual(f.requests, [{ kind: 'device', scopes: ['User.Read', 'Mail.Read'] }, { kind: 'silent', scopes: ['User.Read', 'Mail.Read'] }]);
  // Even a token that somehow carries mail-write scopes cannot write in read-only mode.
  f.control.grant = () => ['User.Read', 'Mail.Read', 'Mail.ReadWrite', 'Mail.Send'];
  assert.deepEqual((await manager.status()).operations, ['me', 'list', 'search', 'read']);
  f.control.grant = () => undefined;
  f.requests.length = 0;
  await assert.rejects(manager.connection(), { code: 'AUTH_REQUIRED' });
  assert.equal(f.requests.length, 1);
});
