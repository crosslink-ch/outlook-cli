import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, lstat } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import * as auth from '../src/auth.js';
import { readPrivate, writePrivate } from '../src/storage.js';
import type { AuthenticationResult, AccountInfo } from '@azure/msal-node';

export const tenantId = '00000000-0000-0000-0000-000000000001';
export const roId = { clientId: '10000000-0000-0000-0000-000000000001', tenantId };
export const rwId = { clientId: '20000000-0000-0000-0000-000000000001', tenantId };
export const account: AccountInfo = { homeAccountId: 'user.tenant', localAccountId: 'user', username: 'person@example.test', tenantId, environment: 'login.microsoftonline.com' };
export function fixture() {
  const requests: any[] = [];
  const seeds: string[] = [];
  let scopes = ['User.Read', 'Mail.Read'];
  let result: AuthenticationResult;
  let fail = false;
  const dependencies = {
    createApp(identity: auth.Identity) {
      let cache = '';
      const acquire = async (req: any) => {
        requests.push({ identity, req });
        cache = JSON.stringify({ client: identity.clientId, account, credential: 'new-private-cache' });
        if (fail) throw new Error('SECRET must never escape');
        return result ?? { accessToken: 'opaque-token', account, scopes, tenantId, uniqueId: 'user', expiresOn: new Date(Date.now() + 3600000) } as AuthenticationResult;
      };
      return {
        getTokenCache: () => ({ deserialize: (data: string) => { cache = data; seeds.push(data); }, serialize: () => cache, getAllAccounts: async () => cache ? [account] : [] }),
        acquireTokenByDeviceCode: acquire,
        acquireTokenSilent: acquire,
      };
    },
    transport: (async () => new Response(JSON.stringify({ id: 'user', mail: account.username }))) as typeof fetch,
  };
  return { dependencies, requests, seeds, setScopes: (value: string[]) => { scopes = value; }, setResult: (value: AuthenticationResult) => { result = value; }, setFail: () => { fail = true; } };
}

test('mode login requests the configured app scopes and atomically saves isolated private connection', async () => {
  const root = await mkdtemp(join(tmpdir(), 'outlook-auth-'));
  const f = fixture();
  assert.equal(typeof auth.AuthManager, 'function', 'mode-aware authentication manager is required');
  const manager = new auth.AuthManager(root, f.dependencies);
  await manager.configure('read-only', roId);
  await manager.configure('read-write', rwId);
  await manager.login('read-only', () => {});
  const first = JSON.parse((await readPrivate('config.json', root))!);
  assert.equal(first.active, 'read-only');
  const ro = first.connections['read-only'];
  assert.deepEqual(f.requests[0].req.scopes, ['User.Read', 'Mail.Read']);
  assert.equal(ro.clientId, roId.clientId);
  f.setScopes(['User.Read', 'Mail.Read', 'Calendars.ReadWrite']);
  await manager.login('read-write', () => {});
  const second = JSON.parse((await readPrivate('config.json', root))!);
  assert.equal(second.active, 'read-write');
  assert.equal(second.connections['read-write'].clientId, rwId.clientId);
  assert.notEqual(second.connections['read-write'].cacheFile, ro.cacheFile);
  assert.equal(await readPrivate(ro.cacheFile, root), JSON.stringify({ client: roId.clientId, account, credential: 'new-private-cache' }));
  assert.deepEqual(f.requests[1].req.scopes, ['User.Read', 'Mail.Read', 'Calendars.ReadWrite', 'Mail.ReadWrite', 'Mail.Send']);
  assert.equal((await lstat(root)).mode & 0o777, 0o700);
  assert.equal((await lstat(join(root, second.connections['read-write'].cacheFile))).mode & 0o777, 0o600);
});

test('failed login validation preserves active identity and every working cache byte', async () => {
  for (const failure of ['missing scope', 'wrong tenant', 'wrong identity', 'expired', 'invalid expiry', 'expires during Graph', 'interrupted', 'Graph failure']) {
    const root = await mkdtemp(join(tmpdir(), 'outlook-rollback-'));
    const f = fixture();
    const manager = new auth.AuthManager(root, f.dependencies);
    await manager.login('read-only', () => {}, roId);
    const before = (await readPrivate('config.json', root))!;
    const cacheFile = JSON.parse(before).connections['read-only'].cacheFile;
    const cacheBefore = await readPrivate(cacheFile, root);
    const result = { accessToken: 'opaque-not-a-jwt', account, scopes: ['User.Read', 'Mail.Read', 'Calendars.ReadWrite'], tenantId, uniqueId: 'user', expiresOn: new Date(Date.now() + 3600000) } as AuthenticationResult;
    if (failure === 'missing scope') result.scopes = ['User.Read', 'Mail.Read'];
    if (failure === 'wrong tenant') result.tenantId = roId.clientId;
    if (failure === 'wrong identity') result.uniqueId = 'another-user';
    if (failure === 'expired') result.expiresOn = new Date(0);
    if (failure === 'invalid expiry') result.expiresOn = new Date('invalid');
    if (failure === 'interrupted') f.setFail();
    if (failure === 'Graph failure') f.dependencies.transport = async () => new Response('{}', { status: 403 });
    if (failure === 'expires during Graph') f.dependencies.transport = async () => { result.expiresOn = new Date(0); return Response.json({ id: 'user' }); };
    f.setResult(result);
    await assert.rejects(manager.login('read-write', () => {}, rwId), error => {
      assert.doesNotMatch(String(error), /SECRET|opaque-not-a-jwt/);
      return true;
    }, failure);
    assert.equal(await readPrivate('config.json', root), before, failure);
    assert.equal(await readPrivate(cacheFile, root), cacheBefore, failure);
  }
});

test('mode app configuration rejects reused client IDs and invalid modes without replacing working config', async () => {
  const root = await mkdtemp(join(tmpdir(), 'outlook-config-'));
  const manager = new auth.AuthManager(root, fixture().dependencies);
  await manager.configure('read-only', roId);
  const before = await readPrivate('config.json', root);
  await assert.rejects(manager.configure('read-write', roId), { code: 'INVALID_CONFIG' });
  await assert.rejects(manager.configure('write' as any, rwId), { code: 'INVALID_INPUT' });
  await assert.rejects(manager.configure('read-write', { ...rwId, tenantId: 'bad' }));
  assert.equal(await readPrivate('config.json', root), before);
});

test('explicit auth use silently verifies only its isolated cache; status fails closed on expiry without switching', async () => {
  const root = await mkdtemp(join(tmpdir(), 'outlook-use-'));
  const f = fixture();
  const manager = new auth.AuthManager(root, f.dependencies);
  await manager.login('read-only', () => {}, roId);
  f.setScopes(['User.Read', 'Mail.Read', 'Calendars.ReadWrite']);
  await manager.login('read-write', () => {}, rwId);
  assert.equal(typeof manager.use, 'function');
  await manager.use('read-only');
  const current = await manager.status();
  assert.equal(current.mode, 'read-only');
  assert.equal(current.username, account.username);
  assert.deepEqual(current.operations, ['me', 'list', 'search', 'read', 'calendar.list', 'calendar.events', 'calendar.get']);
  assert.equal(JSON.parse(f.seeds.at(-1)!).client, roId.clientId);
  const before = await readPrivate('config.json', root);
  f.setFail();
  await assert.rejects(manager.use('read-write'), { code: 'AUTH_REQUIRED' });
  assert.equal(await readPrivate('config.json', root), before);
  const expired = await manager.status();
  assert.equal(expired.signedIn, false);
  assert.deepEqual(expired.operations, []);
  assert.equal(expired.error?.code, 'AUTH_REQUIRED');
  assert.equal(JSON.parse((await readPrivate('config.json', root))!).active, 'read-only');
});

test('legacy one-login cache remains read-only and can migrate alongside a new mode import', async () => {
  const root = await mkdtemp(join(tmpdir(), 'outlook-legacy-'));
  const source = await mkdtemp(join(tmpdir(), 'outlook-source-'));
  await writePrivate('config.json', JSON.stringify({ ...roId, homeAccountId: account.homeAccountId }), root);
  await writePrivate('cache.json', 'legacy-private-cache', root);
  await writePrivate('cache.json', 'import-private-cache', source);
  const f = fixture();
  const manager = new auth.AuthManager(root, f.dependencies);
  const legacy = await manager.status();
  assert.equal(legacy.signedIn, true);
  assert.equal(legacy.mode, 'read-only');
  assert.equal(legacy.username, account.username);
  assert.deepEqual(legacy.operations, ['me', 'list', 'search', 'read']);
  const migrated = (await readPrivate('config.json', root))!;
  await assert.rejects(manager.importCache(join(source, 'cache.json'), 'read-write', rwId), { code: 'INSUFFICIENT_SCOPE' });
  assert.equal(await readPrivate('config.json', root), migrated);
  f.setScopes(['User.Read', 'Mail.Read', 'Calendars.ReadWrite']);
  await manager.importCache(join(source, 'cache.json'), 'read-write', rwId);
  const state = JSON.parse((await readPrivate('config.json', root))!);
  assert.equal(state.version, 2);
  assert.equal(state.connections['read-only'].homeAccountId, account.homeAccountId);
  assert.equal(state.active, 'read-write');
  assert.equal(await readPrivate('cache.json', source), 'import-private-cache');
  await assert.rejects(manager.importCache(join(source, 'cache.json'), 'read-write', rwId), /refusing to overwrite/);
  await manager.logout();
  assert.equal(await readPrivate('config.json', root), undefined);
  for (const connection of Object.values(state.connections) as any[]) assert.equal(await readPrivate(connection.cacheFile, root), undefined);
});

test('compatibility helpers resolve active mode rather than the obsolete shared cache', async () => {
  const root = await mkdtemp(join(tmpdir(), 'outlook-helpers-'));
  const f = fixture();
  f.setScopes(['User.Read', 'Mail.Read', 'Calendars.ReadWrite']);
  await new auth.AuthManager(root, f.dependencies).login('read-write', () => {}, rwId);
  assert.equal((await auth.savedIdentity(root)).clientId, rwId.clientId);
  assert.equal((await (await auth.graphForSavedLogin(root, f.dependencies)).me()).id, 'user');
  assert.equal((await auth.status(root, f.dependencies)).mode, 'read-write');
  await auth.logout(root);
  assert.equal((await auth.status(root, f.dependencies)).signedIn, false);
});

test('a stale in-flight status cannot reactivate its mode after an explicit mode switch', async () => {
  const root = await mkdtemp(join(tmpdir(), 'outlook-race-'));
  const f = fixture();
  const manager = new auth.AuthManager(root, f.dependencies);
  await manager.login('read-only', () => {}, roId);
  f.setScopes(['User.Read', 'Mail.Read', 'Calendars.ReadWrite']);
  await manager.login('read-write', () => {}, rwId);
  await manager.use('read-only');
  let release!: () => void;
  let entered!: () => void;
  const paused = new Promise<void>(resolve => { release = resolve; });
  const ready = new Promise<void>(resolve => { entered = resolve; });
  const factory = f.dependencies.createApp;
  f.dependencies.createApp = identity => {
    const app = factory(identity);
    const silent = app.acquireTokenSilent;
    if (identity.clientId === roId.clientId) app.acquireTokenSilent = async req => { entered(); await paused; return silent(req); };
    return app;
  };
  const stale = manager.connection();
  await ready;
  await manager.use('read-write');
  release();
  await assert.rejects(stale, { code: 'CONCURRENT_AUTH_CHANGE' });
  assert.equal(JSON.parse((await readPrivate('config.json', root))!).active, 'read-write');
});

test('parallel reads reuse an unchanged cache without rewriting configuration or failing each other', async () => {
  const root = await mkdtemp(join(tmpdir(), 'outlook-parallel-read-'));
  const f = fixture();
  const manager = new auth.AuthManager(root, f.dependencies);
  await manager.login('read-only', () => {}, roId);
  const before = await readPrivate('config.json', root);
  const results = await Promise.allSettled(Array.from({ length: 4 }, () => manager.connection()));
  assert.deepEqual(results.map(r => r.status), ['fulfilled', 'fulfilled', 'fulfilled', 'fulfilled']);
  assert.equal(await readPrivate('config.json', root), before);
});
