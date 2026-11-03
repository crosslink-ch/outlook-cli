import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { PublicClientApplication } from '@azure/msal-node';
import { AuthManager } from '../src/auth.js';
import { readPrivate, writePrivate } from '../src/storage.js';

// Synthetic MSAL Python SerializableTokenCache schema, not credentials from a real account.
// The real MSAL Node parser and silent token selection are exercised without any external I/O.
test('real MSAL accepts a Python-format opaque Graph cache only for its exact client and scopes', async () => {
  const tenantId = '00000000-0000-0000-0000-000000000001';
  const clientId = '20000000-0000-0000-0000-000000000001';
  const environment = 'login.microsoftonline.com';
  const homeAccountId = `user.${tenantId}`;
  const now = Math.floor(Date.now() / 1000);
  const target = 'Calendars.ReadWrite Mail.Read User.Read';
  const idToken = ['header', Buffer.from(JSON.stringify({ oid: 'user', sub: 'user', tid: tenantId, preferred_username: 'test@example.test', aud: clientId, exp: now + 3600 })).toString('base64url'), 'signature'].join('.');
  const cache = JSON.stringify({
    Account: { [`${homeAccountId}-${environment}-${tenantId}`]: { home_account_id: homeAccountId, environment, realm: tenantId, local_account_id: 'user', username: 'test@example.test', authority_type: 'MSSTS' } },
    IdToken: { [`${homeAccountId}-${environment}-idtoken-${clientId}-${tenantId}-`]: { home_account_id: homeAccountId, environment, credential_type: 'IdToken', client_id: clientId, realm: tenantId, secret: idToken } },
    AccessToken: { [`${homeAccountId}-${environment}-accesstoken-${clientId}-${tenantId}-${target}`.toLowerCase()]: { home_account_id: homeAccountId, environment, credential_type: 'AccessToken', client_id: clientId, realm: tenantId, target, secret: 'opaque-graph-token-not-jwt', token_type: 'Bearer', cached_at: String(now), expires_on: String(now + 3600), extended_expires_on: String(now + 3600) } },
    RefreshToken: {}, AppMetadata: {},
  });
  const root = await mkdtemp(join(tmpdir(), 'outlook-msal-'));
  const source = await mkdtemp(join(tmpdir(), 'outlook-msal-source-'));
  await writePrivate('cache.json', cache, source);
  let network = 0;
  let graphCalls = 0;
  const manager = new AuthManager(root, {
    createApp: identity => new PublicClientApplication({
      auth: {
        clientId: identity.clientId, authority: `https://${environment}/${identity.tenantId}`,
        cloudDiscoveryMetadata: JSON.stringify({ tenant_discovery_endpoint: `https://${environment}/${tenantId}/v2.0/.well-known/openid-configuration`, metadata: [{ preferred_network: environment, preferred_cache: environment, aliases: [environment] }] }),
        authorityMetadata: JSON.stringify({ authorization_endpoint: `https://${environment}/${tenantId}/oauth2/v2.0/authorize`, token_endpoint: `https://${environment}/${tenantId}/oauth2/v2.0/token`, issuer: `https://${environment}/${tenantId}/v2.0`, end_session_endpoint: `https://${environment}/${tenantId}/oauth2/v2.0/logout` }),
      },
      system: { networkClient: { sendGetRequestAsync: async () => { network++; throw new Error('External I/O forbidden'); }, sendPostRequestAsync: async () => { network++; throw new Error('External I/O forbidden'); } }, loggerOptions: { loggerCallback: () => {}, piiLoggingEnabled: false } },
    }),
    transport: async (_url, init) => { graphCalls++; assert.equal(new Headers(init?.headers).get('Authorization'), 'Bearer opaque-graph-token-not-jwt'); return Response.json({ id: 'user' }); },
  });
  const result = await manager.importCache(join(source, 'cache.json'), 'read-write', { clientId, tenantId });
  assert.equal(result.id, 'user');
  assert.equal((await manager.status()).mode, 'read-write');
  assert.equal(network, 0);
  assert.equal(graphCalls, 2);
  assert.equal(await readPrivate('cache.json', source), cache);
  await manager.logout();
  await assert.rejects(manager.importCache(join(source, 'cache.json'), 'read-write', { clientId: '30000000-0000-0000-0000-000000000001', tenantId }), { code: 'AUTH_REQUIRED' });
  assert.equal(await readPrivate('config.json', root), undefined);
});
