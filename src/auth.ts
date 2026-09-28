import { lstat, readFile } from 'node:fs/promises';
import { PublicClientApplication, type AccountInfo, type ICachePlugin } from '@azure/msal-node';
import { GraphClient } from './graph.js';
import { configRoot, readPrivate, removePrivate, writePrivate } from './storage.js';

const SCOPES = ['User.Read', 'Mail.Read'];
const GUID = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/;

export interface Identity { clientId: string; tenantId: string }
export interface SavedIdentity extends Identity { homeAccountId: string }

export function parseIdentity(clientId: string, tenantId: string): Identity {
  if (!GUID.test(clientId)) throw new Error('Invalid application (client) ID');
  if (!GUID.test(tenantId)) throw new Error('Invalid directory (tenant) ID');
  return { clientId, tenantId };
}

export function chooseAccount<T extends { homeAccountId: string }>(accounts: T[], homeAccountId: string): T {
  const account = accounts.find(a => a.homeAccountId === homeAccountId);
  if (!account) throw new Error('No cached account matches this login; run outlook auth login');
  return account;
}

export async function savedIdentity(root = configRoot()): Promise<SavedIdentity> {
  const data = await readPrivate('config.json', root);
  if (!data) throw new Error('Not signed in; run outlook auth login --client-id <id> --tenant-id <id>');
  let parsed: SavedIdentity;
  try { parsed = JSON.parse(data) as SavedIdentity; } catch { throw new Error('Invalid Outlook CLI configuration'); }
  parseIdentity(parsed.clientId, parsed.tenantId);
  if (!parsed.homeAccountId || typeof parsed.homeAccountId !== 'string') throw new Error('Invalid saved account');
  return parsed;
}

function createApp(identity: Identity, root: string): PublicClientApplication {
  const cachePlugin: ICachePlugin = {
    beforeCacheAccess: async context => {
      const data = await readPrivate('cache.json', root);
      if (data) context.tokenCache.deserialize(data);
    },
    afterCacheAccess: async context => {
      if (context.cacheHasChanged) await writePrivate('cache.json', context.tokenCache.serialize(), root);
    },
  };
  return new PublicClientApplication({
    auth: { clientId: identity.clientId, authority: `https://login.microsoftonline.com/${identity.tenantId}` },
    cache: { cachePlugin },
  });
}

export async function login(identity: Identity, notify: (message: string) => void, root = configRoot()) {
  const app = createApp(identity, root);
  const result = await app.acquireTokenByDeviceCode({ scopes: SCOPES, deviceCodeCallback: response => notify(response.message) });
  if (!result?.accessToken || !result.account) throw new Error('Microsoft sign-in did not complete');
  const me = await new GraphClient(result.accessToken).me();
  await writePrivate('config.json', JSON.stringify({ ...identity, homeAccountId: result.account.homeAccountId }), root);
  return me;
}

export async function graphForSavedLogin(root = configRoot()): Promise<GraphClient> {
  const identity = await savedIdentity(root);
  const app = createApp(identity, root);
  const accounts = await app.getTokenCache().getAllAccounts();
  const account = chooseAccount<AccountInfo>(accounts, identity.homeAccountId);
  try {
    const result = await app.acquireTokenSilent({ account, scopes: SCOPES });
    if (!result?.accessToken) throw new Error('No access token');
    return new GraphClient(result.accessToken);
  } catch {
    throw new Error('Microsoft session expired or revoked; run outlook auth login again');
  }
}

export async function status(root = configRoot()): Promise<{ signedIn: boolean; username?: string; tenantId?: string }> {
  const data = await readPrivate('config.json', root);
  if (!data) return { signedIn: false };
  const identity = await savedIdentity(root);
  const app = createApp(identity, root);
  const account = chooseAccount<AccountInfo>(await app.getTokenCache().getAllAccounts(), identity.homeAccountId);
  return { signedIn: true, username: account.username, tenantId: identity.tenantId };
}

export async function logout(root = configRoot()): Promise<void> {
  await removePrivate('cache.json', root);
  await removePrivate('config.json', root);
}

export async function importCache(source: string, identity: Identity, root = configRoot()) {
  if (await readPrivate('config.json', root) || await readPrivate('cache.json', root)) {
    throw new Error('Already configured; refusing to overwrite another account');
  }
  const info = await lstat(source);
  if (!info.isFile() || info.uid !== process.getuid?.() || (info.mode & 0o077)) {
    throw new Error('Source token cache must be an owner-only regular file');
  }
  const app = createApp(identity, root);
  app.getTokenCache().deserialize(await readFile(source, 'utf8'));
  const accounts = await app.getTokenCache().getAllAccounts();
  if (accounts.length !== 1) throw new Error('Source cache must contain exactly one signed-in account');
  const result = await app.acquireTokenSilent({ account: accounts[0]!, scopes: SCOPES });
  if (!result?.accessToken) throw new Error('Source cache cannot acquire a valid Microsoft Graph token');
  const me = await new GraphClient(result.accessToken).me();
  await writePrivate('cache.json', app.getTokenCache().serialize(), root);
  await writePrivate('config.json', JSON.stringify({ ...identity, homeAccountId: accounts[0]!.homeAccountId }), root);
  return me;
}
