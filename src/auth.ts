import { randomUUID } from 'node:crypto';
import { availableOperations, hasScope, type Grant, type Mode } from './capabilities.js';
import { CliError } from './errors.js';
import type { AuthenticationResult, DeviceCodeRequest, SilentFlowRequest } from '@azure/msal-node';
import { lstat, readFile } from 'node:fs/promises';
import { PublicClientApplication, type AccountInfo } from '@azure/msal-node';
import { GraphClient } from './graph.js';
import { configRoot, readPrivate, removePrivate, replaceConfig, writePrivate } from './storage.js';

const SCOPES = ['User.Read', 'Mail.Read'];
// REQUIRED_SCOPES is the minimum for a valid connection in a mode. MODE_SCOPES is what sign-in asks for:
// read-write additionally requests optional mail-write scopes, whose absence only hides the draft operations.
const REQUIRED_SCOPES: Record<Mode, string[]> = { 'read-only': SCOPES, 'read-write': [...SCOPES, 'Calendars.ReadWrite'] };
const MODE_SCOPES: Record<Mode, string[]> = { 'read-only': SCOPES, 'read-write': [...REQUIRED_SCOPES['read-write'], 'Mail.ReadWrite', 'Mail.Send'] };

interface Connection extends SavedIdentity { cacheFile: string; username: string; meId: string }
interface State {
  version: 2;
  apps: Partial<Record<Mode, Identity>>;
  connections: Partial<Record<Mode, Connection>>;
  active?: Mode;
}
export interface AuthApp {
  getTokenCache(): { deserialize(data: string): void; serialize(): string; getAllAccounts(): Promise<AccountInfo[]> };
  acquireTokenByDeviceCode(request: DeviceCodeRequest): Promise<AuthenticationResult | null>;
  acquireTokenSilent(request: SilentFlowRequest): Promise<AuthenticationResult | null>;
}
export interface AuthDependencies { createApp(identity: Identity): AuthApp; transport: typeof fetch }
const defaultDependencies: AuthDependencies = {
  createApp: identity => new PublicClientApplication({
    auth: { clientId: identity.clientId, authority: `https://login.microsoftonline.com/${identity.tenantId}` },
    system: { loggerOptions: { piiLoggingEnabled: false, loggerCallback: () => {} } },
  }),
  transport: (...args) => fetch(...args),
};

export class AuthManager {
  private readonly snapshots = new WeakMap<State, string | undefined>();
  constructor(readonly root = configRoot(), private readonly deps: AuthDependencies = defaultDependencies) {}

  private async load(): Promise<State> {
    const data = await readPrivate('config.json', this.root);
    const remember = (state: State) => { this.snapshots.set(state, data); return state; };
    if (!data) return remember({ version: 2, apps: {}, connections: {} });
    try {
      const parsed = JSON.parse(data);
      if (parsed.version === undefined) {
        const identity = parseIdentity(parsed.clientId, parsed.tenantId);
        if (typeof parsed.homeAccountId !== 'string' || !parsed.homeAccountId) throw new Error();
        return remember({ version: 2, apps: { 'read-only': identity }, active: 'read-only', connections: { 'read-only': { ...identity, homeAccountId: parsed.homeAccountId, cacheFile: 'cache.json', username: '', meId: '' } } });
      }
      if (parsed.version !== 2 || !parsed.apps || !parsed.connections) throw new Error();
      if (parsed.active !== undefined) parseMode(parsed.active);
      for (const [mode, identity] of Object.entries(parsed.apps) as [Mode, Identity][]) {
        parseMode(mode);
        parseIdentity(identity.clientId, identity.tenantId);
      }
      for (const [mode, c] of Object.entries(parsed.connections) as [Mode, Connection][]) {
        parseMode(mode);
        parseIdentity(c.clientId, c.tenantId);
        if (!c.homeAccountId || typeof c.homeAccountId !== 'string' || !/^(cache\.json|cache-[0-9a-f-]{36}\.json)$/.test(c.cacheFile)) throw new Error();
      }
      return remember(parsed as State);
    } catch { throw new CliError('INVALID_CONFIG', 'Invalid Outlook CLI configuration; existing files were not changed'); }
  }

  private checkApp(state: State, mode: Mode, identity: Identity): void {
    parseMode(mode);
    parseIdentity(identity.clientId, identity.tenantId);
    const other = mode === 'read-only' ? 'read-write' : 'read-only';
    if ([state.apps[other], state.connections[other]].some(app => app?.clientId.toLowerCase() === identity.clientId.toLowerCase())) {
      throw new CliError('INVALID_CONFIG', 'Read-only and read-write modes must use distinct application client IDs');
    }
  }

  private validateResult(result: AuthenticationResult | null, identity: Identity, mode: Mode): Grant {
    if (!result?.accessToken || !result.account?.homeAccountId || !result.expiresOn || !Number.isFinite(result.expiresOn.getTime()) || result.expiresOn.getTime() <= Date.now()) {
      throw new CliError('AUTH_REQUIRED', 'Microsoft session expired or sign-in did not complete; run outlook auth login');
    }
    if (result.tenantId?.toLowerCase() !== identity.tenantId.toLowerCase() || result.account.tenantId?.toLowerCase() !== identity.tenantId.toLowerCase()) {
      throw new CliError('IDENTITY_MISMATCH', 'Microsoft returned a different tenant');
    }
    if (!Array.isArray(result.scopes) || REQUIRED_SCOPES[mode].some(scope => !hasScope(result.scopes, scope))) {
      throw new CliError('INSUFFICIENT_SCOPE', 'Microsoft did not grant all required permissions for this mode');
    }
    return { mode, scopes: result.scopes, expiresAt: result.expiresOn.getTime() };
  }

  // Connections authorized before the optional scopes existed (or without consent for them) cannot silently
  // acquire the full set; they stay valid on the required minimum instead of forcing a new login.
  private async silent(app: AuthApp, account: AccountInfo, mode: Mode): Promise<AuthenticationResult | null> {
    try { return await app.acquireTokenSilent({ account, scopes: MODE_SCOPES[mode] }); }
    catch (error) {
      if (MODE_SCOPES[mode].length === REQUIRED_SCOPES[mode].length) throw error;
      return app.acquireTokenSilent({ account, scopes: REQUIRED_SCOPES[mode] });
    }
  }

  private async verify(result: AuthenticationResult | null, identity: Identity, mode: Mode) {
    const grant = this.validateResult(result, identity, mode);
    const graph = new GraphClient(result!.accessToken, this.deps.transport, grant);
    const me = await graph.me();
    this.validateResult(result, identity, mode); // /me may outlast a near-expiry access token.
    if (!me?.id || me.id !== (result!.uniqueId || result!.account!.localAccountId)) {
      throw new CliError('IDENTITY_MISMATCH', 'Microsoft Graph identity does not match the authenticated account');
    }
    return { graph, me, grant };
  }

  async configure(mode: Mode, identity: Identity): Promise<void> {
    const state = await this.load();
    this.checkApp(state, mode, identity);
    state.apps[mode] = identity;
    await replaceConfig(this.snapshots.get(state), JSON.stringify(state), this.root);
  }

  async login(mode: Mode, notify: (message: string) => void, identity?: Identity) {
    parseMode(mode);
    const state = await this.load();
    const selected = identity ?? state.apps[mode];
    if (!selected) throw new CliError('APP_NOT_CONFIGURED', 'Configure the application and tenant IDs for this mode first');
    this.checkApp(state, mode, selected);
    const app = this.deps.createApp(selected);
    let result: AuthenticationResult | null;
    try { result = await app.acquireTokenByDeviceCode({ scopes: MODE_SCOPES[mode], deviceCodeCallback: response => notify(response.message) }); }
    catch { throw new CliError('AUTH_REQUIRED', 'Microsoft sign-in failed or was cancelled; existing credentials were not changed'); }
    const { me } = await this.verify(result, selected, mode);
    state.apps[mode] = selected;
    await this.commit(state, mode, selected, result!, app, me.id);
    return me;
  }

  private async commit(state: State, mode: Mode, selected: Identity, result: AuthenticationResult, app: AuthApp, meId: string) {
    const oldFile = state.connections[mode]?.cacheFile;
    const cacheFile = `cache-${randomUUID()}.json`;
    await writePrivate(cacheFile, app.getTokenCache().serialize(), this.root);
    state.connections[mode] = { ...selected, homeAccountId: result.account!.homeAccountId, cacheFile, username: result.account!.username, meId };
    state.active = mode;
    try { await replaceConfig(this.snapshots.get(state), JSON.stringify(state), this.root); }
    catch (error) { await removePrivate(cacheFile, this.root); throw error; }
    if (oldFile) await removePrivate(oldFile, this.root).catch(() => {});
  }

  private async cached(state: State, mode: Mode) {
    const identity = state.connections[mode];
    if (!identity) throw new CliError('AUTH_REQUIRED', 'No authorized cached connection for this mode; run outlook auth login');
    const app = this.deps.createApp(identity);
    let result: AuthenticationResult | null;
    let cacheBefore = '';
    try {
      const data = await readPrivate(identity.cacheFile, this.root);
      if (!data) throw new Error('Missing cache');
      app.getTokenCache().deserialize(data);
      cacheBefore = app.getTokenCache().serialize();
      const account = chooseAccount(await app.getTokenCache().getAllAccounts(), identity.homeAccountId);
      result = await this.silent(app, account, mode);
    } catch { throw new CliError('AUTH_REQUIRED', 'Microsoft session expired or revoked; run outlook auth login again'); }
    const verified = await this.verify(result, identity, mode);
    if (result!.account!.homeAccountId !== identity.homeAccountId || (identity.meId && verified.me.id !== identity.meId)) {
      throw new CliError('IDENTITY_MISMATCH', 'Cached Microsoft identity changed; sign in explicitly');
    }
    return { ...verified, identity: { ...identity, username: result!.account!.username, meId: verified.me.id }, app, result: result!, cacheChanged: cacheBefore !== app.getTokenCache().serialize() };
  }

  async use(mode: Mode) {
    parseMode(mode);
    const state = await this.load();
    const session = await this.cached(state, mode);
    await this.commit(state, mode, session.identity, session.result, session.app, session.me.id);
    return session.me;
  }

  async savedIdentity(): Promise<SavedIdentity> {
    const state = await this.load();
    const connection = state.active && state.connections[state.active];
    if (!connection) throw new CliError('AUTH_REQUIRED', 'Not signed in; run outlook auth login');
    return { clientId: connection.clientId, tenantId: connection.tenantId, homeAccountId: connection.homeAccountId };
  }

  async connection() {
    const state = await this.load();
    if (!state.active) throw new CliError('AUTH_REQUIRED', 'Not signed in; run outlook auth login');
    const session = await this.cached(state, state.active);
    const latest = await this.load();
    if (this.snapshots.get(latest) !== this.snapshots.get(state)) {
      throw new CliError('CONCURRENT_AUTH_CHANGE', 'Authentication changed during verification; retry without switching modes');
    }
    if (session.cacheChanged) await this.commit(state, state.active, session.identity, session.result, session.app, session.me.id);
    return { graph: session.graph, grant: session.grant, identity: session.identity, me: session.me };
  }

  async importCache(source: string, mode: Mode = 'read-only', identity?: Identity) {
    parseMode(mode);
    const state = await this.load();
    if (state.connections[mode]) throw new CliError('ALREADY_CONFIGURED', 'Already configured for this mode; refusing to overwrite another account');
    const selected = identity ?? state.apps[mode];
    if (!selected) throw new CliError('APP_NOT_CONFIGURED', 'Configure application and tenant IDs for this mode first');
    this.checkApp(state, mode, selected);
    const info = await lstat(source);
    if (!info.isFile() || info.uid !== process.getuid?.() || (info.mode & 0o077)) {
      throw new CliError('UNSAFE_CACHE', 'Source token cache must be an owner-only regular file');
    }
    const app = this.deps.createApp(selected);
    let result: AuthenticationResult | null;
    try {
      app.getTokenCache().deserialize(await readFile(source, 'utf8'));
      const accounts = await app.getTokenCache().getAllAccounts();
      if (accounts.length !== 1) throw new CliError('INVALID_CACHE', 'Source cache must contain exactly one signed-in account');
      result = await this.silent(app, accounts[0]!, mode);
    } catch (error) {
      if (error instanceof CliError) throw error;
      throw new CliError('AUTH_REQUIRED', 'Source cache cannot acquire a valid token for this application');
    }
    const { me } = await this.verify(result, selected, mode);
    state.apps[mode] = selected;
    await this.commit(state, mode, selected, result!, app, me.id);
    return me;
  }

  async logout(): Promise<void> {
    const state = await this.load();
    await replaceConfig(this.snapshots.get(state), undefined, this.root);
    for (const c of Object.values(state.connections)) await removePrivate(c.cacheFile, this.root);
    await removePrivate('cache.json', this.root);
  }

  async status(): Promise<AuthStatus> {
    try {
      const { identity, grant, me } = await this.connection();
      return { signedIn: true, mode: grant.mode, username: identity.username, clientId: identity.clientId, tenantId: identity.tenantId, homeAccountId: identity.homeAccountId, me, scopes: grant.scopes, expiresAt: grant.expiresAt, operations: availableOperations(grant) };
    } catch (error) {
      return { signedIn: false, operations: [], error: { code: error instanceof CliError ? error.code : 'AUTH_REQUIRED', message: error instanceof CliError ? error.message : 'Cannot validate the saved Microsoft connection' } };
    }
  }
}

export interface AuthStatus {
  signedIn: boolean;
  mode?: Mode;
  username?: string;
  clientId?: string;
  tenantId?: string;
  homeAccountId?: string;
  me?: Awaited<ReturnType<GraphClient['me']>>;
  scopes?: string[];
  expiresAt?: number;
  operations: string[];
  error?: { code: string; message: string };
}

const GUID = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/;

export function parseMode(value: string): Mode {
  if (value !== 'read-only' && value !== 'read-write') throw new CliError('INVALID_INPUT', 'Mode must be read-only or read-write');
  return value;
}

export interface Identity { clientId: string; tenantId: string }
export interface SavedIdentity extends Identity { homeAccountId: string }

export function parseIdentity(clientId: string, tenantId: string): Identity {
  if (!GUID.test(clientId)) throw new CliError('INVALID_INPUT', 'Invalid application (client) ID');
  if (!GUID.test(tenantId)) throw new CliError('INVALID_INPUT', 'Invalid directory (tenant) ID');
  return { clientId, tenantId };
}

export function chooseAccount<T extends { homeAccountId: string }>(accounts: T[], homeAccountId: string): T {
  const account = accounts.find(a => a.homeAccountId === homeAccountId);
  if (!account) throw new Error('No cached account matches this login; run outlook auth login');
  return account;
}

// Compatibility exports retain the original signatures; old login/import default to read-only.
export async function savedIdentity(root = configRoot()): Promise<SavedIdentity> {
  return new AuthManager(root).savedIdentity();
}
export async function login(identity: Identity, notify: (message: string) => void, root = configRoot()) {
  return new AuthManager(root).login('read-only', notify, identity);
}
export async function graphForSavedLogin(root = configRoot(), deps?: AuthDependencies): Promise<GraphClient> {
  return (await new AuthManager(root, deps).connection()).graph;
}
export async function status(root = configRoot(), deps?: AuthDependencies): Promise<AuthStatus> {
  return new AuthManager(root, deps).status();
}
export async function logout(root = configRoot()): Promise<void> {
  return new AuthManager(root).logout();
}
export async function importCache(source: string, identity: Identity, root = configRoot()) {
  return new AuthManager(root).importCache(source, 'read-only', identity);
}
