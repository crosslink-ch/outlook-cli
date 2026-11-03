import { CliError } from './errors.js';

export type Mode = 'read-only' | 'read-write';
export interface Grant { mode: Mode; scopes: string[]; expiresAt: number }

export const OPERATIONS = {
  me: { scopes: ['User.Read'], write: false },
  list: { scopes: ['Mail.Read'], write: false },
  search: { scopes: ['Mail.Read'], write: false },
  read: { scopes: ['Mail.Read'], write: false },
  'calendar.list': { scopes: ['Calendars.Read', 'Calendars.ReadWrite'], write: false },
  'calendar.events': { scopes: ['Calendars.Read', 'Calendars.ReadWrite'], write: false },
  'calendar.get': { scopes: ['Calendars.Read', 'Calendars.ReadWrite'], write: false },
  'calendar.create': { scopes: ['Calendars.ReadWrite'], write: true },
  'calendar.update': { scopes: ['Calendars.ReadWrite'], write: true },
  'calendar.delete': { scopes: ['Calendars.ReadWrite'], write: true },
} as const;
export type Operation = keyof typeof OPERATIONS;

export function hasScope(scopes: string[], required: string): boolean {
  return scopes.some(scope => scope.replace(/^https:\/\/graph\.microsoft\.com\//i, '').toLowerCase() === required.toLowerCase());
}

export function requireOperation(grant: Grant | undefined, operation: Operation): void {
  if (!grant || !Number.isFinite(grant.expiresAt) || grant.expiresAt <= Date.now()) {
    throw new CliError('AUTH_REQUIRED', 'A valid Microsoft session is required; run outlook auth login');
  }
  if (OPERATIONS[operation].write && grant.mode === 'read-only') {
    throw new CliError('READ_ONLY_MODE', 'Calendar mutations require an authenticated read-write connection; use outlook auth login --mode read-write');
  }
  if (!OPERATIONS[operation].scopes.some(scope => hasScope(grant.scopes, scope))) {
    throw new CliError('INSUFFICIENT_SCOPE', `The authenticated token does not grant ${operation}; sign in with the required delegated permissions`);
  }
}

export function availableOperations(grant?: Grant): Operation[] {
  return (Object.keys(OPERATIONS) as Operation[]).filter(operation => {
    try { requireOperation(grant, operation); return true; } catch { return false; }
  });
}
