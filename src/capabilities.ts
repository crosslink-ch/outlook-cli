import { CliError } from './errors.js';

export type Mode = 'read-only' | 'read-write';
export interface Grant { mode: Mode; scopes: string[]; expiresAt: number }

// `scopes` grants an operation when ANY listed scope is present; `all: true` requires EVERY listed scope.
interface OperationSpec { scopes: readonly string[]; all?: boolean; write: boolean }

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
  'draft.create': { scopes: ['Mail.ReadWrite'], write: true },
  'draft.reply': { scopes: ['Mail.ReadWrite'], write: true },
  'draft.update': { scopes: ['Mail.ReadWrite'], write: true },
  'draft.send': { scopes: ['Mail.ReadWrite', 'Mail.Send'], all: true, write: true },
} as const satisfies Record<string, OperationSpec>;
export type Operation = keyof typeof OPERATIONS;

export function hasScope(scopes: string[], required: string): boolean {
  return scopes.some(scope => scope.replace(/^https:\/\/graph\.microsoft\.com\//i, '').toLowerCase() === required.toLowerCase());
}

export function requireOperation(grant: Grant | undefined, operation: Operation): void {
  if (!grant || !Number.isFinite(grant.expiresAt) || grant.expiresAt <= Date.now()) {
    throw new CliError('AUTH_REQUIRED', 'A valid Microsoft session is required; run outlook auth login');
  }
  const spec: OperationSpec = OPERATIONS[operation];
  if (spec.write && grant.mode === 'read-only') {
    throw new CliError('READ_ONLY_MODE', 'Write operations require an authenticated read-write connection; use outlook auth login --mode read-write');
  }
  const granted = (scope: string) => hasScope(grant.scopes, scope);
  if (!(spec.all ? spec.scopes.every(granted) : spec.scopes.some(granted))) {
    throw new CliError('INSUFFICIENT_SCOPE', `The authenticated token does not grant ${operation}; sign in with the required delegated permissions`);
  }
}

export function availableOperations(grant?: Grant): Operation[] {
  return (Object.keys(OPERATIONS) as Operation[]).filter(operation => {
    try { requireOperation(grant, operation); return true; } catch { return false; }
  });
}
