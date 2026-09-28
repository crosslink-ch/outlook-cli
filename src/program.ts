import { Command, InvalidArgumentError } from 'commander';
import { graphForSavedLogin, importCache, login, logout, parseIdentity, status } from './auth.js';
import type { MessageDetail, MessageSummary } from './graph.js';

function limitArg(value: string): number {
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed < 1 || parsed > 100) {
    throw new InvalidArgumentError('must be an integer from 1 to 100');
  }
  return parsed;
}

function compact(value: string | undefined): string {
  return (value || '').replace(/[\r\n\t]+/g, ' ').slice(0, 110);
}

function print(program: Command, value: unknown, format: (value: never) => string): void {
  if (program.opts().json) console.log(JSON.stringify(value, null, 2));
  else console.log(format(value as never));
}

function summary(messages: MessageSummary[]): string {
  if (messages.length === 0) return 'No messages found.';
  return messages.map(m => [
    m.receivedDateTime || '(no date)',
    compact(m.from?.emailAddress?.address),
    compact(m.subject || '(no subject)'),
    `id=${m.id}`,
  ].join('  |  ')).join('\n');
}

function detail(m: MessageDetail): string {
  return [
    `Subject: ${compact(m.subject)}`,
    `From: ${compact(m.from?.emailAddress?.address)}`,
    `Received: ${m.receivedDateTime || ''}`,
    `ID: ${m.id}`,
    '',
    m.body?.content || m.bodyPreview || '',
  ].join('\n');
}

export function createProgram(): Command {
  const program = new Command();
  program.name('outlook').description('Read-only Outlook mail CLI using Microsoft Graph and delegated sign-in')
    .version('0.1.0').option('--json', 'Output machine-readable JSON');
  program.showHelpAfterError('(Run outlook --help for usage.)');
  program.addHelpText('after', '\nExamples:\n  outlook auth login --client-id <app-id> --tenant-id <tenant-id>\n  outlook me --json\n  outlook list --folder inbox --limit 10 --json\n  outlook search "invoice" --limit 10 --json\n  outlook read <message-id> --json\n');

  const auth = program.command('auth').description('Manage the local Microsoft delegated sign-in');
  auth.command('login').description('Sign in with a Microsoft device code; caches tokens privately')
    .requiredOption('--client-id <id>', 'Entra application (client) ID')
    .requiredOption('--tenant-id <id>', 'Entra directory (tenant) ID')
    .action(async (opts: { clientId: string; tenantId: string }) => {
      const me = await login(parseIdentity(opts.clientId, opts.tenantId), message => console.log(message));
      print(program, me, (m: { mail?: string; userPrincipalName?: string }) => `Signed in as ${m.mail || m.userPrincipalName}`);
    });
  auth.command('import-cache').description('Privately import an existing single-account MSAL token cache')
    .requiredOption('--from <path>', 'Path to an owner-only MSAL JSON cache file')
    .requiredOption('--client-id <id>', 'Entra application (client) ID')
    .requiredOption('--tenant-id <id>', 'Entra directory (tenant) ID')
    .action(async (opts: { from: string; clientId: string; tenantId: string }) => {
      const me = await importCache(opts.from, parseIdentity(opts.clientId, opts.tenantId));
      print(program, me, (m: { mail?: string; userPrincipalName?: string }) => `Imported sign-in for ${m.mail || m.userPrincipalName}`);
    });
  auth.command('status').description('Show the locally signed-in account (does not contact Microsoft)')
    .action(async () => {
      const result = await status();
      print(program, result, (s: { signedIn: boolean; username?: string }) => s.signedIn ? `Cached login: ${s.username}` : 'Not signed in');
    });
  auth.command('logout').description('Remove the local account and token cache (does not revoke tenant consent)')
    .action(async () => { await logout(); print(program, { signedIn: false }, () => 'Local sign-in removed.'); });
  auth.addHelpText('after', '\nExamples:\n  outlook auth login --client-id <app-id> --tenant-id <tenant-id>\n  outlook auth status --json\n  outlook auth import-cache --from /path/to/msal-cache.json --client-id <app-id> --tenant-id <tenant-id>\n');

  program.command('me').description('Identify the account authenticated with Microsoft Graph')
    .action(async () => {
      const value = await (await graphForSavedLogin()).me();
      print(program, value, (m: { displayName?: string; mail?: string; userPrincipalName?: string }) =>
        `${m.displayName || ''} <${m.mail || m.userPrincipalName || ''}>`);
    });
  program.command('list').description('List message summaries from a mail folder (no bodies)')
    .option('-f, --folder <name>', 'Well-known folder name or folder ID', 'inbox')
    .option('-n, --limit <count>', 'Maximum messages, 1–100', limitArg, 20)
    .action(async (opts: { folder: string; limit: number }) => {
      const value = await (await graphForSavedLogin()).list(opts.folder, opts.limit);
      print(program, value, summary);
    });
  program.command('search').description('Search your mailbox for a phrase; return message summaries')
    .argument('<query>', 'Search text (phrase; quote it in your shell if it contains spaces)')
    .option('-n, --limit <count>', 'Maximum messages, 1–100', limitArg, 20)
    .action(async (query: string, opts: { limit: number }) => {
      const value = await (await graphForSavedLogin()).search(query, opts.limit);
      print(program, value, summary);
    });
  program.command('read').description('Read one message by ID, including its text body')
    .argument('<id>', 'Message ID returned by list or search')
    .action(async (id: string) => {
      const value = await (await graphForSavedLogin()).read(id);
      print(program, value, detail);
    });
  return program;
}
