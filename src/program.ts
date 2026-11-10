import { Command, CommanderError, InvalidArgumentError } from 'commander';
import { open, stat } from 'node:fs/promises';
import { AuthManager, parseIdentity, parseMode, type Identity } from './auth.js';
import { availableOperations, requireOperation, type Operation } from './capabilities.js';
import { CliError } from './errors.js';
import type { Draft, DraftInput, EventInput, MessageDetail, MessageSummary, Recipient } from './graph.js';

type Session = Awaited<ReturnType<AuthManager['connection']>>;
export interface IO { out(text: string): void; err(text: string): void }
interface ProgramOptions { manager?: AuthManager; session?: Session; authError?: CliError; io?: IO }
const defaultIO: IO = { out: text => process.stdout.write(text), err: text => process.stderr.write(text) };

function limitArg(value: string): number {
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed < 1 || parsed > 100) throw new InvalidArgumentError('must be an integer from 1 to 100');
  return parsed;
}
function compact(value: string | undefined): string { return (value || '').replace(/[\r\n\t]+/g, ' ').slice(0, 110); }
function summary(messages: MessageSummary[]): string {
  return messages.length ? messages.map(m => [m.receivedDateTime || '(no date)', compact(m.from?.emailAddress?.address), compact(m.subject || '(no subject)'), `id=${m.id}`].join('  |  ')).join('\n') : 'No messages found.';
}
function detail(m: MessageDetail): string {
  return [`Subject: ${compact(m.subject)}`, `From: ${compact(m.from?.emailAddress?.address)}`, `Received: ${m.receivedDateTime || ''}`, `ID: ${m.id}`, '', m.body?.content || m.bodyPreview || ''].join('\n');
}
function optionalIdentity(opts: { clientId?: string; tenantId?: string }): Identity | undefined {
  if (opts.clientId === undefined && opts.tenantId === undefined) return undefined;
  if (!opts.clientId || !opts.tenantId) throw new CliError('INVALID_INPUT', 'Supply both --client-id and --tenant-id, or configure the mode first');
  return parseIdentity(opts.clientId, opts.tenantId);
}
function fields(opts: Partial<EventInput>): Partial<EventInput> {
  return Object.fromEntries(['title', 'start', 'end', 'timezone', 'location', 'description'].filter(key => opts[key as keyof EventInput] !== undefined).map(key => [key, opts[key as keyof EventInput]]));
}
// Repeatable and comma-separated; an empty value yields an empty list (clears that field on update).
function recipientArg(value: string, previous: string[] = []): string[] {
  return [...previous, ...value.split(',').map(item => item.trim()).filter(Boolean)];
}
const MAX_BODY_FILE = 1_048_576;
async function bodyArg(opts: { body?: string; bodyFile?: string }, required: boolean): Promise<string | undefined> {
  if (opts.body !== undefined && opts.bodyFile !== undefined) throw new CliError('INVALID_INPUT', 'Supply only one of --body and --body-file');
  if (opts.bodyFile === undefined) {
    if (required && opts.body === undefined) throw new CliError('INVALID_INPUT', 'Supply exactly one of --body and --body-file');
    return opts.body;
  }
  const invalid = new CliError('INVALID_INPUT', '--body-file must be a readable UTF-8 regular file of at most 1 MB');
  let handle;
  // Check the type before opening: opening a FIFO or device could block indefinitely.
  try { if (!(await stat(opts.bodyFile)).isFile()) throw invalid; handle = await open(opts.bodyFile, 'r'); } catch { throw invalid; }
  try {
    const info = await handle.stat();
    if (!info.isFile() || info.size > MAX_BODY_FILE) throw invalid;
    const data = await handle.readFile();
    if (data.length > MAX_BODY_FILE) throw invalid;
    return new TextDecoder('utf-8', { fatal: true }).decode(data);
  } catch { throw invalid; } finally { await handle.close(); }
}
function draftFields(opts: Partial<DraftInput>, body: string | undefined): Partial<DraftInput> {
  const input: Partial<DraftInput> = { ...opts, body };
  return Object.fromEntries((['to', 'cc', 'bcc', 'subject', 'body', 'html'] as const).filter(key => input[key] !== undefined).map(key => [key, input[key]]));
}
function draftText(result: { outcome: string; id: string; draft?: Draft; subject?: string | null; recipients?: Record<'to' | 'cc' | 'bcc', string[]> }): string {
  const list = (items?: Recipient[]) => (items ?? []).map(item => item.emailAddress?.address ?? '');
  const d = result.draft;
  const recipients = result.recipients ?? { to: list(d?.toRecipients), cc: list(d?.ccRecipients), bcc: list(d?.bccRecipients) };
  return [`Draft ${result.outcome}: id=${result.id}`, `To: ${compact(recipients.to.join(', '))}`, `Cc: ${compact(recipients.cc.join(', '))}`, `Bcc: ${compact(recipients.bcc.join(', '))}`,
    `Subject: ${compact((d ? d.subject : result.subject) ?? '')}`, ...(d?.bodyPreview ? [`Preview: ${compact(d.bodyPreview)}`] : [])].join('\n');
}
function report(session: Session | undefined, error?: CliError) {
  if (!session) return { signedIn: false, operations: [], error: { code: error?.code ?? 'AUTH_REQUIRED', message: error?.message ?? 'Not signed in; run outlook auth login' } };
  const { grant, identity, me } = session;
  return { signedIn: true, mode: grant.mode, username: identity.username, clientId: identity.clientId, tenantId: identity.tenantId, homeAccountId: identity.homeAccountId, me, scopes: grant.scopes, expiresAt: grant.expiresAt, operations: availableOperations(grant) };
}

// Synchronous construction remains compatible with callers inspecting basic help.
// Use run() for authenticated, contextual discovery and machine-readable failures.
export function createProgram(options: ProgramOptions = {}): Command {
  const manager = options.manager ?? new AuthManager();
  const io = options.io ?? defaultIO;
  const session = options.session;
  const program = new Command().name('outlook')
    .description('Outlook mail, draft and calendar CLI: authenticated read-only / read-write modes')
    .version('0.3.0').option('--json', 'Output JSON; existing mail/auth commands retain their raw success format').exitOverride();
  program.configureOutput({ writeOut: text => io.out(text), writeErr: () => {}, outputError: () => {} });
  const emit = (value: unknown, human?: string) => {
    // Preserve the machine contract consumed by existing mail automation.
    const legacy = ['me', 'list', 'search', 'read'].includes(program.args[0] ?? '')
      || (program.args[0] === 'auth' && ['status', 'login', 'logout', 'import-cache'].includes(program.args[1] ?? ''));
    io.out((program.opts().json ? JSON.stringify(legacy ? value : { ok: true, data: value }, null, 2) : human ?? JSON.stringify(value, null, 2)) + '\n');
  };
  const graph = (operation: Operation) => {
    requireOperation(session?.grant, operation);
    return session!.graph;
  };
  const accountText = (me: { mail?: string; userPrincipalName?: string }) => `Signed in as ${me.mail || me.userPrincipalName || '(verified account)'}`;
  const statusAction = () => {
    const result = report(session, options.authError);
    emit(result, result.signedIn ? `${result.mode}: ${result.username}\nOperations: ${result.operations.join(', ')}` : 'No verified Microsoft session. Run outlook auth login.');
  };

  const auth = program.command('auth').description('Configure public-client apps and manage private delegated connections');
  auth.command('configure').description('Save nonsecret app/tenant IDs for the next login; does not activate a connection')
    .requiredOption('--mode <mode>', 'read-only or read-write', parseMode)
    .requiredOption('--client-id <id>', 'Entra application (client) ID')
    .requiredOption('--tenant-id <id>', 'Entra directory (tenant) ID')
    .action(async opts => {
      const identity = parseIdentity(opts.clientId, opts.tenantId);
      await manager.configure(opts.mode, identity);
      emit({ mode: opts.mode, ...identity }, `Application configured for ${opts.mode}; active connection unchanged.`);
    });
  auth.command('login').description('Device-code sign-in; activates only after token scopes and Graph identity verification')
    .option('--mode <mode>', 'read-only or read-write', parseMode, 'read-only')
    .option('--client-id <id>', 'First setup: application ID (requires tenant ID)')
    .option('--tenant-id <id>', 'First setup: tenant ID (requires application ID)')
    .action(async opts => {
      const me = await manager.login(opts.mode, message => io.err(message + '\n'), optionalIdentity(opts));
      emit(me, accountText(me));
    });
  auth.command('import-cache').description('Validate and privately import a single-account MSAL Node/Python cache for one mode')
    .option('--mode <mode>', 'read-only or read-write', parseMode, 'read-only')
    .requiredOption('--from <path>', 'Owner-only MSAL JSON cache file; never inline tokens')
    .option('--client-id <id>', 'Application ID, if not configured')
    .option('--tenant-id <id>', 'Tenant ID, if not configured')
    .action(async opts => {
      const me = await manager.importCache(opts.from, opts.mode, optionalIdentity(opts));
      emit(me, accountText(me));
    });
  auth.command('use').description('Explicitly activate a previously authorized mode after silent verification')
    .requiredOption('--mode <mode>', 'read-only or read-write', parseMode)
    .action(async opts => { const me = await manager.use(opts.mode); emit(me, accountText(me)); });
  auth.command('status').description('Verify active identity and granted capabilities with Microsoft (network required)').action(statusAction);
  auth.command('logout').description('Remove all local connections/configuration; does not revoke tenant consent')
    .action(async () => { await manager.logout(); emit({ signedIn: false }, 'Local sign-ins removed.'); });
  program.command('capabilities').description('Verify the active mode, identity and exact available operations').action(statusAction);

  program.command('me').description('Identify the authenticated Microsoft Graph account').action(() => {
    graph('me');
    emit(session!.me, `${session!.me.displayName || ''} <${session!.me.mail || session!.me.userPrincipalName || ''}>`);
  });
  program.command('list').description('List read-only message summaries from a mail folder')
    .option('-f, --folder <name>', 'Well-known folder name or folder ID', 'inbox')
    .option('-n, --limit <count>', 'Maximum messages, 1–100', limitArg, 20)
    .action(async opts => { const value = await graph('list').list(opts.folder, opts.limit); emit(value, summary(value)); });
  program.command('search').description('Search mail summaries for a phrase').argument('<query>', 'Search phrase')
    .option('-n, --limit <count>', 'Maximum messages, 1–100', limitArg, 20)
    .action(async (query, opts) => { const value = await graph('search').search(query, opts.limit); emit(value, summary(value)); });
  program.command('read').description('Read a message text body without changing read/unread status').argument('<id>', 'Message ID')
    .action(async id => { const value = await graph('read').read(id); emit(value, detail(value)); });

  const available = availableOperations(session?.grant);
  // One registry drives both discovery and enforcement: ungranted verbs are hidden and refuse before any option parsing.
  const guardedGroup = (name: 'calendar' | 'draft', description: string) => {
    const group = program.command(name).description(description);
    group.hook('preSubcommand', (_parent, cmd) => {
      if (cmd.name() !== 'help') requireOperation(session?.grant, `${name}.${cmd.name()}` as Operation);
    });
    group.addHelpText('after', session ? `\nActive mode: ${session.grant.mode}. Use outlook capabilities --json for exact permissions.\n` : '\nNo verified connection. Use outlook auth login, then outlook capabilities --json.\n');
    return (verb: string, text: string) => {
      const operation = `${name}.${verb}` as Operation;
      const cmd = group.command(verb, { hidden: !available.includes(operation) }).description(text);
      const help = cmd.helpInformation.bind(cmd);
      cmd.helpInformation = () => { requireOperation(session?.grant, operation); return help(); };
      return cmd;
    };
  };
  const calendarCommand = guardedGroup('calendar', 'Calendar operations available to the currently verified connection');
  calendarCommand('list', 'List calendars (bounded to 1–100)')
    .option('--limit <count>', 'Maximum calendars', limitArg, 100)
    .action(async opts => emit(await graph('calendar.list').calendars(opts.limit)));
  calendarCommand('events', 'List calendarView events in an offset-aware window of at most 366 days')
    .option('--start <datetime>', 'ISO datetime with offset, e.g. 2027-01-01T00:00:00Z')
    .option('--end <datetime>', 'Exclusive ISO window end with offset')
    .option('--limit <count>', 'Maximum events, 1–100', limitArg, 100)
    .option('--calendar <id>', 'Calendar ID; default is primary calendar')
    .action(async opts => emit(await graph('calendar.events').events(opts.start, opts.end, opts.limit, opts.calendar)));
  calendarCommand('get', 'Get one event by its opaque ID')
    .option('--event <id>', 'Event ID').option('--calendar <id>', 'Calendar ID')
    .action(async opts => emit(await graph('calendar.get').event(opts.event, opts.calendar)));
  for (const name of ['create', 'update'] as const) {
    const cmd = calendarCommand(name, name === 'create' ? 'Create an attendee-free timed event and verify by readback' : 'Update named event fields with ETag protection and readback')
      .option('--title <text>', 'Event subject')
      .option('--start <datetime>', 'Local YYYY-MM-DDTHH:mm:ss (no offset)')
      .option('--end <datetime>', 'Local YYYY-MM-DDTHH:mm:ss; later than start')
      .option('--timezone <zone>', 'UTC or IANA timezone; required with start/end')
      .option('--location <text>', 'Location display name')
      .option('--description <text>', 'Plain-text event body')
      .option('--calendar <id>', 'Calendar ID; default is primary calendar');
    if (name === 'update') cmd.option('--event <id>', 'Event ID').option('--notify-attendees', 'Explicitly approve possible meeting update notifications');
    cmd.action(async opts => {
      const client = graph(`calendar.${name}`);
      emit(name === 'create' ? await client.createEvent(fields(opts) as EventInput, opts.calendar) : await client.updateEvent(opts.event, fields(opts), { calendar: opts.calendar, notifyAttendees: opts.notifyAttendees }));
    });
  }
  calendarCommand('delete', 'Delete an event with ETag protection and verify HTTP 404')
    .option('--event <id>', 'Event ID').option('--calendar <id>', 'Calendar ID')
    .option('--notify-attendees', 'Explicitly approve possible meeting cancellation notifications')
    .action(async opts => emit(await graph('calendar.delete').deleteEvent(opts.event, { calendar: opts.calendar, notifyAttendees: opts.notifyAttendees })));

  const draftCommand = guardedGroup('draft', 'Mail draft operations available to the currently verified connection');
  const recipientOptions = (cmd: Command) => cmd
    .option('--to <address>', 'Recipient address; repeatable or comma-separated', recipientArg)
    .option('--cc <address>', 'Cc address; repeatable or comma-separated', recipientArg)
    .option('--bcc <address>', 'Bcc address; repeatable or comma-separated', recipientArg);
  const bodyOptions = (cmd: Command) => cmd
    .option('--body <text>', 'Body text')
    .option('--body-file <path>', 'Read the body from a UTF-8 file (max 1 MB) instead of --body');
  bodyOptions(recipientOptions(draftCommand('create', 'Create a new mail draft (nothing is sent) and verify by readback')).option('--subject <text>', 'Subject, single line'))
    .option('--html', 'Treat the body as HTML instead of plain text')
    .action(async opts => {
      const client = graph('draft.create');
      const result = await client.createDraft(draftFields(opts, await bodyArg(opts, true)) as DraftInput);
      emit(result, draftText(result));
    });
  bodyOptions(draftCommand('reply', 'Create a reply draft to a message (nothing is sent); Graph chooses the recipients')
    .option('--message <id>', 'ID of the message to reply to').option('--all', 'Reply to all recipients'))
    .action(async opts => {
      const client = graph('draft.reply');
      const result = await client.createReplyDraft(opts.message, { all: Boolean(opts.all), body: (await bodyArg(opts, true))! });
      emit(result, draftText(result));
    });
  bodyOptions(recipientOptions(draftCommand('update', 'Replace named fields of an existing draft and verify by readback').option('--draft <id>', 'Draft ID')).option('--subject <text>', 'Subject, single line'))
    .option('--html', 'Treat the body as HTML instead of plain text')
    .action(async opts => {
      const client = graph('draft.update');
      const result = await client.updateDraft(opts.draft, draftFields(opts, await bodyArg(opts, false)));
      emit(result, draftText(result));
    });
  draftCommand('send', 'Send one existing draft by ID; only after the user approved that exact email. Never retried')
    .option('--draft <id>', 'Draft ID')
    .action(async opts => { const result = await graph('draft.send').sendDraft(opts.draft); emit(result, draftText(result)); });
  program.addHelpText('after', '\nStart with outlook auth status --json and outlook capabilities --json.\nHelp only advertises calendar and draft operations granted to the verified active connection.\nModes are selected during authentication, never by a flag on an operation.\n');
  return program;
}

export async function run(args: string[], options: { manager?: AuthManager; io?: IO } = {}): Promise<number> {
  const manager = options.manager ?? new AuthManager();
  const io = options.io ?? defaultIO;
  let session: Session | undefined;
  let authError: CliError | undefined;
  // Auth setup/recovery must not depend on an old active connection succeeding.
  const words = args.filter(arg => arg !== '--json');
  const setup = words[0] === 'auth' && words[1] !== 'status';
  if (!setup && !words.includes('--version') && !words.includes('-V')) {
    try { session = await manager.connection(); }
    catch (error) { authError = error instanceof CliError ? error : new CliError('AUTH_REQUIRED', 'No verified Microsoft session; run outlook auth login'); }
  }
  try {
    await createProgram({ manager, session, authError, io }).parseAsync(args, { from: 'user' });
    return (words[0] === 'capabilities' || (words[0] === 'auth' && words[1] === 'status')) && !session ? 1 : 0;
  } catch (error) {
    if (error instanceof CommanderError && error.exitCode === 0) return 0;
    const safe = error instanceof CliError ? error : error instanceof CommanderError
      ? new CliError('INVALID_INPUT', 'Invalid command arguments; run outlook --help or the command-specific --help')
      : new CliError('INTERNAL_ERROR', 'Operation failed; no sensitive diagnostics are printed');
    if (args.includes('--json')) io.out(JSON.stringify({ ok: false, error: { code: safe.code, message: safe.message, ...(safe.details ? { details: safe.details } : {}) } }) + '\n');
    else io.err(`${safe.code}: ${safe.message}\n${safe.details ? JSON.stringify(safe.details) + '\n' : ''}`);
    return 1;
  }
}
