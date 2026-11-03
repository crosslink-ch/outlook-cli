import { randomUUID } from 'node:crypto';
import { CliError } from './errors.js';
import { requireOperation, type Grant } from './capabilities.js';

const GRAPH_ORIGIN = 'https://graph.microsoft.com';
const GRAPH_ROOT = `${GRAPH_ORIGIN}/v1.0`;
const SUMMARY_FIELDS = 'id,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments';

export interface MessageSummary {
  id: string;
  subject?: string;
  from?: { emailAddress?: { name?: string; address?: string } };
  receivedDateTime?: string;
  bodyPreview?: string;
  isRead?: boolean;
  hasAttachments?: boolean;
}

interface Page<T> {
  value: T[];
  '@odata.nextLink'?: string;
}

export interface MessageDetail extends MessageSummary {
  body?: { contentType: string; content: string };
  toRecipients?: unknown[];
  ccRecipients?: unknown[];
  sentDateTime?: string;
  webLink?: string;
}

export interface Calendar { id: string; name?: string; canEdit?: boolean }
export interface CalendarEvent {
  id: string;
  subject?: string;
  start?: { dateTime: string; timeZone: string };
  end?: { dateTime: string; timeZone: string };
  body?: { contentType: string; content: string };
  location?: { displayName?: string };
  attendees?: unknown[];
  isAllDay?: boolean;
  '@odata.etag'?: string;
}

function encodedId(value: string): string {
  if (typeof value !== 'string' || !value.trim() || value.length > 1024 || value === '.' || value === '..') throw new CliError('INVALID_INPUT', 'Resource ID must be 1–1024 characters, not a dot segment');
  return encodeURIComponent(value);
}
function calendarPath(calendar?: string): string {
  return `${GRAPH_ROOT}/me${calendar === undefined ? '' : `/calendars/${encodedId(calendar)}`}`;
}
function checkLimit(limit: number): void {
  if (!Number.isSafeInteger(limit) || limit < 1 || limit > 100) throw new CliError('INVALID_INPUT', 'Limit must be 1–100');
}

export interface EventInput { title: string; start: string; end: string; timezone: string; location?: string; description?: string }
export interface MutationOptions { calendar?: string; notifyAttendees?: boolean }

function checkTimezone(timezone: string): void {
  try { new Intl.DateTimeFormat('en', { timeZone: timezone }); }
  catch { throw new CliError('INVALID_INPUT', 'Timezone must be UTC or an IANA timezone, such as Europe/Zurich'); }
}

function validLocalTime(value: string): boolean {
  return /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d$/.test(value) && Number.isFinite(Date.parse(`${value}Z`)) && new Date(`${value}Z`).toISOString().slice(0, 19) === value;
}

function eventBody(input: Partial<EventInput>, create = false): Partial<CalendarEvent> {
  const keys = ['title', 'start', 'end', 'timezone', 'location', 'description'];
  if (!input || Object.keys(input).some(key => !keys.includes(key)) || Object.entries(input).some(([, value]) => value !== undefined && typeof value !== 'string')) {
    throw new CliError('INVALID_INPUT', 'Only named event fields are supported; attendees and RSVP actions are not supported');
  }
  if ((create || input.title !== undefined) && (!input.title?.trim() || input.title.length > 255)) throw new CliError('INVALID_INPUT', 'Title must be 1–255 characters');
  if (create || input.start !== undefined || input.end !== undefined || input.timezone !== undefined) {
    if (!input.start || !input.end || !input.timezone || !validLocalTime(input.start) || !validLocalTime(input.end) || input.end <= input.start) {
      throw new CliError('INVALID_INPUT', 'Timed events require increasing start/end as YYYY-MM-DDTHH:mm:ss and a timezone (all three together)');
    }
    checkTimezone(input.timezone);
  }
  if (!create && !Object.values(input).some(value => value !== undefined)) throw new CliError('INVALID_INPUT', 'At least one event field is required');
  return {
    ...(input.title !== undefined ? { subject: input.title } : {}),
    ...(input.start !== undefined ? { start: { dateTime: input.start, timeZone: input.timezone! }, end: { dateTime: input.end!, timeZone: input.timezone! } } : {}),
    ...(input.location !== undefined ? { location: { displayName: input.location } } : {}),
    ...(input.description !== undefined ? { body: { contentType: 'text', content: input.description } } : {}),
  };
}

function verifyEvent(event: CalendarEvent, id: string, expected: Partial<CalendarEvent>): void {
  if (!event || event.id !== id) throw new Error('Wrong event identity');
  for (const key of ['subject', 'isAllDay'] as const) {
    if (expected[key] !== undefined && event[key] !== expected[key]) throw new Error('Event field mismatch');
  }
  for (const key of ['start', 'end'] as const) {
    if (expected[key] && (event[key]?.dateTime.replace(/\.0+$/, '') !== expected[key]!.dateTime || event[key]?.timeZone !== expected[key]!.timeZone)) throw new Error('Event time mismatch');
  }
  if (expected.location && event.location?.displayName !== expected.location.displayName) throw new Error('Location mismatch');
  if (expected.body && event.body?.content !== expected.body.content) throw new Error('Description mismatch');
  if (expected.attendees && (!Array.isArray(event.attendees) || event.attendees.length !== 0)) throw new Error('Unexpected attendees');
}

function mutationFailure(error: unknown, acknowledged: boolean, action: 'updated' | 'deleted', id: string, calendar?: string): CliError {
  if (!acknowledged && error instanceof CliError && error.code === 'GRAPH_ERROR' && Number(error.details?.status) < 500) return error;
  return new CliError('MUTATION_UNVERIFIED', 'Mutation was not verified; inspect the event before retrying', { id, calendar, outcome: acknowledged ? `${action}_but_unverified` : 'outcome_unknown' });
}

export class GraphClient {
  constructor(private readonly token: string, private readonly transport: typeof fetch = fetch, private readonly grant?: Grant) {}

  async createEvent(input: EventInput, calendar?: string) {
    requireOperation(this.grant, 'calendar.create');
    const url = new URL(`${calendarPath(calendar)}/events`);
    const transactionId = randomUUID();
    const body = { ...eventBody(input, true), isAllDay: false, attendees: [], transactionId };
    let id: string | undefined;
    try {
      const created = await this.request<CalendarEvent>(url, { method: 'POST', body });
      if (typeof created?.id !== 'string' || !created.id) throw new Error('Missing created ID');
      id = created.id;
      const event = await this.event(id, calendar, input.timezone);
      verifyEvent(event, id, body);
      return { outcome: 'created' as const, id, transactionId, event };
    } catch (error) {
      if (!id && error instanceof CliError && error.code === 'GRAPH_ERROR' && Number(error.details?.status) < 500) throw error;
      throw new CliError('MUTATION_UNVERIFIED', 'Creation was not verified; reconcile by ID or transactionId before retrying', { outcome: id ? 'created_but_unverified' : 'outcome_unknown', ...(id ? { id } : {}), transactionId, calendar });
    }
  }
  private async preflight(id: string, options: MutationOptions): Promise<CalendarEvent> {
    const event = await this.event(id, options.calendar);
    if (!event || event.id !== id || !Array.isArray(event.attendees)) throw new CliError('INVALID_RESPONSE', 'Cannot establish event identity and attendees; nothing was changed');
    if (event.attendees.length && options.notifyAttendees !== true) throw new CliError('ATTENDEE_NOTIFICATION_REQUIRED', 'Event has attendees; --notify-attendees is required to approve possible notifications');
    if (typeof event['@odata.etag'] !== 'string' || !event['@odata.etag']) throw new CliError('CONCURRENCY_REQUIRED', 'Event has no ETag; refusing an unprotected mutation');
    return event;
  }

  async updateEvent(id: string, input: Partial<EventInput>, options: MutationOptions = {}) {
    requireOperation(this.grant, 'calendar.update');
    const body = eventBody(input);
    const before = await this.preflight(id, options);
    if (before.isAllDay && input.start !== undefined) throw new CliError('INVALID_INPUT', 'Changing all-day event times is not supported; use Outlook');
    let acknowledged = false;
    try {
      await this.request(new URL(`${calendarPath(options.calendar)}/events/${encodedId(id)}`), { method: 'PATCH', body, etag: before['@odata.etag'], ignoreBody: true });
      acknowledged = true;
      const event = await this.event(id, options.calendar, input.timezone);
      verifyEvent(event, id, body);
      return { id, outcome: 'updated' as const, event };
    } catch (error) { throw mutationFailure(error, acknowledged, 'updated', id, options.calendar); }
  }

  async deleteEvent(id: string, options: MutationOptions = {}) {
    requireOperation(this.grant, 'calendar.delete');
    const before = await this.preflight(id, options);
    let acknowledged = false;
    try {
      await this.request(new URL(`${calendarPath(options.calendar)}/events/${encodedId(id)}`), { method: 'DELETE', etag: before['@odata.etag'], ignoreBody: true });
      acknowledged = true;
      try { await this.event(id, options.calendar); }
      catch (error) {
        if (error instanceof CliError && error.details?.status === 404) return { id, outcome: 'deleted' as const, verified: true };
        throw error;
      }
      throw new Error('Deleted event still exists');
    } catch (error) { throw mutationFailure(error, acknowledged, 'deleted', id, options.calendar); }
  }

  async calendars(limit = 100): Promise<Calendar[]> {
    requireOperation(this.grant, 'calendar.list');
    checkLimit(limit);
    const url = new URL(`${GRAPH_ROOT}/me/calendars`);
    url.searchParams.set('$top', String(limit));
    return this.collect<Calendar>(url, limit);
  }

  async events(start: string, end: string, limit = 100, calendar?: string): Promise<CalendarEvent[]> {
    requireOperation(this.grant, 'calendar.events');
    checkLimit(limit);
    const zoned = /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d(?:\.\d+)?(?:Z|[+-]\d\d:\d\d)$/;
    const duration = Date.parse(end) - Date.parse(start);
    if (!zoned.test(start) || !zoned.test(end) || !validLocalTime(start.slice(0, 19)) || !validLocalTime(end.slice(0, 19)) || !Number.isFinite(duration) || duration <= 0 || duration > 366 * 86400_000) {
      throw new CliError('INVALID_INPUT', 'Calendar window requires ISO start/end with offsets, increasing and at most 366 days');
    }
    const url = new URL(`${calendarPath(calendar)}/calendarView`);
    url.searchParams.set('startDateTime', start);
    url.searchParams.set('endDateTime', end);
    url.searchParams.set('$top', String(limit));
    return this.collect<CalendarEvent>(url, limit);
  }

  async event(id: string, calendar?: string, timezone = 'UTC'): Promise<CalendarEvent> {
    requireOperation(this.grant, 'calendar.get');
    checkTimezone(timezone);
    return this.request(new URL(`${calendarPath(calendar)}/events/${encodedId(id)}`), { preferText: true, timezone });
  }

  private get<T>(url: URL, preferText = false): Promise<T> { return this.request(url, { preferText }); }

  private async request<T>(url: URL, options: { method?: string; body?: unknown; preferText?: boolean; timezone?: string; etag?: string; ignoreBody?: boolean } = {}): Promise<T> {
    if (url.origin !== GRAPH_ORIGIN || url.username || url.password || url.hash || (url.pathname !== '/v1.0/me' && !url.pathname.startsWith('/v1.0/me/'))) {
      throw new Error('Refusing a Graph URL outside the signed-in mailbox');
    }
    const prefer = [options.preferText ? 'outlook.body-content-type="text"' : '', options.timezone ? `outlook.timezone="${options.timezone}"` : ''].filter(Boolean).join(', ');
    let response: Response;
    try { response = await this.transport(url.toString(), {
      method: options.method ?? 'GET',
      redirect: 'error',
      signal: AbortSignal.timeout(15_000),
      ...(options.body ? { body: JSON.stringify(options.body) } : {}),
      headers: {
        Authorization: `Bearer ${this.token}`,
        ...(options.body ? { 'Content-Type': 'application/json' } : {}),
        ...(prefer ? { Prefer: prefer } : {}),
        ...(options.etag ? { 'If-Match': options.etag } : {}),
      },
    }); } catch { throw new CliError('NETWORK_ERROR', 'Microsoft Graph request failed or timed out'); }
    if (!response.ok) {
      throw new CliError('GRAPH_ERROR', `Microsoft Graph returned HTTP ${response.status}`, { status: response.status });
    }
    if (options.ignoreBody) return undefined as T;
    try { return await response.json() as T; }
    catch { throw new CliError('INVALID_RESPONSE', 'Microsoft Graph returned invalid JSON'); }
  }

  async me(): Promise<{ id: string; displayName?: string; mail?: string; userPrincipalName?: string }> {
    const url = new URL(`${GRAPH_ROOT}/me`);
    url.searchParams.set('$select', 'id,displayName,mail,userPrincipalName');
    return this.get(url);
  }

  async read(id: string): Promise<MessageDetail> {
    if (!id.trim() || id.length > 1024) throw new Error('Message ID is required (max 1024 characters)');
    const url = new URL(`${GRAPH_ROOT}/me/messages/${encodeURIComponent(id)}`);
    url.searchParams.set('$select', 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,body,bodyPreview,isRead,hasAttachments,webLink');
    return this.get<MessageDetail>(url, true);
  }

  async search(query: string, limit: number): Promise<MessageSummary[]> {
    if (!query.trim() || query.length > 256) throw new Error('Search query must be 1–256 characters');
    if (!Number.isSafeInteger(limit) || limit < 1 || limit > 100) throw new Error('Limit must be 1–100');
    const url = new URL(`${GRAPH_ROOT}/me/messages`);
    url.searchParams.set('$search', `"${query.trim().replaceAll('\\', '\\\\').replaceAll('"', '\\"')}"`);
    url.searchParams.set('$top', String(limit));
    url.searchParams.set('$select', SUMMARY_FIELDS);
    return this.collect(url, limit);
  }

  async list(folder: string, limit: number): Promise<MessageSummary[]> {
    if (!Number.isSafeInteger(limit) || limit < 1 || limit > 100) throw new Error('Limit must be 1–100');
    if (!folder || folder.includes('/') || folder === '.' || folder === '..') throw new Error('Invalid folder');
    const url = new URL(`${GRAPH_ROOT}/me/mailFolders/${encodeURIComponent(folder)}/messages`);
    url.searchParams.set('$top', String(limit));
    url.searchParams.set('$select', SUMMARY_FIELDS);
    return this.collect(url, limit);
  }

  private async collect<T = MessageSummary>(url: URL, limit: number): Promise<T[]> {
    const results: T[] = [];
    const seen = new Set<string>();
    let next: URL | undefined = url;
    while (next && results.length < limit) {
      if (next.pathname !== url.pathname) throw new Error('Refusing a Graph URL outside the signed-in mailbox collection');
      const pageKey = new URL(next);
      pageKey.searchParams.sort();
      if (seen.size >= 20 || seen.has(pageKey.href)) throw new Error('Microsoft Graph returned excessive or repeated pages');
      seen.add(pageKey.href);
      const page: Page<T> = await this.get(next);
      if (!Array.isArray(page.value)) throw new Error('Invalid Microsoft Graph page');
      results.push(...page.value.slice(0, limit - results.length));
      next = page['@odata.nextLink'] ? new URL(page['@odata.nextLink']) : undefined;
    }
    return results;
  }
}
