import { test } from 'node:test';
import assert from 'node:assert/strict';
import { GraphClient } from '../src/graph.js';

const ro = { mode: 'read-only' as const, scopes: ['User.Read', 'Mail.Read'], expiresAt: Date.now() + 3600_000 };
const rw = { ...ro, mode: 'read-write' as const, scopes: [...ro.scopes, 'Calendars.ReadWrite'] };
const input = { title: 'Private test', start: '2027-01-02T10:00:00', end: '2027-01-02T11:00:00', timezone: 'UTC' };

test('direct read-only calendar mutation rejects deterministically before any network', async () => {
  let calls = 0;
  const graph = new GraphClient('opaque', async () => { calls++; throw new Error('network forbidden'); }, ro);
  for (const action of [() => graph.createEvent(input), () => graph.updateEvent('id', { title: 'new' }), () => graph.deleteEvent('id')]) {
    await assert.rejects(async () => action(), { code: 'READ_ONLY_MODE' });
  }
  assert.equal(calls, 0);
});

test('calendar reads use encoded IDs and a bounded calendarView window with safe pagination', async () => {
  const calls: { url: URL; init?: RequestInit }[] = [];
  const transport: typeof fetch = async (url, init) => {
    calls.push({ url: new URL(String(url)), init });
    if (String(url).includes('/events/')) return Response.json({ id: 'event/a', attendees: [] });
    if (String(url).includes('calendarView')) return Response.json(calls.filter(c => c.url.pathname.includes('calendarView')).length === 1
      ? { value: [{ id: 'first' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/calendars/cal%2Fa/calendarView?$skiptoken=next' }
      : { value: [{ id: 'second' }, { id: 'third' }] });
    return Response.json({ value: [{ id: 'cal/a' }] });
  };
  const graph = new GraphClient('opaque', transport, rw);
  assert.deepEqual(await graph.calendars(5), [{ id: 'cal/a' }]);
  assert.deepEqual((await graph.events('2027-01-01T00:00:00Z', '2027-02-01T00:00:00Z', 2, 'cal/a')).map(e => e.id), ['first', 'second']);
  assert.equal((await graph.event('event/a', 'cal/a')).id, 'event/a');
  assert.deepEqual(calls.map(c => c.init?.method), ['GET', 'GET', 'GET', 'GET']);
  assert.equal(calls[1].url.pathname, '/v1.0/me/calendars/cal%2Fa/calendarView');
  assert.equal(calls[1].url.searchParams.get('startDateTime'), '2027-01-01T00:00:00Z');
  assert.equal(calls[1].url.searchParams.get('endDateTime'), '2027-02-01T00:00:00Z');
  assert.equal(calls[1].url.searchParams.get('$top'), '2');
  assert.equal(calls[3].url.pathname, '/v1.0/me/calendars/cal%2Fa/events/event%2Fa');
  await assert.rejects(new GraphClient('opaque', transport, ro).calendars(), { code: 'INSUFFICIENT_SCOPE' });
  await assert.rejects(graph.events('bad', 'also bad'), { code: 'INVALID_INPUT' });
  await assert.rejects(graph.events('2027-03-01T00:00:00Z', '2027-02-01T00:00:00Z'), { code: 'INVALID_INPUT' });
  await assert.rejects(graph.events('2027-01-01T00:00:00Z', '2030-01-01T00:00:00Z'), { code: 'INVALID_INPUT' });
  await assert.rejects(graph.event('..'), { code: 'INVALID_INPUT' });
  assert.equal(calls.length, 4);
});

test('create posts a timed attendee-free event once and verifies the exact resource by GET', async () => {
  const calls: { url: URL; init?: RequestInit }[] = [];
  let body: any;
  const graph = new GraphClient('opaque', async (url, init) => {
    calls.push({ url: new URL(String(url)), init });
    if (init?.method === 'POST') {
      body = JSON.parse(String(init.body));
      return Response.json({ id: 'created/id' }, { status: 201 });
    }
    return Response.json({ ...body, id: 'created/id', start: { ...body.start, dateTime: `${body.start.dateTime}.0000000` } });
  }, rw);
  const result = await graph.createEvent({ ...input, location: 'Desk', description: 'Private note' }, 'calendar/id');
  assert.equal(result.id, 'created/id');
  assert.equal(result.outcome, 'created');
  assert.equal(result.event.subject, input.title);
  assert.match(result.transactionId, /^[a-f0-9-]{36}$/);
  assert.deepEqual(body, { subject: input.title, start: { dateTime: input.start, timeZone: 'UTC' }, end: { dateTime: input.end, timeZone: 'UTC' }, location: { displayName: 'Desk' }, body: { contentType: 'text', content: 'Private note' }, isAllDay: false, attendees: [], transactionId: result.transactionId });
  assert.deepEqual(calls.map(c => [c.init?.method, c.url.pathname]), [['POST', '/v1.0/me/calendars/calendar%2Fid/events'], ['GET', '/v1.0/me/calendars/calendar%2Fid/events/created%2Fid']]);
  assert.equal(new Headers(calls[0].init?.headers).get('Content-Type'), 'application/json');
  assert.equal(new Headers(calls[1].init?.headers).get('Prefer'), 'outlook.body-content-type="text", outlook.timezone="UTC"');
});

test('create validation refuses invitations, impossible dates and missing fields before dispatch', async () => {
  let calls = 0;
  const graph = new GraphClient('opaque', async () => { calls++; return Response.json({ id: 'x' }); }, rw);
  for (const bad of [{}, { ...input, attendees: ['someone'] }, { ...input, start: '2027-02-30T10:00:00' }, { ...input, end: input.start }, { ...input, timezone: 'invalid-zone' }, { ...input, isAllDay: true }, { ...input, title: '' }]) {
    await assert.rejects(graph.createEvent(bad as any), { code: 'INVALID_INPUT' });
  }
  assert.equal(calls, 0);
});

test('create preserves known IDs on every readback failure and does not retry uncertain POSTs', async () => {
  for (const failure of ['disconnect', 'malformed response', 'missing id', 'get failure', 'mismatch', 'wrong id', 'malformed readback', 'server error']) {
    let calls = 0;
    const graph = new GraphClient('PRIVATE-TOKEN', async (_url, init) => {
      calls++;
      if (init?.method === 'POST') {
        if (failure === 'disconnect') throw new Error('PRIVATE-TOKEN');
        if (failure === 'malformed response') return new Response('not JSON PRIVATE-TOKEN', { status: 201 });
        if (failure === 'missing id') return Response.json({});
        if (failure === 'server error') return new Response('PRIVATE-TOKEN', { status: 503 });
        return Response.json({ id: 'cleanup-id' }, { status: 201 });
      }
      if (failure === 'get failure') return Response.json({ error: { code: 'PRIVATE-TOKEN' } }, { status: 403 });
      if (failure === 'malformed readback') return Response.json(null);
      return Response.json({ id: failure === 'wrong id' ? 'other' : 'cleanup-id', subject: 'wrong' });
    }, rw);
    const known = ['get failure', 'mismatch', 'wrong id', 'malformed readback'].includes(failure);
    await assert.rejects(graph.createEvent(input), (error: any) => {
      assert.equal(error.code, 'MUTATION_UNVERIFIED', failure);
      assert.equal(error.details.outcome, known ? 'created_but_unverified' : 'outcome_unknown', failure);
      assert.equal(error.details.id, known ? 'cleanup-id' : undefined);
      assert.match(error.details.transactionId, /^[a-f0-9-]{36}$/);
      assert.doesNotMatch(JSON.stringify(error) + String(error), /PRIVATE-TOKEN/);
      return true;
    });
    assert.equal(calls, known ? 2 : 1, failure);
  }
});

test('calendar paging rejects credential-bearing, cross-origin, different-collection and repeated links before following them', async () => {
  for (const link of ['https://graph.microsoft.com@evil.test/v1.0/me/calendars', 'https://user:pass@graph.microsoft.com/v1.0/me/calendars', 'https://graph.microsoft.com/v1.0/me/messages', 'https://graph.microsoft.com/v1.0/me/calendars?$top=2']) {
    let calls = 0;
    const graph = new GraphClient('token', async () => { calls++; return Response.json({ value: [], '@odata.nextLink': link }); }, rw);
    await assert.rejects(graph.calendars(2));
    assert.equal(calls, 1, link);
  }
});

test('calendar view rejects impossible ISO dates and update refuses time changes to all-day events', async () => {
  let calls = 0;
  const graph = new GraphClient('token', async () => { calls++; return Response.json({ id: 'id', attendees: [], isAllDay: true, '@odata.etag': 'v1' }); }, rw);
  await assert.rejects(graph.events('2027-02-30T00:00:00Z', '2027-03-05T00:00:00Z'), { code: 'INVALID_INPUT' });
  assert.equal(calls, 0);
  await assert.rejects(graph.updateEvent('id', input), { code: 'INVALID_INPUT' });
  assert.equal(calls, 1);
});
