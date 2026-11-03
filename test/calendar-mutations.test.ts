import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import { GraphClient } from '../src/graph.js';

const grant = { mode: 'read-write' as const, scopes: ['User.Read', 'Mail.Read', 'Calendars.ReadWrite'], expiresAt: Date.now() + 3600000 };

test('localhost contract: update GET/PATCH/GET then delete GET/DELETE/GET404 with conditional narrow writes', async t => {
  const requests: { method?: string; url?: string; headers: any; body: string }[] = [];
  let event: any = { id: 'id/a', subject: 'Before', attendees: [], '@odata.etag': 'W/"version1"' };
  const server = createServer(async (req, res) => {
    let body = '';
    for await (const chunk of req) body += chunk;
    requests.push({ method: req.method, url: req.url, headers: req.headers, body });
    res.setHeader('Content-Type', 'application/json');
    if (req.method === 'PATCH') { event = { ...event, ...JSON.parse(body) }; res.end(JSON.stringify(event)); }
    else if (req.method === 'DELETE') { event = undefined; res.writeHead(204).end(); }
    else if (event) res.end(JSON.stringify(event));
    else res.writeHead(404).end('{}');
  });
  await new Promise<void>(resolve => server.listen(0, '127.0.0.1', resolve));
  t.after(() => { server.closeAllConnections(); server.close(); });
  const port = (server.address() as { port: number }).port;
  const transport: typeof fetch = (url, init) => {
    const target = new URL(String(url));
    assert.equal(target.origin, 'https://graph.microsoft.com');
    assert.equal(init?.redirect, 'error');
    return fetch(`http://127.0.0.1:${port}${target.pathname}${target.search}`, init);
  };
  const graph = new GraphClient('test-token', transport, grant);
  const updated = await graph.updateEvent('id/a', { title: 'After' });
  assert.equal(updated.outcome, 'updated');
  assert.equal(updated.event.subject, 'After');
  const deleted = await graph.deleteEvent('id/a');
  assert.deepEqual(deleted, { id: 'id/a', outcome: 'deleted', verified: true });
  assert.deepEqual(requests.map(r => r.method), ['GET', 'PATCH', 'GET', 'GET', 'DELETE', 'GET']);
  for (const req of requests) {
    assert.equal(req.url, '/v1.0/me/events/id%2Fa');
    assert.equal(req.headers.authorization, 'Bearer test-token');
  }
  assert.equal(requests[1].body, '{"subject":"After"}');
  assert.equal(requests[1].headers['if-match'], 'W/"version1"');
  assert.equal(requests[4].headers['if-match'], 'W/"version1"');
  assert.equal(requests[4].body, '');
});

test('attendee or incomplete preflight metadata blocks update/delete unless notification is explicitly approved', async () => {
  for (const method of ['update', 'delete']) {
    for (const event of [{ id: 'id', attendees: [{}], '@odata.etag': 'version' }, { id: 'id' }, { id: 'id', attendees: [] }]) {
      const calls: string[] = [];
      const graph = new GraphClient('token', async (_url, init) => { calls.push(init!.method!); return Response.json(event); }, grant);
      await assert.rejects(method === 'update' ? graph.updateEvent('id', { title: 'new' }) : graph.deleteEvent('id'), (error: any) => {
        assert.ok(['ATTENDEE_NOTIFICATION_REQUIRED', 'INVALID_RESPONSE', 'CONCURRENCY_REQUIRED'].includes(error.code));
        return true;
      });
      assert.deepEqual(calls, ['GET']);
    }
  }
});

test('mutations preserve IDs and distinguish unknown dispatch from acknowledged-but-unverified writes', async () => {
  for (const method of ['update', 'delete']) {
    for (const failure of ['disconnect', 'readback403', 'readback500', 'stale', 'conflict']) {
      let calls = 0;
      const graph = new GraphClient('PRIVATE-TOKEN', async (_url, init) => {
        calls++;
        if (calls === 1) return Response.json({ id: 'id', attendees: [], '@odata.etag': 'v1' });
        if (init?.method !== 'GET') {
          if (failure === 'disconnect') throw new Error('PRIVATE-TOKEN');
          if (failure === 'conflict') return Response.json({}, { status: 412 });
          return new Response(null, { status: 204 });
        }
        if (failure.startsWith('readback')) return Response.json({ error: { code: 'PRIVATE-TOKEN' } }, { status: Number(failure.slice(8)) });
        return Response.json({ id: 'id', subject: 'Old' });
      }, grant);
      await assert.rejects(method === 'update' ? graph.updateEvent('id', { title: 'New' }) : graph.deleteEvent('id'), (error: any) => {
        if (failure === 'conflict') {
          assert.equal(error.code, 'GRAPH_ERROR');
          assert.equal(error.details.status, 412);
        } else {
          assert.equal(error.code, 'MUTATION_UNVERIFIED');
          assert.equal(error.details.id, 'id');
          assert.equal(error.details.outcome, failure === 'disconnect' ? 'outcome_unknown' : `${method}d_but_unverified`);
        }
        assert.doesNotMatch(JSON.stringify(error) + String(error), /PRIVATE-TOKEN/);
        return true;
      });
      assert.equal(calls, ['disconnect', 'conflict'].includes(failure) ? 2 : 3);
    }
  }
});

test('explicit notification approval permits conditional updates to meetings without changing attendees', async () => {
  let event: any = { id: 'meeting', attendees: [{ emailAddress: { address: 'test@example.test' } }], '@odata.etag': 'v1' };
  const methods: string[] = [];
  const graph = new GraphClient('token', async (_url, init) => {
    methods.push(init!.method!);
    if (init?.method === 'PATCH') {
      assert.deepEqual(JSON.parse(String(init.body)), { subject: 'Approved' });
      event.subject = 'Approved';
    }
    return Response.json(event);
  }, grant);
  assert.equal((await graph.updateEvent('meeting', { title: 'Approved' }, { notifyAttendees: true })).outcome, 'updated');
  assert.deepEqual(methods, ['GET', 'PATCH', 'GET']);
});

test('localhost POST contract verifies success and preserves reconciliation markers after a real socket disconnect', async () => {
  for (const disconnect of [false, true]) {
    const methods: string[] = [];
    let stored: any;
    const server = createServer(async (req, res) => {
      methods.push(req.method!);
      assert.equal(req.headers.authorization, 'Bearer test-token');
      if (req.method === 'POST') {
        let raw = '';
        for await (const chunk of req) raw += chunk;
        stored = { ...JSON.parse(raw), id: 'created-id' };
        assert.equal(req.url, '/v1.0/me/events');
        assert.deepEqual(stored.attendees, []);
        if (disconnect) { req.socket.destroy(); return; }
        res.writeHead(201, { 'Content-Type': 'application/json' }).end(JSON.stringify({ id: stored.id }));
      } else {
        assert.equal(req.url, '/v1.0/me/events/created-id');
        res.writeHead(200, { 'Content-Type': 'application/json' }).end(JSON.stringify(stored));
      }
    });
    await new Promise<void>(resolve => server.listen(0, '127.0.0.1', resolve));
    const port = (server.address() as { port: number }).port;
    try {
      const graph = new GraphClient('test-token', (url, init) => {
        const target = new URL(String(url));
        assert.equal(target.origin, 'https://graph.microsoft.com');
        return fetch(`http://127.0.0.1:${port}${target.pathname}${target.search}`, init);
      }, grant);
      const action = graph.createEvent({ title: 'Local contract marker', start: '2027-01-02T10:00:00', end: '2027-01-02T11:00:00', timezone: 'UTC' });
      if (disconnect) {
        await assert.rejects(action, (error: any) => {
          assert.equal(error.details.outcome, 'outcome_unknown');
          assert.equal(error.details.transactionId, stored.transactionId);
          assert.equal(error.details.id, undefined);
          return true;
        });
        assert.deepEqual(methods, ['POST']);
      } else {
        const result = await action;
        assert.equal(result.id, stored.id);
        assert.equal(result.transactionId, stored.transactionId);
        assert.deepEqual(methods, ['POST', 'GET']);
      }
    } finally {
      server.closeAllConnections();
      await new Promise<void>(resolve => server.close(() => resolve()));
    }
  }
});
