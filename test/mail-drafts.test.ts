import { test, type TestContext } from 'node:test';
import assert from 'node:assert/strict';
import { createServer, type IncomingMessage, type ServerResponse } from 'node:http';
import { GraphClient } from '../src/graph.js';

const grant = { mode: 'read-write' as const, scopes: ['User.Read', 'Mail.Read', 'Calendars.ReadWrite', 'Mail.ReadWrite', 'Mail.Send'], expiresAt: Date.now() + 3600000 };
const SELECT = '?%24select=id%2Csubject%2CisDraft%2CtoRecipients%2CccRecipients%2CbccRecipients%2CbodyPreview%2CwebLink%2CconversationId%2CparentFolderId';
const people = (...addresses: string[]) => addresses.map(address => ({ emailAddress: { address } }));
interface Recorded { method?: string; url?: string; headers: IncomingMessage['headers']; body: string }

// Real localhost HTTP recorder: the client still believes it talks to graph.microsoft.com.
async function recorder(t: TestContext, handler: (req: Recorded, res: ServerResponse, raw: IncomingMessage) => void, token = 'test-token') {
  const requests: Recorded[] = [];
  const server = createServer(async (req, res) => {
    let body = '';
    for await (const chunk of req) body += chunk;
    const recorded = { method: req.method, url: req.url, headers: req.headers, body };
    requests.push(recorded);
    res.setHeader('Content-Type', 'application/json');
    handler(recorded, res, req);
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
  return { requests, graph: new GraphClient(token, transport, grant) };
}
const calls = (requests: Recorded[]) => requests.map(r => `${r.method} ${r.url}`);

test('localhost contract: draft create POSTs one message then verifies the returned ID by GET', async t => {
  let stored: any;
  const { requests, graph } = await recorder(t, (req, res) => {
    if (req.method === 'POST') {
      stored = { ...JSON.parse(req.body), id: 'draft/1', isDraft: true, webLink: 'https://outlook.example/draft' };
      // Exchange normalizes address case; verification must not depend on it.
      stored.toRecipients = [{ emailAddress: { address: 'Anna@Example.Test', name: 'Anna' } }];
      res.writeHead(201).end(JSON.stringify(stored));
    } else res.end(JSON.stringify(stored));
  });
  const result = await graph.createDraft({ to: ['anna@example.test'], cc: ['carl@example.test'], bcc: ['bea@example.test'], subject: 'Hello', body: 'Plain text' });
  assert.equal(result.outcome, 'created');
  assert.equal(result.id, 'draft/1');
  assert.equal(result.draft.isDraft, true);
  assert.equal(result.draft.webLink, 'https://outlook.example/draft');
  assert.deepEqual(calls(requests), ['POST /v1.0/me/messages', `GET /v1.0/me/messages/draft%2F1${SELECT}`]);
  assert.deepEqual(JSON.parse(requests[0].body), { subject: 'Hello', body: { contentType: 'text', content: 'Plain text' }, toRecipients: people('anna@example.test'), ccRecipients: people('carl@example.test'), bccRecipients: people('bea@example.test') });
  assert.equal(requests[0].headers['content-type'], 'application/json');
  assert.equal(requests[1].body, '');
  for (const req of requests) assert.equal(req.headers.authorization, 'Bearer test-token');
});

test('draft create sends an HTML body only when requested and empty optional recipient lists otherwise', async t => {
  let stored: any;
  const { requests, graph } = await recorder(t, (req, res) => {
    if (req.method === 'POST') stored = { ...JSON.parse(req.body), id: 'd', isDraft: true };
    res.end(JSON.stringify(stored));
  });
  await graph.createDraft({ to: ['anna@example.test'], subject: 'Hi', body: '<p>Hi</p>', html: true });
  assert.deepEqual(JSON.parse(requests[0].body), { subject: 'Hi', body: { contentType: 'html', content: '<p>Hi</p>' }, toRecipients: people('anna@example.test'), ccRecipients: [], bccRecipients: [] });
});

test('draft input validation rejects unsafe recipients, subjects and bodies before any network call', async () => {
  let network = 0;
  const graph = new GraphClient('token', async () => { network++; return Response.json({ id: 'x', isDraft: true }); }, grant);
  const ok = { to: ['anna@example.test'], subject: 'Hello', body: 'Text' };
  const bad: any[] = [
    { ...ok, to: [] }, { ...ok, to: undefined }, { ...ok, to: 'anna@example.test' },
    { ...ok, to: ['Anna <anna@example.test>'] }, { ...ok, to: ['anna@example.test,bea@example.test'] }, { ...ok, to: ['anna@example.test;bea@example.test'] },
    { ...ok, to: ['anna@example.test\r\nBcc: evil@example.test'] }, { ...ok, to: ['anna @example.test'] }, { ...ok, to: [' anna@example.test'] },
    { ...ok, to: ['anna'] }, { ...ok, to: ['anna@localhost'] }, { ...ok, to: ['a@b@example.test'] }, { ...ok, to: [''] }, { ...ok, cc: [42] }, { ...ok, bcc: ['"x"@example.test'] },
    { ...ok, to: Array.from({ length: 60 }, (_, i) => `to${i}@example.test`), cc: Array.from({ length: 41 }, (_, i) => `cc${i}@example.test`) },
    { ...ok, subject: 'x'.repeat(256) }, { ...ok, subject: 'Hello\nBcc: evil@example.test' }, { ...ok, subject: 'Hello\rthere' }, { ...ok, subject: undefined }, { ...ok, subject: 7 },
    { ...ok, body: 'x'.repeat(1_000_001) }, { ...ok, body: undefined }, { ...ok, body: 5 }, { ...ok, html: 'yes' }, { ...ok, attachments: [{}] }, undefined,
  ];
  for (const input of bad) await assert.rejects(graph.createDraft(input), { code: 'INVALID_INPUT' }, JSON.stringify(input)?.slice(0, 80));
  assert.equal(network, 0);
  const limit = { ...ok, to: Array.from({ length: 100 }, (_, i) => `to${i}@example.test`), subject: 'x'.repeat(255), body: 'x'.repeat(1_000_000) };
  let posted: any;
  const accepting = new GraphClient('token', async (_url, init) => { if (init?.method === 'POST') posted = JSON.parse(String(init.body)); return Response.json({ ...posted, id: 'x', isDraft: true }); }, grant);
  assert.equal((await accepting.createDraft(limit)).outcome, 'created');
});

test('draft create preserves a known ID on every readback failure and never repeats an uncertain POST', async () => {
  for (const failure of ['disconnect', 'malformed response', 'missing id', 'server error', 'rejected', 'get failure', 'not a draft', 'wrong subject', 'wrong recipients', 'wrong id', 'malformed readback']) {
    const methods: string[] = [];
    const graph = new GraphClient('PRIVATE-TOKEN', async (_url, init) => {
      methods.push(init!.method!);
      if (init?.method === 'POST') {
        if (failure === 'disconnect') throw new Error('PRIVATE-TOKEN');
        if (failure === 'malformed response') return new Response('not JSON PRIVATE-TOKEN', { status: 201 });
        if (failure === 'missing id') return Response.json({});
        if (failure === 'server error') return new Response('PRIVATE-TOKEN', { status: 503 });
        if (failure === 'rejected') return Response.json({ error: { message: 'PRIVATE-TOKEN' } }, { status: 400 });
        return Response.json({ id: 'known-id' }, { status: 201 });
      }
      if (failure === 'get failure') return Response.json({ error: { code: 'PRIVATE-TOKEN' } }, { status: 403 });
      if (failure === 'malformed readback') return Response.json(null);
      return Response.json({
        id: failure === 'wrong id' ? 'other' : 'known-id', isDraft: failure !== 'not a draft', subject: failure === 'wrong subject' ? 'Other' : 'Hello',
        toRecipients: people(failure === 'wrong recipients' ? 'evil@example.test' : 'anna@example.test'), ccRecipients: [], bccRecipients: [],
      });
    }, grant);
    const known = !['disconnect', 'malformed response', 'missing id', 'server error', 'rejected'].includes(failure);
    await assert.rejects(graph.createDraft({ to: ['anna@example.test'], subject: 'Hello', body: 'Text' }), (error: any) => {
      if (failure === 'rejected') assert.deepEqual([error.code, error.details], ['GRAPH_ERROR', { status: 400 }]);
      else {
        assert.equal(error.code, 'MUTATION_UNVERIFIED', failure);
        assert.deepEqual(error.details, known ? { outcome: 'created_but_unverified', id: 'known-id' } : { outcome: 'outcome_unknown' }, failure);
      }
      assert.doesNotMatch(JSON.stringify(error) + String(error), /PRIVATE-TOKEN/);
      return true;
    });
    assert.deepEqual(methods, known ? ['POST', 'GET'] : ['POST'], failure);
  }
});

test('localhost contract: reply and reply-all POST a plain-text comment once and read back the recipients Graph chose', async t => {
  let stored: any;
  const { requests, graph } = await recorder(t, (req, res) => {
    if (req.method === 'POST') {
      const all = req.url!.endsWith('/createReplyAll');
      stored = { id: all ? 'reply-all-draft' : 'reply-draft', isDraft: true, subject: 'RE: Original', toRecipients: people('sender@example.test'), ccRecipients: all ? people('other@example.test') : [], bccRecipients: [], conversationId: 'conversation' };
      res.writeHead(201).end(JSON.stringify(stored));
    } else res.end(JSON.stringify(stored));
  });
  const reply = await graph.createReplyDraft('msg/1', { body: 'Thanks' });
  assert.deepEqual(reply, { outcome: 'created', id: 'reply-draft', draft: stored });
  assert.deepEqual(reply.draft.toRecipients, people('sender@example.test'));
  const all = await graph.createReplyDraft('msg/1', { all: true, body: 'Thanks all' });
  assert.equal(all.id, 'reply-all-draft');
  assert.deepEqual(all.draft.ccRecipients, people('other@example.test'));
  assert.deepEqual(calls(requests), [
    'POST /v1.0/me/messages/msg%2F1/createReply', `GET /v1.0/me/messages/reply-draft${SELECT}`,
    'POST /v1.0/me/messages/msg%2F1/createReplyAll', `GET /v1.0/me/messages/reply-all-draft${SELECT}`,
  ]);
  assert.equal(requests[0].body, '{"comment":"<div>Thanks</div>"}');
  assert.equal(requests[2].body, '{"comment":"<div>Thanks all</div>"}');
});

// Graph inserts the createReply comment into the HTML body as markup, so plain text has to be converted.
test('reply drafts keep paragraph spacing and literal characters of the plain-text body', async t => {
  const { requests, graph } = await recorder(t, (req, res) => {
    const stored = { id: 'reply-draft', isDraft: true, toRecipients: people('sender@example.test'), ccRecipients: [], bccRecipients: [] };
    res.writeHead(req.method === 'POST' ? 201 : 200).end(JSON.stringify(stored));
  });
  await graph.createReplyDraft('msg', { body: 'Lieber Reto\r\n\r\nZeile eins\nZeile zwei <b>fett</b> a<b & "c"\n\n\nGruss\n\n' });
  assert.equal(JSON.parse(requests[0].body).comment, [
    '<div>Lieber Reto</div>', '<div><br></div>',
    '<div>Zeile eins</div>', '<div>Zeile zwei &lt;b&gt;fett&lt;/b&gt; a&lt;b &amp; &quot;c&quot;</div>',
    '<div><br></div>', '<div><br></div>', '<div>Gruss</div>',
  ].join(''));
});

test('reply drafts reject HTML and invalid input before any network call and keep partial outcomes', async () => {
  let network = 0;
  const graph = new GraphClient('token', async () => { network++; return Response.json({ id: 'x', isDraft: true }); }, grant);
  for (const [id, input] of [['msg', { body: '<p>x</p>', html: true }], ['msg', { body: 'x'.repeat(1_000_001) }], ['msg', {}], ['msg', { body: 'x', all: 'yes' }], ['msg', { body: 'x', to: ['a@example.test'] }], ['..', { body: 'x' }], ['', { body: 'x' }]] as [string, any][]) {
    await assert.rejects(graph.createReplyDraft(id, input), { code: 'INVALID_INPUT' }, JSON.stringify([id, input]).slice(0, 60));
  }
  assert.equal(network, 0);
  for (const failure of ['disconnect', 'not found', 'readback not draft']) {
    const methods: string[] = [];
    const failing = new GraphClient('PRIVATE-TOKEN', async (_url, init) => {
      methods.push(init!.method!);
      if (init?.method === 'POST') {
        if (failure === 'disconnect') throw new Error('PRIVATE-TOKEN');
        if (failure === 'not found') return Response.json({}, { status: 404 });
        return Response.json({ id: 'reply-id' }, { status: 201 });
      }
      return Response.json({ id: 'reply-id', isDraft: false });
    }, grant);
    await assert.rejects(failing.createReplyDraft('msg', { body: 'x' }), (error: any) => {
      if (failure === 'not found') assert.deepEqual([error.code, error.details], ['GRAPH_ERROR', { status: 404 }]);
      else assert.deepEqual([error.code, error.details], ['MUTATION_UNVERIFIED', failure === 'disconnect' ? { outcome: 'outcome_unknown' } : { outcome: 'created_but_unverified', id: 'reply-id' }]);
      assert.doesNotMatch(JSON.stringify(error) + String(error), /PRIVATE-TOKEN/);
      return true;
    });
    assert.deepEqual(methods, failure === 'readback not draft' ? ['POST', 'GET'] : ['POST']);
  }
});

test('localhost contract: draft update is GET, a PATCH of only the supplied fields, then a verifying GET', async t => {
  let stored: any = { id: 'draft/1', isDraft: true, subject: 'Before', toRecipients: people('old@example.test'), ccRecipients: people('keep@example.test'), bccRecipients: [] };
  const { requests, graph } = await recorder(t, (req, res) => {
    if (req.method === 'PATCH') stored = { ...stored, ...JSON.parse(req.body), body: undefined };
    res.end(JSON.stringify(stored));
  });
  const result = await graph.updateDraft('draft/1', { subject: 'After', to: ['new@example.test'] });
  assert.equal(result.outcome, 'updated');
  assert.equal(result.id, 'draft/1');
  assert.equal(result.draft.subject, 'After');
  assert.deepEqual(result.draft.ccRecipients, people('keep@example.test'));
  assert.deepEqual(calls(requests), [`GET /v1.0/me/messages/draft%2F1${SELECT}`, 'PATCH /v1.0/me/messages/draft%2F1', `GET /v1.0/me/messages/draft%2F1${SELECT}`]);
  assert.equal(requests[1].body, '{"subject":"After","toRecipients":[{"emailAddress":{"address":"new@example.test"}}]}');
  await graph.updateDraft('draft/1', { body: '<b>x</b>', html: true, bcc: [] });
  assert.equal(requests[4].body, '{"body":{"contentType":"html","content":"<b>x</b>"},"bccRecipients":[]}');
  await graph.updateDraft('draft/1', { body: 'plain' });
  assert.equal(requests[7].body, '{"body":{"contentType":"text","content":"plain"}}');
  assert.equal(requests.length, 9);
});

test('draft update changes nothing for sent messages or invalid input, and models partial outcomes', async () => {
  let network = 0;
  const graph = new GraphClient('token', async () => { network++; return Response.json({ id: 'id', isDraft: true }); }, grant);
  for (const [id, input] of [['id', {}], ['id', { html: true }], ['id', { to: ['Anna <anna@example.test>'] }], ['id', { subject: 'a\nb' }], ['id', { body: 'x'.repeat(1_000_001) }], ['id', { importance: 'high' }], ['..', { subject: 'x' }]] as [string, any][]) {
    await assert.rejects(graph.updateDraft(id, input), { code: 'INVALID_INPUT' }, JSON.stringify(input).slice(0, 60));
  }
  assert.equal(network, 0);
  for (const before of [{ id: 'id', isDraft: false }, { id: 'id' }, { id: 'other', isDraft: true }, null]) {
    const methods: string[] = [];
    const sent = new GraphClient('token', async (_url, init) => { methods.push(init!.method!); return Response.json(before); }, grant);
    await assert.rejects(sent.updateDraft('id', { subject: 'x' }), { code: before && before.id === 'id' ? 'NOT_A_DRAFT' : 'INVALID_RESPONSE' });
    assert.deepEqual(methods, ['GET']);
  }
  for (const failure of ['preflight404', 'disconnect', 'rejected', 'server error', 'readback403', 'stale']) {
    const methods: string[] = [];
    const failing = new GraphClient('PRIVATE-TOKEN', async (_url, init) => {
      methods.push(init!.method!);
      if (methods.length === 1) return failure === 'preflight404' ? Response.json({}, { status: 404 }) : Response.json({ id: 'id', isDraft: true, subject: 'Old' });
      if (init?.method === 'PATCH') {
        if (failure === 'disconnect') throw new Error('PRIVATE-TOKEN');
        if (failure === 'rejected') return Response.json({}, { status: 400 });
        if (failure === 'server error') return new Response('PRIVATE-TOKEN', { status: 502 });
        return Response.json({ id: 'id' });
      }
      if (failure === 'readback403') return Response.json({ error: 'PRIVATE-TOKEN' }, { status: 403 });
      return Response.json({ id: 'id', isDraft: true, subject: 'Old' });
    }, grant);
    await assert.rejects(failing.updateDraft('id', { subject: 'New' }), (error: any) => {
      if (failure === 'preflight404') assert.deepEqual([error.code, error.details], ['GRAPH_ERROR', { status: 404 }]);
      else if (failure === 'rejected') assert.deepEqual([error.code, error.details], ['GRAPH_ERROR', { status: 400 }]);
      else assert.deepEqual([error.code, error.details], ['MUTATION_UNVERIFIED', { id: 'id', outcome: ['disconnect', 'server error'].includes(failure) ? 'outcome_unknown' : 'updated_but_unverified' }]);
      assert.doesNotMatch(JSON.stringify(error) + String(error), /PRIVATE-TOKEN/);
      return true;
    });
    assert.deepEqual(methods, failure === 'preflight404' ? ['GET'] : ['readback403', 'stale'].includes(failure) ? ['GET', 'PATCH', 'GET'] : ['GET', 'PATCH'], failure);
  }
});

const draft = { id: 'draft/1', isDraft: true, subject: 'Hello', toRecipients: people('anna@example.test'), ccRecipients: people('carl@example.test'), bccRecipients: people('bea@example.test') };
const SENT = { outcome: 'sent', id: 'draft/1', verified: true, subject: 'Hello', recipients: { to: ['anna@example.test'], cc: ['carl@example.test'], bcc: ['bea@example.test'] } };

test('localhost contract: draft send is GET, one bodiless POST /send answered 202, then a GET that must be 404', async t => {
  let sent = false;
  const { requests, graph } = await recorder(t, (req, res) => {
    if (req.method === 'POST') { sent = true; res.writeHead(202).end(); }
    else if (sent) res.writeHead(404).end('{}');
    else res.end(JSON.stringify(draft));
  });
  assert.deepEqual(await graph.sendDraft('draft/1'), SENT);
  assert.deepEqual(calls(requests), [`GET /v1.0/me/messages/draft%2F1${SELECT}`, 'POST /v1.0/me/messages/draft%2F1/send', `GET /v1.0/me/messages/draft%2F1${SELECT}`]);
  assert.equal(requests[1].body, '');
  assert.equal(requests[1].headers['content-length'], '0');
  assert.equal(requests[1].headers['content-type'], undefined);
  for (const req of requests) assert.equal(req.headers.authorization, 'Bearer test-token');
});

test('draft send is also verified when the same ID reads back as a sent message (immutable IDs)', async () => {
  const methods: string[] = [];
  const graph = new GraphClient('token', async (_url, init) => {
    methods.push(init!.method!);
    if (init?.method === 'POST') return new Response(null, { status: 202 });
    return Response.json({ ...draft, isDraft: methods.length === 1 });
  }, grant);
  assert.deepEqual(await graph.sendDraft('draft/1'), SENT);
  assert.deepEqual(methods, ['GET', 'POST', 'GET']);
});

test('draft send keeps reading back while Exchange still shows the draft, then verifies once it is gone; one POST only', async () => {
  const methods: string[] = [];
  const waits: number[] = [];
  const graph = new GraphClient('token', async (_url, init) => {
    methods.push(init!.method!);
    if (init?.method === 'POST') return new Response(null, { status: 202 });
    if (methods.length <= 4) return Response.json(draft);
    return Response.json({}, { status: 404 });
  }, grant, async ms => { waits.push(ms); });
  assert.deepEqual(await graph.sendDraft('draft/1'), SENT);
  assert.deepEqual(methods, ['GET', 'POST', 'GET', 'GET', 'GET']);
  assert.equal(waits.length, 2);
});

test('localhost contract: send is refused with only the preflight GET for non-drafts and drafts without recipients', async t => {
  let current: any;
  const { requests, graph } = await recorder(t, (_req, res) => { res.end(JSON.stringify(current)); });
  const refusals: [any, string][] = [
    [{ ...draft, isDraft: false }, 'NOT_A_DRAFT'], [{ ...draft, isDraft: undefined }, 'NOT_A_DRAFT'],
    [{ ...draft, toRecipients: [], ccRecipients: [], bccRecipients: [] }, 'INVALID_INPUT'], [{ id: 'draft/1', isDraft: true }, 'INVALID_INPUT'],
    [{ ...draft, toRecipients: [{ emailAddress: {} }], ccRecipients: [], bccRecipients: [] }, 'INVALID_INPUT'],
    [{ ...draft, id: 'another' }, 'INVALID_RESPONSE'], [null, 'INVALID_RESPONSE'],
  ];
  for (const [state, code] of refusals) {
    current = state;
    requests.length = 0;
    await assert.rejects(graph.sendDraft('draft/1'), { code }, JSON.stringify(state));
    assert.deepEqual(calls(requests), [`GET /v1.0/me/messages/draft%2F1${SELECT}`]);
  }
  requests.length = 0;
  await assert.rejects(graph.sendDraft('..'), { code: 'INVALID_INPUT' });
  assert.equal(requests.length, 0);
});

test('localhost contract: a real disconnect on the send POST is outcome_unknown with the ID and exactly one POST', async t => {
  const { requests, graph } = await recorder(t, (req, res, raw) => {
    if (req.method === 'POST') raw.socket.destroy();
    else res.end(JSON.stringify(draft));
  }, 'PRIVATE-TOKEN');
  await assert.rejects(graph.sendDraft('draft/1'), (error: any) => {
    assert.equal(error.code, 'MUTATION_UNVERIFIED');
    assert.deepEqual(error.details, { id: 'draft/1', outcome: 'outcome_unknown' });
    assert.doesNotMatch(JSON.stringify(error) + String(error), /PRIVATE-TOKEN/);
    return true;
  });
  assert.deepEqual(requests.map(r => r.method), ['GET', 'POST']);
});

test('send failures separate not-sent, unknown and acknowledged-but-unverified outcomes without ever resending', async () => {
  const cases: Record<string, string> = {
    'post 400': 'GRAPH_ERROR', 'post 403': 'GRAPH_ERROR', 'post 404': 'GRAPH_ERROR', 'post 429': 'GRAPH_ERROR',
    'post 500': 'outcome_unknown', 'post 503': 'outcome_unknown', 'post throws': 'outcome_unknown',
    'still draft': 'sent_but_unverified', 'readback 403': 'sent_but_unverified', 'readback 500': 'sent_but_unverified', 'readback throws': 'sent_but_unverified',
    'readback wrong id': 'sent_but_unverified', 'readback malformed': 'sent_but_unverified',
  };
  for (const [failure, expected] of Object.entries(cases)) {
    const methods: string[] = [];
    const graph = new GraphClient('PRIVATE-TOKEN', async (_url, init) => {
      methods.push(init!.method!);
      if (methods.length === 1) return Response.json(draft);
      if (init?.method === 'POST') {
        if (failure === 'post throws') throw new Error('PRIVATE-TOKEN');
        if (failure.startsWith('post ')) return Response.json({ error: { message: 'PRIVATE-TOKEN' } }, { status: Number(failure.slice(5)) });
        return new Response(null, { status: 202 });
      }
      if (failure === 'readback throws') throw new Error('PRIVATE-TOKEN');
      if (failure === 'readback malformed') return new Response('PRIVATE-TOKEN not json');
      if (failure === 'readback wrong id') return Response.json({ id: 'other', isDraft: false });
      if (failure.startsWith('readback ')) return Response.json({ error: 'PRIVATE-TOKEN' }, { status: Number(failure.slice(9)) });
      return Response.json(draft);
    }, grant, async () => {});
    await assert.rejects(graph.sendDraft('draft/1'), (error: any) => {
      if (expected === 'GRAPH_ERROR') assert.deepEqual([error.code, error.details], ['GRAPH_ERROR', { status: Number(failure.slice(5)) }], failure);
      else assert.deepEqual([error.code, error.details], ['MUTATION_UNVERIFIED', { id: 'draft/1', outcome: expected }], failure);
      assert.doesNotMatch(JSON.stringify(error) + String(error), /PRIVATE-TOKEN/);
      return true;
    });
    assert.deepEqual(methods, failure.startsWith('post') ? ['GET', 'POST'] : failure === 'still draft' ? ['GET', 'POST', 'GET', 'GET', 'GET', 'GET', 'GET', 'GET'] : ['GET', 'POST', 'GET'], failure);
  }
});

test('every draft method is guarded by the registry before any network call', async () => {
  let network = 0;
  const transport: typeof fetch = async () => { network++; throw new Error('network forbidden'); };
  const input = { to: ['anna@example.test'], subject: 'Hello', body: 'Text' };
  const attempts = (graph: GraphClient) => [() => graph.createDraft(input), () => graph.createReplyDraft('id', { body: 'x' }), () => graph.updateDraft('id', { subject: 'x' }), () => graph.sendDraft('id')];
  const expectAll = async (current: any, codes: (string | undefined)[]) => {
    const results = await Promise.all(attempts(new GraphClient('opaque', transport, current)).map(action => action().then(() => undefined, (error: any) => error.code)));
    assert.deepEqual(results, codes);
  };
  await expectAll({ ...grant, mode: 'read-only' }, Array(4).fill('READ_ONLY_MODE'));
  await expectAll(undefined, Array(4).fill('AUTH_REQUIRED'));
  await expectAll({ ...grant, expiresAt: 1 }, Array(4).fill('AUTH_REQUIRED'));
  await expectAll({ ...grant, scopes: ['User.Read', 'Mail.Read', 'Calendars.ReadWrite'] }, Array(4).fill('INSUFFICIENT_SCOPE'));
  await expectAll({ ...grant, scopes: ['User.Read', 'Mail.Read', 'Calendars.ReadWrite', 'Mail.Send'] }, Array(4).fill('INSUFFICIENT_SCOPE'));
  assert.equal(network, 0);
  // Mail.ReadWrite without Mail.Send drafts freely (network reached) but still cannot send.
  await expectAll({ ...grant, scopes: ['User.Read', 'Mail.Read', 'Calendars.ReadWrite', 'Mail.ReadWrite'] }, ['MUTATION_UNVERIFIED', 'MUTATION_UNVERIFIED', 'NETWORK_ERROR', 'INSUFFICIENT_SCOPE']);
  assert.equal(network, 3);
});

test('there is no direct send, delete, move or attachment surface on the client', () => {
  const methods = Object.getOwnPropertyNames(GraphClient.prototype);
  for (const name of ['createDraft', 'createReplyDraft', 'updateDraft', 'sendDraft']) assert.ok(methods.includes(name), name);
  assert.deepEqual(methods.filter(name => /sendMail|deleteMessage|deleteDraft|move|attach|markRead|folder/i.test(name)), []);
});
