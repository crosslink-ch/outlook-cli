import { test } from 'node:test';
import assert from 'node:assert/strict';
import { GraphClient } from '../src/graph.js';

test('list follows Graph nextLink until the requested limit, without fetching bodies', async () => {
  const urls: URL[] = [];
  const fakeFetch: typeof fetch = async (input, init) => {
    assert.equal(init?.method, 'GET');
    const url = new URL(String(input));
    urls.push(url);
    return new Response(JSON.stringify(urls.length === 1
      ? { value: [{ id: '1' }, { id: '2' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$skiptoken=next' }
      : { value: [{ id: '3' }, { id: '4' }] }), { status: 200 });
  };
  const client = new GraphClient('test-token', fakeFetch);
  const messages = await client.list('inbox', 3);
  assert.deepEqual(messages.map(m => m.id), ['1', '2', '3']);
  assert.equal(urls.length, 2);
  assert.equal(urls[0]!.searchParams.get('$top'), '3');
  assert.equal(urls[0]!.searchParams.get('$select'), 'id,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments');
});

test('search encodes the query and requests only message summaries', async () => {
  let called: URL | undefined;
  const fakeFetch: typeof fetch = async (input) => {
    called = new URL(String(input));
    return new Response(JSON.stringify({ value: [{ id: 'found' }] }), { status: 200 });
  };
  const messages = await new GraphClient('test-token', fakeFetch).search('quarterly "report"', 5);
  assert.deepEqual(messages.map(m => m.id), ['found']);
  assert.equal(called?.pathname, '/v1.0/me/messages');
  assert.equal(called?.searchParams.get('$search'), '"quarterly \\"report\\""');
  assert.equal(called?.searchParams.get('$top'), '5');
});

test('search follows nextLink to fill the requested limit', async () => {
  let calls = 0;
  const fakeFetch: typeof fetch = async () => {
    calls++;
    return new Response(JSON.stringify(calls === 1
      ? { value: [{ id: 'first' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/messages?$skiptoken=next' }
      : { value: [{ id: 'second' }] }), { status: 200 });
  };
  const messages = await new GraphClient('test-token', fakeFetch).search('anything', 2);
  assert.deepEqual(messages.map(m => m.id), ['first', 'second']);
  assert.equal(calls, 2);
});

test('read requests a single encoded message ID and text body without mutation', async () => {
  let called: URL | undefined;
  const fakeFetch: typeof fetch = async (input, init) => {
    called = new URL(String(input));
    assert.equal(init?.method, 'GET');
    assert.equal(new Headers(init?.headers).get('Prefer'), 'outlook.body-content-type="text"');
    return new Response(JSON.stringify({ id: 'a/b', subject: 'Hello', body: { contentType: 'text', content: 'Test' } }), { status: 200 });
  };
  const message = await new GraphClient('test-token', fakeFetch).read('a/b');
  assert.equal(message.body?.content, 'Test');
  assert.equal(called?.pathname, '/v1.0/me/messages/a%2Fb');
});

test('me identifies the authenticated account and rejects a cross-origin page link', async () => {
  const fakeFetch: typeof fetch = async (input) => {
    const url = new URL(String(input));
    if (url.pathname === '/v1.0/me') {
      return new Response(JSON.stringify({ id: 'user-1', mail: 'person@example.com' }), { status: 200 });
    }
    return new Response(JSON.stringify({ value: [{ id: '1' }], '@odata.nextLink': 'https://evil.example/steal' }), { status: 200 });
  };
  const client = new GraphClient('test-token', fakeFetch);
  assert.equal((await client.me()).mail, 'person@example.com');
  await assert.rejects(client.list('inbox', 2), /outside the signed-in mailbox/);
});
