import { test } from 'node:test';
import assert from 'node:assert/strict';
import * as policy from '../src/capabilities.js';
import { GraphClient } from '../src/graph.js';

test('capabilities derive from fresh granted scopes, never just the mode label', async () => {
  const ro = { mode: 'read-only' as const, scopes: ['User.Read', 'Mail.Read'], expiresAt: Date.now() + 10000 };
  const rw = { ...ro, mode: 'read-write' as const };
  let calls = 0;
  const transport: typeof fetch = async () => { calls++; throw new Error('forbidden'); };
  await assert.rejects(new GraphClient('opaque', transport, rw).deleteEvent('id'), { code: 'INSUFFICIENT_SCOPE' });
  await assert.rejects(new GraphClient('opaque', transport).deleteEvent('id'), { code: 'AUTH_REQUIRED' });
  await assert.rejects(new GraphClient('opaque', transport, { ...ro, expiresAt: 1 }).deleteEvent('id'), { code: 'AUTH_REQUIRED' });
  assert.deepEqual(policy.availableOperations(ro), ['me', 'list', 'search', 'read']);
  assert.equal(policy.availableOperations({ ...rw, scopes: [...rw.scopes, 'https://graph.microsoft.com/Calendars.ReadWrite'] }).length, 10);
  assert.deepEqual(policy.availableOperations(undefined), []);
  assert.equal(calls, 0);
});
