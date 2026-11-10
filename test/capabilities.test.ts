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

test('draft operations need read-write mode plus Mail.ReadWrite; sending needs Mail.Send as well', () => {
  const old = { mode: 'read-write' as const, scopes: ['User.Read', 'Mail.Read', 'Calendars.ReadWrite'], expiresAt: Date.now() + 10000 };
  const before = ['me', 'list', 'search', 'read', 'calendar.list', 'calendar.events', 'calendar.get', 'calendar.create', 'calendar.update', 'calendar.delete'];
  assert.deepEqual(policy.availableOperations(old), before);
  assert.throws(() => policy.requireOperation(old, 'draft.create'), { code: 'INSUFFICIENT_SCOPE' });
  const drafts = { ...old, scopes: [...old.scopes, 'Mail.ReadWrite'] };
  assert.deepEqual(policy.availableOperations(drafts), [...before, 'draft.create', 'draft.reply', 'draft.update']);
  assert.throws(() => policy.requireOperation(drafts, 'draft.send'), { code: 'INSUFFICIENT_SCOPE' });
  // Mail.Send alone must not unlock sending: both scopes are required, not either.
  assert.deepEqual(policy.availableOperations({ ...old, scopes: [...old.scopes, 'Mail.Send'] }), before);
  const full = { ...old, scopes: [...drafts.scopes, 'https://graph.microsoft.com/Mail.Send'] };
  assert.deepEqual(policy.availableOperations(full), [...before, 'draft.create', 'draft.reply', 'draft.update', 'draft.send']);
  const readOnly = { ...full, mode: 'read-only' as const };
  for (const operation of ['draft.create', 'draft.reply', 'draft.update', 'draft.send', 'calendar.create'] as const) {
    assert.throws(() => policy.requireOperation(readOnly, operation), (error: any) => {
      assert.equal(error.code, 'READ_ONLY_MODE');
      assert.doesNotMatch(error.message, /calendar/i);
      return true;
    });
  }
  assert.deepEqual(policy.availableOperations(readOnly), ['me', 'list', 'search', 'read', 'calendar.list', 'calendar.events', 'calendar.get']);
});
