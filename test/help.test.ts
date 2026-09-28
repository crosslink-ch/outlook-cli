import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createProgram } from '../src/program.js';

test('help documents the read-only commands and the authentication flow', () => {
  const help = createProgram().helpInformation();
  for (const item of ['auth', 'me', 'list', 'search', 'read', '--json']) {
    assert.match(help, new RegExp(item));
  }
  assert.match(help, /read-only/i);
  assert.doesNotMatch(help, /send|delete message/i);
});
