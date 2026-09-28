import { test } from 'node:test';
import assert from 'node:assert/strict';
import { chooseAccount, parseIdentity } from '../src/auth.js';

test('cached access selects the configured account, not an arbitrary first account', () => {
  const accounts = [{ homeAccountId: 'other' }, { homeAccountId: 'chosen' }];
  assert.equal(chooseAccount(accounts, 'chosen')?.homeAccountId, 'chosen');
  assert.throws(() => chooseAccount(accounts, 'missing'), /No cached account matches/);
});

test('invalid Entra identifiers are rejected before network or cache access', () => {
  assert.throws(() => parseIdentity('not-a-guid', '00000000-0000-0000-0000-000000000000'), /application \(client\) ID/);
});
