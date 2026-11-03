import { test } from 'node:test';
import assert from 'node:assert/strict';
import { run } from '../src/program.js';
import { GraphClient } from '../src/graph.js';
import { AuthManager } from '../src/auth.js';

test('existing mail and auth status JSON remains directly consumable by legacy scripts', async () => {
  const grant = { mode: 'read-only' as const, scopes: ['User.Read', 'Mail.Read'], expiresAt: Date.now() + 3600000 };
  const message = { id: 'message-id', subject: 'Synthetic test' };
  const graph = new GraphClient('synthetic', async input => Response.json(String(input).includes('/messages/message-id') ? message : { value: [message] }), grant);
  const current = { grant, graph, identity: { clientId: 'client', tenantId: 'tenant', homeAccountId: 'home', username: 'person@example.test', meId: 'user', cacheFile: 'unused' }, me: { id: 'user', mail: 'person@example.test' } };
  const manager = { connection: async () => current } as AuthManager;
  const legacyCases = [['list'], ['search', 'synthetic'], ['read', 'message-id'], ['me'], ['auth', 'status']];
  for (const command of legacyCases) for (const args of [[...command, '--json'], ['--json', ...command]]) {
    const output: string[] = [];
    assert.equal(await run(args, { manager, io: { out: text => output.push(text), err: () => {} } }), 0);
    const data = JSON.parse(output.join(''));
    if (command[0] === 'list' || command[0] === 'search') assert.deepEqual(data, [message]);
    else if (command[0] === 'read') assert.deepEqual(data, message);
    else if (command[0] === 'me') assert.deepEqual(data, current.me);
    else { assert.equal(data.signedIn, true); assert.equal(data.mode, 'read-only'); assert.equal(data.username, 'person@example.test'); }
  }
});
