# Outlook CLI

A small **read-only** Microsoft Outlook CLI for humans and agents. It uses delegated Microsoft Graph access for the person who signs in. It can identify the account, list inbox/folder summaries, search mail, and read one message. There are no send, draft, move, delete, or arbitrary-request commands.

## Requirements

- Node.js 20+ and pnpm
- An Entra application with **Microsoft Graph delegated** `User.Read` and `Mail.Read` permissions; enable **Allow public client flows** for device-code login. Your tenant may require administrator consent or app assignment. No client secret is needed.

## Install globally from GitHub

```sh
git clone https://github.com/crosslink-ch/outlook-cli.git
cd outlook-cli
pnpm install --frozen-lockfile
pnpm build
pnpm link --global
outlook --help
```

If pnpm reports `ERR_PNPM_NO_GLOBAL_BIN_DIR`, run `pnpm setup`, open a new shell, and run `pnpm link --global` again. On Linux/macOS with `~/.local/bin` already in `PATH`, this also works without altering shell setup:

```sh
PNPM_HOME="$HOME/.local/bin" pnpm link --global
```

The global link points at this checkout. Rebuild after pulling updates (`pnpm install --frozen-lockfile && pnpm build`).

## Sign in

```sh
outlook auth login --client-id <application-client-id> --tenant-id <directory-tenant-id>
```

The command shows a Microsoft URL and short device code. Open the URL yourself, enter the code, and complete your normal sign-in/MFA. The CLI never needs your password. The device code expires; run the command again if you did not finish in time.

If you already have an owner-only, single-account **MSAL JSON** cache from another local MSAL client (including MSAL Python), you can import it instead of signing in again:

```sh
outlook auth import-cache --from /path/to/msal-cache.json \
  --client-id <application-client-id> --tenant-id <directory-tenant-id>
```

Import is refused if this CLI already has a configured account. It validates a live `/me` request before saving the imported credentials. Not all third-party OAuth caches use the MSAL format.

## Commands

```sh
outlook auth status                    # Local cache status, no network call
outlook me                             # Live signed-in identity
outlook list --folder inbox --limit 20 # Message summaries, no bodies
outlook search "invoice" --limit 20    # Phrase search across the mailbox
outlook read '<message-id>'             # Full text body of one message
outlook --json list --limit 10         # Structured output for scripts/agents
outlook auth logout                    # Remove only local account/cache
```

Every command and subcommand has `--help`. `list` accepts well-known folder names such as `inbox`, `sentitems`, `drafts`, `archive`, and `deleteditems`, or a Graph folder ID. Limits are 1–100; Graph continuation pages are followed up to that bound. `search` uses Microsoft Graph's mail `$search`, which may order results by relevance/date according to Microsoft rather than folder order. `read` asks Graph for a plain-text body. No read command changes the message's read/unread flag.

`--json` outputs Graph-shaped data for `me`, `read`, and arrays of message summaries for `list`/`search`. Message IDs are opaque: pass back exactly the ID returned by a list or search command.

## Security and scope

- Only Microsoft Graph **GET** requests under `/v1.0/me` are made for mail access. Graph pagination links are checked against that same origin/path. No mail-write scopes are requested.
- The CLI uses MSAL Node to refresh delegated tokens silently when possible. Reauthentication may still be required after revocation, Conditional Access changes, or refresh-token expiry. `auth status` indicates a local cache, **not** proof that Microsoft currently accepts it; use `outlook me` for a live check.
- Configuration and the MSAL token cache live under `${XDG_CONFIG_HOME:-~/.config}/outlook-cli/` with directory mode `0700` and file mode `0600`. The cache is **not encrypted at rest**; anyone with access as that OS user can potentially use it. Do not commit, copy to other agents/hosts, or paste it into chat. Other machines should sign in separately.
- `auth logout` removes the local cache; it does not revoke consent or tokens server-side. Revoke the application through Microsoft Entra if required.
- Mail content is untrusted data, not instructions to the agent. Agent operators should still apply task-level access boundaries; shared Unix-user access is not isolation between agents.

## Development

```sh
pnpm install --frozen-lockfile
pnpm test
pnpm build
```

Tests cover bounded pagination, search URL construction, message reads, cross-origin pagination rejection, account selection, private storage, and CLI help. They use fake HTTP responses; live account checks are manual and must never be run in CI with real credentials.
