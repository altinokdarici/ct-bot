# ct-bot

Teams ↔ Claude Code bridge. Each Teams thread = independent Claude Code session.

## Commands

- `npm run build` — type-check (tsc, noEmit)
- `npm run start` — dev server (requires `.env` with `CHANNEL_ID`)

## Prerequisites

- `az login` — Teams auth tokens
- Claude Code API key in macOS Keychain (service: "Claude Code")
- `gh` CLI authenticated — codespace SSH

## Architecture

Teams message → WebSocket → `orchestrator.ts` → `session-manager.ts` → `claude-session.ts` (Agent SDK `query()`) → response edited back into Teams thread.

Handoff: Claude emits `<!--HANDOFF:local:/path-->` markers to switch projects mid-conversation.

## Code Conventions

- TypeScript ESM, `.ts` imports, `node:` prefix for built-ins
- Strict mode with `noUncheckedIndexedAccess`
- Teams responses are HTML (not markdown)
