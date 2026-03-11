# ct-bot

Teams channel → Claude Code bridge. Each thread is an independent Claude Code session.

## Setup

```bash
npm install
az login
npm run setup
```

`setup` checks prerequisites (Node 18+, `az`, Claude Code CLI) and connects to your Teams channel — just post the code it gives you.

## Run

```bash
npm start
```

## How it works

Send a message in the Teams channel → bot picks it up via WebSocket → runs Claude Code (Agent SDK) → edits the response back into the thread.

- **Multi-project**: say "switch to repo-x" and it hands off to a new session
- **WorkIQ**: asks about emails, calendar, docs via M365 integration
- **Memory**: persistent file shared across sessions for repo paths, preferences
- **Full tools**: Read, Edit, Bash, Grep — everything Claude Code can do
