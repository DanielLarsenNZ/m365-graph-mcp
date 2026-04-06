# Architecture

## Overview

The server uses the [Model Context Protocol SDK](https://github.com/modelcontextprotocol/typescript-sdk) over stdio transport. Each incoming tool call is dispatched to one of four handler modules (auth, mail, calendar, todo), which interact with Microsoft Graph via MSAL token acquisition.

```
MCP client (e.g. Claude Code)
        │ stdio
        ▼
   src/index.ts          ← MCP server, tool registry, request dispatch
        │
        ├── src/auth.ts          ← MSAL token management
        ├── src/graph.ts         ← Graph client factory
        ├── src/tools/mail.ts    ← Mail tools & handlers
        ├── src/tools/calendar.ts← Calendar tools & handlers
        └── src/tools/todo.ts    ← Todo tools & handlers (custom fetch)
```

## File-by-file

### `src/index.ts`
Entry point. Responsibilities:
- Loads `.env` if present (development convenience via `dotenv`)
- Constructs the MCP `Server` instance and registers all tool definitions
- Implements the single `CallToolRequestSchema` handler: routes by tool name prefix to the appropriate handler, wraps the result as an MCP text content block, returns errors with `isError: true`
- Connects to stdio transport and starts listening

### `src/auth.ts`
MSAL authentication wrapper. Exports:
- `getAccessToken(deviceCodeCallback?)` — acquires a token silently from cache first; falls back to device code flow and invokes the callback with the verification URI and user code
- `isAuthenticated()` — returns true if an account exists in the token cache
- `signOut()` — removes the account from MSAL and clears the cache file

Uses a singleton `PublicClientApplication` initialized with `AZURE_CLIENT_ID` and `AZURE_TENANT_ID`. Token cache is serialized to/from `TOKEN_CACHE_PATH` on every acquire/clear via MSAL's `ICachePlugin` interface.

OAuth scopes: `Mail.ReadWrite`, `Calendars.ReadWrite`, `Tasks.ReadWrite`, `User.Read`, `offline_access`

### `src/auth-cli.ts`
Standalone CLI for one-time authentication. Runs the same `getAccessToken()` path as the server but with a human-readable console prompt, then exits. Intended to be run once before starting the server: `npm run auth`.

### `src/graph.ts`
Exports `getGraphClient()`, a thin factory that returns a `Client` instance (from `@microsoft/microsoft-graph-client`) configured with an auth provider that calls `getAccessToken()`. Used by mail and calendar handlers.

### `src/tools/mail.ts`
Uses the Graph Client from `src/graph.ts`. Exports:
- `mailTools` — array of MCP tool definition objects (schema + description)
- `buildMailFilter(filter, inferenceClassification)` — merges OData clauses; exported for unit testing
- `handleMailTool(name, args)` — switch/dispatch over tool names

### `src/tools/calendar.ts`
Uses the Graph Client. Exports:
- `calendarTools` — tool definitions
- `handleCalendarTool(name, args)` — dispatch

### `src/tools/todo.ts`
Uses a **custom `gFetch()` wrapper** around native `fetch` rather than the Graph Client library. This was necessary because the Graph Client's URL builder double-encoded Exchange-format list IDs (containing `==`), causing 404 errors. The custom wrapper calls `getAccessToken()` directly and constructs URLs manually with explicit `encodeURIComponent()`.

Key helper functions (all exported for unit testing):
- `gFetch(path, method, body?)` — Graph API fetch with Bearer token; handles 204 No Content
- `listPath(listId)` — `/me/todo/lists/{encodedListId}`
- `taskPath(listId, taskId)` — `/me/todo/lists/{encodedListId}/tasks/{encodedTaskId}`
- `listTasksPath(listId, top, filter?)` — full query path with `$top` and optional `$filter`
- `buildTaskPayload(args)` — constructs the POST/PATCH body; wraps `dueDateTime` and `reminderDateTime` with timezone

## Known Graph API quirks

### Exchange-format ID encoding
Todo list and task IDs for Exchange-connected accounts (AQMk…, AAMk…) contain `=` and `==`. These must be `encodeURIComponent`-encoded in URL path segments. The Graph Client library does not do this correctly, hence the custom `gFetch` approach for Todo.

### `$select` in `list_tasks`
Calling `/me/todo/lists/{id}/tasks?$select=…` returns HTTP 400 for Exchange-format list IDs. The `list_tasks` handler deliberately omits `$select`.

## Testing

Tests live alongside source files: `src/tools/mail.test.ts`, `src/tools/todo.test.ts`.

All testable logic is extracted into exported pure functions (`buildMailFilter`, `listPath`, `taskPath`, `listTasksPath`, `buildTaskPayload`). Tests use Node's built-in `node:test` runner — no additional framework.

Run with:
```bash
npm test
```

The test glob (`src/tools/*.test.ts`) picks up all test files automatically.

### Conventions
- Every new tool must have tests before committing
- Bug fixes require a regression test that fails before the fix and passes after, named to describe the bug
- See `AGENTS.md` for the full rules

## Environment variables

| Variable | Required | Default | Description |
|----------|----------|---------|-------------|
| `AZURE_CLIENT_ID` | yes | — | Azure app registration client ID |
| `AZURE_TENANT_ID` | no | `"common"` | Tenant ID (`"common"` for multi-tenant) |
| `TOKEN_CACHE_PATH` | no | `.token-cache.json` | Path to persist the MSAL token cache |

## Build

TypeScript is compiled to `dist/` targeting ES2022 with `moduleResolution: bundler`. Source maps are emitted. The `tsx` dev dependency is used for tests and `npm run dev` to avoid the compile step during development.
