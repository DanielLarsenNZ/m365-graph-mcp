# dan-m365-graph-mcp

A Model Context Protocol (MCP) server that exposes Microsoft 365 via the Graph API. Gives Claude (and other MCP clients) tools to read and manage your Outlook email, calendar events, and Microsoft Todo tasks.

## Features

- **Mail** — List, read, send, reply, move, delete, and mark emails; filter by Focused/Other inbox
- **Calendar** — List, create, update, and delete events; create Teams meetings; query by date range
- **Todo** — Full CRUD for task lists and tasks; set due dates, reminders, importance, and status
- **Auth** — Device code login, silent token refresh, and sign-out tools built in

## Requirements

- Node.js 18+
- An Azure app registration with the following delegated permissions:
  - `Mail.ReadWrite`, `Calendars.ReadWrite`, `Tasks.ReadWrite`, `User.Read`, `offline_access`

## Getting Started

### 1. Clone and install

```bash
git clone https://github.com/your-org/dan-m365-graph-mcp
cd dan-m365-graph-mcp
npm install
```

### 2. Create an Azure app registration

1. Go to [portal.azure.com](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**
2. Choose **Public client / native** redirect URI: `https://login.microsoftonline.com/common/oauth2/nativeclient`
3. Under **API permissions**, add the delegated MS Graph permissions listed above
4. Copy the **Application (client) ID** and optionally the **Directory (tenant) ID**

### 3. Configure environment

Create a `.env` file in the project root (or set these as environment variables):

```env
AZURE_CLIENT_ID=your-client-id-here
AZURE_TENANT_ID=common          # or your specific tenant ID
TOKEN_CACHE_PATH=.token-cache.json
```

### 4. Authenticate

Run the one-time auth CLI to sign in and cache a token:

```bash
npm run auth
```

Follow the on-screen prompt: visit the verification URL and enter the device code. The token is saved to `TOKEN_CACHE_PATH` and refreshed silently on subsequent uses.

### 5. Build and run

```bash
npm run build   # compile TypeScript → dist/
npm start       # run the MCP server (stdio transport)
```

For development without a build step:

```bash
npm run dev
```

### 6. Connect to Claude Code (or another MCP client)

Add the server to your `.mcp.json`:

```json
{
  "mcpServers": {
    "m365-graph": {
      "type": "stdio",
      "command": "node",
      "args": ["/absolute/path/to/dan-m365-graph-mcp/dist/index.js"],
      "env": {
        "AZURE_CLIENT_ID": "your-client-id",
        "AZURE_TENANT_ID": "your-tenant-id",
        "TOKEN_CACHE_PATH": "/absolute/path/to/.token-cache.json"
      }
    }
  }
}
```

## Available Tools (24 total)

| Category | Tools |
|----------|-------|
| **Auth** | `auth_status`, `auth_login`, `auth_logout` |
| **Mail** | `list_emails`, `get_email`, `send_email`, `reply_email`, `move_email`, `delete_email`, `mark_email_read` |
| **Calendar** | `list_events`, `get_event`, `create_event`, `update_event`, `delete_event`, `list_calendars` |
| **Todo** | `list_todo_lists`, `list_tasks`, `get_task`, `create_task`, `update_task`, `complete_task`, `delete_task`, `create_todo_list` |

See [docs/tools.md](docs/tools.md) for full parameter reference.

## Development

```bash
npm test        # run unit tests (Node built-in test runner)
npm run dev     # run server without compiling
```

See [docs/architecture.md](docs/architecture.md) for codebase internals and conventions.

## Timezone

Default timezone is `Pacific/Auckland`. Pass an explicit `timeZone` parameter to calendar and todo tools to override.
