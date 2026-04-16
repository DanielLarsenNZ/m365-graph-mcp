# Creating an Azure Entra App Registration

This guide walks through creating an Azure Entra (formerly Azure Active Directory) app registration that grants this MCP server delegated access to Microsoft 365 on behalf of a user.

## Prerequisites

- An Azure account with access to the [Azure Portal](https://portal.azure.com)
- Permission to register applications in your Entra tenant (or a personal Microsoft account)

---

## Step 1: Open App Registrations

1. Sign in to the [Azure Portal](https://portal.azure.com).
2. In the top search bar, type **App registrations** and select it.
3. Click **+ New registration**.

---

## Step 2: Register the Application

Fill in the registration form:

| Field | Value |
|-------|-------|
| **Name** | Any descriptive name, e.g. `m365-graph-mcp` |
| **Supported account types** | Choose *Accounts in this organizational directory only* (single tenant) |
| **Redirect URI** | Select **Public client / native (mobile & desktop)** from the dropdown, then enter `https://login.microsoftonline.com/common/oauth2/nativeclient` |

Click **Register**.

---

## Step 3: Copy the Client and Tenant IDs

After registration you land on the app's **Overview** page.

1. Copy the **Application (client) ID** — this is your `AZURE_CLIENT_ID`.
2. Copy the **Directory (tenant) ID** — this is your `AZURE_TENANT_ID`.
   - If you chose *multi-tenant / personal accounts* in Step 2, you can use `common` as the tenant ID instead.

Save these values; you will add them to your `.env` file shortly.

---

## Step 4: Configure the Platform (Public Client)

The device-code flow requires the app to be treated as a public client.

1. In the left menu, click **Authentication**.
2. Scroll down to **Advanced settings**.
3. Set **Allow public client flows** to **Yes**.
4. Click **Save**.

---

## Step 5: Add API Permissions

1. In the left menu, click **API permissions**.
2. Click **+ Add a permission** → **Microsoft Graph** → **Delegated permissions**.
3. Search for and select each of the following permissions:

| Permission | Purpose |
|------------|---------|
| `Mail.ReadWrite` | Read, send, move, and delete email |
| `Calendars.ReadWrite` | Read and manage calendar events |
| `Tasks.ReadWrite` | Read and manage Microsoft Todo tasks |
| `User.Read` | Read the signed-in user's profile |
| `offline_access` | Maintain access via refresh tokens |

4. Click **Add permissions**.

> **Note:** You do **not** need to click *Grant admin consent* for delegated permissions on personal or single-user setups — the user grants consent at sign-in.

---

## Step 6: Verify the Configuration

Your **API permissions** page should list:

```
Microsoft Graph (5)
  ✔ Calendars.ReadWrite   Delegated
  ✔ Mail.ReadWrite        Delegated
  ✔ offline_access        Delegated
  ✔ Tasks.ReadWrite       Delegated
  ✔ User.Read             Delegated
```

The status column may show *Not granted for \<tenant\>* until the user completes the first sign-in — that is expected.

---

## Step 7: Configure Your Environment

Create a `.env` file in the project root (copy from `.env.example`):

```env
AZURE_CLIENT_ID=<paste Application (client) ID here>
AZURE_TENANT_ID=<paste Directory (tenant) ID here, or use "common">
TOKEN_CACHE_PATH=.token-cache.json
```

---

## Next Step

Run the one-time authentication flow:

```bash
npm run auth
```

Follow the on-screen prompt: visit the verification URL and enter the device code shown in the terminal. On success, a token is saved to `TOKEN_CACHE_PATH` and refreshed silently on every subsequent use.
