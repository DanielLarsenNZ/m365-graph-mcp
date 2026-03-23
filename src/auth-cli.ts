/**
 * Run this once from the terminal to authenticate:
 *   npm run auth
 *
 * It will print a device code URL and wait for you to sign in.
 * On success the token is cached at .token-cache.json and the
 * MCP server will use it silently from then on.
 */

import { PublicClientApplication, TokenCacheContext } from "@azure/msal-node";
import { readFileSync, writeFileSync, existsSync } from "fs";
import { resolve, dirname } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const projectRoot = resolve(__dirname, "..");

// Load .env
const envPath = resolve(projectRoot, ".env");
if (existsSync(envPath)) {
  for (const line of readFileSync(envPath, "utf-8").split("\n")) {
    const match = line.match(/^([A-Z_]+)=(.+)$/);
    if (match && !process.env[match[1]]) process.env[match[1]] = match[2].trim();
  }
}

const clientId = process.env.AZURE_CLIENT_ID;
const tenantId = process.env.AZURE_TENANT_ID ?? "common";
const cachePath = resolve(projectRoot, process.env.TOKEN_CACHE_PATH ?? ".token-cache.json");

if (!clientId) {
  console.error("ERROR: AZURE_CLIENT_ID not set in .env");
  process.exit(1);
}

const beforeCacheAccess = async (ctx: TokenCacheContext) => {
  if (existsSync(cachePath)) ctx.tokenCache.deserialize(readFileSync(cachePath, "utf-8"));
};
const afterCacheAccess = async (ctx: TokenCacheContext) => {
  if (ctx.cacheHasChanged) writeFileSync(cachePath, ctx.tokenCache.serialize(), "utf-8");
};

const app = new PublicClientApplication({
  auth: { clientId, authority: `https://login.microsoftonline.com/${tenantId}` },
  cache: { cachePlugin: { beforeCacheAccess, afterCacheAccess } },
});

const SCOPES = [
  "Mail.ReadWrite",
  "Calendars.ReadWrite",
  "Tasks.ReadWrite",
  "User.Read",
  "offline_access",
];

// Check if already authenticated
const accounts = await app.getTokenCache().getAllAccounts();
if (accounts.length > 0) {
  try {
    const silent = await app.acquireTokenSilent({ scopes: SCOPES, account: accounts[0] });
    if (silent?.accessToken) {
      console.log(`✅ Already authenticated as: ${accounts[0].username}`);
      process.exit(0);
    }
  } catch {
    console.log("Cached token expired, re-authenticating...");
  }
}

// Device code flow
console.log("\n🔐 Microsoft 365 Authentication\n");

const result = await app.acquireTokenByDeviceCode({
  scopes: SCOPES,
  deviceCodeCallback: (response) => {
    console.log(response.message);
    console.log("\n" + "─".repeat(50));
    console.log(`  URL:  ${response.verificationUri}`);
    console.log(`  Code: ${response.userCode}`);
    console.log("─".repeat(50) + "\n");
    console.log("Waiting for you to sign in...");
  },
});

if (result?.account) {
  console.log(`\n✅ Authenticated successfully as: ${result.account.username}`);
  console.log(`   Token cached at: ${cachePath}`);
  console.log("   The MCP server will now use this token silently.\n");
} else {
  console.error("Authentication failed — no account returned.");
  process.exit(1);
}
