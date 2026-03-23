import {
  PublicClientApplication,
  DeviceCodeRequest,
  AccountInfo,
  TokenCacheContext,
} from "@azure/msal-node";
import { readFileSync, writeFileSync, existsSync } from "fs";

const SCOPES = [
  "Mail.ReadWrite",
  "Calendars.ReadWrite",
  "Tasks.ReadWrite",
  "User.Read",
  "offline_access",
];

function loadEnv() {
  const clientId = process.env.AZURE_CLIENT_ID;
  const tenantId = process.env.AZURE_TENANT_ID ?? "common";
  const cachePath = process.env.TOKEN_CACHE_PATH ?? ".token-cache.json";
  if (!clientId) throw new Error("AZURE_CLIENT_ID environment variable is required");
  return { clientId, tenantId, cachePath };
}

function makeMsalApp(clientId: string, tenantId: string, cachePath: string) {
  const beforeCacheAccess = async (ctx: TokenCacheContext) => {
    if (existsSync(cachePath)) {
      ctx.tokenCache.deserialize(readFileSync(cachePath, "utf-8"));
    }
  };
  const afterCacheAccess = async (ctx: TokenCacheContext) => {
    if (ctx.cacheHasChanged) {
      writeFileSync(cachePath, ctx.tokenCache.serialize(), "utf-8");
    }
  };

  return new PublicClientApplication({
    auth: { clientId, authority: `https://login.microsoftonline.com/${tenantId}` },
    cache: { cachePlugin: { beforeCacheAccess, afterCacheAccess } },
  });
}

let _app: PublicClientApplication | null = null;
let _cachePath = "";

function getApp() {
  if (!_app) {
    const { clientId, tenantId, cachePath } = loadEnv();
    _cachePath = cachePath;
    _app = makeMsalApp(clientId, tenantId, cachePath);
  }
  return _app;
}

async function getCachedAccount(): Promise<AccountInfo | null> {
  const app = getApp();
  const accounts = await app.getTokenCache().getAllAccounts();
  return accounts[0] ?? null;
}

export async function getAccessToken(deviceCodeCallback?: (url: string, code: string) => void): Promise<string> {
  const app = getApp();
  const account = await getCachedAccount();

  if (account) {
    try {
      const result = await app.acquireTokenSilent({ scopes: SCOPES, account });
      if (result?.accessToken) return result.accessToken;
    } catch {
      // fall through to device code
    }
  }

  // Device code flow
  const request: DeviceCodeRequest = {
    scopes: SCOPES,
    deviceCodeCallback: (response) => {
      if (deviceCodeCallback) {
        deviceCodeCallback(response.verificationUri, response.userCode);
      } else {
        // Write to stderr so it doesn't corrupt MCP stdio
        process.stderr.write(
          `\n[M365 Auth] Visit: ${response.verificationUri}\nCode: ${response.userCode}\n\n`
        );
      }
    },
  };

  const result = await app.acquireTokenByDeviceCode(request);
  if (!result?.accessToken) throw new Error("Authentication failed — no token returned");
  return result.accessToken;
}

export async function isAuthenticated(): Promise<boolean> {
  try {
    const account = await getCachedAccount();
    return account !== null;
  } catch {
    return false;
  }
}

export async function signOut(): Promise<void> {
  const app = getApp();
  const account = await getCachedAccount();
  if (account) {
    await app.getTokenCache().removeAccount(account);
    writeFileSync(_cachePath || ".token-cache.json", "", "utf-8");
  }
}
