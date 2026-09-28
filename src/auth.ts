import { PublicClientApplication, type AccountInfo } from "@azure/msal-node";
import fs from "node:fs";
import path from "node:path";

const CACHE_PATH = path.resolve(".cache/msal-token-cache.json");
const SCOPES = ["Mail.Read", "offline_access", "User.Read"];

function loadCache(): string | undefined {
  try {
    return fs.readFileSync(CACHE_PATH, "utf8");
  } catch {
    return undefined;
  }
}

function saveCache(serialized: string) {
  fs.mkdirSync(path.dirname(CACHE_PATH), { recursive: true });
  fs.writeFileSync(CACHE_PATH, serialized, "utf8");
}

function buildClient(): PublicClientApplication {
  const clientId = process.env.MS_CLIENT_ID;
  const tenantId = process.env.MS_TENANT_ID ?? "common";
  if (!clientId) {
    throw new Error("MS_CLIENT_ID is not set. Copy .env.example to .env and fill it in.");
  }

  const client = new PublicClientApplication({
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
    cache: {
      cachePlugin: {
        beforeCacheAccess: async (ctx) => {
          const cached = loadCache();
          if (cached) ctx.tokenCache.deserialize(cached);
        },
        afterCacheAccess: async (ctx) => {
          if (ctx.cacheHasChanged) saveCache(ctx.tokenCache.serialize());
        },
      },
    },
  });

  return client;
}

/** Interactive device-code login. Run once via `npm run login`. */
export async function deviceCodeLogin(): Promise<void> {
  const client = buildClient();
  const result = await client.acquireTokenByDeviceCode({
    scopes: SCOPES,
    deviceCodeCallback: (response) => {
      console.log(response.message);
    },
  });
  if (!result) throw new Error("Device code login did not return a token.");
  console.log(`Signed in as ${result.account?.username}. Token cached at ${CACHE_PATH}.`);
}

/** Gets a valid access token for Graph calls, using the cached account, refreshing silently. */
export async function getAccessToken(): Promise<string> {
  const client = buildClient();
  const cache = client.getTokenCache();
  const accounts: AccountInfo[] = await cache.getAllAccounts();
  if (accounts.length === 0) {
    throw new Error("No cached Microsoft account found. Run `npm run login` first.");
  }

  const result = await client.acquireTokenSilent({
    account: accounts[0],
    scopes: SCOPES,
  });
  if (!result) throw new Error("Failed to silently acquire a Graph access token.");
  return result.accessToken;
}
