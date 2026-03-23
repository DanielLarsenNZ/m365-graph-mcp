import { Client } from "@microsoft/microsoft-graph-client";
import { getAccessToken } from "./auth.js";

export function getGraphClient(): Client {
  return Client.initWithMiddleware({
    authProvider: {
      getAccessToken: () => getAccessToken(),
    },
  });
}
