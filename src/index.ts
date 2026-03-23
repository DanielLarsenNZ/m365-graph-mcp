import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  ListToolsRequestSchema,
  CallToolRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { mailTools, handleMailTool } from "./tools/mail.js";
import { calendarTools, handleCalendarTool } from "./tools/calendar.js";
import { todoTools, handleTodoTool } from "./tools/todo.js";
import { isAuthenticated, signOut, getAccessToken } from "./auth.js";

// Load .env if present (dev convenience)
try {
  const { existsSync, readFileSync } = await import("fs");
  const envPath = new URL("../.env", import.meta.url).pathname;
  if (existsSync(envPath)) {
    for (const line of readFileSync(envPath, "utf-8").split("\n")) {
      const match = line.match(/^([A-Z_]+)=(.+)$/);
      if (match && !process.env[match[1]]) {
        process.env[match[1]] = match[2].trim();
      }
    }
  }
} catch {
  // ignore
}

const authTools = [
  {
    name: "auth_status",
    description: "Check whether the M365 Graph MCP is authenticated.",
    inputSchema: { type: "object" as const, properties: {} },
  },
  {
    name: "auth_login",
    description:
      "Authenticate with Microsoft 365 via device code flow. Returns a URL and code to complete sign-in.",
    inputSchema: { type: "object" as const, properties: {} },
  },
  {
    name: "auth_logout",
    description: "Sign out and clear the cached M365 token.",
    inputSchema: { type: "object" as const, properties: {} },
  },
];

const allTools = [...authTools, ...mailTools, ...calendarTools, ...todoTools];

const server = new Server(
  { name: "m365-graph", version: "0.1.0" },
  { capabilities: { tools: {} } }
);

server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: allTools,
}));

server.setRequestHandler(CallToolRequestSchema, async (req) => {
  const { name, arguments: args = {} } = req.params;
  const a = args as Record<string, unknown>;

  try {
    let text: string;

    if (name === "auth_status") {
      const authed = await isAuthenticated();
      text = authed
        ? "Authenticated."
        : "Not authenticated. Call auth_login to sign in.";
    } else if (name === "auth_login") {
      let instructions = "";
      await getAccessToken((url, code) => {
        instructions = `Visit: ${url}\nEnter code: ${code}`;
      });
      text = instructions || "Already authenticated (token reused from cache).";
    } else if (name === "auth_logout") {
      await signOut();
      text = "Signed out and token cache cleared.";
    } else if (mailTools.some((t) => t.name === name)) {
      text = await handleMailTool(name, a);
    } else if (calendarTools.some((t) => t.name === name)) {
      text = await handleCalendarTool(name, a);
    } else if (todoTools.some((t) => t.name === name)) {
      text = await handleTodoTool(name, a);
    } else {
      throw new Error(`Unknown tool: ${name}`);
    }

    return { content: [{ type: "text", text }] };
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    return { content: [{ type: "text", text: `Error: ${msg}` }], isError: true };
  }
});

const transport = new StdioServerTransport();
await server.connect(transport);
process.stderr.write("M365 Graph MCP server running on stdio\n");
