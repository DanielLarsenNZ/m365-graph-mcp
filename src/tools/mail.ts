import { getGraphClient } from "../graph.js";

export const mailTools = [
  {
    name: "list_emails",
    description: "List emails from a mail folder (default: Inbox). Supports filtering and pagination.",
    inputSchema: {
      type: "object" as const,
      properties: {
        folder: { type: "string", description: "Folder name or ID (default: inbox)", default: "inbox" },
        top: { type: "number", description: "Max emails to return (default: 20, max: 50)", default: 20 },
        filter: { type: "string", description: "OData filter expression e.g. \"isRead eq false\"" },
        search: { type: "string", description: "Search query string" },
        inferenceClassification: { type: "string", enum: ["focused", "other"], description: "Filter by Focused Inbox category: 'focused' or 'other'. Omit to return both." },
      },
    },
  },
  {
    name: "get_email",
    description: "Get a single email by its ID, including full body.",
    inputSchema: {
      type: "object" as const,
      properties: {
        id: { type: "string", description: "Email message ID" },
      },
      required: ["id"],
    },
  },
  {
    name: "send_email",
    description: "Send a new email.",
    inputSchema: {
      type: "object" as const,
      properties: {
        to: { type: "string", description: "Recipient email address(es), comma-separated" },
        cc: { type: "string", description: "CC email address(es), comma-separated" },
        subject: { type: "string", description: "Email subject" },
        body: { type: "string", description: "Email body (HTML or plain text)" },
        bodyType: { type: "string", enum: ["html", "text"], default: "html", description: "Body content type" },
        from: { type: "string", description: "Sender address — must be an alias or shared mailbox on the account e.g. cowork@larsen.nz" },
        saveToSentItems: { type: "boolean", default: true },
      },
      required: ["to", "subject", "body"],
    },
  },
  {
    name: "reply_email",
    description: "Reply to an email.",
    inputSchema: {
      type: "object" as const,
      properties: {
        id: { type: "string", description: "Message ID to reply to" },
        body: { type: "string", description: "Reply body" },
        bodyType: { type: "string", enum: ["html", "text"], default: "html" },
        from: { type: "string", description: "Sender address override e.g. cowork@larsen.nz" },
        replyAll: { type: "boolean", default: false, description: "Reply to all recipients" },
      },
      required: ["id", "body"],
    },
  },
  {
    name: "move_email",
    description: "Move an email to a different folder.",
    inputSchema: {
      type: "object" as const,
      properties: {
        id: { type: "string", description: "Message ID" },
        destinationFolder: { type: "string", description: "Target folder ID or well-known name (inbox, deleteditems, drafts, sentitems, junkemail)" },
      },
      required: ["id", "destinationFolder"],
    },
  },
  {
    name: "delete_email",
    description: "Delete (trash) an email.",
    inputSchema: {
      type: "object" as const,
      properties: {
        id: { type: "string", description: "Message ID to delete" },
      },
      required: ["id"],
    },
  },
  {
    name: "mark_email_read",
    description: "Mark an email as read or unread.",
    inputSchema: {
      type: "object" as const,
      properties: {
        id: { type: "string", description: "Message ID" },
        isRead: { type: "boolean", description: "true = read, false = unread" },
      },
      required: ["id", "isRead"],
    },
  },
];

function parseAddresses(addresses: string) {
  return addresses.split(",").map((a) => ({
    emailAddress: { address: a.trim() },
  }));
}

/**
 * Build an OData $filter string for list_emails.
 * Merges an optional caller-supplied expression with an optional
 * inferenceClassification clause so both can coexist.
 *
 * @param filter  Raw OData expression from the caller (may be undefined)
 * @param inferenceClassification  "focused" | "other" | undefined
 * @returns Combined filter string, or undefined when no filters are needed
 */
export function buildMailFilter(
  filter: string | undefined,
  inferenceClassification: string | undefined,
): string | undefined {
  const parts: string[] = [];
  if (filter) parts.push(filter);
  if (inferenceClassification) parts.push(`inferenceClassification eq '${inferenceClassification}'`);
  return parts.length > 0 ? parts.join(" and ") : undefined;
}

export async function handleMailTool(name: string, args: Record<string, unknown>): Promise<string> {
  const client = getGraphClient();

  switch (name) {
    case "list_emails": {
      const folder = (args.folder as string) ?? "inbox";
      const top = Math.min((args.top as number) ?? 20, 50);
      let req = client
        .api(`/me/mailFolders/${folder}/messages`)
        .top(top)
        .select("id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments,inferenceClassification")
        .orderby("receivedDateTime DESC");
      const combinedFilter = buildMailFilter(
        args.filter as string | undefined,
        args.inferenceClassification as string | undefined,
      );
      if (combinedFilter) req = req.filter(combinedFilter);
      if (args.search) req = req.query({ $search: `"${args.search}"` });
      const result = await req.get();
      return JSON.stringify(result.value, null, 2);
    }

    case "get_email": {
      const msg = await client
        .api(`/me/messages/${args.id}`)
        .select("id,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,body,hasAttachments")
        .get();
      return JSON.stringify(msg, null, 2);
    }

    case "send_email": {
      const toList = parseAddresses(args.to as string);
      const message: Record<string, unknown> = {
        subject: args.subject,
        body: { contentType: args.bodyType ?? "html", content: args.body },
        toRecipients: toList,
      };
      if (args.cc) message.ccRecipients = parseAddresses(args.cc as string);
      if (args.from) message.from = { emailAddress: { address: args.from } };
      await client.api("/me/sendMail").post({
        message,
        saveToSentItems: args.saveToSentItems !== false,
      });
      return "Email sent successfully.";
    }

    case "reply_email": {
      const endpoint = args.replyAll
        ? `/me/messages/${args.id}/replyAll`
        : `/me/messages/${args.id}/reply`;
      const replyMessage: Record<string, unknown> = {
        body: { contentType: args.bodyType ?? "html", content: args.body },
      };
      if (args.from) replyMessage.from = { emailAddress: { address: args.from } };
      await client.api(endpoint).post({ message: replyMessage });
      return `Reply sent successfully.`;
    }

    case "move_email": {
      const result = await client.api(`/me/messages/${args.id}/move`).post({
        destinationId: args.destinationFolder,
      });
      return JSON.stringify({ success: true, newId: result.id }, null, 2);
    }

    case "delete_email": {
      await client.api(`/me/messages/${args.id}`).delete();
      return "Email deleted.";
    }

    case "mark_email_read": {
      await client.api(`/me/messages/${args.id}`).patch({ isRead: args.isRead });
      return `Email marked as ${args.isRead ? "read" : "unread"}.`;
    }

    default:
      throw new Error(`Unknown mail tool: ${name}`);
  }
}
