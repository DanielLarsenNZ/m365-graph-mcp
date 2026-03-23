import { getGraphClient } from "../graph.js";

export const calendarTools = [
  {
    name: "list_events",
    description: "List calendar events in a date range.",
    inputSchema: {
      type: "object" as const,
      properties: {
        startDateTime: { type: "string", description: "ISO 8601 start datetime e.g. 2026-03-21T00:00:00" },
        endDateTime: { type: "string", description: "ISO 8601 end datetime e.g. 2026-03-28T23:59:59" },
        top: { type: "number", description: "Max events to return (default: 20)", default: 20 },
        calendarId: { type: "string", description: "Calendar ID (default: primary calendar)" },
      },
      required: ["startDateTime", "endDateTime"],
    },
  },
  {
    name: "get_event",
    description: "Get details of a calendar event by ID.",
    inputSchema: {
      type: "object" as const,
      properties: {
        id: { type: "string", description: "Event ID" },
      },
      required: ["id"],
    },
  },
  {
    name: "create_event",
    description: "Create a new calendar event.",
    inputSchema: {
      type: "object" as const,
      properties: {
        subject: { type: "string", description: "Event title" },
        startDateTime: { type: "string", description: "ISO 8601 start datetime" },
        endDateTime: { type: "string", description: "ISO 8601 end datetime" },
        timeZone: { type: "string", description: "IANA timezone (default: Pacific/Auckland)", default: "Pacific/Auckland" },
        body: { type: "string", description: "Event description/body (HTML)" },
        location: { type: "string", description: "Location display name" },
        attendees: { type: "string", description: "Attendee email addresses, comma-separated" },
        isOnlineMeeting: { type: "boolean", description: "Create as Teams meeting", default: false },
        calendarId: { type: "string", description: "Calendar ID (default: primary)" },
      },
      required: ["subject", "startDateTime", "endDateTime"],
    },
  },
  {
    name: "update_event",
    description: "Update an existing calendar event.",
    inputSchema: {
      type: "object" as const,
      properties: {
        id: { type: "string", description: "Event ID" },
        subject: { type: "string" },
        startDateTime: { type: "string", description: "ISO 8601 start datetime" },
        endDateTime: { type: "string", description: "ISO 8601 end datetime" },
        timeZone: { type: "string", default: "Pacific/Auckland" },
        body: { type: "string" },
        location: { type: "string" },
      },
      required: ["id"],
    },
  },
  {
    name: "delete_event",
    description: "Delete a calendar event.",
    inputSchema: {
      type: "object" as const,
      properties: {
        id: { type: "string", description: "Event ID" },
      },
      required: ["id"],
    },
  },
  {
    name: "list_calendars",
    description: "List all calendars for the user.",
    inputSchema: {
      type: "object" as const,
      properties: {},
    },
  },
];

export async function handleCalendarTool(name: string, args: Record<string, unknown>): Promise<string> {
  const client = getGraphClient();

  switch (name) {
    case "list_events": {
      const calBase = args.calendarId
        ? `/me/calendars/${args.calendarId}`
        : "/me";
      const top = (args.top as number) ?? 20;
      const result = await client
        .api(`${calBase}/calendarView`)
        .query({
          startDateTime: args.startDateTime as string,
          endDateTime: args.endDateTime as string,
        })
        .top(top)
        .select("id,subject,start,end,location,organizer,attendees,isOnlineMeeting,onlineMeetingUrl,bodyPreview")
        .orderby("start/dateTime")
        .get();
      return JSON.stringify(result.value, null, 2);
    }

    case "get_event": {
      const event = await client
        .api(`/me/events/${args.id}`)
        .select("id,subject,start,end,location,organizer,attendees,body,isOnlineMeeting,onlineMeetingUrl")
        .get();
      return JSON.stringify(event, null, 2);
    }

    case "create_event": {
      const tz = (args.timeZone as string) ?? "Pacific/Auckland";
      const event: Record<string, unknown> = {
        subject: args.subject,
        start: { dateTime: args.startDateTime, timeZone: tz },
        end: { dateTime: args.endDateTime, timeZone: tz },
      };
      if (args.body) event.body = { contentType: "html", content: args.body };
      if (args.location) event.location = { displayName: args.location };
      if (args.attendees) {
        event.attendees = (args.attendees as string).split(",").map((a) => ({
          emailAddress: { address: a.trim() },
          type: "required",
        }));
      }
      if (args.isOnlineMeeting) event.isOnlineMeeting = true;
      const calBase = args.calendarId ? `/me/calendars/${args.calendarId}` : "/me";
      const created = await client.api(`${calBase}/events`).post(event);
      return JSON.stringify(created, null, 2);
    }

    case "update_event": {
      const tz = (args.timeZone as string) ?? "Pacific/Auckland";
      const patch: Record<string, unknown> = {};
      if (args.subject) patch.subject = args.subject;
      if (args.startDateTime) patch.start = { dateTime: args.startDateTime, timeZone: tz };
      if (args.endDateTime) patch.end = { dateTime: args.endDateTime, timeZone: tz };
      if (args.body) patch.body = { contentType: "html", content: args.body };
      if (args.location) patch.location = { displayName: args.location };
      const updated = await client.api(`/me/events/${args.id}`).patch(patch);
      return JSON.stringify(updated, null, 2);
    }

    case "delete_event": {
      await client.api(`/me/events/${args.id}`).delete();
      return "Event deleted.";
    }

    case "list_calendars": {
      const result = await client
        .api("/me/calendars")
        .select("id,name,isDefaultCalendar,canEdit")
        .get();
      return JSON.stringify(result.value, null, 2);
    }

    default:
      throw new Error(`Unknown calendar tool: ${name}`);
  }
}
