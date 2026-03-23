import { getAccessToken } from "../auth.js";

const GRAPH = "https://graph.microsoft.com/v1.0";

async function gFetch(path: string, method = "GET", body?: unknown) {
  const token = await getAccessToken();
  const opts: RequestInit = {
    method,
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
  };
  if (body !== undefined) opts.body = JSON.stringify(body);
  const r = await fetch(`${GRAPH}${path}`, opts);
  if (r.status === 204) return null;
  const text = await r.text();
  if (!r.ok) throw new Error(`Graph ${r.status}: ${text}`);
  return text ? JSON.parse(text) : null;
}

export const todoTools = [
  {
    name: "list_todo_lists",
    description: "List all Microsoft Todo task lists.",
    inputSchema: {
      type: "object" as const,
      properties: {},
    },
  },
  {
    name: "list_tasks",
    description: "List tasks in a Todo list.",
    inputSchema: {
      type: "object" as const,
      properties: {
        listId: { type: "string", description: "Task list ID (use list_todo_lists to find IDs)" },
        filter: { type: "string", description: "OData filter e.g. \"status ne 'completed'\"" },
        top: { type: "number", description: "Max tasks to return (default: 50)", default: 50 },
      },
      required: ["listId"],
    },
  },
  {
    name: "get_task",
    description: "Get a specific task by ID.",
    inputSchema: {
      type: "object" as const,
      properties: {
        listId: { type: "string", description: "Task list ID" },
        taskId: { type: "string", description: "Task ID" },
      },
      required: ["listId", "taskId"],
    },
  },
  {
    name: "create_task",
    description: "Create a new task in a Todo list.",
    inputSchema: {
      type: "object" as const,
      properties: {
        listId: { type: "string", description: "Task list ID" },
        title: { type: "string", description: "Task title" },
        body: { type: "string", description: "Task notes/body" },
        dueDateTime: { type: "string", description: "Due date ISO 8601 e.g. 2026-03-25T00:00:00" },
        importance: { type: "string", enum: ["low", "normal", "high"], default: "normal" },
        reminderDateTime: { type: "string", description: "Reminder datetime ISO 8601" },
      },
      required: ["listId", "title"],
    },
  },
  {
    name: "update_task",
    description: "Update an existing task.",
    inputSchema: {
      type: "object" as const,
      properties: {
        listId: { type: "string", description: "Task list ID" },
        taskId: { type: "string", description: "Task ID" },
        title: { type: "string" },
        body: { type: "string" },
        dueDateTime: { type: "string", description: "Due date ISO 8601" },
        importance: { type: "string", enum: ["low", "normal", "high"] },
        status: { type: "string", enum: ["notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"] },
      },
      required: ["listId", "taskId"],
    },
  },
  {
    name: "complete_task",
    description: "Mark a task as completed.",
    inputSchema: {
      type: "object" as const,
      properties: {
        listId: { type: "string", description: "Task list ID" },
        taskId: { type: "string", description: "Task ID" },
      },
      required: ["listId", "taskId"],
    },
  },
  {
    name: "delete_task",
    description: "Delete a task.",
    inputSchema: {
      type: "object" as const,
      properties: {
        listId: { type: "string", description: "Task list ID" },
        taskId: { type: "string", description: "Task ID" },
      },
      required: ["listId", "taskId"],
    },
  },
  {
    name: "create_todo_list",
    description: "Create a new Todo task list.",
    inputSchema: {
      type: "object" as const,
      properties: {
        displayName: { type: "string", description: "Name for the new list" },
      },
      required: ["displayName"],
    },
  },
];

function buildTaskPayload(args: Record<string, unknown>) {
  const payload: Record<string, unknown> = {};
  if (args.title) payload.title = args.title;
  if (args.body) payload.body = { content: args.body, contentType: "text" };
  if (args.dueDateTime) {
    payload.dueDateTime = { dateTime: args.dueDateTime, timeZone: "Pacific/Auckland" };
  }
  if (args.importance) payload.importance = args.importance;
  if (args.status) payload.status = args.status;
  if (args.reminderDateTime) {
    payload.reminderDateTime = { dateTime: args.reminderDateTime, timeZone: "Pacific/Auckland" };
    payload.isReminderOn = true;
  }
  return payload;
}

export function listPath(listId: string) {
  return `/me/todo/lists/${encodeURIComponent(listId)}`;
}

export function taskPath(listId: string, taskId: string) {
  return `${listPath(listId)}/tasks/${encodeURIComponent(taskId)}`;
}

export function listTasksPath(listId: string, top: number, filter?: string) {
  let path = `${listPath(listId)}/tasks?$top=${top}`;
  if (filter) path += `&$filter=${encodeURIComponent(filter)}`;
  return path;
}

export async function handleTodoTool(name: string, args: Record<string, unknown>): Promise<string> {
  switch (name) {
    case "list_todo_lists": {
      const result = await gFetch("/me/todo/lists");
      return JSON.stringify(result.value, null, 2);
    }

    case "list_tasks": {
      const result = await gFetch(
        listTasksPath(args.listId as string, (args.top as number) ?? 50, args.filter as string | undefined)
      );
      return JSON.stringify(result.value, null, 2);
    }

    case "get_task": {
      const task = await gFetch(taskPath(args.listId as string, args.taskId as string));
      return JSON.stringify(task, null, 2);
    }

    case "create_task": {
      const task = await gFetch(`${listPath(args.listId as string)}/tasks`, "POST", buildTaskPayload(args));
      return JSON.stringify(task, null, 2);
    }

    case "update_task": {
      const task = await gFetch(taskPath(args.listId as string, args.taskId as string), "PATCH", buildTaskPayload(args));
      return JSON.stringify(task, null, 2);
    }

    case "complete_task": {
      const task = await gFetch(taskPath(args.listId as string, args.taskId as string), "PATCH", { status: "completed" });
      return JSON.stringify(task, null, 2);
    }

    case "delete_task": {
      await gFetch(taskPath(args.listId as string, args.taskId as string), "DELETE");
      return "Task deleted.";
    }

    case "create_todo_list": {
      const list = await gFetch("/me/todo/lists", "POST", { displayName: args.displayName });
      return JSON.stringify(list, null, 2);
    }

    default:
      throw new Error(`Unknown todo tool: ${name}`);
  }
}
