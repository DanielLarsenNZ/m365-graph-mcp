# Tool Reference

Full parameter reference for all 24 MCP tools exposed by the server.

---

## Auth

### `auth_status`
Check whether the server currently holds a valid cached token.

_No parameters._

Returns `"Authenticated."` or `"Not authenticated. Call auth_login to sign in."`

---

### `auth_login`
Authenticate via Microsoft device code flow.

_No parameters._

Returns a prompt to visit a URL and enter a code. If a cached token already exists and can be refreshed silently, returns a confirmation that the cached token was reused.

---

### `auth_logout`
Sign out and clear the cached token.

_No parameters._

---

## Mail

### `list_emails`
List emails from a mail folder.

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `folder` | string | `"inbox"` | Folder name or well-known ID |
| `top` | number | `20` | Max emails to return (max 50) |
| `filter` | string | — | OData filter expression, e.g. `"isRead eq false"` |
| `search` | string | — | Free-text search query |
| `inferenceClassification` | `"focused"` \| `"other"` | — | Scope to Focused or Other inbox |

`filter` and `inferenceClassification` are automatically merged with `and` when both are provided.

Fields returned per message: `id`, `subject`, `from`, `receivedDateTime`, `isRead`, `bodyPreview`, `hasAttachments`, `inferenceClassification`.

---

### `get_email`
Retrieve a single email with its full body.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `id` | string | yes | Message ID |

Fields returned: `id`, `subject`, `from`, `toRecipients`, `ccRecipients`, `receivedDateTime`, `isRead`, `body`, `hasAttachments`.

---

### `send_email`
Send a new email.

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `to` | string | yes | — | Recipient address(es), comma-separated |
| `subject` | string | yes | — | Email subject |
| `body` | string | yes | — | Email body |
| `cc` | string | no | — | CC address(es), comma-separated |
| `bodyType` | `"html"` \| `"text"` | no | `"html"` | Body content type |
| `from` | string | no | — | Sender override (must be an alias/shared mailbox on the account) |
| `saveToSentItems` | boolean | no | `true` | Whether to save to Sent Items |

---

### `reply_email`
Reply to an email.

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `id` | string | yes | — | Message ID to reply to |
| `body` | string | yes | — | Reply body |
| `bodyType` | `"html"` \| `"text"` | no | `"html"` | Body content type |
| `from` | string | no | — | Sender address override |
| `replyAll` | boolean | no | `false` | Reply to all recipients |

---

### `move_email`
Move an email to a different folder.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `id` | string | yes | Message ID |
| `destinationFolder` | string | yes | Target folder ID or well-known name: `inbox`, `deleteditems`, `drafts`, `sentitems`, `junkemail` |

---

### `delete_email`
Delete (trash) an email.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `id` | string | yes | Message ID |

---

### `mark_email_read`
Mark an email as read or unread.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `id` | string | yes | Message ID |
| `isRead` | boolean | yes | `true` = read, `false` = unread |

---

## Calendar

### `list_events`
List calendar events within a date range.

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `startDateTime` | string | yes | — | ISO 8601 start datetime, e.g. `"2026-04-01T00:00:00"` |
| `endDateTime` | string | yes | — | ISO 8601 end datetime |
| `top` | number | no | `20` | Max events to return |
| `calendarId` | string | no | — | Specific calendar ID; omit for primary calendar |

Fields returned per event: `id`, `subject`, `start`, `end`, `location`, `organizer`, `attendees`, `isOnlineMeeting`, `onlineMeetingUrl`, `bodyPreview`.

---

### `get_event`
Get full details of a calendar event.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `id` | string | yes | Event ID |

Fields returned: `id`, `subject`, `start`, `end`, `location`, `organizer`, `attendees`, `body`, `isOnlineMeeting`, `onlineMeetingUrl`.

---

### `create_event`
Create a new calendar event.

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `subject` | string | yes | — | Event title |
| `startDateTime` | string | yes | — | ISO 8601 start time |
| `endDateTime` | string | yes | — | ISO 8601 end time |
| `timeZone` | string | no | `"Pacific/Auckland"` | IANA timezone |
| `body` | string | no | — | Event description (HTML) |
| `location` | string | no | — | Location display name |
| `attendees` | string | no | — | Attendee email(s), comma-separated |
| `isOnlineMeeting` | boolean | no | `false` | Create as Teams meeting |
| `calendarId` | string | no | — | Target calendar; omit for primary |

---

### `update_event`
Update an existing calendar event. Only fields provided are changed.

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `id` | string | yes | — | Event ID |
| `subject` | string | no | — | New title |
| `startDateTime` | string | no | — | New start time (ISO 8601) |
| `endDateTime` | string | no | — | New end time (ISO 8601) |
| `timeZone` | string | no | `"Pacific/Auckland"` | Timezone for datetime updates |
| `body` | string | no | — | New description |
| `location` | string | no | — | New location |

---

### `delete_event`
Delete a calendar event.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `id` | string | yes | Event ID |

---

### `list_calendars`
List all calendars for the signed-in user.

_No parameters._

Fields returned per calendar: `id`, `name`, `isDefaultCalendar`, `canEdit`.

---

## Todo

> **Note on IDs:** Microsoft Todo list and task IDs in Exchange-format accounts (AQMk…, AAMk…) contain `=` and `==` characters. The server handles URL-encoding automatically — pass raw IDs as returned by `list_todo_lists` and `list_tasks`.

### `list_todo_lists`
List all Todo task lists.

_No parameters._

---

### `list_tasks`
List tasks in a Todo list.

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `listId` | string | yes | — | Task list ID |
| `filter` | string | no | — | OData filter, e.g. `"status ne 'completed'"` |
| `top` | number | no | `50` | Max tasks to return |

---

### `get_task`
Get a specific task.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `listId` | string | yes | Task list ID |
| `taskId` | string | yes | Task ID |

---

### `create_task`
Create a new task.

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `listId` | string | yes | — | Task list ID |
| `title` | string | yes | — | Task title |
| `body` | string | no | — | Task notes |
| `dueDateTime` | string | no | — | Due date, ISO 8601 |
| `importance` | `"low"` \| `"normal"` \| `"high"` | no | `"normal"` | Priority |
| `reminderDateTime` | string | no | — | Reminder datetime, ISO 8601 |

Datetime fields are stored with timezone `Pacific/Auckland` unless the server default is overridden.

---

### `update_task`
Update an existing task. Only fields provided are changed.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `listId` | string | yes | Task list ID |
| `taskId` | string | yes | Task ID |
| `title` | string | no | New title |
| `body` | string | no | New notes |
| `dueDateTime` | string | no | New due date (ISO 8601) |
| `importance` | `"low"` \| `"normal"` \| `"high"` | no | New priority |
| `status` | `"notStarted"` \| `"inProgress"` \| `"completed"` \| `"waitingOnOthers"` \| `"deferred"` | no | New status |

---

### `complete_task`
Mark a task as completed (convenience wrapper around `update_task`).

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `listId` | string | yes | Task list ID |
| `taskId` | string | yes | Task ID |

---

### `delete_task`
Delete a task.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `listId` | string | yes | Task list ID |
| `taskId` | string | yes | Task ID |

---

### `create_todo_list`
Create a new Todo task list.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `displayName` | string | yes | Name for the new list |
