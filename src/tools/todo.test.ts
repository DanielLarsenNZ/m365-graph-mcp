/**
 * Regression tests for Todo URL construction bugs.
 *
 * Bug 1 — Graph client library double-encodes path segments.
 *   Exchange-format (AQMk) list IDs contain `==` which is invalid raw in a
 *   URL path segment. The fix was to switch from the Graph client library to
 *   raw fetch with manual encodeURIComponent.
 *
 * Bug 2 — $select causes HTTP 400 ParseUri from Graph Todo API when combined
 *   with Exchange-format list IDs. Fix: removed $select entirely from
 *   list_tasks requests.
 *
 * These tests cover the pure URL-builder functions (no auth / network needed).
 */

import { test, describe } from "node:test";
import assert from "node:assert/strict";
import { listPath, taskPath, listTasksPath } from "./todo.js";

// ── Test IDs ──────────────────────────────────────────────────────────────────

// Exchange-format compound ID (well-known "defaultList"). Contains `==` at end.
const AQMK_LIST =
  "AQMkADU0NjVmYjU0LTc2YTQtNGFiYgAtOWZhZi1kYTY3MDVmZDdiMmEALgAAA_H_Bigm4yhKmxgpW3a7QnwBAEb0kVanRpNKsseJdkShLA4AAAIBEgAAAA==";

// Standard Graph Todo list ID. Contains `-` and trailing `=`.
const AAMK_LIST =
  "AAMkADU0NjVmYjU0LTc2YTQtNGFiYi05ZmFmLWRhNjcwNWZkN2IyYQAuAAAAAADh-gYoJuMoSpsYKVt2u0J8AQBG9JFWp0aTSrLHiXZEoSwOAAAAAz2zAAA=";

// Task ID containing trailing `=`.
const TASK_ID =
  "AAMkADU0NjVmYjU0LTc2YTQtNGFiYi05ZmFmLWRhNjcwNWZkN2IyYQBGAAAAAADh-gYoJuMoSpsYKVt2u0J8BwBG9JFWp0aTSrLHiXZEoSwOAAAAAAESAABG9JFWp0aTSrLHiXZEoSwOAAKqgKm0AAA=";

// ── Helpers ───────────────────────────────────────────────────────────────────

/** Extract the list ID segment from a todo list path. */
function extractListSegment(path: string) {
  const match = path.match(/\/me\/todo\/lists\/([^/]+)/);
  assert.ok(match, `Expected /me/todo/lists/{id} in: ${path}`);
  return match[1];
}

/** Extract the task ID segment from a todo task path. */
function extractTaskSegment(path: string) {
  const match = path.match(/\/tasks\/([^/?]+)/);
  assert.ok(match, `Expected /tasks/{id} in: ${path}`);
  return match[1];
}

// ── Regression: Bug 1 — AQMk IDs must be percent-encoded in path ─────────────

describe("Regression Bug 1: Exchange-format (AQMk) list ID encoding", () => {
  test("listPath: trailing == is encoded as %3D%3D", () => {
    const path = listPath(AQMK_LIST);
    const segment = extractListSegment(path);

    assert.ok(!segment.includes("=="), `Raw '==' must not appear in path segment — got: ${segment}`);
    assert.ok(segment.endsWith("%3D%3D"), `Trailing '==' must become '%3D%3D' — got: ${segment}`);
  });

  test("listPath: AAMk trailing = is encoded as %3D", () => {
    const path = listPath(AAMK_LIST);
    const segment = extractListSegment(path);

    assert.ok(!segment.endsWith("="), `Raw '=' must not be at end of path segment — got: ${segment}`);
    assert.ok(segment.endsWith("%3D"), `Trailing '=' must become '%3D' — got: ${segment}`);
  });

  test("listPath: encoded segment round-trips via decodeURIComponent", () => {
    const path = listPath(AQMK_LIST);
    const segment = extractListSegment(path);
    assert.equal(decodeURIComponent(segment), AQMK_LIST);
  });

  test("taskPath: both list ID and task ID are encoded", () => {
    const path = taskPath(AQMK_LIST, TASK_ID);
    const listSeg = extractListSegment(path);
    const taskSeg = extractTaskSegment(path);

    assert.equal(decodeURIComponent(listSeg), AQMK_LIST, "List ID must round-trip");
    assert.equal(decodeURIComponent(taskSeg), TASK_ID, "Task ID must round-trip");
    assert.ok(!listSeg.includes("=="), "Raw == must not appear in list segment");
    assert.ok(!taskSeg.includes("=") || taskSeg.includes("%3D"), "Raw = must not appear in task segment");
  });

  test("taskPath: AAMk IDs also round-trip correctly", () => {
    const path = taskPath(AAMK_LIST, TASK_ID);
    const listSeg = extractListSegment(path);
    const taskSeg = extractTaskSegment(path);

    assert.equal(decodeURIComponent(listSeg), AAMK_LIST);
    assert.equal(decodeURIComponent(taskSeg), TASK_ID);
  });
});

// ── Regression: Bug 2 — $select must not appear in list_tasks paths ───────────

describe("Regression Bug 2: $select must not appear in list_tasks paths", () => {
  test("listTasksPath: no $select in path with AQMk list ID", () => {
    const path = listTasksPath(AQMK_LIST, 50);
    assert.ok(!path.includes("$select"), `$select must not appear — got: ${path}`);
  });

  test("listTasksPath: no $select in path with AAMk list ID", () => {
    const path = listTasksPath(AAMK_LIST, 50);
    assert.ok(!path.includes("$select"), `$select must not appear — got: ${path}`);
  });

  test("listTasksPath: no $select even when filter is provided", () => {
    const path = listTasksPath(AQMK_LIST, 20, "status ne 'completed'");
    assert.ok(!path.includes("$select"), `$select must not appear with filter — got: ${path}`);
  });
});

// ── Correct behaviour ─────────────────────────────────────────────────────────

describe("Correct URL structure", () => {
  test("listPath: produces /me/todo/lists/{encodedId}", () => {
    const path = listPath(AAMK_LIST);
    assert.ok(path.startsWith("/me/todo/lists/"), `Expected /me/todo/lists/ prefix — got: ${path}`);
  });

  test("taskPath: produces /me/todo/lists/{listId}/tasks/{taskId}", () => {
    const path = taskPath(AAMK_LIST, TASK_ID);
    assert.match(path, /\/me\/todo\/lists\/.+\/tasks\/.+/);
  });

  test("listTasksPath: $top is included", () => {
    const path = listTasksPath(AAMK_LIST, 25);
    assert.ok(path.includes("$top=25"), `Expected $top=25 — got: ${path}`);
  });

  test("listTasksPath: filter is URL-encoded", () => {
    const path = listTasksPath(AAMK_LIST, 50, "status ne 'completed'");
    assert.ok(path.includes("$filter="), "Filter must be present");
    assert.ok(!path.includes(" "), "Spaces must be encoded (no raw spaces in URL)");
  });

  test("listTasksPath: no filter param when filter is undefined", () => {
    const path = listTasksPath(AAMK_LIST, 50);
    assert.ok(!path.includes("$filter"), "No $filter when not specified");
  });

  test("listTasksPath: list ID is encoded in the tasks path", () => {
    const path = listTasksPath(AQMK_LIST, 50);
    const listSeg = extractListSegment(path);
    assert.equal(decodeURIComponent(listSeg), AQMK_LIST, "List ID must round-trip in tasks path");
  });
});
