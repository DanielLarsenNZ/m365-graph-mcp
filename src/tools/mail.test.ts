/**
 * Unit tests for mail tool helpers.
 *
 * These tests cover pure functions only — no auth or network required.
 */

import { test, describe } from "node:test";
import assert from "node:assert/strict";
import { buildMailFilter } from "./mail.js";

// ── buildMailFilter ───────────────────────────────────────────────────────────

describe("buildMailFilter", () => {
  // No args → no filter
  test("returns undefined when both args are undefined", () => {
    assert.equal(buildMailFilter(undefined, undefined), undefined);
  });

  // Caller filter only
  test("returns caller filter unchanged when no inferenceClassification", () => {
    assert.equal(buildMailFilter("isRead eq false", undefined), "isRead eq false");
  });

  // inferenceClassification only — focused
  test("returns focused clause when only inferenceClassification is 'focused'", () => {
    assert.equal(
      buildMailFilter(undefined, "focused"),
      "inferenceClassification eq 'focused'",
    );
  });

  // inferenceClassification only — other
  test("returns other clause when only inferenceClassification is 'other'", () => {
    assert.equal(
      buildMailFilter(undefined, "other"),
      "inferenceClassification eq 'other'",
    );
  });

  // Both — must be joined with " and "
  test("joins caller filter and inferenceClassification with ' and '", () => {
    const result = buildMailFilter("isRead eq false", "other");
    assert.equal(result, "isRead eq false and inferenceClassification eq 'other'");
  });

  test("joins caller filter and 'focused' with ' and '", () => {
    const result = buildMailFilter("hasAttachments eq true", "focused");
    assert.equal(result, "hasAttachments eq true and inferenceClassification eq 'focused'");
  });

  // Empty string filter — treated as falsy, not included
  test("ignores empty string caller filter", () => {
    assert.equal(
      buildMailFilter("", "other"),
      "inferenceClassification eq 'other'",
    );
  });

  test("returns undefined for empty string filter and no inferenceClassification", () => {
    assert.equal(buildMailFilter("", undefined), undefined);
  });
});
