import { describe, expect, it } from "vitest";

import { stampToTags, tagsToRef, tagsToStamp, TAG_ID, TAG_VERSION, TAG_STAMP } from "@/lib/office/tags";
import type { ExportStamp } from "@/lib/types";

const stamp: ExportStamp = {
  chartId: "PLT-7Q2F",
  version: 3,
  timestamp: "2026-07-20T10:00:00.000Z",
  appVersion: "0.1.0",
};

describe("stampToTags", () => {
  it("writes id, version, and the full JSON stamp", () => {
    const tags = stampToTags(stamp);
    expect(tags[TAG_ID]).toBe("PLT-7Q2F");
    expect(tags[TAG_VERSION]).toBe("3");
    expect(JSON.parse(tags[TAG_STAMP])).toEqual(stamp);
  });
});

describe("tagsToRef", () => {
  it("round-trips a stamp through the tags", () => {
    expect(tagsToRef(stampToTags(stamp))).toEqual({ chartId: "PLT-7Q2F", version: 3 });
  });

  it("prefers the explicit id/version tags", () => {
    expect(tagsToRef({ [TAG_ID]: "PLT-ABCD", [TAG_VERSION]: "2" })).toEqual({
      chartId: "PLT-ABCD",
      version: 2,
    });
  });

  it("is case-insensitive on tag keys (PowerPoint upper-cases them)", () => {
    expect(tagsToRef({ plott_id: "PLT-ABCD", plott_version: "5" })).toEqual({
      chartId: "PLT-ABCD",
      version: 5,
    });
  });

  it("falls back to the JSON stamp when id/version tags are absent", () => {
    expect(tagsToRef({ [TAG_STAMP]: JSON.stringify(stamp) })).toEqual({
      chartId: "PLT-7Q2F",
      version: 3,
    });
  });

  it("returns an id with undefined version when the version tag is missing/invalid", () => {
    expect(tagsToRef({ [TAG_ID]: "PLT-ABCD" })).toEqual({ chartId: "PLT-ABCD", version: undefined });
    expect(tagsToRef({ [TAG_ID]: "PLT-ABCD", [TAG_VERSION]: "x" })).toEqual({
      chartId: "PLT-ABCD",
      version: undefined,
    });
  });

  it("returns null for tags with no Plott identity", () => {
    expect(tagsToRef({})).toBeNull();
    expect(tagsToRef({ SOMETHING: "else" })).toBeNull();
  });
});

describe("tagsToStamp", () => {
  it("returns the full stamp from the JSON tag", () => {
    expect(tagsToStamp(stampToTags(stamp))).toEqual(stamp);
  });

  it("returns null for a malformed stamp", () => {
    expect(tagsToStamp({ [TAG_STAMP]: "not json" })).toBeNull();
    expect(tagsToStamp({})).toBeNull();
  });
});
