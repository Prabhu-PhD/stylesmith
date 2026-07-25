import { describe, it, expect } from "vitest";
import { computeSignature, canonicalize } from "../src/core/fingerprint/signature";
import { clusterBySignature } from "../src/core/fingerprint/cluster";
import { matchFormatting } from "../src/core/fingerprint/match";

describe("computeSignature", () => {
  it("normalises colours (case, #) and quantises numbers", () => {
    expect(canonicalize("#0057B8")).toBe("0057b8");
    expect(canonicalize(0.16667)).toBe("0.1700");
    expect(canonicalize(0.15)).toBe("0.1500");
  });

  it("produces the same hash for equivalent formatting regardless of key order", () => {
    const a = computeSignature({ "shape.fill": "#0057B8", "text.fontSize": 14 });
    const b = computeSignature({ "text.fontSize": 14, "shape.fill": "#0057b8" });
    expect(a.hash).toBe(b.hash);
  });

  it("differs when a value differs beyond the quantum", () => {
    const a = computeSignature({ "geometry.adjustments[0]": 0.15 });
    const b = computeSignature({ "geometry.adjustments[0]": 0.18 });
    expect(a.hash).not.toBe(b.hash);
  });
});

describe("clusterBySignature", () => {
  it("groups identical formatting and sorts largest first, dropping singletons", () => {
    const navy = { "shape.fill": "#0A1F44", "shape.borderWeight": 1.5 };
    const white = { "shape.fill": "#FFFFFF" };
    const items = [
      { item: "a", formatting: navy },
      { item: "b", formatting: navy },
      { item: "c", formatting: navy },
      { item: "d", formatting: white },
      { item: "e", formatting: white },
      { item: "f", formatting: { "shape.fill": "#123456" } }, // singleton
    ];
    const clusters = clusterBySignature(items);
    expect(clusters).toHaveLength(2);
    expect(clusters[0]?.size).toBe(3);
    expect(clusters[0]?.members).toEqual(["a", "b", "c"]);
    expect(clusters[1]?.size).toBe(2);
  });

  it("respects a custom minSize", () => {
    const items = [{ item: 1, formatting: { "shape.fill": "#000" } }];
    expect(clusterBySignature(items, { minSize: 1 })).toHaveLength(1);
  });
});

describe("matchFormatting", () => {
  const style = { "shape.fill": "#0057B8", "geometry.adjustments[0]": 0.15, "text.fontSize": 14 };

  it("detects an exact match", () => {
    const r = matchFormatting({ ...style }, style);
    expect(r.exact).toBe(true);
    expect(r.near).toBe(false);
    expect(r.deviations).toEqual([]);
  });

  it("detects a near match and reports the specific deviation", () => {
    const shape = { "shape.fill": "#0057B8", "geometry.adjustments[0]": 0.18, "text.fontSize": 14 };
    const r = matchFormatting(shape, style);
    expect(r.exact).toBe(false);
    expect(r.near).toBe(true);
    expect(r.deviations).toEqual([{ path: "geometry.adjustments[0]", shapeValue: 0.18, styleValue: 0.15 }]);
  });

  it("treats a large colour difference as a non-match, not near", () => {
    const shape = { "shape.fill": "#FF0000", "geometry.adjustments[0]": 0.15, "text.fontSize": 14 };
    const r = matchFormatting(shape, style);
    expect(r.near).toBe(false);
    expect(r.deviations[0]?.path).toBe("shape.fill");
  });

  it("treats a close colour (one channel off) as near", () => {
    const shape = { "shape.fill": "#0057BC", "geometry.adjustments[0]": 0.15, "text.fontSize": 14 };
    const r = matchFormatting(shape, style);
    expect(r.near).toBe(true);
  });

  it("a missing property breaks near", () => {
    const shape = { "shape.fill": "#0057B8", "text.fontSize": 14 }; // no adjustments
    const r = matchFormatting(shape, style);
    expect(r.near).toBe(false);
  });
});
