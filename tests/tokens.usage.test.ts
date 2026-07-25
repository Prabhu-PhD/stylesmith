import { describe, it, expect } from "vitest";
import type { Style, StyleSmithDocument, Token, ValueKind } from "../src/core/schema/types";
import { computeTokenUsage, directStyleUsageCount } from "../src/core/tokens/usage";

const ref = (tokenId: string): ValueKind => ({ kind: "tokenRef", tokenId });
const lit = (value: string | number): ValueKind => ({ kind: "literal", value });

function doc(tokens: Token[], styles: Style[]): StyleSmithDocument {
  return {
    schemaVersion: 1,
    tokenSets: [{ id: "set", name: "Default", origin: null, tokens }],
    styles,
  };
}

function style(id: string, layers: Partial<Style["layers"]>): Style {
  return {
    id,
    name: id,
    origin: null,
    basedOn: null,
    layers: { text: null, shape: null, geometry: null, table: null, ...layers },
  };
}

describe("computeTokenUsage", () => {
  it("finds a direct style reference and reports its property path", () => {
    const d = doc(
      [{ id: "X", name: "brand", type: "color", value: lit("#000") }],
      [style("s1", { shape: { fill: ref("X") } })],
    );
    const usage = computeTokenUsage(d, "X");
    expect(usage.directStyles).toEqual([{ styleId: "s1", properties: ["shape.fill"] }]);
    expect(usage.affectedStyleCount).toBe(1);
    expect(usage.aliasedBy).toEqual([]);
  });

  it("labels geometry adjustment references by index", () => {
    const d = doc(
      [{ id: "R", name: "radius-md", type: "radius", value: lit(0.15) }],
      [style("s1", { geometry: { adjustments: [lit(0.1), ref("R")] } })],
    );
    const usage = computeTokenUsage(d, "R");
    expect(usage.directStyles).toEqual([{ styleId: "s1", properties: ["geometry.adjustments[1]"] }]);
  });

  it("counts a style twice-referencing a token once, with both paths", () => {
    const d = doc(
      [{ id: "C", name: "ink", type: "color", value: lit("#111") }],
      [style("s1", { text: { color: ref("C") }, shape: { borderColor: ref("C") } })],
    );
    const usage = computeTokenUsage(d, "C");
    expect(usage.directStyles).toHaveLength(1);
    expect(usage.directStyles[0]?.properties.sort()).toEqual(["shape.borderColor", "text.color"]);
    expect(usage.affectedStyleCount).toBe(1);
  });

  it("includes transitively-affected styles through an alias chain", () => {
    // primitive P <- semantic S ; style bound to S. Editing P must affect the style.
    const tokens: Token[] = [
      { id: "P", name: "blue-600", type: "color", value: lit("#0057B8") },
      { id: "S", name: "surface", type: "color", value: ref("P") },
    ];
    const d = doc(tokens, [style("s1", { shape: { fill: ref("S") } })]);

    const usage = computeTokenUsage(d, "P");
    expect(usage.directStyles).toEqual([]); // no style references P directly
    expect(usage.aliasedBy).toEqual(["S"]); // S aliases P
    expect(usage.affectedStyleIds).toEqual(["s1"]); // ... so s1 is affected transitively
    expect(usage.affectedStyleCount).toBe(1);
  });

  it("directStyleUsageCount counts only direct references", () => {
    const tokens: Token[] = [
      { id: "P", name: "p", type: "color", value: lit("#000") },
      { id: "S", name: "s", type: "color", value: ref("P") },
    ];
    const d = doc(tokens, [style("s1", { shape: { fill: ref("S") } })]);
    expect(directStyleUsageCount(d, "P")).toBe(0);
    expect(directStyleUsageCount(d, "S")).toBe(1);
  });
});
