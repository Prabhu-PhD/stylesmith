import { describe, it, expect } from "vitest";
import type { Style, Token, ValueKind } from "../src/core/schema/types";
import type { ResolveContext } from "../src/core/tokens/resolve";
import {
  resolveStyleToTargets,
  valuesEqual,
  planApply,
  computeDrift,
  type ResolvedStyle,
  type Snapshot,
  type CurrentFormatting,
} from "../src/core/styles/diff";

const lit = (value: string | number): ValueKind => ({ kind: "literal", value });

function style(layers: Partial<Style["layers"]>): Style {
  return {
    id: "s1",
    name: "S",
    origin: null,
    basedOn: null,
    layers: { text: null, shape: null, geometry: null, table: null, ...layers },
  };
}

const emptyCtx: ResolveContext = { tokens: new Map<string, Token>() };

describe("resolveStyleToTargets", () => {
  it("resolves literals across layers with correct paths", () => {
    const s = style({
      text: { fontSize: lit(18) },
      shape: { fill: lit("#000000") },
      geometry: { adjustments: [lit(0.15)] },
    });
    const { targets, unresolved } = resolveStyleToTargets(s, emptyCtx);
    expect(unresolved).toEqual([]);
    expect(targets.map((t) => `${t.path}=${t.value}`).sort()).toEqual([
      "geometry.adjustments[0]=0.15",
      "shape.fill=#000000",
      "text.fontSize=18",
    ]);
  });

  it("reports an unresolvable themeColor (no resolver, no cache)", () => {
    const s = style({ shape: { fill: { kind: "themeColor", slot: "accent1" } } });
    const { targets, unresolved } = resolveStyleToTargets(s, emptyCtx);
    expect(targets).toEqual([]);
    expect(unresolved).toHaveLength(1);
    expect(unresolved[0]?.path).toBe("shape.fill");
    expect(unresolved[0]?.reason).toBe("theme-unresolved");
  });

  it("uses a themeColor's cached hex when the palette is unreadable (hex-now)", () => {
    const s = style({ shape: { fill: { kind: "themeColor", slot: "accent1", cached: "#0057B8" } } });
    const { targets } = resolveStyleToTargets(s, emptyCtx);
    expect(targets[0]?.value).toBe("#0057B8");
  });
});

describe("valuesEqual", () => {
  it("treats numbers within epsilon as equal", () => {
    expect(valuesEqual(0.15, 0.150001)).toBe(true);
    expect(valuesEqual(0.15, 0.16)).toBe(false);
  });
  it("normalises colour strings (case + leading #)", () => {
    expect(valuesEqual("#FFFFFF", "ffffff")).toBe(true);
    expect(valuesEqual("none", "None")).toBe(true);
    expect(valuesEqual("Arial", "arial ")).toBe(true);
  });
});

describe("planApply", () => {
  const resolved: ResolvedStyle = {
    targets: [
      { path: "shape.fill", layer: "shape", property: "fill", value: "#0057B8" },
      { path: "text.fontSize", layer: "text", property: "fontSize", value: 18 },
    ],
    unresolved: [],
  };

  it("applies everything when there is no snapshot", () => {
    const plan = planApply(resolved, null, {}, true);
    expect(plan.toApply).toHaveLength(2);
    expect(plan.preserved).toEqual([]);
    expect(plan.nextSnapshot).toEqual({ "shape.fill": "#0057B8", "text.fontSize": 18 });
  });

  it("preserves a property the user diverged from the snapshot", () => {
    const snapshot: Snapshot = { "shape.fill": "#0057B8", "text.fontSize": 18 };
    const current: CurrentFormatting = { "shape.fill": "#FF0000", "text.fontSize": 18 };
    const plan = planApply(resolved, snapshot, current, true);
    expect(plan.preserved).toEqual(["shape.fill"]);
    expect(plan.toApply.map((t) => t.path)).toEqual(["text.fontSize"]);
    // preserved keeps its OLD baseline so it stays preserved next time
    expect(plan.nextSnapshot["shape.fill"]).toBe("#0057B8");
  });

  it("ignores overrides when preserve is off", () => {
    const snapshot: Snapshot = { "shape.fill": "#0057B8" };
    const current: CurrentFormatting = { "shape.fill": "#FF0000" };
    const plan = planApply(resolved, snapshot, current, false);
    expect(plan.preserved).toEqual([]);
    expect(plan.toApply).toHaveLength(2);
  });

  it("re-applies a property the user left on-style", () => {
    const snapshot: Snapshot = { "shape.fill": "#0057B8" };
    const current: CurrentFormatting = { "shape.fill": "#0057B8" };
    const plan = planApply(resolved, snapshot, current, true);
    expect(plan.toApply.map((t) => t.path)).toContain("shape.fill");
  });
});

describe("computeDrift", () => {
  it("lists properties whose current value differs from the target", () => {
    const resolved: ResolvedStyle = {
      targets: [
        { path: "shape.fill", layer: "shape", property: "fill", value: "#0057B8" },
        { path: "text.fontSize", layer: "text", property: "fontSize", value: 18 },
      ],
      unresolved: [],
    };
    const current: CurrentFormatting = { "shape.fill": "#FF0000", "text.fontSize": 18 };
    expect(computeDrift(resolved, current)).toEqual(["shape.fill"]);
  });
});
