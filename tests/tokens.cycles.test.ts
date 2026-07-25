import { describe, it, expect } from "vitest";
import type { Token, ValueKind } from "../src/core/schema/types";
import { checkTokenValueWritable, findCycles } from "../src/core/tokens/cycles";

function tokenMap(...tokens: Token[]): Map<string, Token> {
  return new Map(tokens.map((t) => [t.id, t]));
}
const ref = (tokenId: string): ValueKind => ({ kind: "tokenRef", tokenId });
const lit = (value: string | number): ValueKind => ({ kind: "literal", value });

describe("checkTokenValueWritable (write-time cycle guard)", () => {
  it("accepts a literal", () => {
    expect(checkTokenValueWritable(tokenMap(), "a", lit("#fff"))).toEqual({ ok: true });
  });

  it("rejects a direct self-reference", () => {
    expect(checkTokenValueWritable(tokenMap(), "a", ref("a"))).toEqual({ ok: false, reason: "cycle" });
  });

  it("rejects a two-node cycle (a -> b already exists, set b -> a)", () => {
    // Existing: a -> b. Now attempt to set b's value to ref(a).
    const a: Token = { id: "a", name: "a", type: "color", value: ref("b") };
    const b: Token = { id: "b", name: "b", type: "color", value: lit("#000") };
    const result = checkTokenValueWritable(tokenMap(a, b), "b", ref("a"));
    expect(result).toEqual({ ok: false, reason: "cycle" });
  });

  it("allows a valid new alias (b -> a where a is a literal)", () => {
    const a: Token = { id: "a", name: "a", type: "color", value: lit("#000") };
    const b: Token = { id: "b", name: "b", type: "color", value: lit("#111") };
    expect(checkTokenValueWritable(tokenMap(a, b), "b", ref("a"))).toEqual({ ok: true });
  });

  it("treats a dangling reference as writable (resolution handles it via cache)", () => {
    expect(checkTokenValueWritable(tokenMap(), "a", ref("does-not-exist"))).toEqual({ ok: true });
  });

  it("rejects an over-deep chain", () => {
    const chain = tokenMap(
      { id: "b", name: "b", type: "color", value: ref("c") },
      { id: "c", name: "c", type: "color", value: ref("d") },
      { id: "d", name: "d", type: "color", value: lit("#000") },
    );
    // Setting a -> b, with maxDepth 1, exceeds the limit before terminating.
    expect(checkTokenValueWritable(chain, "a", ref("b"), 1)).toEqual({
      ok: false,
      reason: "depth-exceeded",
    });
  });
});

describe("findCycles (whole-set validation)", () => {
  it("returns [] for an acyclic set", () => {
    const tokens = tokenMap(
      { id: "a", name: "a", type: "color", value: ref("b") },
      { id: "b", name: "b", type: "color", value: lit("#000") },
    );
    expect(findCycles(tokens)).toEqual([]);
  });

  it("finds a two-node cycle exactly once", () => {
    const tokens = tokenMap(
      { id: "a", name: "a", type: "color", value: ref("b") },
      { id: "b", name: "b", type: "color", value: ref("a") },
    );
    const cycles = findCycles(tokens);
    expect(cycles).toHaveLength(1);
    expect([...(cycles[0] ?? [])].sort()).toEqual(["a", "b"]);
  });

  it("finds a self-cycle", () => {
    const tokens = tokenMap({ id: "a", name: "a", type: "color", value: ref("a") });
    expect(findCycles(tokens)).toEqual([["a"]]);
  });
});
