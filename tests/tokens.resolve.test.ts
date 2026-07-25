import { describe, it, expect } from "vitest";
import type { Token, ValueKind } from "../src/core/schema/types";
import { resolveValue, type ResolveContext } from "../src/core/tokens/resolve";

/** Build a token map from tokens. */
function tokenMap(...tokens: Token[]): Map<string, Token> {
  return new Map(tokens.map((t) => [t.id, t]));
}

const lit = (value: string | number): ValueKind => ({ kind: "literal", value });
const ref = (tokenId: string, cached?: string | number): ValueKind =>
  cached === undefined ? { kind: "tokenRef", tokenId } : { kind: "tokenRef", tokenId, cached };

const ctxOf = (tokens: Map<string, Token>, extra: Partial<ResolveContext> = {}): ResolveContext => ({
  tokens,
  ...extra,
});

describe("resolveValue", () => {
  it("resolves a literal directly", () => {
    const r = resolveValue(lit("#FFFFFF"), ctxOf(tokenMap()));
    expect(r).toEqual({ ok: true, value: "#FFFFFF", provenance: "literal", degraded: false });
  });

  it("resolves a tokenRef to its literal value (token-sourced)", () => {
    const t: Token = { id: "t1", name: "brand", type: "color", value: lit("#0057B8") };
    const r = resolveValue(ref("t1"), ctxOf(tokenMap(t)));
    expect(r.ok).toBe(true);
    if (r.ok) {
      expect(r.value).toBe("#0057B8");
      expect(r.provenance).toBe("token");
      expect(r.degraded).toBe(false);
    }
  });

  it("follows a deep alias chain a -> b -> c -> literal", () => {
    const a: Token = { id: "a", name: "a", type: "color", value: ref("b") };
    const b: Token = { id: "b", name: "b", type: "color", value: ref("c") };
    const c: Token = { id: "c", name: "c", type: "color", value: lit("#123456") };
    const r = resolveValue(ref("a"), ctxOf(tokenMap(a, b, c)));
    expect(r.ok && r.value).toBe("#123456");
    expect(r.ok && r.provenance).toBe("token");
  });

  it("falls back to the cached literal when the token is missing (degraded)", () => {
    const r = resolveValue(ref("gone", "#CACHED0"), ctxOf(tokenMap()));
    expect(r).toEqual({
      ok: true,
      value: "#CACHED0",
      provenance: "cache",
      degraded: true,
      reason: "missing-token",
    });
  });

  it("fails when the token is missing and no cache exists", () => {
    const r = resolveValue(ref("gone"), ctxOf(tokenMap()));
    expect(r).toEqual({ ok: false, reason: "missing-token", tokenId: "gone" });
  });

  it("detects a cycle at resolve time (no cache -> failure)", () => {
    const a: Token = { id: "a", name: "a", type: "color", value: ref("b") };
    const b: Token = { id: "b", name: "b", type: "color", value: ref("a") };
    const r = resolveValue(ref("a"), ctxOf(tokenMap(a, b)));
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.reason).toBe("cycle");
  });

  it("a cycle with a cached fallback degrades rather than breaking", () => {
    const a: Token = { id: "a", name: "a", type: "color", value: ref("b") };
    const b: Token = { id: "b", name: "b", type: "color", value: ref("a") };
    const r = resolveValue(ref("a", "#SAFE00"), ctxOf(tokenMap(a, b)));
    expect(r.ok && r.value).toBe("#SAFE00");
    expect(r.ok && r.provenance).toBe("cache");
    expect(r.ok && r.reason).toBe("cycle");
  });

  it("enforces the depth limit", () => {
    const chain = tokenMap(
      { id: "a", name: "a", type: "color", value: ref("b") },
      { id: "b", name: "b", type: "color", value: ref("c") },
      { id: "c", name: "c", type: "color", value: ref("d") },
      { id: "d", name: "d", type: "color", value: lit("#DEEP00") },
    );
    const r = resolveValue(ref("a"), ctxOf(chain, { maxDepth: 2 }));
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.reason).toBe("depth-exceeded");
  });

  it("resolves a themeColor via the injected theme resolver", () => {
    const theme = (slot: string): string | undefined => (slot === "accent1" ? "#0057B8" : undefined);
    const r = resolveValue({ kind: "themeColor", slot: "accent1" }, ctxOf(tokenMap(), { theme }));
    expect(r).toEqual({ ok: true, value: "#0057B8", provenance: "themeColor", degraded: false });
  });

  it("a tokenRef terminating at a themeColor reports themeColor provenance", () => {
    const theme = (slot: string): string | undefined => (slot === "accent2" ? "#00A" : undefined);
    const t: Token = { id: "t", name: "semantic", type: "color", value: { kind: "themeColor", slot: "accent2" } };
    const r = resolveValue(ref("t"), ctxOf(tokenMap(t), { theme }));
    expect(r.ok && r.value).toBe("#00A");
    expect(r.ok && r.provenance).toBe("themeColor");
  });

  it("an unresolved themeColor fails, or degrades to cache when present", () => {
    const noCache = resolveValue({ kind: "themeColor", slot: "accent1" }, ctxOf(tokenMap()));
    expect(noCache.ok).toBe(false);
    if (!noCache.ok) expect(noCache.reason).toBe("theme-unresolved");

    const cached = resolveValue(
      { kind: "themeColor", slot: "accent1", cached: "#PREV00" },
      ctxOf(tokenMap()),
    );
    expect(cached.ok && cached.value).toBe("#PREV00");
    expect(cached.ok && cached.provenance).toBe("cache");
    expect(cached.ok && cached.reason).toBe("theme-unresolved");
  });
});
