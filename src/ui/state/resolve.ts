/**
 * UI-side value resolution for display (swatches, binding chips). Wraps core's
 * pure resolveValue with the document's tokens and the (best-effort) deck theme
 * resolver. UI-only convenience — no mutation.
 */
import { indexTokens, resolveValue, type ResolveContext } from "../../core/tokens/resolve";
import { createDeckThemeResolver } from "../../office/theme";
import type { LiteralPrimitive, StyleSmithDocument, Token, ValueKind } from "../../core/schema/types";

export function makeResolveContext(doc: StyleSmithDocument): ResolveContext {
  return { tokens: indexTokens(doc), theme: createDeckThemeResolver() };
}

/** Resolve a value to its concrete literal, or null if unresolvable. */
export function resolveLiteral(doc: StyleSmithDocument, value: ValueKind): LiteralPrimitive | null {
  const r = resolveValue(value, makeResolveContext(doc));
  return r.ok ? r.value : null;
}

/** The token referenced by a value, if it is a tokenRef. */
export function tokenFor(doc: StyleSmithDocument, value: ValueKind): Token | null {
  if (value.kind !== "tokenRef") return null;
  for (const set of doc.tokenSets) {
    const t = set.tokens.find((tok) => tok.id === value.tokenId);
    if (t) return t;
  }
  return null;
}
