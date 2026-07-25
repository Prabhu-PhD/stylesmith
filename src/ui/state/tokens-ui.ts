/**
 * UI-only helpers for grouping and displaying tokens (UX §5.1). The four list
 * categories map onto the token types.
 */
import type { StyleSmithDocument, Token, TokenType, ValueKind } from "../../core/schema/types";
import { resolveLiteral, tokenFor } from "./resolve";

export type TokenCategory = "Colour" | "Type" | "Geometry" | "Stroke";

export const CATEGORY_ORDER: TokenCategory[] = ["Colour", "Type", "Geometry", "Stroke"];

export function categoryOf(type: TokenType): TokenCategory {
  if (type === "color") return "Colour";
  if (type === "radius") return "Geometry";
  if (type === "strokeWeight" || type === "dashStyle") return "Stroke";
  return "Type"; // fontFamily, fontSize, fontWeight, lineSpacing, spacing
}

/** A short label for how a token's value is expressed. */
export function tokenValueLabel(doc: StyleSmithDocument, value: ValueKind): string {
  switch (value.kind) {
    case "literal":
      return String(value.value);
    case "themeColor":
      return value.slot;
    case "tokenRef": {
      const target = tokenFor(doc, value);
      return `→ ${target?.name ?? "token"}`;
    }
  }
}

/** The resolved concrete value of a token, as text (or "—"). */
export function tokenResolvedText(doc: StyleSmithDocument, token: Token): string {
  const r = resolveLiteral(doc, token.value);
  return r === null ? "—" : String(r);
}
