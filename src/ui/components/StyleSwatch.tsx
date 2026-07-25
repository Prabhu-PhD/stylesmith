import { tokens } from "@fluentui/react-components";
import type { Style, StyleSmithDocument, ValueKind } from "../../core/schema/types";
import { resolveLiteral } from "../state/resolve";

/**
 * A live mini-preview of a style's fill, border and corner radius. The radius is
 * deliberate — it advertises the geometry layer at a glance (UX spec §4.1).
 * Uses inline colours because it renders the style's DATA, not panel chrome.
 */
export function StyleSwatch({
  doc,
  style,
  size = 28,
}: {
  doc: StyleSmithDocument;
  style: Style;
  size?: number;
}): JSX.Element {
  const resolve = (v: ValueKind | undefined): string | number | null =>
    v ? resolveLiteral(doc, v) : null;

  const fill = resolve(style.layers.shape?.fill);
  const border = resolve(style.layers.shape?.borderColor);
  const adj0 = style.layers.geometry?.adjustments[0];
  const radiusFrac = adj0 ? Number(resolve(adj0) ?? 0) : 0;
  const radius = Math.max(0, Math.min(0.5, radiusFrac)) * size;

  const hasFill = typeof fill === "string" && fill.toLowerCase() !== "none";

  return (
    <div
      role="img"
      aria-label={`${style.name} preview`}
      style={{
        width: size,
        height: size,
        flexShrink: 0,
        borderRadius: radius,
        backgroundColor: hasFill ? String(fill) : tokens.colorNeutralBackground3,
        border:
          typeof border === "string"
            ? `1.5px solid ${border}`
            : `1px solid ${tokens.colorNeutralStroke2}`,
        boxSizing: "border-box",
      }}
    />
  );
}
