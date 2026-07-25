import { makeStyles, tokens, shorthands, Body1, Caption1, mergeClasses } from "@fluentui/react-components";
import type { Token, StyleSmithDocument } from "../../core/schema/types";
import { resolveLiteral } from "../state/resolve";
import { categoryOf, tokenValueLabel } from "../state/tokens-ui";

const useStyles = makeStyles({
  row: {
    display: "flex",
    alignItems: "center",
    ...shorthands.gap(tokens.spacingHorizontalM),
    minHeight: "36px",
    ...shorthands.padding(tokens.spacingVerticalXS, tokens.spacingHorizontalS),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorNeutralBackground2,
    cursor: "pointer",
    ...shorthands.border("1px", "solid", "transparent"),
    ":hover": { backgroundColor: tokens.colorNeutralBackground2Hover },
    ":focus-visible": { ...shorthands.outline("2px", "solid", tokens.colorStrokeFocus2) },
  },
  glyph: {
    width: "18px",
    textAlign: "center",
    flexShrink: 0,
    fontFamily: tokens.fontFamilyMonospace,
    color: tokens.colorNeutralForeground3,
  },
  swatch: { width: "16px", height: "16px", flexShrink: 0, ...shorthands.borderRadius(tokens.borderRadiusSmall), ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2) },
  body: { flexGrow: 1, minWidth: 0, display: "flex", flexDirection: "column" },
  name: { whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" },
  value: { display: "flex", alignItems: "center", ...shorthands.gap("4px"), color: tokens.colorNeutralForeground3, fontFamily: tokens.fontFamilyMonospace },
  count: { color: tokens.colorNeutralForeground3, flexShrink: 0 },
});

const GLYPH: Record<string, string> = { Type: "Aa", Geometry: "⌒", Stroke: "─" };

export function TokenRow({
  doc,
  token,
  styleCount,
  onOpen,
}: {
  doc: StyleSmithDocument;
  token: Token;
  styleCount: number;
  onOpen: () => void;
}): JSX.Element {
  const s = useStyles();
  const category = categoryOf(token.type);
  const resolved = resolveLiteral(doc, token.value);
  const isTheme = token.value.kind === "themeColor";

  return (
    <div
      role="button"
      tabIndex={0}
      className={s.row}
      onClick={onOpen}
      onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); onOpen(); } }}
      aria-label={`${token.name} token, ${styleCount} ${styleCount === 1 ? "style" : "styles"}`}
    >
      {category === "Colour" ? (
        <span
          className={s.swatch}
          style={{ backgroundColor: typeof resolved === "string" ? resolved : tokens.colorNeutralBackground3 }}
          aria-hidden
        />
      ) : (
        <span className={s.glyph} aria-hidden>{GLYPH[category] ?? "•"}</span>
      )}
      <div className={s.body}>
        <Body1 className={s.name}>{token.name}</Body1>
        <div className={mergeClasses(s.value)}>
          {tokenValueLabel(doc, token.value)}
          {isTheme && <span aria-label="deck theme colour">🎨</span>}
        </div>
      </div>
      <Caption1 className={s.count}>{styleCount} {styleCount === 1 ? "style" : "styles"}</Caption1>
    </div>
  );
}
