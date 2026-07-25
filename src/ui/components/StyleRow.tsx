import { makeStyles, tokens, shorthands, Body1, Caption1, mergeClasses } from "@fluentui/react-components";
import type { Style, StyleSmithDocument } from "../../core/schema/types";
import { StyleSwatch } from "./StyleSwatch";
import { LayerSummary } from "./LayerSummary";

const useStyles = makeStyles({
  row: {
    display: "flex",
    alignItems: "center",
    ...shorthands.gap(tokens.spacingHorizontalM),
    minHeight: "40px",
    ...shorthands.padding(tokens.spacingVerticalXS, tokens.spacingHorizontalS),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorNeutralBackground2,
    cursor: "pointer",
    ...shorthands.border("1px", "solid", "transparent"),
  },
  rowHover: {
    ":hover": { backgroundColor: tokens.colorNeutralBackground2Hover },
    ":focus-visible": {
      ...shorthands.outline("2px", "solid", tokens.colorStrokeFocus2),
    },
  },
  body: { flexGrow: 1, minWidth: 0, display: "flex", flexDirection: "column" },
  name: {
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  meta: { display: "flex", alignItems: "center", ...shorthands.gap(tokens.spacingHorizontalS) },
  count: { color: tokens.colorNeutralForeground3 },
  chevron: { color: tokens.colorNeutralForeground3, flexShrink: 0 },
});

export function StyleRow({
  doc,
  style,
  usage,
  onOpen,
}: {
  doc: StyleSmithDocument;
  style: Style;
  usage: number;
  onOpen: () => void;
}): JSX.Element {
  const s = useStyles();
  return (
    <div
      role="button"
      tabIndex={0}
      className={mergeClasses(s.row, s.rowHover)}
      onClick={onOpen}
      onKeyDown={(e) => {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          onOpen();
        }
      }}
      aria-label={`${style.name} style, ${usage} ${usage === 1 ? "shape" : "shapes"}`}
    >
      <StyleSwatch doc={doc} style={style} />
      <div className={s.body}>
        <Body1 className={s.name}>{style.name}</Body1>
        <div className={s.meta}>
          <Caption1 className={s.count}>{usage} {usage === 1 ? "shape" : "shapes"}</Caption1>
          <LayerSummary style={style} />
        </div>
      </div>
      <span className={s.chevron} aria-hidden>›</span>
    </div>
  );
}
