import { makeStyles, tokens, shorthands, Tooltip } from "@fluentui/react-components";
import type { Style } from "../../core/schema/types";

const useStyles = makeStyles({
  row: {
    display: "flex",
    ...shorthands.gap("6px"),
    alignItems: "center",
  },
  glyph: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    lineHeight: "1",
  },
});

const LAYERS: { key: keyof Style["layers"]; glyph: string; label: string }[] = [
  { key: "text", glyph: "Aa", label: "Text" },
  { key: "shape", glyph: "▢", label: "Shape" },
  { key: "geometry", glyph: "⌒", label: "Geometry" },
  { key: "table", glyph: "▦", label: "Table" },
];

/** Compact glyphs showing which of the four layers a style defines (UX §4.1). */
export function LayerSummary({ style }: { style: Style }): JSX.Element {
  const s = useStyles();
  const present = LAYERS.filter((l) => style.layers[l.key] != null);
  return (
    <div className={s.row}>
      {present.map((l) => (
        <Tooltip key={l.key} content={`${l.label} layer defined`} relationship="label">
          <span className={s.glyph} aria-label={`${l.label} layer`}>
            {l.glyph}
          </span>
        </Tooltip>
      ))}
    </div>
  );
}
