import { makeStyles, tokens, shorthands, Subtitle2, Body1, Button } from "@fluentui/react-components";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    textAlign: "center",
    ...shorthands.gap(tokens.spacingVerticalM),
    ...shorthands.padding(tokens.spacingVerticalXXL, tokens.spacingHorizontalL),
  },
  glyphs: { fontSize: "28px", letterSpacing: "4px", color: tokens.colorNeutralForeground3 },
  help: { color: tokens.colorNeutralForeground2, maxWidth: "240px" },
  actions: { display: "flex", flexDirection: "column", ...shorthands.gap(tokens.spacingVerticalS), width: "100%", maxWidth: "220px" },
});

/**
 * First-run empty state (UX §4.3) — the most important screen, because every
 * existing deck opens here. Scan is the primary onboarding path (adoption).
 */
export function EmptyState({
  shapeCount,
  onScan,
  onCreate,
  canCreate,
}: {
  shapeCount: number;
  onScan: () => void;
  onCreate: () => void;
  canCreate: boolean;
}): JSX.Element {
  const s = useStyles();
  return (
    <div className={s.root}>
      <div className={s.glyphs} aria-hidden>▢▢▢</div>
      <Subtitle2>No styles yet</Subtitle2>
      <Body1 className={s.help}>
        {shapeCount > 0
          ? `This deck has ${shapeCount} shapes. StyleSmith can find groups that already share formatting.`
          : "Create a style from a formatted shape, then apply it across the deck."}
      </Body1>
      <div className={s.actions}>
        <Button appearance="primary" onClick={onScan} disabled={shapeCount === 0}>
          Scan this deck
        </Button>
        <Button appearance="secondary" onClick={onCreate} disabled={!canCreate}>
          Create from selection
        </Button>
      </div>
    </div>
  );
}
