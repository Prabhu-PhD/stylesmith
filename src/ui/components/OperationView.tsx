import {
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Button,
  ProgressBar,
  Body1,
  Caption1,
  makeStyles,
  tokens,
  shorthands,
} from "@fluentui/react-components";
import { useStore } from "../state/store";

const useStyles = makeStyles({
  content: { display: "flex", flexDirection: "column", ...shorthands.gap(tokens.spacingVerticalS), minWidth: "260px" },
  undo: { color: tokens.colorNeutralForeground3 },
});

/**
 * Progress + completion for a sweep (UX §6.1–6.2). Chunked execution is
 * architectural, so progress is a first-class state, not a spinner. Cancel stops
 * at a chunk boundary and the completion state reports what was done.
 */
export function OperationView(): JSX.Element {
  const s = useStyles();
  const op = useStore((st) => st.operation);
  const dismiss = useStore((st) => st.dismissOperation);
  if (!op) return <></>;

  const running = op.phase === "running";
  const fraction = op.total > 0 ? op.done / op.total : undefined;

  return (
    <Dialog
      open
      modalType={running ? "alert" : "modal"}
      onOpenChange={(_, d) => { if (!d.open && !running) dismiss(); }}
    >
      <DialogSurface>
        <DialogBody>
          <DialogTitle>{running ? `Applying "${op.styleName}"…` : "✓ Done"}</DialogTitle>
          <DialogContent>
            <div className={s.content}>
              {running ? (
                <>
                  <ProgressBar {...(fraction !== undefined ? { value: fraction } : {})} max={1} thickness="large" />
                  <Body1>{op.done} / {op.total || "…"}</Body1>
                </>
              ) : (
                <>
                  <Body1>
                    Updated {op.result?.applied ?? 0} {(op.result?.applied ?? 0) === 1 ? "shape" : "shapes"}.
                  </Body1>
                  {op.result && op.result.skippedGeometryShapes > 0 && (
                    <Caption1>Geometry skipped on {op.result.skippedGeometryShapes} shapes (no adjustable handles).</Caption1>
                  )}
                  {op.result?.cancelled && <Caption1>Cancelled — this was a partial apply.</Caption1>}
                  <Caption1 className={s.undo}>Ctrl+Z to undo</Caption1>
                </>
              )}
            </div>
          </DialogContent>
          <DialogActions>
            {running ? (
              <Button onClick={op.cancel}>Cancel</Button>
            ) : (
              <Button appearance="primary" onClick={dismiss}>Close</Button>
            )}
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
