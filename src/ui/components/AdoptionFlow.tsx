import { useEffect, useState } from "react";
import {
  makeStyles,
  tokens,
  shorthands,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Button,
  Input,
  Checkbox,
  Spinner,
  Body1,
  Caption1,
  Divider,
} from "@fluentui/react-components";
import { useStore } from "../state/store";

const useStyles = makeStyles({
  scanning: { display: "flex", justifyContent: "center", ...shorthands.padding(tokens.spacingVerticalXXL, "0") },
  list: { display: "flex", flexDirection: "column", ...shorthands.gap(tokens.spacingVerticalM), maxHeight: "50vh", overflowY: "auto" },
  card: {
    display: "flex", flexDirection: "column", ...shorthands.gap(tokens.spacingVerticalXS),
    ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalM),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorNeutralBackground2,
  },
  cardHead: { display: "flex", alignItems: "center", justifyContent: "space-between" },
  summary: { color: tokens.colorNeutralForeground3, fontFamily: tokens.fontFamilyMonospace, fontSize: tokens.fontSizeBase200 },
  nameRow: { display: "flex", alignItems: "center", ...shorthands.gap(tokens.spacingHorizontalS) },
  heading: { color: tokens.colorNeutralForeground3 },
  near: { display: "flex", flexDirection: "column", ...shorthands.gap(tokens.spacingVerticalXS) },
  nearRow: { display: "flex", alignItems: "center", justifyContent: "space-between", ...shorthands.gap(tokens.spacingHorizontalS) },
  nearInfo: { display: "flex", flexDirection: "column", minWidth: 0 },
  dev: { color: tokens.colorPaletteYellowForeground2, fontFamily: tokens.fontFamilyMonospace, fontSize: tokens.fontSizeBase200 },
  nearActions: { display: "flex", ...shorthands.gap(tokens.spacingHorizontalXS), flexShrink: 0 },
});

interface Draft { name: string; include: boolean }

export function AdoptionFlow(): JSX.Element {
  const s = useStyles();
  const open = useStore((st) => st.adoptionOpen);
  const scanning = useStore((st) => st.adoptionScanning);
  const scanned = useStore((st) => st.adoptionScannedCount);
  const clusters = useStore((st) => st.adoptionClusters);
  const nearMatches = useStore((st) => st.adoptionNearMatches);
  const cancelAdoption = useStore((st) => st.cancelAdoption);
  const adoptClusters = useStore((st) => st.adoptClusters);
  const linkNearMatch = useStore((st) => st.linkNearMatch);
  const skipNearMatch = useStore((st) => st.skipNearMatch);

  const [drafts, setDrafts] = useState<Record<string, Draft>>({});

  useEffect(() => {
    setDrafts((prev) => {
      const next: Record<string, Draft> = {};
      for (const c of clusters) next[c.id] = prev[c.id] ?? { name: "", include: true };
      return next;
    });
  }, [clusters]);

  if (!open) return <></>;

  const selected = clusters.filter((c) => drafts[c.id]?.include && (drafts[c.id]?.name.trim() ?? "") !== "");
  const includedButUnnamed = clusters.some((c) => drafts[c.id]?.include && (drafts[c.id]?.name.trim() ?? "") === "");

  return (
    <Dialog open modalType="modal" onOpenChange={(_, d) => { if (!d.open) cancelAdoption(); }}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>Adopt existing shapes</DialogTitle>
          <DialogContent>
            {scanning ? (
              <div className={s.scanning}><Spinner label="Scanning deck…" /></div>
            ) : (
              <div className={s.list}>
                {clusters.length === 0 && nearMatches.length === 0 && (
                  <Body1>No groups of matching shapes were found among {scanned} unlinked shapes.</Body1>
                )}

                {clusters.length > 0 && (
                  <>
                    <Caption1 className={s.heading}>{clusters.length} GROUPS OF MATCHING SHAPES</Caption1>
                    {clusters.map((c) => (
                      <div key={c.id} className={s.card}>
                        <div className={s.cardHead}>
                          <Body1>{c.size} shapes</Body1>
                          <Checkbox
                            checked={drafts[c.id]?.include ?? true}
                            onChange={(_, d) => setDrafts((p) => ({ ...p, [c.id]: { name: p[c.id]?.name ?? "", include: !!d.checked } }))}
                            label="Include"
                          />
                        </div>
                        <span className={s.summary}>{c.summary}</span>
                        <div className={s.nameRow}>
                          <Input
                            size="small"
                            placeholder="Name this style…"
                            value={drafts[c.id]?.name ?? ""}
                            onChange={(_, d) => setDrafts((p) => ({ ...p, [c.id]: { name: d.value, include: p[c.id]?.include ?? true } }))}
                          />
                        </div>
                      </div>
                    ))}
                  </>
                )}

                {nearMatches.length > 0 && (
                  <>
                    <Divider />
                    <Caption1 className={s.heading}>{nearMatches.length} NEARLY MATCH AN EXISTING STYLE</Caption1>
                    <div className={s.near}>
                      {nearMatches.map((m) => (
                        <div key={m.shapeId} className={s.nearRow}>
                          <div className={s.nearInfo}>
                            <Caption1>Slide {m.slideIndex + 1} → {m.styleName}</Caption1>
                            <span className={s.dev}>
                              {m.deviations.map((d) => `${d.path.split(".").pop()} ${d.shapeValue} (style: ${d.styleValue})`).join(", ")}
                            </span>
                          </div>
                          <div className={s.nearActions}>
                            <Button size="small" onClick={() => void linkNearMatch(m.shapeId, m.styleId, false)}>Link</Button>
                            <Button size="small" appearance="subtle" onClick={() => void linkNearMatch(m.shapeId, m.styleId, true)}>Fix</Button>
                            <Button size="small" appearance="subtle" onClick={() => skipNearMatch(m.shapeId)}>Skip</Button>
                          </div>
                        </div>
                      ))}
                    </div>
                  </>
                )}
              </div>
            )}
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={cancelAdoption}>Close</Button>
            <Button
              appearance="primary"
              disabled={scanning || selected.length === 0 || includedButUnnamed}
              onClick={() => void adoptClusters(selected.map((c) => ({ id: c.id, name: drafts[c.id]?.name ?? "" })))}
            >
              Create {selected.length} {selected.length === 1 ? "style" : "styles"}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
