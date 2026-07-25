import { useEffect, useState, lazy, Suspense } from "react";
import {
  FluentProvider,
  makeStyles,
  tokens,
  shorthands,
  Title3,
  Badge,
  TabList,
  Tab,
  Input,
  Button,
  Caption1,
  Body1,
  Spinner,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Field,
} from "@fluentui/react-components";
import { useOfficeTheme } from "./theme/useOfficeTheme";
import { useStore, unlinkedShapeCount, type View } from "./state/store";
import { useKeyboardShortcuts } from "./useKeyboardShortcuts";
import { ErrorBoundary } from "./components/ErrorBoundary";
import { StylesView } from "./views/StylesView";
import { StyleDetail } from "./views/StyleDetail";
import { ApplyModal } from "./components/ApplyModal";
import { OperationView } from "./components/OperationView";

// Deferred surfaces — not needed for the first paint of the Styles view.
const TokensView = lazy(() => import("./views/TokensView").then((m) => ({ default: m.TokensView })));
const TokenDetail = lazy(() => import("./views/TokenDetail").then((m) => ({ default: m.TokenDetail })));
const AdoptionFlow = lazy(() => import("./components/AdoptionFlow").then((m) => ({ default: m.AdoptionFlow })));
const SpikePanel = lazy(() => import("./components/SpikePanel").then((m) => ({ default: m.SpikePanel })));

const useStyles = makeStyles({
  provider: { minHeight: "100vh", backgroundColor: tokens.colorNeutralBackground1 },
  root: { maxWidth: "320px", display: "flex", flexDirection: "column", minHeight: "100vh" },
  header: { display: "flex", alignItems: "center", ...shorthands.gap(tokens.spacingHorizontalS), ...shorthands.padding(tokens.spacingVerticalM, tokens.spacingHorizontalM, tokens.spacingVerticalXS) },
  title: { flexGrow: 1 },
  nav: { ...shorthands.padding("0", tokens.spacingHorizontalM) },
  search: { ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalM) },
  banner: {
    display: "flex", alignItems: "center", justifyContent: "space-between",
    ...shorthands.gap(tokens.spacingHorizontalS),
    ...shorthands.margin("0", tokens.spacingHorizontalM, tokens.spacingVerticalS),
    ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalM),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorPaletteYellowBackground2,
    color: tokens.colorNeutralForeground1,
  },
  notice: {
    display: "flex", alignItems: "center", justifyContent: "space-between",
    ...shorthands.gap(tokens.spacingHorizontalS),
    ...shorthands.margin("0", tokens.spacingHorizontalM, tokens.spacingVerticalS),
    ...shorthands.padding(tokens.spacingVerticalXS, tokens.spacingHorizontalM),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorNeutralBackground3,
  },
  content: { flexGrow: 1, minHeight: 0 },
  footer: { ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalM), ...shorthands.borderTop("1px", "solid", tokens.colorNeutralStroke2) },
  centered: { display: "flex", justifyContent: "center", ...shorthands.padding(tokens.spacingVerticalXXL, "0") },
});

export function App(): JSX.Element {
  const { theme } = useOfficeTheme();
  const s = useStyles();

  const init = useStore((st) => st.init);
  const status = useStore((st) => st.status);
  const error = useStore((st) => st.error);
  const hostAvailable = useStore((st) => st.hostAvailable);
  const view = useStore((st) => st.view);
  const setView = useStore((st) => st.setView);
  const search = useStore((st) => st.search);
  const setSearch = useStore((st) => st.setSearch);
  const selectedStyleId = useStore((st) => st.selectedStyleId);
  const selectStyle = useStore((st) => st.selectStyle);
  const selectedTokenId = useStore((st) => st.selectedTokenId);
  const selectToken = useStore((st) => st.selectToken);
  const shapes = useStore((st) => st.shapes);
  const refreshSelection = useStore((st) => st.refreshSelection);
  const startAdoption = useStore((st) => st.startAdoption);
  const createFromSelection = useStore((st) => st.createFromSelection);
  const applyStyleId = useStore((st) => st.applyStyleId);
  const notice = useStore((st) => st.notice);
  const setNotice = useStore((st) => st.setNotice);

  const [createOpen, setCreateOpen] = useState(false);
  const [createName, setCreateName] = useState("");

  useKeyboardShortcuts(() => document.getElementById("styles-search")?.focus());
  useEffect(() => { void init(); }, [init]);

  const unlinked = unlinkedShapeCount(shapes);

  async function openCreate(): Promise<void> {
    await refreshSelection();
    setCreateName("");
    setCreateOpen(true);
  }

  async function confirmCreate(): Promise<void> {
    const id = await createFromSelection(createName.trim());
    setCreateOpen(false);
    if (id) selectStyle(id);
  }

  return (
    <FluentProvider theme={theme} className={s.provider}>
      <ErrorBoundary>
      <div className={s.root}>
        <div className={s.header}>
          <Title3 className={s.title}>StyleSmith</Title3>
          <Badge appearance="tint" color={hostAvailable ? "success" : "warning"}>
            {hostAvailable ? "connected" : "no host"}
          </Badge>
        </div>

        <div className={s.nav}>
          <TabList selectedValue={view} onTabSelect={(_, d) => { setView(d.value as View); selectStyle(null); selectToken(null); }}>
            <Tab value="styles">Styles</Tab>
            <Tab value="tokens">Tokens</Tab>
            <Tab value="debug">Debug</Tab>
          </TabList>
        </div>

        {status === "loading" && <div className={s.centered}><Spinner label="Loading…" /></div>}
        {status === "error" && <div className={s.centered}><Body1>{error}</Body1></div>}

        {status === "ready" && (
          <>
            {view === "styles" && !selectedStyleId && (
              <div className={s.search}>
                <Input
                  value={search}
                  onChange={(_, d) => setSearch(d.value)}
                  placeholder="Search styles"
                  input={{ id: "styles-search", "aria-label": "Search styles" }}
                  contentBefore={<span aria-hidden>🔍</span>}
                />
              </div>
            )}

            {view === "styles" && !selectedStyleId && unlinked > 0 && (
              <div className={s.banner}>
                <Caption1>⚠ {unlinked} unlinked shapes</Caption1>
                <Button size="small" onClick={() => void startAdoption()}>Scan</Button>
              </div>
            )}

            {notice && (
              <div className={s.notice}>
                <Caption1>{notice}</Caption1>
                <Button size="small" appearance="subtle" onClick={() => setNotice(null)} aria-label="Dismiss">✕</Button>
              </div>
            )}

            <div className={s.content}>
              {view === "styles" && (
                selectedStyleId ? (
                  <StyleDetail styleId={selectedStyleId} onBack={() => selectStyle(null)} />
                ) : (
                  <StylesView
                    onOpenStyle={(id) => selectStyle(id)}
                    onCreate={() => void openCreate()}
                    onScan={() => void startAdoption()}
                  />
                )
              )}
              {view === "tokens" && (
                <Suspense fallback={<div className={s.centered}><Spinner size="tiny" label="Loading…" /></div>}>
                  {selectedTokenId ? (
                    <TokenDetail tokenId={selectedTokenId} onBack={() => selectToken(null)} />
                  ) : (
                    <TokensView />
                  )}
                </Suspense>
              )}
              {view === "debug" && (
                <Suspense fallback={<div className={s.centered}><Spinner size="tiny" label="Loading…" /></div>}>
                  <div style={{ padding: 12 }}><SpikePanel /></div>
                </Suspense>
              )}
            </div>

            {view === "styles" && !selectedStyleId && (
              <div className={s.footer}>
                <Button appearance="primary" style={{ width: "100%" }} onClick={() => void openCreate()}>
                  + New from selection
                </Button>
              </div>
            )}
          </>
        )}
      </div>

      {applyStyleId && <ApplyModal key={applyStyleId} />}
      <OperationView />
      <Suspense fallback={null}><AdoptionFlow /></Suspense>

      <Dialog open={createOpen} onOpenChange={(_, d) => setCreateOpen(d.open)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>New style from selection</DialogTitle>
            <DialogContent>
              <Field label="Style name" hint={hostAvailable ? "Captures the selected shape's fill, border, geometry and text." : "Connect PowerPoint to capture a shape's formatting."}>
                <Input value={createName} onChange={(_, d) => setCreateName(d.value)} autoFocus onKeyDown={(e) => { if (e.key === "Enter" && createName.trim()) void confirmCreate(); }} />
              </Field>
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setCreateOpen(false)}>Cancel</Button>
              <Button appearance="primary" disabled={createName.trim() === ""} onClick={() => void confirmCreate()}>Create</Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
      </ErrorBoundary>
    </FluentProvider>
  );
}
