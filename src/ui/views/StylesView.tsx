import { makeStyles, tokens, shorthands } from "@fluentui/react-components";
import { useStore, countShapesForStyle } from "../state/store";
import { StyleRow } from "../components/StyleRow";
import { EmptyState } from "../components/EmptyState";

const useStyles = makeStyles({
  list: {
    display: "flex",
    flexDirection: "column",
    ...shorthands.gap(tokens.spacingVerticalXS),
    ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalM),
    overflowY: "auto",
  },
});

export function StylesView({
  onOpenStyle,
  onCreate,
  onScan,
}: {
  onOpenStyle: (id: string) => void;
  onCreate: () => void;
  onScan: () => void;
}): JSX.Element {
  const s = useStyles();
  const doc = useStore((st) => st.doc);
  const shapes = useStore((st) => st.shapes);
  const search = useStore((st) => st.search).trim().toLowerCase();
  const canCreate = useStore((st) => st.selectionIds.length > 0 || !st.hostAvailable);

  if (!doc) return <></>;

  if (doc.styles.length === 0) {
    return <EmptyState shapeCount={shapes.length} onScan={onScan} onCreate={onCreate} canCreate={canCreate} />;
  }

  const visible = search
    ? doc.styles.filter((style) => style.name.toLowerCase().includes(search))
    : doc.styles;

  return (
    <div className={s.list}>
      {visible.map((style) => (
        <StyleRow
          key={style.id}
          doc={doc}
          style={style}
          usage={countShapesForStyle(shapes, style.id)}
          onOpen={() => onOpenStyle(style.id)}
        />
      ))}
    </div>
  );
}
