import { Component, type ErrorInfo, type ReactNode } from "react";
import { Button, Body1, Subtitle2, Caption1, tokens } from "@fluentui/react-components";

interface Props {
  children: ReactNode;
}
interface State {
  error: Error | null;
}

/**
 * Catches render-time exceptions so a bug never leaves the panel blank
 * (ACX.10: surface errors in plain language). Recoverable via "Reload panel".
 */
export class ErrorBoundary extends Component<Props, State> {
  state: State = { error: null };

  static getDerivedStateFromError(error: Error): State {
    return { error };
  }

  componentDidCatch(error: Error, info: ErrorInfo): void {
    // eslint-disable-next-line no-console
    console.error("StyleSmith panel error:", error, info.componentStack);
  }

  render(): ReactNode {
    const { error } = this.state;
    if (!error) return this.props.children;
    return (
      <div
        style={{
          display: "flex",
          flexDirection: "column",
          gap: tokens.spacingVerticalM,
          padding: tokens.spacingVerticalXXL,
          maxWidth: 320,
        }}
      >
        <Subtitle2>Something went wrong</Subtitle2>
        <Body1>The panel hit an unexpected error. Your deck was not changed.</Body1>
        <Caption1 style={{ color: tokens.colorNeutralForeground3, fontFamily: tokens.fontFamilyMonospace }}>
          {error.message}
        </Caption1>
        <Button appearance="primary" onClick={() => window.location.reload()}>
          Reload panel
        </Button>
      </div>
    );
  }
}
