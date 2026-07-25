/**
 * Thin wrapper over PowerPoint.run plus shared error/cancellation helpers. ALL
 * Office JS lives under office/; nothing outside this directory imports the
 * Office/PowerPoint globals directly.
 */

/** True when the PowerPoint host globals are present (sideloaded, not a bare browser tab). */
export function isHostAvailable(): boolean {
  return typeof PowerPoint !== "undefined" && typeof Office !== "undefined";
}

/** Run a batch against the PowerPoint host. Throws if not running inside PowerPoint. */
export async function runPowerPoint<T>(
  fn: (ctx: PowerPoint.RequestContext) => Promise<T>,
): Promise<T> {
  if (!isHostAvailable()) {
    throw new Error("Not running inside PowerPoint — sideload the add-in and open a deck.");
  }
  return PowerPoint.run(fn);
}

/** Thrown when a long-running operation is aborted at a checkpoint. */
export class OperationCancelled extends Error {
  constructor() {
    super("Operation cancelled");
    this.name = "OperationCancelled";
  }
}

/** Throw {@link OperationCancelled} if the signal is aborted. Call at sync boundaries. */
export function throwIfAborted(signal?: AbortSignal): void {
  if (signal?.aborted) throw new OperationCancelled();
}

/** A readable message for an Office JS / OfficeExtension error, for surfacing to users. */
export function officeErrorMessage(e: unknown): string {
  if (typeof OfficeExtension !== "undefined" && e instanceof OfficeExtension.Error) {
    return `${e.code}: ${e.message}`;
  }
  if (e instanceof Error) return e.message;
  return String(e);
}
