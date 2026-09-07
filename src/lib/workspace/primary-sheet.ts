// Client-safe helper: pick the sheet the app should reason over.
// Falls back to the first sheet for older inspector payloads that predate
// worksheet classification (backward compatible).
export type MinimalInspectorSheet = {
  name: string;
  headers?: string[];
  worksheetType?: string;
  likelyPrimaryTable?: boolean;
  rank?: number;
};

export type MinimalInspector = {
  sheets?: MinimalInspectorSheet[];
  primaryDataSheet?: string | null;
} | null;

export function primarySheetOf<T extends MinimalInspectorSheet>(
  inspector: { sheets?: T[]; primaryDataSheet?: string | null } | null | undefined,
): T | undefined {
  const sheets = inspector?.sheets ?? [];
  if (sheets.length === 0) return undefined;
  const named = inspector?.primaryDataSheet
    ? sheets.find((s) => s.name === inspector.primaryDataSheet)
    : undefined;
  if (named) return named;
  const flagged = sheets.find((s) => s.likelyPrimaryTable);
  if (flagged) return flagged;
  const data = sheets.filter((s) => s.worksheetType === "DATA");
  if (data.length) return [...data].sort((a, b) => (b.rank ?? 0) - (a.rank ?? 0))[0];
  return sheets[0];
}
