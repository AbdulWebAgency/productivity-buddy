// Deterministic action suggestions based on uploaded files + inspector.
import type { InspectorReport } from "./inspector.server";
import { normalizeHeader } from "@/lib/excel/engine/shared/headers";
import { primarySheetOf } from "./primary-sheet";

export type Suggestion = { label: string; prompt: string };

export type SuggestionInput = {
  files: { name: string; inspector: InspectorReport | null }[];
};

function firstKey(files: SuggestionInput["files"]): string | null {
  for (const f of files) {
    const k = primarySheetOf(f.inspector)?.likelyKeys[0];
    if (k) return k;
  }
  return null;
}

function sharedHeaders(files: SuggestionInput["files"]): string[] {
  if (files.length < 2) return [];
  const perFile = files.map((f) =>
    new Set((primarySheetOf(f.inspector)?.headers ?? []).map(normalizeHeader)),
  );
  const base = perFile[0];
  const shared: string[] = [];
  for (const h of base) if (perFile.slice(1).every((s) => s.has(h))) shared.push(h);
  return shared;
}

export function suggestActions(input: SuggestionInput): Suggestion[] {
  const out: Suggestion[] = [];
  const key = firstKey(input.files);
  const shared = sharedHeaders(input.files);

  if (input.files.length >= 2) {
    out.push({
      label: `Create Master Sheet`,
      prompt: `Create a master sheet by merging all uploaded files.`,
    });
  }
  if (input.files.length >= 1) {
    const target = input.files[input.files.length - 1]?.name ?? "file";
    out.push({
      label: `Bulk Lookup`,
      prompt: `Bulk lookup in ${target}:\n`,
    });
  }
  if (input.files.length >= 2 && (key || shared.length)) {
    const on = key ?? shared[0];
    out.push({
      label: `Find rows missing between files`,
      prompt: `Compare the two files using ${on} and list rows that are only in one of them.`,
    });
    out.push({
      label: `Merge files on ${on}`,
      prompt: `Merge all files on ${on} and highlight rows only present in one file.`,
    });
  }

  if (input.files.some((f) => (primarySheetOf(f.inspector)?.duplicateKeyValues ?? 0) > 0)) {
    out.push({
      label: `Remove duplicate rows`,
      prompt: `Remove duplicate rows${key ? ` based on ${key}` : ""}.`,
    });
  }
  out.push({
    label: `Generate a summary sheet`,
    prompt: `Add a summary sheet describing the data (row counts, numeric column stats).`,
  });
  if (primarySheetOf(input.files[0]?.inspector)?.blankRows) {
    out.push({
      label: `Clean up blank rows`,
      prompt: `Remove blank rows from the workbook.`,
    });
  }
  return out.slice(0, 5);
}
