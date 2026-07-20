// Convert engine / planner errors into human-readable guidance.

export function friendlyError(err: unknown): string {
  const raw = err instanceof Error ? err.message : String(err);
  const low = raw.toLowerCase();
  if (low.includes("needs clarification") || low.includes("shared key column")) {
    return "I couldn't find a shared column to match on. Please pick which column should be used, or rename the columns so they match.";
  }
  if (low.includes("no plan available")) {
    return "There's no plan to run yet. Tell me what you'd like to do first.";
  }
  if (low.includes("key column") && low.includes("not found")) {
    return `A column I was told to match on doesn't exist in one of the files. ${raw}`;
  }
  if (low.includes("no input files")) {
    return "This workspace has no uploaded files yet. Upload one or more spreadsheets to begin.";
  }
  if (low.includes("planner produced no operations")) {
    return "I couldn't turn that into a concrete plan. Try rephrasing (e.g. \"merge on Registration Number\" or \"find missing students\").";
  }
  if (low.includes("unauthorized") || low.includes("forbidden")) {
    return "You don't have access to that workspace.";
  }
  if (low.includes("zod") || low.includes("invalid_type") || low.includes("expected")) {
    return "I couldn't understand that request. Could you rephrase it more directly?";
  }
  return raw;
}

// Human-readable summary of run stats for the assistant reply.
export function summarizeRun(stats: Record<string, unknown>, warnings: string[]): string {
  const ops = Array.isArray(stats.ops) ? (stats.ops as Record<string, unknown>[]) : [];
  const lines: string[] = [];
  for (const op of ops) {
    if (op.status === "skipped") {
      lines.push(`• Skipped ${op.op}: ${op.reason ?? "unknown reason"}`);
      continue;
    }
    if (op.op === "merge") {
      lines.push(`• Merged ${op.inputFiles ?? "?"} files on "${op.keyColumn}" — ${op.merged ?? 0} rows (${op.unmatched ?? 0} unmatched).`);
    } else if (op.op === "dedupe") {
      lines.push(`• Removed ${op.duplicatesRemoved ?? 0} duplicate rows (${op.rowsOut ?? 0} kept).`);
    } else if (op.op === "diff") {
      const onlyInA = op.onlyInA ?? op.missingInB ?? 0;
      const onlyInB = op.onlyInB ?? op.missingInA ?? 0;
      lines.push(
        `• Compared **${op.fileA}** vs **${op.fileB}** on "${op.keyColumn}" — **${onlyInA}** only in ${op.fileA}, **${onlyInB}** only in ${op.fileB}, ${op.changed ?? 0} changed. See sheets "Only in ${op.fileA}" and "Only in ${op.fileB}".`,
      );
    } else if (op.op === "intersection") {
      lines.push(`• Found ${op.common ?? 0} rows present in both files (matched on "${op.keyColumn}").`);
    } else if (op.op === "summary") {
      lines.push(`• Generated a summary sheet.`);
    } else if (op.op === "master_merge") {
      lines.push(
        `• Built master sheet on "${op.keyColumn}" — ${op.rows ?? 0} rows across ${op.files ?? 0} files (${op.joinType}${op.duplicates ? `, ${op.duplicates} duplicate keys` : ""}).`,
      );
    } else if (op.op === "bulk_lookup") {
      lines.push(
        `• Bulk lookup: ${op.matched ?? 0}/${op.queries ?? 0} matched (${op.notFound ?? 0} not found, ${op.rowsReturned ?? 0} rows returned).`,
      );
    } else {
      lines.push(`• Ran ${op.op}.`);
    }

  }
  if (warnings.length) lines.push(`\nNotes: ${warnings.join("; ")}`);
  return lines.join("\n") || "Done.";
}
