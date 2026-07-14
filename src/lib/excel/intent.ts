// Intent classifier. Maps natural-language phrases to a canonical intent
// concept BEFORE we choose a deterministic engine op.

export type Intent =
  | "intersection" // rows in both datasets
  | "difference" // rows only in one (either direction, or specified side)
  | "merge" // union across files on a key
  | "dedupe"
  | "summary"
  | "clean"
  | "formula_help"
  | "capabilities" // "what can you do"
  | "unknown";

export type IntentMatch = {
  intent: Intent;
  side?: "A" | "B" | "either"; // for difference: only in A, only in B, or either side
  confidence: number; // 0..1
  matchedPhrase?: string;
};

// Ordered phrase lists — first match wins.
const PATTERNS: { intent: Intent; side?: IntentMatch["side"]; re: RegExp }[] = [
  // Capabilities / help
  { intent: "capabilities", re: /\b(what can you do|help|capabilities|features|how do you work|what.*can.*(you|this|ledgerly).*do)\b/i },

  // Intersection — "in both", "common", "shared", "matching"
  { intent: "intersection", re: /\b(present|available|exist|found|appear|listed)\s+(in\s+)?both\b/i },
  { intent: "intersection", re: /\bin\s+both(\s+files|\s+sheets|\s+workbooks)?\b/i },
  { intent: "intersection", re: /\b(common|shared|matching|overlapping|repeated)\s+(students|rows|records|entries|people|names|items|values)\b/i },
  { intent: "intersection", re: /\b(same\s+(people|students|rows|records))\b/i },
  { intent: "intersection", re: /\b(who\s+is\s+in\s+both|which\s+are\s+in\s+both)\b/i },
  { intent: "intersection", re: /\b(repeated\s+across\s+(files|sheets))\b/i },
  { intent: "intersection", re: /\bintersection\b/i },

  // Difference variants
  { intent: "difference", side: "A", re: /\bonly\s+in\s+(file\s*a|the\s+first|first\s+file)\b/i },
  { intent: "difference", side: "B", re: /\bonly\s+in\s+(file\s*b|the\s+second|second\s+file)\b/i },
  { intent: "difference", re: /\bmissing\s+(students|rows|records|entries|people|names)?\b/i },
  { intent: "difference", re: /\b(new\s+joinees|new\s+students|newly\s+added|joined\s+recently)\b/i },
  { intent: "difference", re: /\b(not\s+in|absent\s+from|missing\s+from)\b/i },
  { intent: "difference", re: /\b(who\s+joined|who.?s\s+new)\b/i },
  { intent: "difference", re: /\b(difference|differ|diff|compare)\b/i },
  { intent: "difference", re: /\bwhich\s+(students|rows|records|people).*(missing|absent|not)\b/i },

  // Merge
  { intent: "merge", re: /\b(merge|combine|consolidate|join|unify|union)\b/i },
  { intent: "merge", re: /\bappend\b/i },

  // Dedupe
  { intent: "dedupe", re: /\b(dedup|deduplicate|duplicate|duplicates|unique|uniq)\b/i },

  // Summary
  { intent: "summary", re: /\b(summar|overview|report|statistics|stats|describe)\b/i },

  // Cleaning
  { intent: "clean", re: /\b(clean|clean\s*up|tidy|blank\s+rows|empty\s+rows|trim|normalize)\b/i },

  // Formula help
  { intent: "formula_help", re: /\b(formula|vlookup|xlookup|sumif|countif|explain\s+the?\s+formula)\b/i },
];

export function classifyIntent(text: string): IntentMatch {
  const s = (text ?? "").trim();
  if (!s) return { intent: "unknown", confidence: 0 };
  for (const p of PATTERNS) {
    const m = s.match(p.re);
    if (m) {
      return {
        intent: p.intent,
        side: p.side,
        confidence: 0.85,
        matchedPhrase: m[0],
      };
    }
  }
  return { intent: "unknown", confidence: 0 };
}

// Detect follow-ups like "use ID NO instead" / "match on Name" so the planner
// can regenerate against a user-selected key.
export function extractKeyOverride(text: string): string | null {
  const patterns: RegExp[] = [
    /\buse\s+["']?([A-Za-z0-9 _\-.]{1,60}?)["']?\s+(?:instead|to\s+match|as\s+key|for\s+key)\b/i,
    /\b(?:match|join|key)\s+on\s+["']?([A-Za-z0-9 _\-.]{1,60}?)["']?\s*$/i,
    /\b(?:by|using)\s+["']?([A-Za-z0-9 _\-.]{1,60}?)["']?\s*$/i,
  ];
  for (const re of patterns) {
    const m = text.match(re);
    if (m?.[1]) return m[1].trim();
  }
  return null;
}
