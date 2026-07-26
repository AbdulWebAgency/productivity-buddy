// Barrel that triggers self-registration of every operation. Import this
// exactly once (from runPlan) so the registry is populated before dispatch.
import "./dedupe";
import "./summary";
import "./intersection";
import "./diff";
import "./merge";
import "./bulkLookup";
import "./masterMerge";
import "./highlight";
