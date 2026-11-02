import {
  ARTIFACT_LINK,
  ATTACHMENT_LINK,
  HYPERLINK_LINK,
  PARENT_LINK_FORWARD,
  PARENT_LINK_REVERSE,
  PREDECESSOR_LINK,
  RELATED_LINK,
  REMOTE_LINK_PREFIX,
  SUCCESSOR_LINK
} from "./constants";

/** One entry from Azure DevOps' work item relation types list (`_apis/wit/workitemrelationtypes`). */
export interface RelationTypeInfo {
  referenceName?: string;
  name?: string;
  attributes?: Record<string, unknown>;
}

export interface LinkTypeOption {
  rel: string;
  label: string;
  group: string;
}

const LINK_LABELS: Record<string, string> = {
  [RELATED_LINK]: "Related",
  [PREDECESSOR_LINK]: "Predecessor",
  [SUCCESSOR_LINK]: "Successor",
  [PARENT_LINK_REVERSE]: "Parent",
  [PARENT_LINK_FORWARD]: "Child",
  "System.LinkTypes.Duplicate-Forward": "Duplicate",
  "System.LinkTypes.Duplicate-Reverse": "Duplicate Of",
  "Microsoft.VSTS.Common.TestedBy-Forward": "Tested By",
  "Microsoft.VSTS.Common.TestedBy-Reverse": "Tests",
  "Microsoft.VSTS.Common.Affects-Forward": "Affects",
  "Microsoft.VSTS.Common.Affects-Reverse": "Affected By",
  [`${REMOTE_LINK_PREFIX}Related`]: "Remote Related",
  [HYPERLINK_LINK]: "Hyperlink",
  [ARTIFACT_LINK]: "Artifact",
  [ATTACHMENT_LINK]: "Attachment"
};

const LINK_GROUPS: Array<{ name: string; matches: (rel: string) => boolean }> = [
  { name: "Related", matches: (rel) => rel === RELATED_LINK },
  { name: "Dependency", matches: (rel) => rel.startsWith("System.LinkTypes.Dependency-") },
  { name: "Duplicate", matches: (rel) => rel.startsWith("System.LinkTypes.Duplicate-") },
  { name: "Test", matches: (rel) => rel.startsWith("Microsoft.VSTS.Common.TestedBy-") },
  { name: "Affects", matches: (rel) => rel.startsWith("Microsoft.VSTS.Common.Affects-") }
];
const OTHER_GROUP = "Other";
const GROUP_ORDER = [...LINK_GROUPS.map((group) => group.name), OTHER_GROUP];

/** Parent/child comes from indent, remote links need another organization, and Test Plans owns the test case links. */
const EXCLUDED_LINK_PREFIXES = ["System.LinkTypes.Hierarchy-", REMOTE_LINK_PREFIX, "Microsoft.VSTS.TestCase."];

export const DEFAULT_LINK_TYPES: LinkTypeOption[] = [
  { rel: RELATED_LINK, label: "Related", group: "Related" },
  { rel: PREDECESSOR_LINK, label: "Predecessor", group: "Dependency" },
  { rel: SUCCESSOR_LINK, label: "Successor", group: "Dependency" }
];

export function isHierarchyLink(rel: string): boolean {
  return rel === PARENT_LINK_FORWARD || rel === PARENT_LINK_REVERSE;
}

function readableRel(rel: string): string {
  return rel
    .replace(/^.*\./, "")
    .replace(/-(Forward|Reverse)$/, "")
    .replace(/([a-z])([A-Z])/g, "$1 $2");
}

/** Display name for a relation, preferring the names Azure DevOps returned for this organization. */
export function relationLabel(rel: string, names?: Record<string, string>): string {
  return names?.[rel] || LINK_LABELS[rel] || readableRel(rel);
}

export function relationTypeNames(types: RelationTypeInfo[]): Record<string, string> {
  const names: Record<string, string> = {};
  for (const type of types) {
    const rel = String(type.referenceName || "").trim();
    const name = String(type.name || "").trim();
    if (rel && name) {
      names[rel] = name;
    }
  }
  return names;
}

function linkGroup(rel: string): string {
  return LINK_GROUPS.find((group) => group.matches(rel))?.name || OTHER_GROUP;
}

/** Work item link types that can be added from the task pane, grouped and ordered for the picker. */
export function addableLinkTypes(types: RelationTypeInfo[]): LinkTypeOption[] {
  const options: LinkTypeOption[] = [];
  const seen = new Set<string>();
  for (const type of types) {
    const rel = String(type.referenceName || "").trim();
    const attributes = type.attributes || {};
    if (
      !rel ||
      seen.has(rel) ||
      attributes.usage !== "workItemLink" ||
      attributes.enabled === false ||
      attributes.remote === true ||
      EXCLUDED_LINK_PREFIXES.some((prefix) => rel.startsWith(prefix))
    ) {
      continue;
    }
    seen.add(rel);
    options.push({ rel, label: String(type.name || "").trim() || relationLabel(rel), group: linkGroup(rel) });
  }
  if (options.length === 0) {
    return DEFAULT_LINK_TYPES.slice();
  }
  return options.sort(
    (left, right) =>
      GROUP_ORDER.indexOf(left.group) - GROUP_ORDER.indexOf(right.group) || left.label.localeCompare(right.label)
  );
}

/** Accepts `https://…`, `http://…`, or a bare host such as `example.com/page` (treated as https). */
export function normalizeHyperlinkUrl(raw: string): string {
  const trimmed = raw.trim();
  if (!trimmed) {
    throw new Error("Enter a web address to link.");
  }
  const withScheme = /^[a-z][a-z0-9+.-]*:/i.test(trimmed) ? trimmed : `https://${trimmed}`;
  let parsed: URL;
  try {
    parsed = new URL(withScheme);
  } catch {
    throw new Error(`"${trimmed}" isn't a valid web address.`);
  }
  if ((parsed.protocol !== "http:" && parsed.protocol !== "https:") || !parsed.hostname) {
    throw new Error("Hyperlinks must start with http:// or https://.");
  }
  return parsed.href;
}
