import { describe, expect, it } from "vitest";
import {
  DEFAULT_LINK_TYPES,
  RelationTypeInfo,
  addableLinkTypes,
  normalizeHyperlinkUrl,
  relationLabel,
  relationTypeNames
} from "./linkTypes";

function linkType(referenceName: string, name: string, attributes: Record<string, unknown> = {}): RelationTypeInfo {
  return { referenceName, name, attributes: { usage: "workItemLink", enabled: true, ...attributes } };
}

const ORG_TYPES: RelationTypeInfo[] = [
  linkType("System.LinkTypes.Hierarchy-Forward", "Child"),
  linkType("System.LinkTypes.Hierarchy-Reverse", "Parent"),
  linkType("Microsoft.VSTS.Common.TestedBy-Reverse", "Tests"),
  linkType("Microsoft.VSTS.Common.TestedBy-Forward", "Tested By"),
  linkType("System.LinkTypes.Duplicate-Reverse", "Duplicate Of"),
  linkType("System.LinkTypes.Duplicate-Forward", "Duplicate"),
  linkType("System.LinkTypes.Dependency-Forward", "Successor"),
  linkType("System.LinkTypes.Dependency-Reverse", "Predecessor"),
  linkType("System.LinkTypes.Related", "Related"),
  linkType("Microsoft.VSTS.TestCase.SharedStepReferencedBy-Forward", "Test Case"),
  linkType("System.LinkTypes.Remote.Related", "Remote Related", { remote: true }),
  linkType("Contoso.LinkTypes.Blocks-Forward", "Blocks"),
  linkType("Contoso.LinkTypes.Retired", "Retired", { enabled: false }),
  { referenceName: "Hyperlink", name: "Hyperlink", attributes: { usage: "resourceLink" } },
  { referenceName: "ArtifactLink", name: "Artifact Link", attributes: { usage: "resourceLink" } }
];

describe("addable link types", () => {
  it("keeps work item links, drops hierarchy, remote, test case, disabled, and resource links, and groups the rest", () => {
    expect(addableLinkTypes(ORG_TYPES).map((option) => `${option.group}: ${option.label}`)).toEqual([
      "Related: Related",
      "Dependency: Predecessor",
      "Dependency: Successor",
      "Duplicate: Duplicate",
      "Duplicate: Duplicate Of",
      "Test: Tested By",
      "Test: Tests",
      "Other: Blocks"
    ]);
  });

  it("includes Affects only when the organization has it", () => {
    const withAffects = [
      ...ORG_TYPES,
      linkType("Microsoft.VSTS.Common.Affects-Forward", "Affects"),
      linkType("Microsoft.VSTS.Common.Affects-Reverse", "Affected By")
    ];
    expect(addableLinkTypes(withAffects).filter((option) => option.group === "Affects").map((option) => option.label)).toEqual([
      "Affected By",
      "Affects"
    ]);
    expect(addableLinkTypes(ORG_TYPES).some((option) => option.group === "Affects")).toBe(false);
  });

  it("falls back to Related, Predecessor, and Successor when nothing usable comes back", () => {
    expect(addableLinkTypes([])).toEqual(DEFAULT_LINK_TYPES);
    expect(addableLinkTypes([linkType("System.LinkTypes.Hierarchy-Forward", "Child")])).toEqual(DEFAULT_LINK_TYPES);
  });
});

describe("relation labels", () => {
  it("prefers organization names, then built-in names, then a readable reference name", () => {
    const names = relationTypeNames(ORG_TYPES);
    expect(relationLabel("Contoso.LinkTypes.Blocks-Forward", names)).toBe("Blocks");
    expect(relationLabel("Microsoft.VSTS.Common.TestedBy-Forward")).toBe("Tested By");
    expect(relationLabel("System.LinkTypes.Duplicate-Reverse")).toBe("Duplicate Of");
    expect(relationLabel("Contoso.LinkTypes.DependsOn-Reverse")).toBe("Depends On");
  });
});

describe("hyperlink URLs", () => {
  it("accepts http and https, and treats a bare host as https", () => {
    expect(normalizeHyperlinkUrl(" https://example.com/spec ")).toBe("https://example.com/spec");
    expect(normalizeHyperlinkUrl("http://intranet/page")).toBe("http://intranet/page");
    expect(normalizeHyperlinkUrl("example.com/doc")).toBe("https://example.com/doc");
  });

  it("rejects empty, non-web, and malformed addresses", () => {
    expect(() => normalizeHyperlinkUrl("  ")).toThrow(/Enter a web address/);
    expect(() => normalizeHyperlinkUrl("javascript:alert(1)")).toThrow(/http/);
    expect(() => normalizeHyperlinkUrl("file:///C:/secret.txt")).toThrow(/http/);
    expect(() => normalizeHyperlinkUrl("https://")).toThrow();
  });
});
