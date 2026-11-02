import { describe, expect, it } from "vitest";
import { isSupportedQueryId, WellKnownQueries } from "./constants";
import {
  dateToAdoValue,
  dateToYmd,
  fieldToCell,
  formatIdentityCell,
  idsFromQueryResult,
  isTreeQuery,
  mergeColumns,
  normalizeOrgUrl,
  parseWorkItemId,
  parseDateOnly,
  stripHtml,
  valuesEqual,
  withTitleLevels
} from "./fields";
import { applyRevisionConflicts, buildPublishPlan } from "./publish";
import { buildTreeLayout, childOutlineRanges, collapseTitleLevels, inferTreeParents, treeLevelFromParents, validateTree } from "./tree";
import { TableRow, WorkbookMetadata } from "./types";

function metadata(overrides: Partial<WorkbookMetadata> = {}): WorkbookMetadata {
  return {
    version: 2,
    orgUrl: "https://dev.azure.com/contoso",
    project: "Fabrikam",
    action: "OpenQuery",
    listType: "flat",
    columns: [
      { referenceName: "System.Id", name: "ID" },
      { referenceName: "System.WorkItemType", name: "Work Item Type" },
      { referenceName: "System.Title", name: "Title" },
      { referenceName: "System.State", name: "State" }
    ],
    loadedAt: "2026-09-14T00:00:00.000Z",
    revisions: { "101": 3 },
    snapshot: {
      "101": {
        "System.Id": 101,
        "System.WorkItemType": "Bug",
        "System.Title": "Original title",
        "System.State": "New"
      }
    },
    parents: {},
    ...overrides
  };
}

function row(rowIndex: number, values: TableRow["values"], extra: Partial<TableRow> = {}): TableRow {
  return { rowIndex, id: parseWorkItemId(values["System.Id"]), values, ...extra };
}

describe("query support", () => {
  it("rejects unsaved and well-known query ids", () => {
    expect(isSupportedQueryId(WellKnownQueries.AssignedToMe)).toBe(false);
    expect(isSupportedQueryId("aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee")).toBe(true);
  });
});

describe("field conversion", () => {
  it("flattens identity objects and HTML", () => {
    expect(fieldToCell({ displayName: "Ada Lovelace", uniqueName: "ada@contoso.com" })).toBe(
      "Ada Lovelace <ada@contoso.com>"
    );
    expect(formatIdentityCell({ displayName: "Ada Lovelace", uniqueName: "ada@contoso.com" })).toBe(
      "Ada Lovelace <ada@contoso.com>"
    );
    expect(stripHtml("<div>Line 1<br/>Line 2</div>")).toBe("Line 1\nLine 2");
    expect(fieldToCell("<div>Line 1<br/>Line 2</div>", "System.Description")).toBe("Line 1\nLine 2");
    expect(valuesEqual(" New ", "New")).toBe(true);
    expect(parseWorkItemId("42")).toBe(42);
    expect(normalizeOrgUrl("https://dev.azure.com/org/")).toBe("https://dev.azure.com/org");
  });

  it("stores date fields as Excel serials from ISO timestamps", () => {
    const serial = fieldToCell("2026-09-18T04:00:00.000Z", "Microsoft.VSTS.Scheduling.TargetDate");
    expect(typeof serial).toBe("number");
    expect(dateToYmd(serial)).toBe("2026-09-18");
    expect(dateToAdoValue(serial)).toBe("2026-09-18T00:00:00.000Z");
    expect(valuesEqual(serial, "2026-09-18T04:00:00Z", "Microsoft.VSTS.Scheduling.TargetDate")).toBe(true);
    expect(dateToYmd(fieldToCell("2026-09-18T04:00:00:00Z", "Microsoft.VSTS.Scheduling.TargetDate"))).toBe("2026-09-18");
    expect(parseDateOnly("Thu, 18 Sep 2026 00:00:00 GMT")).toEqual({ year: 2026, month: 9, day: 18 });
  });

  it("extracts ids from tree query relations", () => {
    expect(
      idsFromQueryResult({
        workItemRelations: [{ target: { id: 1 } }, { target: { id: 1 } }, { target: { id: 2 } }]
      })
    ).toEqual([1, 2]);
  });

  it("always includes id, rev, type, and title columns", () => {
    const columns = mergeColumns([{ referenceName: "System.State", name: "State" }]);
    expect(columns.map((column) => column.referenceName)).toEqual([
      "System.Id",
      "System.Rev",
      "System.WorkItemType",
      "System.Title",
      "System.State"
    ]);
  });
});

describe("tree layout", () => {
  const treeResult = {
    queryType: "tree",
    workItemRelations: [
      { target: { id: 10 } },
      { rel: "System.LinkTypes.Hierarchy-Forward", source: { id: 10 }, target: { id: 11 } },
      { rel: "System.LinkTypes.Hierarchy-Forward", source: { id: 11 }, target: { id: 12 } },
      { target: { id: 20 } }
    ]
  };

  it("detects tree queries and builds parent/depth maps", () => {
    expect(isTreeQuery(treeResult)).toBe(true);
    const layout = buildTreeLayout(treeResult);
    expect(layout?.order).toEqual([10, 11, 12, 20]);
    expect(layout?.parentById[11]).toBe(10);
    expect(layout?.parentById[12]).toBe(11);
    expect(layout?.depthById[12]).toBe(3);
    expect(layout?.maxDepth).toBeGreaterThanOrEqual(3);
  });

  it("keeps a single Title column for tree lists", () => {
    const columns = withTitleLevels(mergeColumns([{ referenceName: "System.State", name: "State" }]), 2);
    expect(columns.map((column) => column.name)).toEqual([
      "ID",
      "Rev",
      "Work Item Type",
      "Title",
      "State"
    ]);
    expect(columns.filter((column) => column.titleLevel)).toEqual([]);
  });

  it("infers parents from Title levels and flags missing parents", () => {
    const rows = inferTreeParents([
      row(0, { "System.Id": 10, "System.WorkItemType": "Epic", "System.Title": "Epic" }, { treeLevel: 1 }),
      row(1, { "System.Id": 11, "System.WorkItemType": "User Story", "System.Title": "Story" }, { treeLevel: 2 }),
      row(2, { "System.Id": "", "System.WorkItemType": "Task", "System.Title": "New task" }, { treeLevel: 3 })
    ]);
    expect(rows[1].parentId).toBe(10);
    expect(rows[2].parentRowIndex).toBe(1);
    expect(validateTree(rows, true)).toEqual([]);
  });

  it("computes outline groups and parent-chain depth", () => {
    expect(childOutlineRanges([1, 2, 3, 2, 1])).toEqual([
      { start: 2, end: 2 },
      { start: 1, end: 3 }
    ]);
    expect(treeLevelFromParents(12, { "12": 11, "11": 10 })).toBe(3);
    expect(treeLevelFromParents(10, { "12": 11, "11": 10 })).toBe(1);
  });

  it("reads legacy Title N headers without keeping those columns", () => {
    const collapsed = collapseTitleLevels(
      ["ID", "Title 1", "Title 2"],
      [11, "", "Child"],
      [{ name: "ID", referenceName: "System.Id" }]
    );
    expect(collapsed).toEqual({ title: "Child", treeLevel: 2 });
  });
});

describe("publish plan", () => {
  it("skips unchanged rows and patches edited fields", () => {
    const plan = buildPublishPlan(
      [
        row(0, {
          "System.Id": 101,
          "System.WorkItemType": "Bug",
          "System.Title": "Original title",
          "System.State": "New"
        }),
        row(1, {
          "System.Id": 101,
          "System.WorkItemType": "Bug",
          "System.Title": "Updated title",
          "System.State": "Active"
        })
      ],
      metadata()
    );

    expect(plan.skipCount).toBe(1);
    expect(plan.updateCount).toBe(1);
    expect(plan.changes[1].operations).toEqual([
      { op: "add", path: "/fields/System.Title", value: "Updated title" },
      { op: "add", path: "/fields/System.State", value: "Active" }
    ]);
  });

  it("rejects changing the work item type of an existing row", () => {
    const plan = buildPublishPlan(
      [
        row(0, {
          "System.Id": 101,
          "System.WorkItemType": "User Story",
          "System.Title": "Updated title",
          "System.State": "New"
        })
      ],
      metadata()
    );

    expect(plan.updateCount).toBe(0);
    expect(plan.changes[0].kind).toBe("error");
    expect(plan.changes[0].message).toContain("Changing the work item type isn't supported from Excel");
    expect(plan.changes[0].message).toContain("Bug");
  });

  it("treats a type that differs only in case as unchanged", () => {
    const plan = buildPublishPlan(
      [
        row(0, {
          "System.Id": 101,
          "System.WorkItemType": "bug",
          "System.Title": "Original title",
          "System.State": "New"
        })
      ],
      metadata()
    );

    expect(plan.changes[0].kind).toBe("skip");
  });

  it("creates rows without an id when type and title are present", () => {
    const plan = buildPublishPlan(
      [
        row(0, {
          "System.Id": "",
          "System.WorkItemType": "User Story",
          "System.Title": "New story",
          "System.State": "New"
        })
      ],
      metadata()
    );

    expect(plan.createCount).toBe(1);
    expect(plan.changes[0].kind).toBe("create");
    expect(plan.changes[0].workItemType).toBe("User Story");
    expect(plan.changes[0].operations?.some((operation) => operation.path === "/fields/System.Title")).toBe(true);
    expect(plan.changes[0].operations?.some((operation) => operation.path === "/fields/System.WorkItemType")).toBe(false);
  });

  it("rejects new rows whose work item type is not in the project process", () => {
    const plan = buildPublishPlan(
      [
        row(0, {
          "System.Id": "",
          "System.WorkItemType": "Epic",
          "System.Title": "New epic"
        })
      ],
      metadata({ workItemTypes: ["Bug", "User Story"] })
    );

    expect(plan.createCount).toBe(0);
    expect(plan.changes[0].kind).toBe("error");
    expect(plan.changes[0].message).toMatch(/process template/);
  });

  it("rejects states that are not valid for the work item type", () => {
    const plan = buildPublishPlan(
      [
        row(0, {
          "System.Id": "",
          "System.WorkItemType": "User Story",
          "System.Title": "New story",
          "System.State": "Resolved"
        })
      ],
      metadata({
        workItemTypes: ["Bug", "User Story"],
        statesByType: {
          Bug: ["New", "Active", "Resolved", "Closed"],
          "User Story": ["New", "Active", "Closed"]
        }
      })
    );

    expect(plan.createCount).toBe(0);
    expect(plan.changes[0].kind).toBe("error");
    expect(plan.changes[0].message).toMatch(/valid state/);
  });

  it("rejects picklist values that are not in the process list", () => {
    const plan = buildPublishPlan(
      [
        row(0, {
          "System.Id": "",
          "System.WorkItemType": "Bug",
          "System.Title": "New bug",
          "Microsoft.VSTS.Common.Priority": "9"
        })
      ],
      metadata({
        workItemTypes: ["Bug"],
        picklists: { "Microsoft.VSTS.Common.Priority": ["1", "2", "3", "4"] }
      })
    );
    expect(plan.createCount).toBe(0);
    expect(plan.changes[0].kind).toBe("error");
    expect(plan.changes[0].message).toMatch(/valid value/);
  });

  it("rejects area paths that are not in the project tree", () => {
    const plan = buildPublishPlan(
      [
        row(0, {
          "System.Id": "",
          "System.WorkItemType": "Bug",
          "System.Title": "New bug",
          "System.AreaPath": "Other\\Team"
        })
      ],
      metadata({
        workItemTypes: ["Bug"],
        areaPaths: ["Fabrikam", "Fabrikam\\Team A"]
      })
    );
    expect(plan.changes[0].kind).toBe("error");
    expect(plan.changes[0].message).toMatch(/area path/);
  });

  it("publishes Markdown HTML fields as HTML without rewriting unchanged descriptions", () => {
    const withHtml = metadata({
      htmlFields: ["System.Description"],
      snapshot: {
        "101": {
          "System.Id": 101,
          "System.WorkItemType": "Bug",
          "System.Title": "Original title",
          "System.State": "New",
          "System.Description": "Keep **this**"
        }
      }
    });
    const unchanged = buildPublishPlan(
      [
        row(0, {
          "System.Id": 101,
          "System.WorkItemType": "Bug",
          "System.Title": "Original title",
          "System.State": "New",
          "System.Description": "Keep **this**"
        })
      ],
      withHtml
    );
    expect(unchanged.skipCount).toBe(1);

    const changed = buildPublishPlan(
      [
        row(0, {
          "System.Id": 101,
          "System.WorkItemType": "Bug",
          "System.Title": "Original title",
          "System.State": "New",
          "System.Description": "Now **bold**"
        })
      ],
      withHtml
    );
    expect(changed.updateCount).toBe(1);
    expect(changed.changes[0].operations).toEqual([
      { op: "add", path: "/fields/System.Description", value: "<div>Now <b>bold</b></div>" }
    ]);
  });

  it("marks revision mismatches as conflicts", () => {
    const planned = buildPublishPlan(
      [
        row(0, {
          "System.Id": 101,
          "System.WorkItemType": "Bug",
          "System.Title": "Updated title",
          "System.State": "New"
        })
      ],
      metadata()
    );

    const withConflicts = applyRevisionConflicts(planned, metadata(), { "101": 9 });
    expect(withConflicts.changes[0].kind).toBe("conflict");
    expect(withConflicts.updateCount).toBe(0);
  });

  it("treats parent changes as publishable even when fields are unchanged", () => {
    const plan = buildPublishPlan(
      [
        row(0, { "System.Id": 10, "System.WorkItemType": "Epic", "System.Title": "Epic" }, { treeLevel: 1 }),
        row(
          1,
          { "System.Id": 101, "System.WorkItemType": "Bug", "System.Title": "Original title", "System.State": "New" },
          { treeLevel: 2 }
        )
      ],
      metadata({
        listType: "tree",
        treeLevels: 2,
        parents: { "101": 99 },
        snapshot: {
          "10": { "System.Id": 10, "System.WorkItemType": "Epic", "System.Title": "Epic" },
          "101": {
            "System.Id": 101,
            "System.WorkItemType": "Bug",
            "System.Title": "Original title",
            "System.State": "New"
          }
        }
      })
    );

    const child = plan.changes.find((change) => change.id === 101);
    expect(child?.kind).toBe("update");
    expect(child?.desiredParentId).toBe(10);
    expect(child?.previousParentId).toBe(99);
    expect(child?.message).toContain("parent");
  });
});
