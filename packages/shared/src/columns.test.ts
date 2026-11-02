import { describe, expect, it } from "vitest";
import {
  addWorkbookColumn,
  addableProcessFields,
  allowedWorkItemTypeNames,
  canRemoveColumn,
  flattenClassificationPaths,
  insertWorkItemRow,
  isAllowedPath,
  isAllowedPicklistValue,
  isAllowedState,
  isAllowedTransition,
  isAllowedWorkItemType,
  buildColumnLayout,
  moveWorkbookColumn,
  removeWorkbookColumn
} from "./columns";
import {
  applyParentTitles,
  fieldReferenceNames,
  identityLabelsFromPicks,
  identityLabelsFromSnapshot,
  mergeColumns,
  queryWebUrl,
  sameOrganization,
  withTitleLevels
} from "./fields";
import { PARENT_TITLE_FIELD } from "./constants";

describe("process template work item types", () => {
  it("omits disabled types and sorts names", () => {
    expect(
      allowedWorkItemTypeNames([
        { name: "User Story", isDisabled: false },
        { name: "Bug", isDisabled: false },
        { name: "Legacy", isDisabled: true },
        { name: "" }
      ])
    ).toEqual(["Bug", "User Story"]);
  });

  it("rejects types that are not in the project process", () => {
    expect(isAllowedWorkItemType("Bug", ["Bug", "User Story"])).toBe(true);
    expect(isAllowedWorkItemType("Epic", ["Bug", "User Story"])).toBe(false);
    expect(isAllowedWorkItemType("Bug", undefined)).toBe(true);
  });

  it("validates state against the work item type's process states", () => {
    const statesByType = {
      Bug: ["New", "Active", "Resolved", "Closed"],
      "User Story": ["New", "Active", "Closed"]
    };
    expect(isAllowedState("Active", "Bug", statesByType)).toBe(true);
    expect(isAllowedState("Resolved", "User Story", statesByType)).toBe(false);
    expect(isAllowedState("", "Bug", statesByType)).toBe(true);
    expect(isAllowedState("Active", "Bug", undefined)).toBe(true);
  });
});

describe("column chooser", () => {
  const columns = mergeColumns([{ referenceName: "System.State", name: "State" }]);

  it("locks id, rev, type, and title", () => {
    expect(columns.filter((column) => canRemoveColumn(column)).map((column) => column.referenceName)).toEqual([
      "System.State"
    ]);
  });

  it("adds a process field and will not duplicate it", () => {
    const withPriority = addWorkbookColumn(columns, {
      referenceName: "Microsoft.VSTS.Common.Priority",
      name: "Priority"
    });
    expect(withPriority.map((column) => column.referenceName)).toContain("Microsoft.VSTS.Common.Priority");
    expect(addWorkbookColumn(withPriority, { referenceName: "Microsoft.VSTS.Common.Priority", name: "Priority" })).toHaveLength(
      withPriority.length
    );
  });

  it("removes optional columns and refuses required ones", () => {
    const withPriority = addWorkbookColumn(columns, {
      referenceName: "Microsoft.VSTS.Common.Priority",
      name: "Priority"
    });
    expect(removeWorkbookColumn(withPriority, "Microsoft.VSTS.Common.Priority").map((column) => column.referenceName)).not.toContain(
      "Microsoft.VSTS.Common.Priority"
    );
    expect(() => removeWorkbookColumn(columns, "System.Title")).toThrow(/required/);
  });

  it("only offers process fields that are not already in the list", () => {
    const addable = addableProcessFields(
      [
        { referenceName: "System.State", name: "State" },
        { referenceName: "Microsoft.VSTS.Common.Priority", name: "Priority" },
        { referenceName: "System.History", name: "History" }
      ],
      columns
    );
    expect(addable.map((field) => field.referenceName)).toEqual([
      "Ado.ParentTitle",
      "Microsoft.VSTS.Common.Priority"
    ]);
  });

  it("offers Parent Title when it is not already a column", () => {
    const addable = addableProcessFields([], columns);
    expect(addable.map((field) => field.referenceName)).toContain("Ado.ParentTitle");
    const withParent = addWorkbookColumn(columns, {
      referenceName: "Ado.ParentTitle",
      name: "Parent Title"
    });
    expect(addableProcessFields([], withParent).map((field) => field.referenceName)).not.toContain("Ado.ParentTitle");
    expect(canRemoveColumn(withParent.find((column) => column.referenceName === "Ado.ParentTitle")!)).toBe(true);
  });
});

describe("insert work item row", () => {
  it("inserts a typed empty row below the selection", () => {
    const columns = mergeColumns([{ referenceName: "System.State", name: "State" }]);
    const rows = insertWorkItemRow(
      [
        {
          rowIndex: 0,
          id: 10,
          values: { "System.Id": 10, "System.WorkItemType": "Bug", "System.Title": "Existing" }
        }
      ],
      "User Story",
      columns,
      1,
      false,
      "New"
    );
    expect(rows).toHaveLength(2);
    expect(rows[1].values["System.WorkItemType"]).toBe("User Story");
    expect(rows[1].values["System.Title"]).toBe("");
    expect(rows[1].values["System.State"]).toBe("New");
    expect(rows[1].id).toBeUndefined();
  });
});

describe("column order and process values", () => {
  it("moves an optional column next to work item type", () => {
    const columns = withTitleLevels(
      mergeColumns([
        { referenceName: "System.State", name: "State" },
        { referenceName: "System.AssignedTo", name: "Assigned To" }
      ]),
      2
    );
    const moved = moveWorkbookColumn(columns, "System.AssignedTo", -1);
    expect(moved.map((column) => column.referenceName)).toEqual([
      "System.Id",
      "System.Rev",
      "System.WorkItemType",
      "System.Title",
      "System.AssignedTo",
      "System.State"
    ]);
  });

  it("flattens classification nodes and validates paths", () => {
    const paths = flattenClassificationPaths({
      name: "Fabrikam",
      children: [{ name: "Team A", children: [{ name: "Web" }] }, { name: "Team B" }]
    });
    expect(paths).toEqual(["Fabrikam", "Fabrikam\\Team A", "Fabrikam\\Team A\\Web", "Fabrikam\\Team B"]);
    expect(isAllowedPath("Fabrikam\\Team A", paths)).toBe(true);
    expect(isAllowedPath("Other\\Team", paths)).toBe(false);
    expect(isAllowedPath("", paths)).toBe(true);
  });

  it("allows suggested picklist values outside the list", () => {
    expect(isAllowedPicklistValue("1", ["1", "2", "3"])).toBe(true);
    expect(isAllowedPicklistValue("9", ["1", "2", "3"])).toBe(false);
    expect(isAllowedPicklistValue("custom", ["1", "2"], true)).toBe(true);
  });

  it("checks state transitions when the process defines them", () => {
    const transitions = {
      Bug: {
        New: ["Active", "Closed"],
        Active: ["Resolved", "Closed"]
      }
    };
    expect(isAllowedTransition("New", "Active", "Bug", transitions)).toBe(true);
    expect(isAllowedTransition("New", "Resolved", "Bug", transitions)).toBe(false);
    expect(isAllowedTransition("Active", "Active", "Bug", transitions)).toBe(true);
    expect(isAllowedTransition("", "New", "Bug", transitions)).toBe(true);
  });
});

describe("identity cell lists", () => {
  it("formats unique people for an Excel dropdown", () => {
    expect(
      identityLabelsFromPicks([
        { displayName: "Ada Lovelace", uniqueName: "ada@contoso.com" },
        { displayName: "Ada Lovelace", uniqueName: "ada@contoso.com" },
        { displayName: "Grace Hopper" }
      ])
    ).toEqual(["Ada Lovelace <ada@contoso.com>", "Grace Hopper"]);
  });

  it("reuses people already on the sheet", () => {
    expect(
      identityLabelsFromSnapshot(
        {
          "1": { "System.AssignedTo": "Ada Lovelace <ada@contoso.com>" },
          "2": { "System.AssignedTo": "Ada Lovelace <ada@contoso.com>" }
        },
        [{ referenceName: "System.AssignedTo", name: "Assigned To", identity: true }]
      )
    ).toEqual(["Ada Lovelace <ada@contoso.com>"]);
  });
});

describe("parent title column", () => {
  it("does not request the synthetic field from Azure DevOps", () => {
    const names = fieldReferenceNames([
      { referenceName: "System.State", name: "State" },
      { referenceName: PARENT_TITLE_FIELD, name: "Parent Title" }
    ]);
    expect(names).toContain("System.Parent");
    expect(names).not.toContain(PARENT_TITLE_FIELD);
  });

  it("fills parent titles from a title map", () => {
    const titles = new Map([
      [1, "Epic"],
      [2, "Story"]
    ]);
    const items = applyParentTitles(
      [
        { id: 2, fields: { "System.Title": "Story", "System.Parent": 1 } },
        { id: 3, fields: { "System.Title": "Task", "System.Parent": 2 } },
        { id: 4, fields: { "System.Title": "Orphan" } }
      ],
      titles
    );
    expect(items[0].fields?.[PARENT_TITLE_FIELD]).toBe("Epic");
    expect(items[1].fields?.[PARENT_TITLE_FIELD]).toBe("Story");
    expect(items[2].fields?.[PARENT_TITLE_FIELD]).toBe("");
  });
});

describe("buildColumnLayout", () => {
  const current = [
    { referenceName: "System.Id", name: "ID" },
    { referenceName: "System.WorkItemType", name: "Work Item Type" },
    { referenceName: "System.Title", name: "Title" },
    { referenceName: "System.State", name: "State" },
    { referenceName: "System.Tags", name: "Tags" }
  ];
  const addable = [
    { referenceName: "Microsoft.VSTS.Common.Priority", name: "Priority" },
    { referenceName: "System.AssignedTo", name: "Assigned To", isIdentity: true }
  ];

  it("applies adds, removes, and moves in one layout", () => {
    const change = buildColumnLayout(
      current,
      [
        "System.Id",
        "System.WorkItemType",
        "System.Title",
        "System.AssignedTo",
        "Microsoft.VSTS.Common.Priority",
        "System.State"
      ],
      addable
    );
    expect(change.columns.map((column) => column.referenceName)).toEqual([
      "System.Id",
      "System.WorkItemType",
      "System.Title",
      "System.AssignedTo",
      "Microsoft.VSTS.Common.Priority",
      "System.State"
    ]);
    expect(change.added.map((column) => column.referenceName)).toEqual([
      "System.AssignedTo",
      "Microsoft.VSTS.Common.Priority"
    ]);
    expect(change.added[0].identity).toBe(true);
    expect(change.removed.map((column) => column.referenceName)).toEqual(["System.Tags"]);
    expect(change.reordered).toBe(false);
  });

  it("detects a reorder of existing columns", () => {
    const change = buildColumnLayout(
      current,
      ["System.State", "System.Id", "System.WorkItemType", "System.Title", "System.Tags"],
      addable
    );
    expect(change.reordered).toBe(true);
    expect(change.added).toEqual([]);
    expect(change.removed).toEqual([]);
  });

  it("refuses to drop required columns or add unknown fields", () => {
    expect(() => buildColumnLayout(current, ["System.WorkItemType", "System.Title", "System.State"], addable)).toThrow(
      /ID column is required/
    );
    expect(() =>
      buildColumnLayout(current, [...current.map((column) => column.referenceName), "Custom.Missing"], addable)
    ).toThrow(/not available/);
  });

  it("keeps tree Title levels right after Work Item Type", () => {
    const tree = [
      { referenceName: "System.Id", name: "ID" },
      { referenceName: "System.WorkItemType", name: "Work Item Type" },
      { referenceName: "System.Title", name: "Title 1", titleLevel: 1 },
      { referenceName: "System.Title", name: "Title 2", titleLevel: 2 },
      { referenceName: "System.State", name: "State" }
    ];
    const change = buildColumnLayout(tree, ["System.State", "System.WorkItemType", "System.Id"], addable);
    expect(change.columns.map((column) => column.name)).toEqual(["State", "Work Item Type", "Title 1", "Title 2", "ID"]);
  });
});

describe("queryWebUrl", () => {
  it("links to the saved query's results page", () => {
    expect(queryWebUrl("https://dev.azure.com/contoso/", "Fabrikam Fiber", "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee")).toBe(
      "https://dev.azure.com/contoso/Fabrikam%20Fiber/_queries/query/aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee/"
    );
  });
});

describe("sameOrganization", () => {
  it("ignores case, trailing slashes, and the legacy visualstudio.com host", () => {
    expect(sameOrganization("https://dev.azure.com/Contoso/", "https://dev.azure.com/contoso")).toBe(true);
    expect(sameOrganization("https://contoso.visualstudio.com", "https://dev.azure.com/contoso")).toBe(true);
    expect(sameOrganization("https://dev.azure.com/contoso", "https://dev.azure.com/fabrikam")).toBe(false);
  });
});

describe("mapConcurrent", () => {
  it("preserves order with capped concurrency", async () => {
    const { mapConcurrent } = await import("./fields");
    const seen: number[] = [];
    const results = await mapConcurrent([3, 1, 2], 2, async (value) => {
      seen.push(value);
      await new Promise((resolve) => setTimeout(resolve, value * 5));
      return value * 10;
    });
    expect(results).toEqual([30, 10, 20]);
    expect(seen).toHaveLength(3);
  });
});
