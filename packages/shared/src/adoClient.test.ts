import { afterEach, describe, expect, it, vi } from "vitest";
import {
  AdoClientError,
  AdoWorkItemClient,
  adoErrorDetail,
  collectQueryTree,
  flattenQueries,
  isQueryFolder,
  QueryNode,
  queryPickerRows,
  queryPickerTree,
  SavedQuery
} from "./adoClient";

describe("removeWorkItemRelation", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  function stubFetch() {
    const fetchMock = vi.fn(async () => new Response(JSON.stringify({ id: 101 }), { status: 200 }));
    vi.stubGlobal("fetch", fetchMock);
    return fetchMock;
  }

  it("tests the revision before removing so a shifted index can't remove the wrong relation", async () => {
    const fetchMock = stubFetch();
    const client = new AdoWorkItemClient("https://dev.azure.com/contoso", { type: "pat", value: "x" });
    await client.removeWorkItemRelation(101, 2, 7);
    const [, init] = fetchMock.mock.calls[0] as unknown as [string, RequestInit];
    expect(JSON.parse(String(init.body))).toEqual([
      { op: "test", path: "/rev", value: 7 },
      { op: "remove", path: "/relations/2" }
    ]);
  });

  it("removes without a revision test when none is known", async () => {
    const fetchMock = stubFetch();
    const client = new AdoWorkItemClient("https://dev.azure.com/contoso", { type: "pat", value: "x" });
    await client.removeWorkItemRelation(101, 0);
    const [, init] = fetchMock.mock.calls[0] as unknown as [string, RequestInit];
    expect(JSON.parse(String(init.body))).toEqual([{ op: "remove", path: "/relations/0" }]);
  });

  it("adds a hyperlink with its comment, and omits an empty comment", async () => {
    const fetchMock = stubFetch();
    const client = new AdoWorkItemClient("https://dev.azure.com/contoso", { type: "pat", value: "x" });
    await client.addWorkItemHyperlink(101, "https://example.com/spec", " Design spec ");
    await client.addWorkItemHyperlink(101, "https://example.com/other", "  ");
    const bodies = fetchMock.mock.calls.map((call) => JSON.parse(String((call as unknown as [string, RequestInit])[1].body)));
    expect(bodies[0]).toEqual([
      {
        op: "add",
        path: "/relations/-",
        value: { rel: "Hyperlink", url: "https://example.com/spec", attributes: { comment: "Design spec" } }
      }
    ]);
    expect(bodies[1][0].value).toEqual({ rel: "Hyperlink", url: "https://example.com/other" });
  });
});

describe("adoErrorDetail", () => {
  it("returns the message from an Azure DevOps error body", () => {
    const error = new AdoClientError("failed", 400, JSON.stringify({ message: "TF201036: Relation already exists." }));
    expect(adoErrorDetail(error)).toBe("TF201036: Relation already exists.");
  });

  it("returns undefined for bodies without a message", () => {
    expect(adoErrorDetail(new AdoClientError("failed", 400, "<html>"))).toBeUndefined();
    expect(adoErrorDetail(new Error("plain"))).toBeUndefined();
  });
});

function folder(name: string, children: QueryNode[], extra: Partial<QueryNode> = {}): QueryNode {
  return {
    id: extra.id || `${name}-id`,
    name,
    path: extra.path || name,
    isFolder: true,
    hasChildren: children.length > 0 || extra.hasChildren === true,
    children,
    ...extra
  };
}

function query(name: string, id = `${name}-id`): QueryNode {
  return { id, name, path: name, isFolder: false, queryType: "flat" };
}

/** Mimic WIT Get with $depth=2: include two nested levels, then a folder stub with no children. */
function depth2Slice(node: QueryNode, remaining = 2): QueryNode {
  if (!isQueryFolder(node)) {
    return { ...node };
  }
  if (remaining <= 0) {
    const { children: _children, hasChildren: _has, ...stub } = node;
    return { ...stub, isFolder: true };
  }
  return {
    ...node,
    children: (node.children || []).map((child) => depth2Slice(child, remaining - 1))
  };
}

function expandAll(root: QueryNode, fullTreeById: Map<string, QueryNode>): SavedQuery[] {
  const saved: SavedQuery[] = [];
  const seen = new Set<string>();
  const pending: string[] = [];
  seen.add(`fetch:${root.id}`);
  collectQueryTree(depth2Slice(root), saved, seen, pending);
  while (pending.length > 0) {
    const id = pending.shift()!;
    if (seen.has(`fetch:${id}`)) {
      continue;
    }
    seen.add(`fetch:${id}`);
    const next = fullTreeById.get(id);
    if (next) {
      collectQueryTree(depth2Slice(next), saved, seen, pending);
    }
  }
  return saved;
}

function indexFolders(node: QueryNode, map = new Map<string, QueryNode>()): Map<string, QueryNode> {
  if (node.id && isQueryFolder(node)) {
    map.set(node.id, node);
  }
  for (const child of node.children || []) {
    indexFolders(child, map);
  }
  return map;
}

describe("query tree flattening", () => {
  it("collects nested saved queries and skips folders", () => {
    const saved: SavedQuery[] = [];
    flattenQueries(
      folder("My Queries", [
        query("Assigned to me", "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"),
        folder("Sprint", [query("Sprint bugs", "bbbbbbbb-bbbb-cccc-dddd-eeeeeeeeeeee")])
      ]),
      saved
    );
    expect(saved.map((item) => item.name)).toEqual(["Assigned to me", "Sprint bugs"]);
  });

  it("queues unfetched folder stubs even when hasChildren is omitted", () => {
    const saved: SavedQuery[] = [];
    const pending: string[] = [];
    collectQueryTree(
      folder("Shared Queries", [
        query("All work", "cccccccc-bbbb-cccc-dddd-eeeeeeeeeeee"),
        { id: "folder-guid", name: "Deep", path: "Shared Queries/Deep", isFolder: true }
      ]),
      saved,
      new Set(["fetch:Shared Queries-id"]),
      pending
    );
    expect(saved.map((item) => item.id)).toEqual(["cccccccc-bbbb-cccc-dddd-eeeeeeeeeeee"]);
    expect(pending).toEqual(["folder-guid"]);
  });

  it("walks a 6-level parent/child folder tree across $depth=2 fetches", () => {
    const deep = query("Leaf query", "leaf-query-id");
    const tree = folder("Shared Queries", [
      folder("L1", [folder("L2", [folder("L3", [folder("L4", [folder("L5", [deep])])])])])
    ]);
    const saved = expandAll(tree, indexFolders(tree));
    expect(saved.map((item) => item.name)).toEqual(["Leaf query"]);
    expect(saved[0].id).toBe("leaf-query-id");
  });

  it("treats hasChildren without queryType as a folder", () => {
    expect(isQueryFolder({ name: "Folder", hasChildren: true })).toBe(true);
    expect(isQueryFolder({ name: "Query", queryType: "flat" })).toBe(false);
  });
});

describe("query picker tree", () => {
  it("nests folders above their queries", () => {
    const tree = queryPickerTree([
      { id: "2", name: "Stories", path: "Shared Queries/Current/Stories" },
      { id: "1", name: "Assigned to me", path: "My Queries/Assigned to me" },
      { id: "3", name: "Bugs", path: "Shared Queries/Current/Bugs" },
      { id: "4", name: "Backlog", path: "Shared Queries/Backlog" }
    ]);
    expect(tree.map((node) => node.name)).toEqual(["My Queries", "Shared Queries"]);
    expect(tree[0].folder).toBe(true);
    expect(tree[0].children.map((node) => ({ name: node.name, id: node.id, folder: node.folder }))).toEqual([
      { name: "Assigned to me", id: "1", folder: false }
    ]);
    expect(tree[1].children.map((node) => node.name)).toEqual(["Backlog", "Current"]);
    expect(tree[1].children[1].folder).toBe(true);
    expect(tree[1].children[1].children.map((node) => node.name)).toEqual(["Bugs", "Stories"]);
  });

  it("sorts siblings alphabetically and indents folders above their queries", () => {
    const rows = queryPickerRows([
      { id: "2", name: "Stories", path: "Shared Queries/Current/Stories" },
      { id: "1", name: "Assigned to me", path: "My Queries/Assigned to me" },
      { id: "3", name: "Bugs", path: "Shared Queries/Current/Bugs" },
      { id: "4", name: "Backlog", path: "Shared Queries/Backlog" }
    ]);
    expect(rows.map((row) => ({ label: row.label.trim(), disabled: row.disabled, id: row.id }))).toEqual([
      { label: "My Queries", disabled: true, id: "" },
      { label: "Assigned to me", disabled: false, id: "1" },
      { label: "Shared Queries", disabled: true, id: "" },
      { label: "Backlog", disabled: false, id: "4" },
      { label: "Current", disabled: true, id: "" },
      { label: "Bugs", disabled: false, id: "3" },
      { label: "Stories", disabled: false, id: "2" }
    ]);
    expect(rows[1].label.startsWith("  ")).toBe(true);
    expect(rows[5].label.startsWith("    ")).toBe(true);
  });
});
