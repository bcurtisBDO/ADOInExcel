import { MIN_TREE_LEVELS, PARENT_LINK_FORWARD, PARENT_TITLE_FIELD } from "./constants";
import { isTreeQuery, parseWorkItemId } from "./fields";
import { QueryResultLike, TableRow, TreeLayout, WorkItemLike } from "./types";

export function buildTreeLayout(result: QueryResultLike | undefined | null): TreeLayout | null {
  if (!result || !isTreeQuery(result)) {
    return null;
  }

  const parentById = new Map<number, number>();
  const children = new Map<number, number[]>();
  const roots: number[] = [];
  const seen = new Set<number>();

  for (const relation of result.workItemRelations || []) {
    const targetId = relation.target?.id;
    const sourceId = relation.source?.id;
    if (targetId) {
      seen.add(targetId);
    }
    if (sourceId) {
      seen.add(sourceId);
    }

    if (targetId && sourceId && isHierarchyRelation(relation.rel)) {
      if (!parentById.has(targetId)) {
        parentById.set(targetId, sourceId);
        const list = children.get(sourceId) || [];
        list.push(targetId);
        children.set(sourceId, list);
      }
    } else if (targetId && !sourceId) {
      roots.push(targetId);
    }
  }

  if (roots.length === 0) {
    for (const id of seen) {
      if (!parentById.has(id)) {
        roots.push(id);
      }
    }
  }

  return walkTree(roots, children, parentById, seen);
}

export function buildTreeFromParentField(items: WorkItemLike[]): TreeLayout | null {
  const parentById = new Map<number, number>();
  const children = new Map<number, number[]>();
  const ids = items.map((item) => item.id).filter((id): id is number => Boolean(id));
  const idSet = new Set(ids);

  for (const item of items) {
    if (!item.id) {
      continue;
    }
    const parent = parseWorkItemId(item.fields?.["System.Parent"]);
    if (parent && idSet.has(parent) && parent !== item.id) {
      parentById.set(item.id, parent);
      const list = children.get(parent) || [];
      list.push(item.id);
      children.set(parent, list);
    }
  }

  if (parentById.size === 0) {
    return null;
  }

  const roots = ids.filter((id) => !parentById.has(id));
  return walkTree(roots, children, parentById, new Set(ids));
}

function isHierarchyRelation(rel?: string): boolean {
  if (!rel) {
    return true;
  }
  return rel === PARENT_LINK_FORWARD || rel.toLowerCase().includes("hierarchy");
}

function walkTree(
  roots: number[],
  children: Map<number, number[]>,
  parentById: Map<number, number>,
  seen: Set<number>
): TreeLayout {
  const order: number[] = [];
  const depthById: Record<number, number> = {};
  const visiting = new Set<number>();

  const walk = (id: number, depth: number) => {
    if (visiting.has(id) || depthById[id]) {
      return;
    }
    visiting.add(id);
    order.push(id);
    depthById[id] = depth;
    for (const child of children.get(id) || []) {
      walk(child, depth + 1);
    }
    visiting.delete(id);
  };

  for (const root of roots) {
    walk(root, 1);
  }
  for (const id of seen) {
    if (!depthById[id]) {
      walk(id, 1);
    }
  }

  const maxDepth = Math.max(1, ...Object.values(depthById), MIN_TREE_LEVELS);
  const parents: Record<number, number> = {};
  parentById.forEach((parent, child) => {
    parents[child] = parent;
  });

  return {
    order,
    parentById: parents,
    depthById,
    maxDepth
  };
}

export function orderItems(items: WorkItemLike[], layout: TreeLayout): WorkItemLike[] {
  const byId = new Map<number, WorkItemLike>();
  for (const item of items) {
    if (item.id) {
      byId.set(item.id, item);
    }
  }
  const ordered: WorkItemLike[] = [];
  const used = new Set<number>();
  for (const id of layout.order) {
    const item = byId.get(id);
    if (item) {
      ordered.push(item);
      used.add(id);
    }
  }
  for (const item of items) {
    if (item.id && !used.has(item.id)) {
      ordered.push(item);
    }
  }
  return ordered;
}

export function inferTreeParents(rows: TableRow[]): TableRow[] {
  const lastAtLevel: Record<number, TableRow> = {};
  return rows.map((row) => {
    const level = row.treeLevel || 1;
    const next: TableRow = { ...row, parentId: undefined, parentRowIndex: undefined };
    if (level > 1) {
      const parent = lastAtLevel[level - 1];
      if (parent) {
        next.parentRowIndex = parent.rowIndex;
        next.parentId = parent.id;
      }
    }
    lastAtLevel[level] = next;
    for (const key of Object.keys(lastAtLevel)) {
      if (Number(key) > level) {
        delete lastAtLevel[Number(key)];
      }
    }
    return next;
  });
}

/** Resolve Parent Title from titles already on the sheet (no Azure DevOps call). */
export function fillParentTitlesOnRows(rows: TableRow[]): TableRow[] {
  const titles = new Map<number, string>();
  for (const row of rows) {
    if (row.id) {
      titles.set(row.id, String(row.values["System.Title"] ?? "").trim());
    }
  }
  return inferTreeParents(rows).map((row) => {
    const parentTitle = row.parentId ? titles.get(row.parentId) || "" : "";
    return {
      ...row,
      values: {
        ...row.values,
        [PARENT_TITLE_FIELD]: parentTitle
      }
    };
  });
}

export function validateTree(rows: TableRow[], isTree: boolean): string[] {
  if (!isTree) {
    return [];
  }

  const errors: string[] = [];
  let previousLevel = 0;
  let sawItem = false;

  for (const row of rows) {
    const title = String(row.values["System.Title"] ?? "").trim();
    const type = String(row.values["System.WorkItemType"] ?? "").trim();
    const empty = !title && !type && !row.id;
    if (empty) {
      continue;
    }

    const level = row.treeLevel || 1;
    if (!title) {
      errors.push(`Row ${row.rowIndex + 2}: tree items need a Title.`);
    }
    if (!sawItem && level !== 1) {
      errors.push(`Row ${row.rowIndex + 2}: the first work item must be at the top level (not indented).`);
    }
    if (sawItem && level > previousLevel + 1) {
      errors.push(`Row ${row.rowIndex + 2}: indent at most one level deeper than the row above.`);
    }
    if (level > 1 && row.parentRowIndex === undefined) {
      errors.push(`Row ${row.rowIndex + 2}: child has no parent. Indent it under a shallower row.`);
    }
    sawItem = true;
    previousLevel = level;
  }

  return errors;
}

export function collapseTitleLevels(
  headerValues: string[],
  line: (string | number | boolean)[],
  columns: { name: string; referenceName: string; titleLevel?: number }[]
): { title: string; treeLevel?: number; error?: string } {
  const filled: Array<{ level: number; title: string }> = [];
  const seen = new Set<number>();
  const addFilled = (level: number, title: string) => {
    if (!title || seen.has(level)) {
      return;
    }
    seen.add(level);
    filled.push({ level, title });
  };
  for (const column of columns) {
    if (!column.titleLevel) {
      continue;
    }
    const index = headerValues.findIndex((name) => name === column.name);
    if (index < 0) {
      continue;
    }
    addFilled(column.titleLevel, String(line[index] ?? "").trim());
  }
  headerValues.forEach((name, index) => {
    const match = /^Title\s+(\d+)$/i.exec(String(name).trim());
    if (match) {
      addFilled(Number(match[1]), String(line[index] ?? "").trim());
    }
  });

  if (filled.length > 1) {
    return {
      title: filled[0].title,
      treeLevel: filled[0].level,
      error: "Enter text in only one Title column per row."
    };
  }
  if (filled.length === 1) {
    return { title: filled[0].title, treeLevel: filled[0].level };
  }

  const titleIndex = headerValues.findIndex((name) => name === "Title" || name === "System.Title");
  if (titleIndex >= 0) {
    return { title: String(line[titleIndex] ?? "").trim() };
  }
  return { title: "" };
}

export function shiftTreeLevel(row: TableRow, delta: number, maxLevel: number): TableRow {
  const current = row.treeLevel || 1;
  const next = Math.min(maxLevel, Math.max(1, current + delta));
  return { ...row, treeLevel: next };
}

export function treeLevelFromParents(id: number | undefined, parents: Record<string, number>): number | undefined {
  if (!id) {
    return undefined;
  }
  let depth = 1;
  let current = id;
  const seen = new Set<number>();
  while (parents[String(current)]) {
    if (seen.has(current)) {
      break;
    }
    seen.add(current);
    current = parents[String(current)];
    depth += 1;
  }
  return depth;
}

export function childOutlineRanges(levels: number[]): Array<{ start: number; end: number }> {
  const ranges: Array<{ start: number; end: number; level: number }> = [];
  for (let index = 0; index < levels.length; index += 1) {
    const level = levels[index] || 1;
    let end = index;
    for (let next = index + 1; next < levels.length; next += 1) {
      if ((levels[next] || 1) <= level) {
        break;
      }
      end = next;
    }
    if (end > index) {
      ranges.push({ start: index + 1, end, level });
    }
  }
  ranges.sort((left, right) => right.level - left.level || left.start - right.start);
  return ranges.map(({ start, end }) => ({ start, end }));
}

export function parentsFromLayout(layout: TreeLayout | null | undefined): Record<string, number> {
  if (!layout) {
    return {};
  }
  const parents: Record<string, number> = {};
  for (const [child, parent] of Object.entries(layout.parentById)) {
    parents[child] = parent;
  }
  return parents;
}

export function parentsEqual(left: number | undefined | null, right: number | undefined | null): boolean {
  return (left || 0) === (right || 0);
}
