import { isAllowedPath, isAllowedPicklistValue, isAllowedState, isAllowedTransition, isAllowedWorkItemType } from "./columns";
import { dateToAdoValue, isDateTimeField, isIsoDateString, isReadOnlyField, parseWorkItemId, valuesEqual } from "./fields";
import { isHtmlField, markdownToHtml } from "./html";
import { inferTreeParents, parentsEqual, validateTree } from "./tree";
import { JsonPatchOperation, PublishChange, PublishPlan, TableRow, WorkbookMetadata } from "./types";

function isEmptyRow(row: TableRow): boolean {
  const title = String(row.values["System.Title"] ?? "").trim();
  const type = String(row.values["System.WorkItemType"] ?? "").trim();
  if (title || type || row.id) {
    return false;
  }
  return Object.entries(row.values).every(([key, value]) => {
    if (key === "System.Id" || key === "System.Rev") {
      return true;
    }
    return value === "" || value === undefined;
  });
}

function toPatchValue(value: unknown, referenceName: string, metadata?: WorkbookMetadata): unknown {
  if (value === "") {
    return null;
  }
  if (isDateTimeField(referenceName) || isIsoDateString(value)) {
    return dateToAdoValue(value);
  }
  if (typeof value === "string" && isHtmlField(referenceName, metadata?.htmlFields)) {
    return markdownToHtml(value);
  }
  return value;
}

export function operationsForRow(
  values: TableRow["values"],
  original: Record<string, unknown> | undefined,
  isCreate: boolean,
  metadata?: WorkbookMetadata
): JsonPatchOperation[] {
  const operations: JsonPatchOperation[] = [];

  for (const [referenceName, value] of Object.entries(values)) {
    if (isReadOnlyField(referenceName) || referenceName.startsWith("__")) {
      continue;
    }
    if (isCreate && (value === "" || value === undefined)) {
      continue;
    }

    const previous = original ? original[referenceName] : undefined;
    if (!isCreate && valuesEqual(previous, value, referenceName)) {
      continue;
    }

    operations.push({
      op: "add",
      path: `/fields/${referenceName}`,
      value: toPatchValue(value, referenceName, metadata)
    });
  }

  return operations;
}

export function buildPublishPlan(rows: TableRow[], metadata: WorkbookMetadata): PublishPlan {
  const isTree = metadata.listType === "tree" || (metadata.treeLevels || 0) > 1;
  const withParents = isTree ? inferTreeParents(rows) : rows;
  const errors = validateTree(withParents, isTree);
  const changes: PublishChange[] = [];
  let updateCount = 0;
  let createCount = 0;
  let skipCount = 0;

  for (const row of withParents) {
    const id = row.id ?? parseWorkItemId(row.values["System.Id"]);
    const workItemType = String(row.values["System.WorkItemType"] ?? "").trim();
    const title = String(row.values["System.Title"] ?? "").trim();
    const previousParentId = id ? metadata.parents[String(id)] : undefined;
    const parentChanged = isTree && !parentsEqual(previousParentId, row.parentId);
    const parentFields = isTree
      ? {
          desiredParentId: row.parentId ?? null,
          desiredParentRowIndex: row.parentRowIndex,
          previousParentId: previousParentId ?? null
        }
      : {};

    if (!id) {
      if (isEmptyRow(row)) {
        skipCount += 1;
        changes.push({
          kind: "skip",
          rowIndex: row.rowIndex,
          message: "Empty row"
        });
        continue;
      }

      if (!workItemType || !title) {
        changes.push({
          kind: "error",
          rowIndex: row.rowIndex,
          workItemType,
          message: "New work items require Work Item Type and Title.",
          ...parentFields
        });
        continue;
      }

      if (!isAllowedWorkItemType(workItemType, metadata.workItemTypes)) {
        changes.push({
          kind: "error",
          rowIndex: row.rowIndex,
          workItemType,
          message: `"${workItemType}" is not a work item type in this project's process template.`,
          ...parentFields
        });
        continue;
      }

      const state = String(row.values["System.State"] ?? "").trim();
      if (!isAllowedState(state, workItemType, metadata.statesByType)) {
        changes.push({
          kind: "error",
          rowIndex: row.rowIndex,
          workItemType,
          message: `"${state}" is not a valid state for ${workItemType}.`,
          ...parentFields
        });
        continue;
      }

      const processError = processFieldError(row, metadata, workItemType, "");
      if (processError) {
        changes.push({
          kind: "error",
          rowIndex: row.rowIndex,
          workItemType,
          message: processError,
          ...parentFields
        });
        continue;
      }

      const operations = operationsForRow(row.values, undefined, true, metadata);
      createCount += 1;
      changes.push({
        kind: "create",
        rowIndex: row.rowIndex,
        workItemType,
        operations,
        message: `Create ${workItemType}`,
        ...parentFields
      });
      continue;
    }

    const original = metadata.snapshot[String(id)];
    const loadedType = original ? String(original["System.WorkItemType"] ?? "").trim() : "";
    if (loadedType && workItemType.toLowerCase() !== loadedType.toLowerCase()) {
      changes.push({
        kind: "error",
        rowIndex: row.rowIndex,
        id,
        workItemType,
        message: `Work item ${id} is a ${loadedType}, but the row says "${workItemType}". Changing the work item type isn't supported from Excel. Change it in Azure DevOps, or set the cell back to ${loadedType}.`,
        ...parentFields
      });
      continue;
    }

    const state = String(row.values["System.State"] ?? "").trim();
    if (workItemType && !isAllowedState(state, workItemType, metadata.statesByType)) {
      changes.push({
        kind: "error",
        rowIndex: row.rowIndex,
        id,
        workItemType,
        message: `"${state}" is not a valid state for ${workItemType}.`,
        ...parentFields
      });
      continue;
    }

    const previousState = original ? String(original["System.State"] ?? "").trim() : "";
    const processError = processFieldError(row, metadata, workItemType, previousState);
    if (processError) {
      changes.push({
        kind: "error",
        rowIndex: row.rowIndex,
        id,
        workItemType,
        message: processError,
        ...parentFields
      });
      continue;
    }

    const operations = operationsForRow(row.values, original, false, metadata);
    if (operations.length === 0 && !parentChanged) {
      skipCount += 1;
      changes.push({
        kind: "skip",
        rowIndex: row.rowIndex,
        id,
        message: "No changes"
      });
      continue;
    }

    updateCount += 1;
    changes.push({
      kind: "update",
      rowIndex: row.rowIndex,
      id,
      operations,
      message: parentChanged && operations.length === 0
        ? "Update parent"
        : `Update ${operations.length} field(s)${parentChanged ? " and parent" : ""}`,
      ...parentFields
    });
  }

  return { changes, updateCount, createCount, skipCount, errors };
}

export function applyRevisionConflicts(
  plan: PublishPlan,
  metadata: WorkbookMetadata,
  serverRevisions: Record<string, number>
): PublishPlan {
  const changes = plan.changes.map((change) => {
    if (change.kind !== "update" || !change.id) {
      return change;
    }

    const localRev = metadata.revisions[String(change.id)];
    const serverRev = serverRevisions[String(change.id)];
    if (localRev !== undefined && serverRev !== undefined && localRev !== serverRev) {
      return {
        ...change,
        kind: "conflict" as const,
        operations: undefined,
        message: `Conflict on work item ${change.id}: Excel has rev ${localRev}, Azure DevOps has rev ${serverRev}. Refresh, then re-apply your edits.`
      };
    }

    return change;
  });

  return {
    ...plan,
    changes,
    updateCount: changes.filter((change) => change.kind === "update").length
  };
}

export function parentUrl(orgUrl: string, parentId: number): string {
  return `${orgUrl.replace(/\/+$/, "")}/_apis/wit/workItems/${parentId}`;
}

function processFieldError(
  row: TableRow,
  metadata: WorkbookMetadata,
  workItemType: string,
  previousState: string
): string | undefined {
  const state = String(row.values["System.State"] ?? "").trim();
  if (workItemType && !isAllowedTransition(previousState, state, workItemType, metadata.transitionsByType)) {
    return `"${state}" is not a valid next state from ${previousState || "a new work item"} for ${workItemType}.`;
  }

  const area = String(row.values["System.AreaPath"] ?? "").trim();
  if (!isAllowedPath(area, metadata.areaPaths)) {
    return `"${area}" is not an area path in this project.`;
  }
  const iteration = String(row.values["System.IterationPath"] ?? "").trim();
  if (!isAllowedPath(iteration, metadata.iterationPaths)) {
    return `"${iteration}" is not an iteration path in this project.`;
  }

  const suggested = new Set((metadata.suggestedPicklists || []).map((name) => name.toLowerCase()));
  for (const [referenceName, raw] of Object.entries(row.values)) {
    if (referenceName === "System.State" || referenceName === "System.WorkItemType") {
      continue;
    }
    const allowed = metadata.picklists?.[referenceName];
    if (!allowed || allowed.length === 0) {
      continue;
    }
    if (
      !isAllowedPicklistValue(
        String(raw ?? ""),
        allowed,
        suggested.has(referenceName.toLowerCase())
      )
    ) {
      return `"${String(raw ?? "").trim()}" is not a valid value for ${referenceName}.`;
    }
  }
  return undefined;
}
