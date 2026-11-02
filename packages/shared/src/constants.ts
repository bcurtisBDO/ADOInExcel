export const WORKBOOK_METADATA_VERSION = 2;
export const WORK_ITEMS_SHEET = "Work Items";
export const META_SHEET = "_AdoOpenInExcel";
export const TABLE_NAME = "AdoWorkItems";
export const PICKLIST_SHEET = "_AdoPickLists";
export const ADO_API_VERSION = "7.1";
export const WORK_ITEM_BATCH_SIZE = 200;
export const MAX_WORK_ITEMS = 10000;
export const MIN_TREE_LEVELS = 2;
export const MAX_TREE_LEVELS = 15;
export const PARENT_LINK_REVERSE = "System.LinkTypes.Hierarchy-Reverse";
export const PARENT_LINK_FORWARD = "System.LinkTypes.Hierarchy-Forward";
export const RELATED_LINK = "System.LinkTypes.Related";
export const PREDECESSOR_LINK = "System.LinkTypes.Dependency-Reverse";
export const SUCCESSOR_LINK = "System.LinkTypes.Dependency-Forward";
export const ATTACHMENT_LINK = "AttachedFile";
export const HYPERLINK_LINK = "Hyperlink";
export const ARTIFACT_LINK = "ArtifactLink";
export const REMOTE_LINK_PREFIX = "System.LinkTypes.Remote.";

/** Synthetic display column: parent work item title (resolved client-side; not an ADO field). */
export const PARENT_TITLE_FIELD = "Ado.ParentTitle";
export const PARENT_TITLE_COLUMN = { referenceName: PARENT_TITLE_FIELD, name: "Parent Title" } as const;

export const DEFAULT_COLUMNS: Array<{ referenceName: string; name: string }> = [
  { referenceName: "System.Id", name: "ID" },
  { referenceName: "System.WorkItemType", name: "Work Item Type" },
  { referenceName: "System.Title", name: "Title" },
  { referenceName: "System.State", name: "State" },
  { referenceName: "System.AssignedTo", name: "Assigned To" },
  { referenceName: "System.AreaPath", name: "Area Path" },
  { referenceName: "System.IterationPath", name: "Iteration Path" },
  { referenceName: "System.Tags", name: "Tags" }
];

export const REQUIRED_FIELDS = ["System.Id", "System.Rev", "System.WorkItemType", "System.Title"];

export const LOCKED_COLUMN_FIELDS = new Set(["System.Id", "System.Rev", "System.WorkItemType", "System.Title"]);

export const CHOOSER_EXCLUDED_FIELDS = new Set([
  "System.History",
  "System.ExternalLinkCount",
  "System.HyperLinkCount",
  "System.AttachedFileCount",
  "System.RelatedLinkCount",
  "System.RemoteLinkCount",
  "System.CommentCount",
  "System.Parent",
  "System.Watermark",
  "System.AuthorizedAs",
  "System.AuthorizedDate",
  "System.RevisedDate",
  "System.BoardColumn",
  "System.BoardColumnDone",
  "System.BoardLane"
]);

export const IDENTITY_FIELDS = new Set(["System.AssignedTo"]);

export const TREE_PATH_FIELDS = new Set(["System.AreaPath", "System.IterationPath"]);

export const MAX_CLASSIFICATION_PATHS = 5000;
export const MAX_IDENTITY_PICKLIST = 1000;

export const READ_ONLY_FIELDS = new Set([
  "System.Id",
  "System.Rev",
  "System.RevisedDate",
  "System.ChangedDate",
  "System.ChangedBy",
  "System.CreatedDate",
  "System.CreatedBy",
  "System.AuthorizedAs",
  "System.AuthorizedDate",
  "System.Watermark",
  "System.CommentCount",
  "System.Parent",
  PARENT_TITLE_FIELD,
  "System.BoardColumn",
  "System.BoardColumnDone",
  "System.BoardLane",
  "System.WorkItemType",
  "Microsoft.VSTS.Common.StateChangeDate",
  "Microsoft.VSTS.Common.ActivatedDate",
  "Microsoft.VSTS.Common.ClosedDate",
  "Microsoft.VSTS.Common.ResolvedDate"
]);

export namespace WellKnownQueries {
  export const AssignedToMe = "A2108D31-086C-4FB0-AFDA-097E4CC46DF4";
  export const UnsavedWorkItems = "B7A26A56-EA87-4C97-A504-3F028808BB9F";
  export const FollowedWorkItems = "202230E0-821E-401D-96D1-24A7202330D0";
  export const CreatedBy = "53FB153F-C52C-42F1-90B6-CA17FC3561A8";
  export const SearchResults = "2CBF5136-1AE5-4948-B59A-36F526D9AC73";
  export const CustomWiql = "08E20883-D56C-4461-88EB-CE77C0C7936D";
  export const RecycleBin = "2650C586-0DE4-4156-BA0E-14BCFB664CCA";
}

export const QUERY_EXCLUSION_LIST = [
  WellKnownQueries.AssignedToMe,
  WellKnownQueries.UnsavedWorkItems,
  WellKnownQueries.FollowedWorkItems,
  WellKnownQueries.CreatedBy,
  WellKnownQueries.SearchResults,
  WellKnownQueries.CustomWiql,
  WellKnownQueries.RecycleBin
];

export function isSupportedQueryId(queryId: string | undefined | null): boolean {
  return Boolean(queryId && QUERY_EXCLUSION_LIST.indexOf(queryId.toUpperCase()) === -1);
}
