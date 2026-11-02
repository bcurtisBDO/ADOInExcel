import { ATTACHMENT_LINK, HYPERLINK_LINK, REMOTE_LINK_PREFIX } from "./constants";
import { parentIdFromRelations, stripHtml, workItemIdFromRelationUrl, workItemWebUrl } from "./fields";
import { isHierarchyLink, relationLabel } from "./linkTypes";
import {
  WorkItemCommentInfo,
  WorkItemDetails,
  WorkItemLike,
  WorkItemLinkInfo,
  WorkItemRelation
} from "./types";

function attributeText(relation: WorkItemRelation, key: string): string | undefined {
  const value = relation.attributes?.[key];
  return typeof value === "string" && value.trim() ? value.trim() : undefined;
}

/** Remote links point at another organization's API, so the web URL is built from that organization. */
function remoteWorkItemWebUrl(apiUrl: string, id: number): string {
  const base = apiUrl.split(/\/_apis\//i)[0];
  return `${base}/_workitems/edit/${id}`;
}

export function workItemLinksFromRelations(
  relations: WorkItemRelation[] | undefined,
  orgUrl: string,
  project: string,
  names?: Record<string, string>
): { links: WorkItemLinkInfo[]; attachments: WorkItemLinkInfo[] } {
  const links: WorkItemLinkInfo[] = [];
  const attachments: WorkItemLinkInfo[] = [];
  (relations || []).forEach((relation, index) => {
    const rel = String(relation.rel || "");
    const url = String(relation.url || "");
    if (!rel || !url || isHierarchyLink(rel)) {
      return;
    }
    if (rel === ATTACHMENT_LINK) {
      const name = attributeText(relation, "name");
      attachments.push({ index, rel, label: name || "Attachment", url, name });
      return;
    }
    if (rel === HYPERLINK_LINK) {
      links.push({ index, rel, label: relationLabel(rel, names), url, name: attributeText(relation, "comment") || url });
      return;
    }
    const id = workItemIdFromRelationUrl(url);
    let webUrl = url;
    if (id) {
      webUrl = rel.startsWith(REMOTE_LINK_PREFIX) ? remoteWorkItemWebUrl(url, id) : workItemWebUrl(orgUrl, project, id);
    }
    links.push({ index, rel, label: relationLabel(rel, names), url: webUrl, id, name: attributeText(relation, "name") });
  });
  return { links, attachments };
}

/** Relabels links with the relation names Azure DevOps returned for this organization. */
export function withRelationNames(details: WorkItemDetails, names: Record<string, string>): WorkItemDetails {
  return {
    ...details,
    links: details.links.map((link) => ({ ...link, label: relationLabel(link.rel, names) }))
  };
}

export function mapWorkItemComments(
  raw: Array<{ id?: number; text?: string; createdBy?: { displayName?: string }; createdDate?: string }>
): WorkItemCommentInfo[] {
  return raw
    .map((comment) => ({
      id: Number(comment.id) || 0,
      text: stripHtml(String(comment.text || "")).trim(),
      createdBy: String(comment.createdBy?.displayName || "").trim() || "Unknown",
      createdDate: String(comment.createdDate || "")
    }))
    .filter((comment) => comment.id && comment.text);
}

export function workItemDetailsFromSources(
  item: WorkItemLike,
  orgUrl: string,
  project: string,
  comments: WorkItemCommentInfo[],
  parent?: { id: number; title: string }
): WorkItemDetails {
  const id = item.id || 0;
  const { links, attachments } = workItemLinksFromRelations(item.relations, orgUrl, project);
  const parentId = parent?.id || parentIdFromRelations(item);
  return {
    id,
    rev: typeof item.rev === "number" ? item.rev : undefined,
    title: String(item.fields?.["System.Title"] || "").trim(),
    webUrl: workItemWebUrl(orgUrl, project, id),
    parentId: parentId || undefined,
    parentTitle: parent?.title || "",
    parentUrl: parentId ? workItemWebUrl(orgUrl, project, parentId) : undefined,
    links,
    attachments,
    comments
  };
}

export function commentHtml(text: string): string {
  const escaped = text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
  return `<div>${escaped.replace(/\n/g, "<br/>")}</div>`;
}
