import { describe, expect, it } from "vitest";
import { ATTACHMENT_LINK, PARENT_LINK_REVERSE, RELATED_LINK, SUCCESSOR_LINK } from "./constants";
import { workItemWebUrl } from "./fields";
import { commentHtml, withRelationNames, workItemDetailsFromSources, workItemLinksFromRelations } from "./workItemDetails";

describe("work item web URL", () => {
  it("builds the Azure DevOps edit URL", () => {
    expect(workItemWebUrl("https://dev.azure.com/org/", "My Project", 42)).toBe(
      "https://dev.azure.com/org/My%20Project/_workitems/edit/42"
    );
  });
});

describe("work item links and comments", () => {
  it("splits related links from attachments and skips parent/child", () => {
    const { links, attachments } = workItemLinksFromRelations(
      [
        { rel: PARENT_LINK_REVERSE, url: "https://dev.azure.com/org/_apis/wit/workItems/1" },
        { rel: RELATED_LINK, url: "https://dev.azure.com/org/_apis/wit/workItems/9" },
        { rel: SUCCESSOR_LINK, url: "https://dev.azure.com/org/_apis/wit/workItems/8" },
        {
          rel: ATTACHMENT_LINK,
          url: "https://dev.azure.com/org/_apis/wit/attachments/abc",
          attributes: { name: "spec.pdf" }
        },
        { rel: "Hyperlink", url: "https://example.com/doc" }
      ],
      "https://dev.azure.com/org",
      "Proj"
    );
    expect(links.map((item) => item.label)).toEqual(["Related", "Successor", "Hyperlink"]);
    expect(links[0].id).toBe(9);
    expect(links[0].url).toBe("https://dev.azure.com/org/Proj/_workitems/edit/9");
    expect(links[2].url).toBe("https://example.com/doc");
    expect(attachments).toEqual([
      expect.objectContaining({
        label: "spec.pdf",
        url: "https://dev.azure.com/org/_apis/wit/attachments/abc"
      })
    ]);
  });

  it("labels other link types and shows hyperlink comments, artifact names, and remote work items", () => {
    const { links } = workItemLinksFromRelations(
      [
        { rel: "Microsoft.VSTS.Common.TestedBy-Forward", url: "https://dev.azure.com/org/_apis/wit/workItems/30" },
        { rel: "Hyperlink", url: "https://example.com/spec", attributes: { comment: "Design spec" } },
        { rel: "ArtifactLink", url: "vstfs:///Git/Commit/abc", attributes: { name: "Fixed in Commit" } },
        { rel: "System.LinkTypes.Remote.Related", url: "https://dev.azure.com/other/_apis/wit/workItems/77" }
      ],
      "https://dev.azure.com/org",
      "Proj",
      { "System.LinkTypes.Remote.Related": "Remote Related" }
    );
    expect(links.map((link) => link.label)).toEqual(["Tested By", "Hyperlink", "Artifact", "Remote Related"]);
    expect(links[0].url).toBe("https://dev.azure.com/org/Proj/_workitems/edit/30");
    expect(links[1].name).toBe("Design spec");
    expect(links[2]).toEqual(expect.objectContaining({ name: "Fixed in Commit", url: "vstfs:///Git/Commit/abc" }));
    expect(links[3]).toEqual(expect.objectContaining({ id: 77, url: "https://dev.azure.com/other/_workitems/edit/77" }));
  });

  it("relabels loaded links with organization names", () => {
    const details = workItemDetailsFromSources(
      { id: 5, relations: [{ rel: "Contoso.LinkTypes.Blocks-Forward", url: "https://dev.azure.com/org/_apis/wit/workItems/6" }] },
      "https://dev.azure.com/org",
      "Proj",
      []
    );
    expect(details.links[0].label).toBe("Blocks");
    expect(withRelationNames(details, { "Contoso.LinkTypes.Blocks-Forward": "Blocks Release" }).links[0].label).toBe(
      "Blocks Release"
    );
  });

  it("builds details from a work item and comments", () => {
    const details = workItemDetailsFromSources(
      {
        id: 5,
        fields: { "System.Title": "Fix login", "System.Parent": 1 },
        relations: [{ rel: RELATED_LINK, url: "https://dev.azure.com/org/_apis/wit/workItems/6" }]
      },
      "https://dev.azure.com/org",
      "Proj",
      [{ id: 1, text: "Looks good", createdBy: "Ada", createdDate: "2026-09-14" }],
      { id: 1, title: "Login epic" }
    );
    expect(details.webUrl).toBe("https://dev.azure.com/org/Proj/_workitems/edit/5");
    expect(details.title).toBe("Fix login");
    expect(details.parentId).toBe(1);
    expect(details.parentTitle).toBe("Login epic");
    expect(details.parentUrl).toBe("https://dev.azure.com/org/Proj/_workitems/edit/1");
    expect(details.links).toHaveLength(1);
    expect(details.comments[0].text).toBe("Looks good");
  });

  it("wraps comments as escaped HTML", () => {
    expect(commentHtml('Say <b>hi</b> & "ok"')).toBe(
      "<div>Say &lt;b&gt;hi&lt;/b&gt; &amp; &quot;ok&quot;</div>"
    );
    expect(commentHtml("line1\nline2")).toBe("<div>line1<br/>line2</div>");
  });
});
