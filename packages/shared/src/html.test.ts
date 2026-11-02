import { describe, expect, it } from "vitest";
import { htmlToMarkdown, looksLikeHtml, markdownToHtml } from "./html";

describe("HTML and Markdown", () => {
  it("converts Azure DevOps HTML into editable Markdown", () => {
    expect(htmlToMarkdown("<div>Line 1<br/>Line 2</div>")).toBe("Line 1\nLine 2");
    expect(htmlToMarkdown("<div>Say <b>hello</b> and <i>goodbye</i></div>")).toBe("Say **hello** and _goodbye_");
    expect(htmlToMarkdown('<div><a href="https://contoso">Docs</a></div>')).toBe("[Docs](https://contoso)");
    expect(htmlToMarkdown("<ul><li>One</li><li>Two</li></ul>")).toBe("- One\n- Two");
    expect(looksLikeHtml("<div>Hello</div>")).toBe(true);
    expect(looksLikeHtml("plain text")).toBe(false);
  });

  it("converts Markdown back to Azure DevOps HTML", () => {
    expect(markdownToHtml("Line 1\nLine 2")).toBe("<div>Line 1</div><div>Line 2</div>");
    expect(markdownToHtml("Say **hello** and *goodbye*")).toBe("<div>Say <b>hello</b> and <i>goodbye</i></div>");
    expect(markdownToHtml("Say **hello** and _goodbye_")).toBe("<div>Say <b>hello</b> and <i>goodbye</i></div>");
    expect(markdownToHtml("[Docs](https://contoso)")).toBe('<div><a href="https://contoso">Docs</a></div>');
    expect(markdownToHtml("- One\n- Two")).toBe("<ul><li>One</li><li>Two</li></ul>");
  });

  it("leaves existing HTML unchanged on publish", () => {
    expect(markdownToHtml("<div>Already HTML</div>")).toBe("<div>Already HTML</div>");
  });

  it("round-trips nested bold and italic without dropping either style", () => {
    expect(htmlToMarkdown("<b><i>bold italic</i></b>")).toBe("**_bold italic_**");
    expect(htmlToMarkdown("<i><b>bold italic</b></i>")).toBe("_**bold italic**_");
    expect(markdownToHtml("**_bold italic_**")).toBe("<div><b><i>bold italic</i></b></div>");
    expect(markdownToHtml("_**bold italic**_")).toBe("<div><i><b>bold italic</b></i></div>");
    expect(markdownToHtml("***bold italic***")).toBe("<div><b><i>bold italic</i></b></div>");
    expect(markdownToHtml(htmlToMarkdown("<div><b><i>bold italic</i></b></div>"))).toBe(
      "<div><b><i>bold italic</i></b></div>"
    );
  });
});
