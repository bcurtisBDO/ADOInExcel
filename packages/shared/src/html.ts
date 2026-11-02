const HTML_MARKERS =
  /<\/?(div|p|br|ul|ol|li|b|i|strong|em|span|a|table|tr|td|th|h[1-6]|font|blockquote|pre|code)\b/i;

export const KNOWN_HTML_FIELDS = new Set([
  "System.Description",
  "Microsoft.VSTS.TCM.ReproSteps",
  "Microsoft.VSTS.TCM.SystemInfo",
  "Microsoft.VSTS.TCM.Steps"
]);

export function looksLikeHtml(value: string): boolean {
  return HTML_MARKERS.test(value);
}

export function isHtmlField(referenceName?: string, htmlFields?: string[]): boolean {
  if (!referenceName) {
    return false;
  }
  if (htmlFields && htmlFields.length > 0) {
    return htmlFields.some((name) => name.toLowerCase() === referenceName.toLowerCase());
  }
  return KNOWN_HTML_FIELDS.has(referenceName);
}

export function htmlToMarkdown(html: string): string {
  let text = html.replace(/\r\n/g, "\n").replace(/<!--[\s\S]*?-->/g, "");
  text = text.replace(/<br\s*\/?>/gi, "\n");
  text = text.replace(/<h[1-6][^>]*>([\s\S]*?)<\/h[1-6]>/gi, (_match, inner: string) => `\n## ${inner}\n`);
  text = text.replace(
    /<a\s+[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi,
    (_match, href: string, inner: string) => `[${stripTags(inner)}](${href})`
  );
  text = text.replace(/<code[^>]*>([\s\S]*?)<\/code>/gi, (_match, inner: string) => `\`${stripTags(inner)}\``);
  text = replaceInnermostEmphasis(text);
  text = text.replace(/<li[^>]*>\s*/gi, "\n- ");
  text = text.replace(/<\/li>/gi, "");
  text = text.replace(/<\/?(ul|ol)[^>]*>/gi, "\n");
  text = text.replace(/<\/(div|p|tr|blockquote)>/gi, "\n");
  text = text.replace(/<(div|p|tr|blockquote|span|font)[^>]*>/gi, "");
  text = decodeEntities(stripTags(text));
  return text.replace(/\u00a0/g, " ").replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
}

function replaceInnermostEmphasis(html: string): string {
  const innermost = /<(strong|b|em|i)(?:\s[^>]*)?>([^<]*)<\/\1>/gi;
  let current = html;
  for (let pass = 0; pass < 32; pass += 1) {
    const next = current.replace(innermost, (_match, tag: string, inner: string) => {
      const name = String(tag).toLowerCase();
      if (name === "strong" || name === "b") {
        return `**${inner}**`;
      }
      return `_${inner}_`;
    });
    if (next === current) {
      break;
    }
    current = next;
  }
  return current;
}

export function markdownToHtml(markdown: string): string {
  const trimmed = markdown.replace(/\r\n/g, "\n").trim();
  if (!trimmed) {
    return "";
  }
  if (looksLikeHtml(trimmed)) {
    return trimmed;
  }

  const blocks: string[] = [];
  let list: { kind: "ul" | "ol"; items: string[] } | undefined;

  const flushList = () => {
    if (!list) {
      return;
    }
    const tag = list.kind;
    blocks.push(`<${tag}>${list.items.map((item) => `<li>${inlineMarkdown(item)}</li>`).join("")}</${tag}>`);
    list = undefined;
  };

  for (const line of trimmed.split("\n")) {
    const unordered = line.match(/^\s*[-*]\s+(.*)$/);
    const ordered = line.match(/^\s*\d+\.\s+(.*)$/);
    const heading = line.match(/^#{1,6}\s+(.*)$/);
    if (unordered) {
      if (!list || list.kind !== "ul") {
        flushList();
        list = { kind: "ul", items: [] };
      }
      list.items.push(unordered[1]);
      continue;
    }
    if (ordered) {
      if (!list || list.kind !== "ol") {
        flushList();
        list = { kind: "ol", items: [] };
      }
      list.items.push(ordered[1]);
      continue;
    }
    flushList();
    if (!line.trim()) {
      continue;
    }
    if (heading) {
      blocks.push(`<div><b>${inlineMarkdown(heading[1])}</b></div>`);
      continue;
    }
    blocks.push(`<div>${inlineMarkdown(line)}</div>`);
  }
  flushList();
  return blocks.join("");
}

function inlineMarkdown(text: string): string {
  let value = escapeHtml(text);
  value = value.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2">$1</a>');
  value = value.replace(/`([^`]+)`/g, "<code>$1</code>");
  value = value.replace(/\*\*\*([^*]+)\*\*\*/g, "<b><i>$1</i></b>");
  value = value.replace(/___([^_]+)___/g, "<b><i>$1</i></b>");
  value = value.replace(/\*\*([^*]+)\*\*/g, "<b>$1</b>");
  value = value.replace(/__([^_]+)__/g, "<b>$1</b>");
  value = value.replace(/_([^_]+)_/g, "<i>$1</i>");
  value = value.replace(/\*([^*]+)\*/g, "<i>$1</i>");
  return value;
}

function stripTags(value: string): string {
  return value.replace(/<[^>]+>/g, "");
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function decodeEntities(value: string): string {
  return value
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/g, "'");
}
