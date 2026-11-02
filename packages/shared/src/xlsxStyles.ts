/**
 * Top-aligns every cell format in an xlsx `styles.xml`, including the Normal style that new cells use.
 * The free SheetJS build can't write cell styles, so the downloaded workbook is patched after it's written.
 */
export function topAlignStylesXml(xml: string): string {
  return xml.replace(
    /<(cellStyleXfs|cellXfs)\b([^>]*)>([\s\S]*?)<\/\1>/g,
    (_match, tag: string, attrs: string, body: string) =>
      `<${tag}${attrs}>${body.replace(
        /<xf\b([^>]*?)\s*\/>/g,
        (_xf, xfAttrs: string) =>
          `<xf${xfAttrs.replace(/\s+applyAlignment="[^"]*"/, "")} applyAlignment="1"><alignment vertical="top"/></xf>`
      )}</${tag}>`
  );
}
