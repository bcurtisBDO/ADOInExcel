import { describe, expect, it } from "vitest";
import { topAlignStylesXml } from "./xlsxStyles";

const SHEETJS_STYLES =
  '<styleSheet><fonts count="1"><font><sz val="12"/></font></fonts>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
  '<xf numFmtId="60" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="0"/></cellXfs>' +
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

describe("topAlignStylesXml", () => {
  it("top-aligns the Normal style and every cell format, including date formats", () => {
    const xml = topAlignStylesXml(SHEETJS_STYLES);
    expect(xml).toContain(
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment vertical="top"/></xf></cellStyleXfs>'
    );
    expect(xml).toContain(
      '<xf numFmtId="60" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment vertical="top"/></xf></cellXfs>'
    );
    expect(xml.match(/<alignment vertical="top"\/>/g)).toHaveLength(3);
  });

  it("leaves fonts and named cell styles alone", () => {
    const xml = topAlignStylesXml(SHEETJS_STYLES);
    expect(xml).toContain('<fonts count="1"><font><sz val="12"/></font></fonts>');
    expect(xml).toContain('<cellStyle name="Normal" xfId="0" builtinId="0"/>');
  });
});
