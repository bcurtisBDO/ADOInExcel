import { describe, expect, it } from "vitest";
import { emptyProcessCatalog } from "./columns";
import {
  buildMetadata,
  catalogFields,
  packClassificationPathGrid,
  packMetadataSheet,
  parseMetadataSheet,
  unpackClassificationPathGrid
} from "./workbookMeta";

describe("workbook metadata", () => {
  it("persists classification paths from the process catalog", () => {
    const catalog = {
      ...emptyProcessCatalog(),
      types: ["Bug"],
      statesByType: { Bug: ["New"] },
      areaPaths: ["Fabrikam", "Fabrikam\\Team A"],
      iterationPaths: ["Fabrikam\\Sprint 1"]
    };
    const metadata = buildMetadata({
      orgUrl: "https://dev.azure.com/contoso",
      project: "Fabrikam",
      action: "OpenQuery",
      columns: [{ referenceName: "System.Id", name: "ID" }],
      items: [],
      catalog
    });
    expect(metadata.areaPaths).toEqual(["Fabrikam", "Fabrikam\\Team A"]);
    expect(metadata.iterationPaths).toEqual(["Fabrikam\\Sprint 1"]);
    expect(catalogFields(catalog).areaPaths).toEqual(catalog.areaPaths);
  });

  it("stores classification paths as a grid so they are not limited to one Excel cell", () => {
    const areaPaths = Array.from({ length: 2000 }, (_, index) => `Fabrikam\\Team\\Path ${index}`);
    const iterationPaths = Array.from({ length: 25 }, (_, index) => `Fabrikam\\Sprint ${index}`);
    expect(JSON.stringify({ areaPaths, iterationPaths }).length).toBeGreaterThan(32767);
    const packed = packClassificationPathGrid(areaPaths, iterationPaths);
    expect(packed[0]).toEqual(["areaPaths", "iterationPaths"]);
    expect(packed).toHaveLength(2001);
    expect(unpackClassificationPathGrid(packed)).toEqual({ areaPaths, iterationPaths });
  });

  it("chunks metadata JSON so Excel cells stay under the 32k limit", () => {
    const snapshot: Record<string, Record<string, string>> = {};
    for (let index = 0; index < 400; index += 1) {
      snapshot[String(index)] = {
        "System.Title": `Work item ${index} with a reasonably long title for bulk edit ${"x".repeat(80)}`
      };
    }
    const metadata = buildMetadata({
      orgUrl: "https://dev.azure.com/contoso",
      project: "Fabrikam",
      action: "OpenQuery",
      columns: [{ referenceName: "System.Id", name: "ID" }, { referenceName: "System.Title", name: "Title" }],
      items: []
    });
    metadata.snapshot = snapshot;
    const packed = packMetadataSheet(metadata);
    expect(packed.length).toBeGreaterThan(1);
    expect(packed.every((row) => String(row[0]).length <= 24000 + 5)).toBe(true);
    expect(String(packed[0][0]).startsWith("ADO1:")).toBe(true);
    const parsed = parseMetadataSheet(packed);
    expect(parsed?.project).toBe("Fabrikam");
    expect(Object.keys(parsed?.snapshot || {})).toHaveLength(400);
  });

  it("reads metadata when Excel prefixes or splits the JSON cell", () => {
    const metadata = buildMetadata({
      orgUrl: "https://dev.azure.com/contoso",
      project: "Fabrikam",
      action: "OpenQuery",
      columns: [{ referenceName: "System.Id", name: "ID" }],
      items: []
    });
    const packed = packMetadataSheet(metadata);
    packed[0][0] = `'${packed[0][0]}`;
    packed.splice(1, 0, ["", "areaPaths", ""]);
    expect(parseMetadataSheet(packed)?.project).toBe("Fabrikam");
  });

  it("stops reading JSON at the first empty metadata cell", () => {
    const metadata = buildMetadata({
      orgUrl: "https://dev.azure.com/contoso",
      project: "Fabrikam",
      action: "OpenQuery",
      columns: [{ referenceName: "System.Id", name: "ID" }],
      items: []
    });
    const packed = packMetadataSheet(metadata);
    packed.push(["", "leftover-area", ""]);
    packed.push(["ADO1:{\"not\":\"json that would break parse\"}", "", ""]);
    expect(parseMetadataSheet(packed)?.project).toBe("Fabrikam");
  });

  it("still reads the older single-cell JSON path blob", () => {
    expect(
      unpackClassificationPathGrid([
        ['{"areaPaths":["Fabrikam\\\\Web"],"iterationPaths":["Fabrikam\\\\Sprint 1"]}']
      ])
    ).toEqual({
      areaPaths: ["Fabrikam\\Web"],
      iterationPaths: ["Fabrikam\\Sprint 1"]
    });
  });
});
