import { describe, it, expect } from "vitest";
import { toCsv } from "../csvExporter";

describe("toCsv", () => {
  it("converts array of objects to CSV string", () => {
    const data = [
      { name: "List A", itemCount: 10 },
      { name: "List B", itemCount: 20 },
    ];
    const csv = toCsv(data);
    expect(csv).toBe("name,itemCount\nList A,10\nList B,20");
  });

  it("escapes commas and quotes", () => {
    const data = [{ name: 'Hello, "World"', value: "ok" }];
    const csv = toCsv(data);
    expect(csv).toBe('name,value\n"Hello, ""World""",ok');
  });

  it("returns empty string for empty array", () => {
    expect(toCsv([])).toBe("");
  });
});
