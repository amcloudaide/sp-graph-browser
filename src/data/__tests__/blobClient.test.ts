import { describe, it, expect } from "vitest";
import { resolveDscUrl } from "../blobClient";

describe("resolveDscUrl", () => {
  it("resolves PowerShell expression to actual URL", () => {
    const raw = "https://$($OrganizationName.Split('.')[0]).sharepoint.com/sites/IT";
    const resolved = resolveDscUrl(raw, "myehn");
    expect(resolved).toBe("https://myehn.sharepoint.com/sites/IT");
  });

  it("returns plain URLs unchanged", () => {
    const url = "https://myehn.sharepoint.com/sites/IT";
    expect(resolveDscUrl(url, "myehn")).toBe(url);
  });

  it("handles root site URL", () => {
    const raw = "https://$($OrganizationName.Split('.')[0]).sharepoint.com/";
    expect(resolveDscUrl(raw, "myehn")).toBe("https://myehn.sharepoint.com/");
  });
});
