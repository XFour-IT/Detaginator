import { describe, expect, it } from "vitest";
import { stripHtmlTags } from "../src/stripHtml";

describe("stripHtmlTags", () => {
  it("removes HTML tags", () => {
    expect(stripHtmlTags("<p>Hello <b>world</b></p>")).toBe("Hello world");
  });
});
