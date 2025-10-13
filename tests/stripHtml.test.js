require("ts-node/register");
const { strict: assert } = require("assert");
const { stripHtml } = require("../src/commands/commands.ts");

describe("stripHtml", () => {
  it("replaces &nbsp; within paragraph tags with spaces", () => {
    const input = "<p>Hello&nbsp;World</p>";
    const expected = "Hello World";
    assert.equal(stripHtml(input), expected);
  });

  it("replaces &nbsp; outside paragraph tags with spaces", () => {
    const input = "Hello&nbsp;World";
    const expected = "Hello World";
    assert.equal(stripHtml(input), expected);
  });

  it("converts list items to bullet points", () => {
    const input = "<ul><li>First</li><li>Second</li></ul>";
    const expected = "- First\n- Second";
    assert.equal(stripHtml(input), expected);
  });

  it("removes bold tags", () => {
    const input = "<p>Hello <b>World</b></p>";
    const expected = "Hello World";
    assert.equal(stripHtml(input), expected);
  });

  it("removes italic tags", () => {
    const input = "<p>Hello <i>World</i></p>";
    const expected = "Hello World";
    assert.equal(stripHtml(input), expected);
  });

  it("decodes common HTML entities", () => {
    const input = "&lt;div&gt;Fish &amp; Chips&#39;&lt;/div&gt;";
    const expected = "<div>Fish & Chips'</div>";
    assert.equal(stripHtml(input), expected);
  });
});
