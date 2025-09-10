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
});
