/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Excel, Office, console */

Office.onReady(() => {
  // Office.js is ready.
});

/**
 * Remove HTML tags from the selected range.
 * @param event
 */
async function removeTagsFromSelection(event: Office.AddinCommands.Event) {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      await cleanRange(range);
    });
  } catch (error) {
    console.error(error);
  } finally {
    event.completed();
  }
}

/**
 * Remove HTML tags from the active worksheet.
 * @param event
 */
async function removeTagsFromWorksheet(event: Office.AddinCommands.Event) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.getActiveWorksheet();
      const used = sheet.getUsedRangeOrNullObject();
      await context.sync();
      if (!used.isNullObject) {
        await cleanRange(used);
      }
    });
  } catch (error) {
    console.error(error);
  } finally {
    event.completed();
  }
}

/**
 * Remove HTML tags from all worksheets in the workbook.
 * @param event
 */
async function removeTagsFromDocument(event: Office.AddinCommands.Event) {
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items");
      await context.sync();
      const ranges: Excel.Range[] = [];
      for (const sheet of sheets.items) {
        ranges.push(sheet.getUsedRangeOrNullObject());
      }
      await context.sync();
      for (const range of ranges) {
        if (!range.isNullObject) {
          await cleanRange(range);
        }
      }
    });
  } catch (error) {
    console.error(error);
  } finally {
    event.completed();
  }
}

/**
 * Cleans all cells within a range by stripping HTML tags and keeping text within
 * paragraph tags.
 * @param range
 */
async function cleanRange(range: Excel.Range) {
  range.load("values");
  await range.context.sync();
  const values = range.values as (string | number | boolean)[][];
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const val = values[r][c];
      if (typeof val === "string") {
        const cleaned = stripHtml(val);
        if (cleaned !== val) {
          values[r][c] = cleaned;
        }
      }
    }
  }
  range.values = values;
  await range.context.sync();
}

/**
 * Remove HTML tags from a string, keeping text inside <p> tags when present.
 * @param value
 * @returns cleaned string
 */
function stripHtml(value: string): string {
  if (!/<[^>]+>/i.test(value)) {
    return value;
  }
  const paragraphRegex = /<p[^>]*>(.*?)<\/p>/gi;
  const paragraphs: string[] = [];
  let match: RegExpExecArray | null;
  while ((match = paragraphRegex.exec(value)) !== null) {
    paragraphs.push(match[1]);
  }
  if (paragraphs.length > 0) {
    return paragraphs
      .map((p) => p.replace(/<[^>]+>/g, ""))
      .join("\n")
      .trim();
  }
  return value.replace(/<[^>]+>/g, "");
}

Office.actions.associate("removeTagsFromSelection", removeTagsFromSelection);
Office.actions.associate("removeTagsFromWorksheet", removeTagsFromWorksheet);
Office.actions.associate("removeTagsFromDocument", removeTagsFromDocument);
