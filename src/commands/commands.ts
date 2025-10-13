/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Excel, Office, console */

if (typeof Office !== "undefined") {
  Office.onReady(() => {
    // Office.js is ready.
  });
}

/**
 * Remove HTML tags from the selected range.
 * @param event - the event object provided by the Office runtime
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
 * @param event - the event object provided by the Office runtime
 */
async function removeTagsFromWorksheet(event: Office.AddinCommands.Event) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
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
 * @param event - the event object provided by the Office runtime
 */
async function removeTagsFromWorkbook(event: Office.AddinCommands.Event) {
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

// /** Cleans all cells in a workbook by stripping HTML tags and keeping text within
//  * paragraph tags.
//  * Calls the cleanSheet function for each sheet in the workbook.
//  * @param workbook - the Excel workbook to clean
//  */
// function cleanWorkbook(workbook: Excel.Workbook) {
//   const sheets = workbook.worksheets;
//   sheets.load("items");
//   return sheets.context.sync().then(() => {
//     const promises: Promise<void>[] = [];
//     for (const sheet of sheets.items) {
//       promises.push(cleanSheet(sheet));
//     }
//     return Promise.all(promises).then(() => {});
//   });
// }

// /**
//  * Cleans all cells in a worksheet by stripping HTML tags and keeping text within
//  * paragraph tags.
//  * Calls the cleanRange function for the used range of the sheet.
//  * @param sheet - the Excel worksheet to clean
//  */
// function cleanSheet(sheet: Excel.Worksheet) {
//   const usedRange = sheet.getUsedRangeOrNullObject();
//   return usedRange.load("address").context.sync().then(() => {
//     if (!usedRange.isNullObject) {
//       return cleanRange(usedRange);
//     }
//   });
// }

/**
 * Cleans all cells within a range by stripping HTML tags and keeping text within
 * paragraph tags.
 * @param range - the Excel range to clean
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
 * Remove HTML tags from a string while keeping a basic representation of the
 * formatting. Supports common block tags like <p>, <ul>, <ol>, and <li>, as
 * well as inline formatting tags such as <b>, <strong>, <i>, and <em>.
 *
 * Block level tags are converted to newlines and list items are prefixed with
 * "- ". Inline formatting is removed rather than represented with text
 * markers so that the result is plain text suitable for Excel.
 *
 * @param value - the string to clean
 * @returns cleaned string with minimal formatting markers
 */
const htmlEntityMap: Record<string, string> = {
  "&lt;": "<",
  "&gt;": ">",
  "&amp;": "&",
  "&quot;": "\"",
  "&apos;": "'",
  "&#39;": "'",
};

function decodeHtmlEntities(text: string): string {
  let decoded = text.replace(/&(lt|gt|amp|quot|apos|#39);/gi, (entity) => {
    const lower = entity.toLowerCase();
    return htmlEntityMap[lower] ?? entity;
  });

  decoded = decoded.replace(/&#(\d+);/g, (_match, code) => {
    const charCode = Number.parseInt(code, 10);
    return Number.isNaN(charCode) ? _match : String.fromCharCode(charCode);
  });

  decoded = decoded.replace(/&#x([0-9a-f]+);/gi, (_match, code) => {
    const charCode = Number.parseInt(code, 16);
    return Number.isNaN(charCode) ? _match : String.fromCharCode(charCode);
  });

  return decoded;
}

export function stripHtml(value: string): string {
  const nbspRegex = /&nbsp;/gi;

  let text = value;

  // Replace non-breaking spaces early so inner recursive calls don't need to
  // handle them separately.
  text = text.replace(nbspRegex, " ");

  // Convert list items to lines prefixed with "- "
  text = text.replace(/<li[^>]*>([\s\S]*?)<\/li>/gi, (_match, inner) => `\n- ${stripHtml(inner)}`);

  // Remove list containers but keep their content (already handled above).
  text = text.replace(/<\/?(ul|ol)[^>]*>/gi, "");

  // Convert paragraph and div tags to newlines.
  text = text.replace(/<\/?(p|div)[^>]*>/gi, "\n");

  // Convert <br> tags to newlines.
  text = text.replace(/<br\s*\/?\s*>/gi, "\n");

  // Remove any remaining tags.
  text = text.replace(/<[^>]+>/g, "");

  // Decode HTML entities to their literal counterparts.
  text = decodeHtmlEntities(text);

  // Normalize whitespace around newlines and trim the result.
  text = text
    .split("\n")
    .map((line) => line.trim())
    .filter((line) => line.length > 0)
    .join("\n");

  return text.trim();
}

if (typeof Office !== "undefined") {
  Office.actions.associate("removeTagsFromSelection", removeTagsFromSelection);
  Office.actions.associate("removeTagsFromWorksheet", removeTagsFromWorksheet);
  Office.actions.associate("removeTagsFromWorkbook", removeTagsFromWorkbook);
}
