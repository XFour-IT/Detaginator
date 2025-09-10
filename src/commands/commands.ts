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
 * Remove HTML tags from a string, keeping text inside <p> tags when present.
 * @param value - the string to clean
 * @returns cleaned string
 */
export function stripHtml(value: string): string {
  const paragraphRegex = /<p[^>]*>(.*?)<\/p>/gi;
  const nbspRegex = /&nbsp;/gi;

  const hasParagraph = paragraphRegex.test(value);
  paragraphRegex.lastIndex = 0;

  if (!hasParagraph) {
    return value.replace(nbspRegex, " ");
  }
  const paragraphs: string[] = [];
  let match: RegExpExecArray | null;
  while ((match = paragraphRegex.exec(value)) !== null) {
    paragraphs.push(match[1]);
  }
  if (paragraphs.length > 0) {
    return paragraphs
      .map((p) => p.replace(/<[^>]+>/g, "").replace(nbspRegex, " "))
      .join("\n")
      .trim();
  }
  return value.replace(/<[^>]+>/g, "").replace(nbspRegex, " ");
}

if (typeof Office !== "undefined") {
  Office.actions.associate("removeTagsFromSelection", removeTagsFromSelection);
  Office.actions.associate("removeTagsFromWorksheet", removeTagsFromWorksheet);
  Office.actions.associate("removeTagsFromWorkbook", removeTagsFromWorkbook);
}
