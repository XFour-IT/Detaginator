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
 * Remove HTML tags from the currently selected range in Excel.
 *
 * This command handler is intended to be associated with an Office Add-in
 * command (for example, a ribbon button). It obtains the currently
 * selected range and strips HTML from every cell inside the range.
 *
 * @param event - The add-in command event supplied by the Office runtime.
 *   The handler must call `event.completed()` when finished to let the
 *   runtime know the action has completed.
 *
 * @returns A promise that resolves when the operation has finished. The
 *   Office runtime is notified of completion via `event.completed()` in
 *   the finally block so callers do not need to await the returned
 *   promise for runtime bookkeeping.
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
 * Remove HTML tags from every used cell in the active worksheet.
 *
 * This command handler finds the active worksheet's used range and, if it
 * exists, strips HTML tags from each cell in that range.
 *
 * @param event - The add-in command event supplied by the Office runtime.
 *   Must be completed by calling `event.completed()` when the handler is
 *   finished.
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
 * Remove HTML tags from every used cell across all worksheets in the
 * current workbook.
 *
 * The handler enumerates all worksheets, collects their used ranges, and
 * then strips HTML from the cells in each non-empty range. This is
 * executed as a single logical operation from the user's perspective
 * (triggered by an add-in command).
 *
 * @param event - The add-in command event supplied by the Office runtime.
 *   `event.completed()` will be invoked once processing is finished.
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
 * Remove HTML markup from every string cell inside the provided range.
 *
 * Non-string cell values (numbers, booleans, errors, etc.) are left
 * untouched. The function reads the range values, applies `stripHtml` to
 * any string cells, writes the possibly-updated values back to the range,
 * and synchronizes the context.
 *
 * @param range - The Excel range whose cell values should be cleaned.
 *
 * @remarks
 * This function relies on the Excel JavaScript API's batching model and
 * therefore must be called from within an `Excel.run` callback or when a
 * valid context is available on the range.
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
 * A small mapping of commonly-encoded HTML entities to their literal
 * character equivalents.
 *
 * Keys are lower-cased to simplify lookup when decoding case-insensitive
 * entity names (the decoder normalizes matches to lower-case before lookup).
 *
 * @internal
 */
const htmlEntityMap: Record<string, string> = {
  "&lt;": "<",
  "&gt;": ">",
  "&amp;": "&",
  "&quot;": '"',
  "&apos;": "'",
  "&#39;": "'",
};

/**
 * Decode HTML character entities in a string.
 *
 * This function handles named entities included in `htmlEntityMap` as well
 * as numeric character references (decimal and hexadecimal forms).
 *
 * @param text - The text potentially containing HTML entities.
 * @returns The text with entities replaced by their corresponding
 *   characters. Unknown or malformed entities are left untouched.
 */
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

/**
 * Strip HTML markup from a string while preserving a minimal readable
 * representation of block structure.
 *
 * Behavior summary:
 * - <li> items become lines prefixed with "- ".
 * - <ul> and <ol> containers are removed (their items are already handled).
 * - <p>, <div>, and <br> become newlines.
 * - Remaining tags (including inline formatting like <b>/<i>) are removed.
 * - Common HTML entities (named and numeric) are decoded into characters.
 * - Leading/trailing whitespace for each line is trimmed and empty lines are
 *   removed.
 *
 * This function is safe for use on user-visible text cells and produces
 * plain text suitable for Excel cells.
 *
 * @param value - The input string that may contain HTML markup.
 * @returns A cleaned string with HTML removed and simple list/paragraph
 *   structure retained as newlines and dashes.
 */
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
