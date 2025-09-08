/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Excel console Office */

function stripHtmlTags(text: string): string {
  return text.replace(/<[^>]*>/g, "");
}

async function removeTagsFromRange(range: Excel.Range, context: Excel.RequestContext): Promise<void> {
  range.load(["values", "isNullObject"]);
  await context.sync();
  if ((range as any).isNullObject) {
    return;
  }

  const values = range.values as any[][];
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const cell = values[r][c];
      if (typeof cell === "string") {
        values[r][c] = stripHtmlTags(cell);
      }
    }
  }
  range.values = values;
  await context.sync();
}

export async function removeAllTags(event: Office.AddinCommands.Event): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items");
      await context.sync();

      for (const sheet of sheets.items) {
        const used = sheet.getUsedRangeOrNullObject();
        await removeTagsFromRange(used, context);
      }
    });
  } catch (error) {
    console.error(error);
  }
  event.completed();
}

export async function removeTagsFromWorksheet(event: Office.AddinCommands.Event): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const used = sheet.getUsedRangeOrNullObject();
      await removeTagsFromRange(used, context);
    });
  } catch (error) {
    console.error(error);
  }
  event.completed();
}

export async function removeTagsFromSelection(event: Office.AddinCommands.Event): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      await removeTagsFromRange(range, context);
    });
  } catch (error) {
    console.error(error);
  }
  event.completed();
}
