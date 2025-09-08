import { stripHtmlTags } from "./stripHtml";

Office.onReady(() => {
  // Office.js is ready
});

async function cleanRange(range: Excel.Range): Promise<void> {
  range.load(["values", "valueTypes"]);
  await range.context.sync();

  const { values, valueTypes } = range;
  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      if (valueTypes[i][j] === Excel.RangeValueType.string) {
        const text = values[i][j] as string;
        values[i][j] = stripHtmlTags(text);
      }
    }
  }
  range.values = values;
  await range.context.sync();
}

export async function cleanSelection(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    await cleanRange(range);
  });
}

export async function cleanWorksheet(): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRangeOrNullObject();
    range.load("address");
    await context.sync();

    if (!range.isNullObject) {
      await cleanRange(range);
    }
  });
}

export async function cleanWorkbook(): Promise<void> {
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items");
    await context.sync();

    for (const sheet of sheets.items) {
      const range = sheet.getUsedRangeOrNullObject();
      range.load("address");
      await context.sync();
      if (!range.isNullObject) {
        await cleanRange(range);
      }
    }
  });
}

(globalThis as any).cleanSelection = cleanSelection;
(globalThis as any).cleanWorksheet = cleanWorksheet;
(globalThis as any).cleanWorkbook = cleanWorkbook;
