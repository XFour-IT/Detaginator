import { stripHtmlTags } from "./stripHtml";

Office.onReady(() => {
  // Office.js is ready
});

export async function cleanSelection(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["values", "valueTypes"]);
    await context.sync();

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
    await context.sync();
  });
}
