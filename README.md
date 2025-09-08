# Begone-Tags

Begone-Tags is an Excel add-in that strips HTML tags from selected cells in a workbook.

## Development

Install dependencies, lint, and run the test suite:

```bash
npm ci
npm run lint
npm test
```

Build the TypeScript sources:

```bash
npm run build
```

## Usage

Sideload `manifest.xml` into Excel for Windows or Excel for the Web. Open the add-in's task pane and choose **Clean Selection** to remove HTML tags from the currently selected cells.
