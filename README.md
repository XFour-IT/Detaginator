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

Create a release package for Microsoft Partner submission:

```bash
npm run build:release
```

## Usage

Sideload `manifest.xml` into Excel for Windows or Excel for the Web. Use the **Tag Tools** group on the Home ribbon to remove HTML tags from the current selection, the active worksheet, or the entire workbook.
