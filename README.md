# Doc to HTML Converter

A local React app to convert lab `.docx` reports into HTML templates for quick copy/paste into your internal system.

## What It Does

- Batch upload `.docx` files (drag/drop or file picker).
- Extract key report fields from Word XML:
- `Test Name`
- `Method`
- `Result`
- `Units`
- `Bio. Ref. Interval`
- Build formatted HTML output per file.
- Preview, edit, and copy generated HTML.
- Copy file name without `.docx`.
- Download all generated HTML as a single file.

## Recent Behavior Fixes

- Reduced unwanted blank spacing in pasted output by tightening `<br>` usage and using margin-based spacing in generated blocks.
- Fixed table alignment by rendering header and data in a single fixed-layout table with shared column widths.
- Updated filename copy action to exclude `.docx` extension (case-insensitive).

## Tech Stack

- React 19
- `react-scripts` (Create React App)
- `JSZip` loaded from CDN at runtime for parsing `.docx` archives

## Run Locally

```bash
npm install
npm start
```

Open `http://localhost:3000`.

## Build

```bash
npm run build
```

Production output is generated in `build/`.

## Main File

- `src/lab-html-generator.jsx`

## Notes

- Processing is fully client-side in the browser.
- No server upload is required for conversion logic in this app.
