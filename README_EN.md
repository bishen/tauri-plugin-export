# tauri-plugin-export

Document export plugin (Excel + DOCX) for Tauri apps, cross-platform support.

## Supported Platforms

| Platform | Architecture | Status |
|----------|--------------|--------|
| Windows | x86_64, aarch64 | ✅ |
| macOS | x86_64, aarch64 | ✅ |
| Linux | x86_64, aarch64 | ✅ |
| Android | arm64-v8a, armeabi-v7a, x86_64 | ✅ |
| iOS | arm64, x86_64 (simulator) | ✅ |

## Features

### Excel Export

- **Multiple Sheets**: Export multiple worksheets to a single Excel file
- **Complex Headers**: Support multi-row headers with automatic cell merging (colspan/rowspan)
- **Auto Styling**: Automatic styling for headers and data cells (borders, alignment, background)
- **Type Conversion**: Automatic conversion of JSON values to Excel cell content

### DOCX Report Export

- **Section Structure**: Cover, TOC, multi-level headings, body text, tables, appendix tables
- **Typography**: Preset styles for forestry reports (font size, typeface, line spacing, indent)
- **Header/Footer**: Page header (project name) and footer (organization + page number)
- **Word Heading Styles**: Uses built-in Heading1-4 styles for automatic TOC generation

## Installation

### Rust Dependency

Add to `src-tauri/Cargo.toml`:

```toml
[dependencies]
tauri-plugin-export = { path = "../tauri-plugin-export" }
```

### JavaScript

Use the API from guest-js directly, or copy to your project.

## Setup

### 1. Register Plugin

In `src-tauri/src/lib.rs`:

```rust
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_export::init())
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
```

### 2. Configure Permissions

Add to `src-tauri/capabilities/default.json`:

```json
{
  "permissions": [
    "export:default"
  ]
}
```

## Usage

### Excel Export

```typescript
import { invoke } from '@tauri-apps/api/core';

const result = await invoke('plugin:export|export', {
  request: {
    path: 'D:/output/report.xlsx',
    sheets: [{
      name: 'Sheet1',
      headers: [
        [{ text: 'Name' }, { text: 'Value' }]
      ],
      rows: [['Item1', '100'], ['Item2', '200']]
    }]
  }
});
```

### DOCX Report Export

```typescript
const result = await invoke('plugin:export|export_docx', {
  params: {
    sections: [
      { id: 'cover', type: 'cover', content: '<p>Report Title</p>' },
      { id: 'ch1', type: 'heading', title: 'Chapter 1', content: '<p>...</p>' }
    ],
    outputPath: 'D:/output/report.docx',
    reportName: 'Report',
    headerText: 'Report Title',
    footerText: 'Organization',
  }
});
```

## API Reference

### Excel: `export(request)` → `ExportResult`

| Field | Type | Description |
|-------|------|-------------|
| path | string | Output file path |
| sheets | SheetData[] | Array of sheets |

### DOCX: `export_docx(params)` → `DocxExportResult`

| Field | Type | Description |
|-------|------|-------------|
| sections | ReportSection[] | Report section tree |
| docxStyle | DocxStyle? | Style config (optional, defaults to A4) |
| outputPath | string? | Output path |
| reportName | string? | Report name |
| headerText | string? | Page header text |
| footerText | string? | Page footer left text |

## License

MIT

Copyright (c) 2026 BiShen <bishen@live.com>
SuanJinShan™ (https://www.suanjinshan.com/)
