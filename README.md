# tauri-plugin-export

文档导出插件（Excel + DOCX），支持全平台。

[English](./README_EN.md)

## 支持的平台

| 平台 | 架构 | 状态 |
|------|------|------|
| Windows | x86_64, aarch64 | ✅ |
| macOS | x86_64, aarch64 | ✅ |
| Linux | x86_64, aarch64 | ✅ |
| Android | arm64-v8a, armeabi-v7a, x86_64 | ✅ |
| iOS | arm64, x86_64 (模拟器) | ✅ |

## 功能特性

### Excel 导出
- **多工作表**: 支持导出多个工作表到单个 Excel 文件
- **复杂表头**: 支持多行表头，自动处理单元格合并（colspan/rowspan）
- **自动样式**: 表头和数据单元格自动应用样式（边框、居中、背景色）
- **类型转换**: 自动将 JSON 值转换为 Excel 单元格内容

### DOCX 报告导出
- **章节结构**: 支持封面、目录、多级标题、正文、表格、附表
- **排版规范**: 预设林业报告排版样式（字号、字体、行距、缩进）
- **页眉页脚**: 支持页眉（项目名称）和页脚（调查单位 + 页码）
- **Word 标题样式**: 使用 Word 内置 Heading1-4 样式，支持自动生成目录

## 安装

### Rust 依赖

在 `src-tauri/Cargo.toml` 中添加:

```toml
[dependencies]
tauri-plugin-export = { path = "../tauri-plugin-export" }
```

### JavaScript 依赖

直接使用 guest-js 中的 API，或复制到项目中使用。

## 配置

### 1. 注册插件

在 `src-tauri/src/lib.rs` 中:

```rust
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_export::init())
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
```

### 2. 配置权限

在 `src-tauri/capabilities/default.json` 中添加:

```json
{
  "permissions": [
    "export:default"
  ]
}
```

## 使用方法

### Excel 导出

```typescript
import { invoke } from '@tauri-apps/api/core';

const result = await invoke('plugin:export|export', {
  request: {
    path: 'D:/output/因子一览表.xlsx',
    sheets: [{
      name: '因子一览表',
      headers: [
        [{ text: '小班属性', colspan: 5 }, { text: '乔木', colspan: 10 }],
        [{ text: '省' }, { text: '市' }, { text: '县' }, { text: '小班号' }]
      ],
      rows: [
        ['云南省', '昆明市', '富民县', '001'],
      ]
    }]
  }
});
```

### DOCX 报告导出

```typescript
import { invoke } from '@tauri-apps/api/core';

const result = await invoke('plugin:export|export_docx', {
  params: {
    sections: [
      { id: 'cover', type: 'cover', content: '<p>森林资源调查报告</p>' },
      { id: 'toc', type: 'toc', title: '目录' },
      {
        id: 'ch1', type: 'heading', title: '基本情况',
        content: '<p>项目区位于...</p>',
        children: [
          { id: 'ch1-1', type: 'heading', title: '地理位置', content: '<p>...</p>' }
        ]
      }
    ],
    outputPath: 'D:/output/调查报告.docx',
    reportName: '森林资源调查报告',
    headerText: '森林资源调查报告',
    footerText: '调查单位名称',
  }
});
```

## API 参考

### Excel: `export(request)` → `ExportResult`

| 字段 | 类型 | 说明 |
|------|------|------|
| path | string | 输出文件路径 |
| sheets | SheetData[] | 工作表列表 |

### DOCX: `export_docx(params)` → `DocxExportResult`

| 字段 | 类型 | 说明 |
|------|------|------|
| sections | ReportSection[] | 报告章节结构 |
| docxStyle | DocxStyle? | 样式配置（可选，默认 A4/仿宋/黑体） |
| outputPath | string? | 输出路径 |
| reportName | string? | 报告名称 |
| headerText | string? | 页眉文字 |
| footerText | string? | 页脚左侧文字 |

## 许可证

MIT

Copyright (c) 2026 BiShen <bishen@live.com>
算金山™ (https://www.suanjinshan.com/)
