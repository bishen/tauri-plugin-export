import { invoke } from '@tauri-apps/api/core';

// ========== Excel 导出 ==========

/** 表头单元格 */
export interface HeaderCell {
  /** 显示文本 */
  text: string;
  /** 列合并数，默认1 */
  colspan?: number;
  /** 行合并数，默认1 */
  rowspan?: number;
}

/** 工作表数据 */
export interface SheetData {
  /** 工作表名称 */
  name: string;
  /** 标题行（可选，大号字体居中合并显示在第一行） */
  title?: string;
  /** 多行表头 */
  headers: HeaderCell[][];
  /** 数据行 */
  rows: (string | number | boolean | null)[][];
}

/** Excel 导出请求 */
export interface ExportRequest {
  /** 输出文件路径 */
  path: string;
  /** 工作表列表 */
  sheets: SheetData[];
}

/** 导出结果 */
export interface ExportResult {
  success: boolean;
  path: string;
  error?: string;
}

/**
 * 导出数据到 Excel 文件
 * @param request 导出请求
 * @returns 导出结果
 */
export async function exportExcel(request: ExportRequest): Promise<ExportResult> {
  return invoke<ExportResult>('plugin:export|export', { request });
}

// ========== DOCX 报告导出 ==========

/** 报告表格数据 */
export interface DocxTableData {
  headers: string[];
  rows: string[][];
}

/** 报告章节 */
export interface ReportSection {
  id: string;
  type: string;
  title?: string;
  content?: string;
  level?: number;
  children?: ReportSection[];
  data?: DocxTableData;
  src?: string;
  caption?: string;
  auto?: boolean;
}

/** DOCX 页边距 */
export interface DocxMargin {
  top?: number;
  bottom?: number;
  left?: number;
  right?: number;
}

/** DOCX 字体配置 */
export interface DocxFont {
  body?: string;
  heading?: string;
  size?: number;
}

/** DOCX 样式配置 */
export interface DocxStyle {
  pageSize?: string;
  margin?: DocxMargin;
  font?: DocxFont;
  lineSpacing?: number;
}

/** DOCX 导出参数 */
export interface DocxExportParams {
  sections: ReportSection[];
  docxStyle?: DocxStyle | null;
  outputPath?: string;
  reportName?: string;
  headerText?: string;
  footerText?: string;
}

/** DOCX 导出结果 */
export interface DocxExportResult {
  success: boolean;
  path: string;
  message: string;
}

/**
 * 导出报告到 DOCX 文件
 * @param params 导出参数
 * @returns 导出结果
 */
export async function exportDocx(params: DocxExportParams): Promise<DocxExportResult> {
  return invoke<DocxExportResult>('plugin:export|export_docx', { params });
}
