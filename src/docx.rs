use crate::Error;
use serde::{Deserialize, Serialize};
use std::path::PathBuf;

/// 报告章节（从前端传入的 JSON 结构）
#[derive(Debug, Deserialize)]
pub struct ReportSection {
    pub id: String,
    #[serde(rename = "type")]
    pub sec_type: String,
    pub title: Option<String>,
    pub content: Option<String>,
    pub level: Option<u32>,
    pub children: Option<Vec<ReportSection>>,
    pub data: Option<TableData>,
    pub src: Option<String>,
    pub caption: Option<String>,
    pub auto: Option<bool>,
}

#[derive(Debug, Deserialize)]
pub struct TableData {
    pub headers: Vec<String>,
    pub rows: Vec<Vec<String>>,
}

/// DOCX 样式配置
#[derive(Debug, Deserialize)]
pub struct DocxStyle {
    #[serde(rename = "pageSize", default = "default_page_size")]
    pub page_size: String,
    #[serde(default)]
    pub margin: DocxMargin,
    #[serde(default)]
    pub font: DocxFont,
    #[serde(rename = "lineSpacing", default = "default_line_spacing")]
    pub line_spacing: f64,
}

fn default_page_size() -> String { "A4".into() }
fn default_line_spacing() -> f64 { 1.5 }

#[derive(Debug, Deserialize)]
pub struct DocxMargin {
    #[serde(default = "default_margin_top")]
    pub top: f64,
    #[serde(default = "default_margin_bottom")]
    pub bottom: f64,
    #[serde(default = "default_margin_left")]
    pub left: f64,
    #[serde(default = "default_margin_right")]
    pub right: f64,
}

fn default_margin_top() -> f64 { 3.5 }
fn default_margin_bottom() -> f64 { 2.5 }
fn default_margin_left() -> f64 { 3.0 }
fn default_margin_right() -> f64 { 2.5 }

impl Default for DocxMargin {
    fn default() -> Self {
        Self { top: 3.5, bottom: 2.5, left: 3.0, right: 2.5 }
    }
}

#[derive(Debug, Deserialize)]
pub struct DocxFont {
    #[serde(default = "default_body_font")]
    pub body: String,
    #[serde(default = "default_heading_font")]
    pub heading: String,
    #[serde(default = "default_font_size")]
    pub size: u32,
}

fn default_body_font() -> String { "仿宋".into() }
fn default_heading_font() -> String { "黑体".into() }
fn default_font_size() -> u32 { 12 } // 小四号 12pt

impl Default for DocxFont {
    fn default() -> Self {
        Self { body: "仿宋".into(), heading: "黑体".into(), size: 12 }
    }
}

/// 导出参数
#[derive(Debug, Deserialize)]
pub struct ExportParams {
    pub sections: Vec<ReportSection>,
    #[serde(rename = "docxStyle")]
    pub docx_style: Option<DocxStyle>,
    #[serde(rename = "outputPath")]
    pub output_path: Option<String>,
    #[serde(rename = "reportName")]
    pub report_name: Option<String>,
    /// 页眉文字（项目名称）
    #[serde(rename = "headerText", default)]
    pub header_text: Option<String>,
    /// 页脚左侧文字（调查单位）
    #[serde(rename = "footerText", default)]
    pub footer_text: Option<String>,
}

/// 导出结果
#[derive(Debug, Serialize)]
pub struct DocxExportResult {
    pub success: bool,
    pub path: String,
    pub message: String,
}

/// cm → twips (1 cm = 567 twips)
fn cm_to_twips(cm: f64) -> i32 {
    (cm * 567.0) as i32
}

/// pt → half-points (DOCX uses half-points for font size)
fn pt_to_half_points(pt: u32) -> u32 {
    pt * 2
}

/// 简易 HTML → 纯文本（去除标签）
fn strip_html(html: &str) -> String {
    let mut result = String::new();
    let mut in_tag = false;
    for ch in html.chars() {
        match ch {
            '<' => in_tag = true,
            '>' => in_tag = false,
            _ if !in_tag => result.push(ch),
            _ => {}
        }
    }
    // 处理 HTML 实体
    result
        .replace("&nbsp;", " ")
        .replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
}

/// 从 HTML 中提取段落文本列表
fn html_to_paragraphs(html: &str) -> Vec<String> {
    if html.is_empty() {
        return vec![];
    }
    // 按 <p>...</p> 分段
    let mut paragraphs = Vec::new();
    let mut remaining = html;

    while let Some(start) = remaining.find("<p") {
        if let Some(tag_end) = remaining[start..].find('>') {
            let after_tag = start + tag_end + 1;
            if let Some(end) = remaining[after_tag..].find("</p>") {
                let text = strip_html(&remaining[after_tag..after_tag + end]);
                let trimmed = text.trim();
                if !trimmed.is_empty() {
                    paragraphs.push(trimmed.to_string());
                }
                remaining = &remaining[after_tag + end + 4..];
            } else {
                break;
            }
        } else {
            break;
        }
    }

    // 没有 <p> 标签时，将整体作为一个段落
    if paragraphs.is_empty() {
        let text = strip_html(html).trim().to_string();
        if !text.is_empty() {
            paragraphs.push(text);
        }
    }

    paragraphs
}

/// 生成 DOCX 文件
///
/// 使用 docx-rs 库构建文档，将报告章节结构转为 DOCX 格式
#[tauri::command]
pub async fn export_docx(params: ExportParams) -> Result<DocxExportResult, Error> {
    use docx_rs::*;

    let style = params.docx_style.unwrap_or(DocxStyle {
        page_size: "A4".into(),
        margin: DocxMargin::default(),
        font: DocxFont::default(),
        line_spacing: 1.5,
    });

    let body_font = &style.font.body;
    let heading_font = &style.font.heading;
    let body_size = pt_to_half_points(style.font.size); // 小四号 12pt = 24 half-points
    let heading1_size = pt_to_half_points(16); // 三号 16pt
    let heading2_size = pt_to_half_points(14); // 四号 14pt
    let heading3_size = pt_to_half_points(12); // 小四号 12pt
    let hf_size: usize = 21; // 页眉页脚：五号 10.5pt = 21 half-points
    let line_spacing_val = (style.line_spacing * 240.0) as u32; // 240 twips = single spacing

    let mut docx = Docx::new();

    // ── 页面设置 ──
    // A4 = 11906 × 16838 twips
    docx = docx
        .page_size(11906, 16838)
        .page_margin(
            PageMargin::new()
                .top(cm_to_twips(style.margin.top))
                .bottom(cm_to_twips(style.margin.bottom))
                .left(cm_to_twips(style.margin.left))
                .right(cm_to_twips(style.margin.right))
                .header(cm_to_twips(2.0))  // 页眉距顶 2cm
                .footer(cm_to_twips(1.5))  // 页脚距底 1.5cm
        );

    // ── 页眉（项目名称，居中，宋体五号，下方横线）──
    if let Some(header_text) = &params.header_text {
        if !header_text.is_empty() {
            let header_run = Run::new()
                .add_text(header_text)
                .size(hf_size)
                .fonts(RunFonts::new().east_asia("宋体"));
            let header_para = Paragraph::new()
                .add_run(header_run)
                .align(AlignmentType::Center);
            docx = docx.header(Header::new().add_paragraph(header_para));
        }
    }

    // ── 页脚（左：调查单位，右：页码）──
    {
        let footer_left_text = params.footer_text.as_deref().unwrap_or("");
        // 用 Tab 实现左右分列：左侧文字 + Tab + 右侧页码
        let footer_left_run = Run::new()
            .add_text(footer_left_text)
            .size(hf_size)
            .fonts(RunFonts::new().east_asia("宋体"));

        let tab_run = Run::new()
            .add_tab()
            .size(hf_size);

        // 页码域: - {PAGE} -
        let page_prefix = Run::new()
            .add_text("- ")
            .size(hf_size)
            .fonts(RunFonts::new().east_asia("宋体"));
        let page_begin = Run::new()
            .add_field_char(FieldCharType::Begin, false)
            .size(hf_size);
        let page_instr = Run::new()
            .add_instr_text(InstrText::Unsupported("PAGE".into()))
            .size(hf_size);
        let page_end = Run::new()
            .add_field_char(FieldCharType::End, false)
            .size(hf_size);
        let page_suffix = Run::new()
            .add_text(" -")
            .size(hf_size)
            .fonts(RunFonts::new().east_asia("宋体"));

        let footer_para = Paragraph::new()
            .add_run(footer_left_run)
            .add_run(tab_run)
            .add_run(page_prefix)
            .add_run(page_begin)
            .add_run(page_instr)
            .add_run(page_end)
            .add_run(page_suffix)
            .align(AlignmentType::Both);
        docx = docx.footer(Footer::new().add_paragraph(footer_para));
    }

    // ── 定义 Word 内置标题样式（Heading1-4），使 Word 目录和导航窗格可用 ──
    {
        use docx_rs::{Style, StyleType, RunFonts, AlignmentType, LineSpacing};
        // 章标题：三号 16pt，居中，段前段后各1行(240twips×2=480)
        let h1 = Style::new("Heading1", StyleType::Paragraph)
            .name("heading 1")
            .based_on("Normal")
            .next("Normal")
            .size(heading1_size as usize)
            .bold()
            .fonts(RunFonts::new().east_asia(heading_font))
            .align(AlignmentType::Center)
            .line_spacing(LineSpacing::new().before(480).after(480))
            .outline_lvl(0);
        // 节标题：四号 14pt，左对齐，段前段后各0.5行(120twips)
        let h2 = Style::new("Heading2", StyleType::Paragraph)
            .name("heading 2")
            .based_on("Normal")
            .next("Normal")
            .size(heading2_size as usize)
            .bold()
            .fonts(RunFonts::new().east_asia(heading_font))
            .line_spacing(LineSpacing::new().before(120).after(120))
            .outline_lvl(1);
        // 条标题：小四号 12pt，左对齐，段前段后各0.3行(72twips)
        let h3 = Style::new("Heading3", StyleType::Paragraph)
            .name("heading 3")
            .based_on("Normal")
            .next("Normal")
            .size(heading3_size as usize)
            .bold()
            .fonts(RunFonts::new().east_asia(heading_font))
            .line_spacing(LineSpacing::new().before(72).after(72))
            .outline_lvl(2);
        // 款标题：小四号，加粗
        let h4 = Style::new("Heading4", StyleType::Paragraph)
            .name("heading 4")
            .based_on("Normal")
            .next("Normal")
            .size(body_size as usize)
            .bold()
            .fonts(RunFonts::new().east_asia(heading_font))
            .outline_lvl(3);
        docx = docx.add_style(h1).add_style(h2).add_style(h3).add_style(h4);
    }

    let chinese_nums = ["一","二","三","四","五","六","七","八","九","十","十一","十二","十三","十四","十五"];

    let mut heading_counter: usize = 0; // 顶层 heading 计数

    for section in &params.sections {
        match section.sec_type.as_str() {
            "cover" => {
                let paragraphs = html_to_paragraphs(section.content.as_deref().unwrap_or(""));
                for (i, text) in paragraphs.iter().enumerate() {
                    let size = if i == 0 { pt_to_half_points(26) } else { body_size };
                    let font = if i == 0 { heading_font } else { body_font };
                    let run = Run::new()
                        .add_text(text)
                        .size(size as usize)
                        .fonts(RunFonts::new().east_asia(font));
                    docx = docx.add_paragraph(
                        Paragraph::new().add_run(run).align(AlignmentType::Center)
                    );
                }
                docx = docx.add_paragraph(
                    Paragraph::new().add_run(Run::new().add_break(BreakType::Page))
                );
            }
            "toc" => {
                let run = Run::new()
                    .add_text("目  录")
                    .size(heading1_size as usize)
                    .fonts(RunFonts::new().east_asia(heading_font))
                    .bold();
                docx = docx.add_paragraph(
                    Paragraph::new().add_run(run).align(AlignmentType::Center)
                );
                docx = docx.add_paragraph(
                    Paragraph::new().add_run(
                        Run::new().add_text("（目录请在 Word 中通过 引用 → 目录 生成）")
                            .size(body_size as usize)
                            .fonts(RunFonts::new().east_asia(body_font))
                    ).align(AlignmentType::Center)
                );
                docx = docx.add_paragraph(
                    Paragraph::new().add_run(Run::new().add_break(BreakType::Page))
                );
            }
            "heading" => {
                heading_counter += 1;
                let num_path = vec![heading_counter];
                docx = render_heading(
                    docx, section, &num_path, &chinese_nums,
                    heading_font, body_font,
                    heading1_size, heading2_size, heading3_size,
                    body_size, line_spacing_val,
                );
            }
            "appendix-table" => {
                let title = section.title.as_deref().unwrap_or("");
                let run = Run::new()
                    .add_text(title)
                    .size(heading2_size as usize)
                    .fonts(RunFonts::new().east_asia(heading_font))
                    .bold();
                docx = docx.add_paragraph(
                    Paragraph::new().add_run(run).align(AlignmentType::Center)
                );
                if let Some(data) = &section.data {
                    docx = add_table_to_docx(docx, data, body_font, body_size);
                }
            }
            _ => {}
        }
    }

    // 确定输出路径
    let output_path = if let Some(p) = &params.output_path {
        PathBuf::from(p)
    } else {
        let name = params.report_name.as_deref().unwrap_or("report");
        let dir = dirs::document_dir().unwrap_or_else(|| dirs::home_dir().unwrap_or_default());
        dir.join(format!("{}.docx", name))
    };

    // 确保目录存在
    if let Some(parent) = output_path.parent() {
        std::fs::create_dir_all(parent).map_err(|e| Error::Docx(format!("创建目录失败: {}", e)))?;
    }

    // 写入文件
    let file = std::fs::File::create(&output_path)
        .map_err(|e| Error::Docx(format!("创建文件失败: {}", e)))?;

    docx.build()
        .pack(file)
        .map_err(|e| Error::Docx(format!("写入 DOCX 失败: {}", e)))?;

    let path_str = output_path.to_string_lossy().to_string();
    log::info!("[DOCX] 导出成功: {}", path_str);

    Ok(DocxExportResult {
        success: true,
        path: path_str.clone(),
        message: format!("已导出到 {}", path_str),
    })
}

/// 递归渲染 heading 节点（支持 4 级标题）
fn render_heading(
    mut docx: docx_rs::Docx,
    node: &ReportSection,
    num_path: &[usize],
    chinese_nums: &[&str],
    heading_font: &str,
    body_font: &str,
    h1_size: u32,
    h2_size: u32,
    h3_size: u32,
    body_size: u32,
    line_spacing_val: u32,
) -> docx_rs::Docx {
    use docx_rs::*;

    let title = node.title.as_deref().unwrap_or("");
    let depth = num_path.len(); // 1=章, 2=节, 3=条, 4=款

    // 生成编号文本
    let num_text = if depth == 1 {
        let idx = num_path[0];
        let cn = chinese_nums.get(idx - 1).unwrap_or(&"");
        format!("第{}章", cn)
    } else {
        num_path.iter().map(|n| n.to_string()).collect::<Vec<_>>().join(".")
    };

    // 根据层级决定字号、对齐、Word 样式 ID
    let (font_size, center, style_id) = match depth {
        1 => (h1_size, true, "Heading1"),
        2 => (h2_size, false, "Heading2"),
        3 => (h3_size, false, "Heading3"),
        _ => (body_size, false, "Heading4"),
    };

    // 标题段落（应用 Word 内置标题样式）
    let heading_text = format!("{} {}", num_text, title);
    let run = Run::new()
        .add_text(&heading_text)
        .size(font_size as usize)
        .fonts(RunFonts::new().east_asia(heading_font))
        .bold();
    let mut p = Paragraph::new()
        .add_run(run)
        .style(style_id);
    if center {
        p = p.align(AlignmentType::Center);
    }
    docx = docx.add_paragraph(p);

    // 本级正文内容（首行缩进2字符、1.5倍行距）
    if let Some(content) = &node.content {
        let paragraphs = html_to_paragraphs(content);
        for text in &paragraphs {
            let run = Run::new()
                .add_text(text)
                .size(body_size as usize)
                .fonts(RunFonts::new().east_asia(body_font));
            let indent_val = (body_size as i32 / 2) * 20 * 2; // 2字符缩进
            let p = Paragraph::new()
                .add_run(run)
                .line_spacing(LineSpacing::new().line(line_spacing_val as i32))
                .indent(Some(0), None, Some(indent_val), None);
            docx = docx.add_paragraph(p);
        }
    }

    // 递归渲染子节点
    if let Some(children) = &node.children {
        for (i, child) in children.iter().enumerate() {
            let mut child_path = num_path.to_vec();
            child_path.push(i + 1);
            docx = render_heading(
                docx, child, &child_path, chinese_nums,
                heading_font, body_font,
                h1_size, h2_size, h3_size,
                body_size, line_spacing_val,
            );
        }
    }

    docx
}

/// 向文档添加表格（100%页宽、完整边框、表头加粗底色）
fn add_table_to_docx(
    mut docx: docx_rs::Docx,
    data: &TableData,
    font: &str,
    _size: u32,
) -> docx_rs::Docx {
    use docx_rs::*;

    let tbl_size: u32 = 21; // 表格用五号字 10.5pt = 21 half-points

    // 表格边框：细实线
    let border = |pos| {
        TableBorder::new(pos)
            .size(4)
            .color("000000")
            .border_type(BorderType::Single)
    };
    let borders = TableBorders::new()
        .set(border(TableBorderPosition::Top))
        .set(border(TableBorderPosition::Bottom))
        .set(border(TableBorderPosition::Left))
        .set(border(TableBorderPosition::Right))
        .set(border(TableBorderPosition::InsideH))
        .set(border(TableBorderPosition::InsideV));

    // 单元格内边距 (twips): 上40 右80 下40 左80
    let cell_margins = TableCellMargins::new().margin(40, 80, 40, 80);

    // 列数
    let col_count = data.headers.len();

    let mut table = Table::new(vec![])
        .width(5000, WidthType::Pct) // 100% 页面宽度 (5000 = 100% in pct50ths)
        .align(TableAlignmentType::Center)
        .set_borders(borders)
        .margins(cell_margins);

    // 均分列宽
    if col_count > 0 {
        let col_w = 5000 / col_count;
        let grid: Vec<usize> = vec![col_w; col_count];
        table = table.set_grid(grid);
    }

    // 表头行（加粗 + 浅灰底色 + 居中）
    let header_cells: Vec<TableCell> = data.headers.iter().map(|h| {
        let run = Run::new()
            .add_text(h)
            .size(tbl_size as usize)
            .fonts(RunFonts::new().east_asia(font))
            .bold();
        TableCell::new()
            .add_paragraph(
                Paragraph::new().add_run(run).align(AlignmentType::Center)
            )
            .shading(Shading::new().fill("D9E2F3")) // 浅蓝灰底色
            .vertical_align(VAlignType::Center)
    }).collect();
    table = table.add_row(TableRow::new(header_cells));

    // 数据行
    for row in &data.rows {
        let cells: Vec<TableCell> = row.iter().enumerate().map(|(i, cell)| {
            let run = Run::new()
                .add_text(cell)
                .size(tbl_size as usize)
                .fonts(RunFonts::new().east_asia(font));
            // 第一列左对齐，其余居中
            let align = if i == 0 { AlignmentType::Left } else { AlignmentType::Center };
            TableCell::new()
                .add_paragraph(Paragraph::new().add_run(run).align(align))
                .vertical_align(VAlignType::Center)
        }).collect();
        table = table.add_row(TableRow::new(cells));
    }

    docx = docx.add_table(table);
    // 表后空一行
    docx = docx.add_paragraph(Paragraph::new());
    docx
}
