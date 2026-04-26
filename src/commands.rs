use crate::{Error, ExportRequest, ExportResult, HeaderCell};
use rust_xlsxwriter::{Format, Workbook, Worksheet};

/// 导出数据到 Excel 文件
#[tauri::command]
pub async fn export(request: ExportRequest) -> Result<ExportResult, Error> {
    let mut workbook = Workbook::new();
    
    // 标题样式（大号字体居中）
    let title_format = Format::new()
        .set_bold()
        .set_font_size(16.0)
        .set_align(rust_xlsxwriter::FormatAlign::Center)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter);
    
    // 表头样式
    let header_format = Format::new()
        .set_bold()
        .set_text_wrap()
        .set_align(rust_xlsxwriter::FormatAlign::Center)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
        .set_border(rust_xlsxwriter::FormatBorder::Thin)
        .set_background_color(0xDEE7F1);
    
    // 数据单元格样式
    let cell_format = Format::new()
        .set_align(rust_xlsxwriter::FormatAlign::Center)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
        .set_border(rust_xlsxwriter::FormatBorder::Thin);
    
    for sheet_data in &request.sheets {
        let worksheet = workbook.add_worksheet();
        worksheet.set_name(&sheet_data.name)?;
        
        // 标题行偏移
        let title_offset: usize = if sheet_data.title.is_some() { 1 } else { 0 };
        
        // 写入标题行
        if let Some(title) = &sheet_data.title {
            let max_cols = get_max_cols(&sheet_data.headers);
            if max_cols > 1 {
                worksheet.merge_range(0, 0, 0, max_cols - 1, title, &title_format)?;
            } else {
                worksheet.write_string_with_format(0, 0, title, &title_format)?;
            }
            worksheet.set_row_height(0, 30.0)?;
        }
        
        // 写入表头
        let header_rows = write_headers(worksheet, &sheet_data.headers, &header_format, title_offset)?;
        
        // 设置表头行高
        for row in 0..header_rows {
            worksheet.set_row_height((title_offset + row) as u32, 22.0)?;
        }
        
        // 写入数据行
        for (row_idx, row) in sheet_data.rows.iter().enumerate() {
            let excel_row = (title_offset + header_rows + row_idx) as u32;
            // 设置数据行高
            worksheet.set_row_height(excel_row, 20.0)?;
            for (col_idx, cell) in row.iter().enumerate() {
                let value = json_to_string(cell);
                worksheet.write_string_with_format(excel_row, col_idx as u16, &value, &cell_format)?;
            }
        }
        
        // 自动调整列宽（设置最小宽度）
        for col in 0..get_max_cols(&sheet_data.headers) {
            worksheet.set_column_width(col, 12.0)?;
        }
    }
    
    workbook.save(&request.path)?;
    
    log::info!("[Excel] 导出成功: {}", request.path);
    
    Ok(ExportResult {
        success: true,
        path: request.path,
        error: None,
    })
}

/// 写入多行表头，处理合并单元格
fn write_headers(
    worksheet: &mut Worksheet,
    headers: &[Vec<HeaderCell>],
    format: &Format,
    row_offset: usize,
) -> Result<usize, Error> {
    // 跟踪哪些单元格已被占用（由于合并）
    let mut occupied: std::collections::HashSet<(u32, u16)> = std::collections::HashSet::new();
    
    for (row_idx, header_row) in headers.iter().enumerate() {
        let excel_row = (row_offset + row_idx) as u32;
        let mut col: u16 = 0;
        
        for cell in header_row {
            // 跳过被占用的列
            while occupied.contains(&(excel_row, col)) {
                col += 1;
            }
            
            let start_col = col;
            let end_col = col + cell.colspan - 1;
            let end_row = excel_row + cell.rowspan as u32 - 1;
            
            // 标记占用的单元格
            for r in excel_row..=end_row {
                for c in start_col..=end_col {
                    occupied.insert((r, c));
                }
            }
            
            // 写入并合并
            if cell.colspan > 1 || cell.rowspan > 1 {
                worksheet.merge_range(
                    excel_row, start_col,
                    end_row, end_col,
                    &cell.text,
                    format,
                )?;
            } else {
                worksheet.write_string_with_format(excel_row, col, &cell.text, format)?;
            }
            
            col = end_col + 1;
        }
    }
    
    Ok(headers.len())
}

/// 获取表头最大列数
fn get_max_cols(headers: &[Vec<HeaderCell>]) -> u16 {
    headers.iter()
        .map(|row| row.iter().map(|c| c.colspan).sum::<u16>())
        .max()
        .unwrap_or(1)
}

/// JSON 值转字符串
fn json_to_string(value: &serde_json::Value) -> String {
    match value {
        serde_json::Value::Null => "-".to_string(),
        serde_json::Value::Bool(b) => if *b { "是" } else { "否" }.to_string(),
        serde_json::Value::Number(n) => n.to_string(),
        serde_json::Value::String(s) => s.clone(),
        _ => value.to_string(),
    }
}
