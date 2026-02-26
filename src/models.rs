use serde::{Deserialize, Serialize};

/// 表头单元格（支持合并）
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct HeaderCell {
    /// 显示文本
    pub text: String,
    /// 列合并数（colspan），默认1
    #[serde(default = "default_span")]
    pub colspan: u16,
    /// 行合并数（rowspan），默认1
    #[serde(default = "default_span")]
    pub rowspan: u16,
}

fn default_span() -> u16 { 1 }

/// 单个工作表
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SheetData {
    /// 工作表名称
    pub name: String,
    /// 多行表头（支持多层嵌套）
    pub headers: Vec<Vec<HeaderCell>>,
    /// 数据行
    pub rows: Vec<Vec<serde_json::Value>>,
}

/// 导出请求
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ExportRequest {
    /// 输出文件路径
    pub path: String,
    /// 工作表列表
    pub sheets: Vec<SheetData>,
}

/// 导出结果
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ExportResult {
    /// 是否成功
    pub success: bool,
    /// 输出路径
    pub path: String,
    /// 错误信息
    pub error: Option<String>,
}
