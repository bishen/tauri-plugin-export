const COMMANDS: &[&str] = &["export", "export_docx"];

fn main() {
    tauri_plugin::Builder::new(COMMANDS).build();
}
