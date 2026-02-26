use tauri::{
    plugin::{Builder, TauriPlugin},
    Runtime,
};

mod commands;
pub mod docx;
mod error;
mod models;

pub use error::Error;
pub use models::*;

pub fn init<R: Runtime>() -> TauriPlugin<R> {
    Builder::new("export")
        .invoke_handler(tauri::generate_handler![commands::export, docx::export_docx])
        .build()
}
