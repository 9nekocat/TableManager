use thiserror::Error;

#[derive(Debug, Error)]
pub enum CoreError {
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    #[error("Spreadsheet error: {0}")]
    Spreadsheet(#[from] calamine::Error),

    #[error("JSON error: {0}")]
    Json(#[from] serde_json::Error),

    #[error("Excel write error: {0}")]
    Xlsx(#[from] rust_xlsxwriter::XlsxError),

    #[error("{0}")]
    Message(String),
}

pub type Result<T> = std::result::Result<T, CoreError>;
