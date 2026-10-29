use thiserror::Error;

#[derive(Error, Debug, Clone, PartialEq, Eq)]
pub enum RpyxlError {
    #[error("Invalid cell coordinate '{0}'. Must be like 'A1' or 1-based (row, column).")]
    InvalidCoordinate(String),

    #[error("Worksheet '{0}' does not exist")]
    SheetNotFound(String),

    #[error("Worksheet '{0}' already exists")]
    DuplicateSheetTitle(String),

    #[error("Invalid worksheet title '{0}'. Title cannot be empty, exceed 31 characters, start/end with single quotes, or contain any of : \\ / ? * [ ]")]
    InvalidSheetTitle(String),

    #[error("Worksheet index {0} out of range (workbook has {1} sheets)")]
    SheetIndexOutOfBounds(usize, usize),

    #[error("Cannot remove the only worksheet from workbook")]
    CannotRemoveLastSheet,

    #[error("Invalid range string '{0}'. Expected format like 'A1:B10'.")]
    InvalidRange(String),

    #[error("Row index {0} is out of bounds (Excel allows rows 1 to 1,048,576)")]
    RowOutOfBounds(u32),

    #[error("Column index {0} is out of bounds (Excel allows columns 1 to 16,384)")]
    ColumnOutOfBounds(u32),

    #[error("Value error: {0}")]
    ValueError(String),

    #[error("Type error: {0}")]
    TypeError(String),

    #[error("IO error: {0}")]
    IoError(String),

    #[error("ZIP bomb detected: {0}")]
    ZipBombDetected(String),

    #[error("XML security violation: {0}")]
    XmlSecurityViolation(String),

    #[error("Invalid XLSX file format: {0}")]
    InvalidFileFormat(String),
}

pub type RpyxlResult<T> = Result<T, RpyxlError>;
