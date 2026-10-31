use crate::coordinate::{get_column_letter, tuple_to_coordinate};
use crate::error::RpyxlResult;

#[derive(Debug, Clone, PartialEq, Default)]
pub enum CellValue {
    #[default]
    Empty,
    String(String),
    Int(i64),
    Float(f64),
    Bool(bool),
    Date(String),
    DateTime(String),
    Formula {
        formula: String,
        cached_value: Option<Box<CellValue>>,
    },
}

impl CellValue {
    pub fn formula(formula: impl Into<String>) -> Self {
        CellValue::Formula {
            formula: formula.into(),
            cached_value: None,
        }
    }

    pub fn formula_with_cached(formula: impl Into<String>, cached: CellValue) -> Self {
        CellValue::Formula {
            formula: formula.into(),
            cached_value: Some(Box::new(cached)),
        }
    }

    pub fn data_type(&self) -> &'static str {
        match self {
            CellValue::Empty => "n",
            CellValue::String(_) => "s",
            CellValue::Int(_) | CellValue::Float(_) => "n",
            CellValue::Bool(_) => "b",
            CellValue::Date(_) | CellValue::DateTime(_) => "d",
            CellValue::Formula { .. } => "f",
        }
    }

    pub fn is_empty(&self) -> bool {
        matches!(self, CellValue::Empty)
    }

    pub fn resolve_value(&self, data_only: bool) -> Option<&CellValue> {
        match self {
            CellValue::Formula { cached_value, .. } => {
                if data_only {
                    cached_value.as_deref()
                } else {
                    Some(self)
                }
            }
            _ => Some(self),
        }
    }
}

/// OOXML element fragments for the native rich-style components.
/// Fragments are validated by the OOXML writer before being persisted.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub struct RichStyle {
    pub font: Option<String>,
    pub fill: Option<String>,
    pub border: Option<String>,
    pub alignment: Option<String>,
    pub protection: Option<String>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Cell {
    pub row: u32,
    pub column: u32,
    pub value: CellValue,
    pub number_format: Option<String>,
    pub rich_style: Option<std::sync::Arc<RichStyle>>,
    /// One-based original XF index; NonZero keeps provenance compact.
    pub source_xf_id: Option<std::num::NonZeroU32>,
    /// Explicit XF zero overrides row/column defaults; absence can inherit them.
    pub source_xf_explicit: bool,
    pub comment: Option<String>,
    pub hyperlink: Option<String>,
}

impl Cell {
    pub fn new(row: u32, column: u32, value: CellValue) -> Self {
        Self {
            row,
            column,
            value,
            number_format: None,
            rich_style: None,
            source_xf_id: None,
            source_xf_explicit: false,
            comment: None,
            hyperlink: None,
        }
    }

    pub fn source_xf_index(&self) -> Option<usize> {
        self.source_xf_id.map(|id| (id.get() - 1) as usize)
    }

    pub fn coordinate(&self) -> RpyxlResult<String> {
        tuple_to_coordinate(self.row, self.column)
    }

    pub fn column_letter(&self) -> RpyxlResult<String> {
        get_column_letter(self.column)
    }

    pub fn data_type(&self) -> &'static str {
        self.value.data_type()
    }

    pub fn is_date(&self) -> bool {
        matches!(self.value, CellValue::Date(_) | CellValue::DateTime(_))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_creation_and_properties() {
        let cell = Cell::new(1, 1, CellValue::String("hello".to_string()));
        assert_eq!(cell.row, 1);
        assert_eq!(cell.column, 1);
        assert_eq!(cell.coordinate().unwrap(), "A1");
        assert_eq!(cell.column_letter().unwrap(), "A");
        assert_eq!(cell.data_type(), "s");

        let num_cell = Cell::new(5, 27, CellValue::Int(42));
        assert_eq!(num_cell.coordinate().unwrap(), "AA5");
        assert_eq!(num_cell.data_type(), "n");

        let form_cell = Cell::new(
            1,
            2,
            CellValue::formula_with_cached("=A1+1", CellValue::Int(43)),
        );
        assert_eq!(form_cell.data_type(), "f");
        assert_eq!(
            form_cell.value.resolve_value(true),
            Some(&CellValue::Int(43))
        );
    }
}
