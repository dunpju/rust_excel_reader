#[cfg(feature = "serde")]
use serde::Serialize;

#[derive(Clone, PartialEq, Debug)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct DataValidation {
    /// Allow blank cells
    pub allow_blank: bool,

    /// Error message
    pub error_message: Option<String>,

    /// Error title
    pub error_title: Option<String>,

    /// Formula 1
    pub formula1: Option<String>,

    /// Formula 2
    pub formula2: Option<String>,

    /// Operator
    pub operator: Option<String>,

    /// Input message
    pub prompt: Option<String>,

    /// Input message title
    pub prompt_title: Option<String>,

    /// Show drop down list
    pub show_drop_down: bool,

    /// Show error message
    pub show_error_message: bool,

    /// Show input message
    pub show_input_message: bool,

    /// Sequence of references (cell ranges)
    pub sqref: String,

    /// Data validation type
    pub r#type: String,
}

impl DataValidation {
    pub(crate) fn from_raw(raw: crate::raw::spreadsheet::sheet::worksheet::data_validation::XlsxDataValidation) -> Self {
        Self {
            allow_blank: raw.allow_blank.unwrap_or(false),
            error_message: raw.error_message,
            error_title: raw.error_title,
            formula1: raw.formula1,
            formula2: raw.formula2,
            operator: raw.operator,
            prompt: raw.prompt,
            prompt_title: raw.prompt_title,
            show_drop_down: raw.show_drop_down.unwrap_or(false),
            show_error_message: raw.show_error_message.unwrap_or(false),
            show_input_message: raw.show_input_message.unwrap_or(false),
            sqref: raw.sqref.unwrap_or_default(),
            r#type: raw.r#type.unwrap_or_default(),
        }
    }
}
