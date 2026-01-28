use anyhow::bail;
use quick_xml::events::{BytesStart, Event};
use std::io::Read;

use crate::{excel::XmlReader, helper::string_to_bool};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.datavalidation?view=openxml-3.0.1
///
/// This element represents a data validation rule applied to a range of cells.
///
/// Example:
    /// ```xml
    /// <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1:A10">
    ///     <formula1>"Option1,Option2,Option3"</formula1>
    /// </dataValidation>
    /// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxDataValidation {
    /// extLst (Future Feature Data Storage Area)	Not supported

    /// Child Elements
    /// formula1 (Formula 1)
    pub formula1: Option<String>,

    /// formula2 (Formula 2)
    pub formula2: Option<String>,

    /// Attributes
    /// allowBlank (Allow Blank)
    pub allow_blank: Option<bool>,

    /// error (Error Message)
    pub error_message: Option<String>,

    /// errorTitle (Error Title)
    pub error_title: Option<String>,

    /// operator (Operator)
    pub operator: Option<String>,

    /// prompt (Input Message)
    pub prompt: Option<String>,

    /// promptTitle (Input Message Title)
    pub prompt_title: Option<String>,

    /// showDropDown (Show Drop Down List)
    pub show_drop_down: Option<bool>,

    /// showErrorMessage (Show Error Message)
    pub show_error_message: Option<bool>,

    /// showInputMessage (Show Input Message)
    pub show_input_message: Option<bool>,

    /// sqref (Sequence of References)
    pub sqref: Option<String>,

    /// type (Data Validation Type)
    pub r#type: Option<String>,
}

impl XlsxDataValidation {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let mut data_validation = Self {
            formula1: None,
            formula2: None,
            allow_blank: None,
            error_message: None,
            error_title: None,
            operator: None,
            prompt: None,
            prompt_title: None,
            show_drop_down: None,
            show_error_message: None,
            show_input_message: None,
            sqref: None,
            r#type: None,
        };

        // Parse attributes
        let attributes = e.attributes();
        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"allowBlank" => {
                            data_validation.allow_blank = string_to_bool(&string_value);
                        }
                        b"error" => {
                            data_validation.error_message = Some(string_value);
                        }
                        b"errorTitle" => {
                            data_validation.error_title = Some(string_value);
                        }
                        b"operator" => {
                            data_validation.operator = Some(string_value);
                        }
                        b"prompt" => {
                            data_validation.prompt = Some(string_value);
                        }
                        b"promptTitle" => {
                            data_validation.prompt_title = Some(string_value);
                        }
                        b"showDropDown" => {
                            data_validation.show_drop_down = string_to_bool(&string_value);
                        }
                        b"showErrorMessage" => {
                            data_validation.show_error_message = string_to_bool(&string_value);
                        }
                        b"showInputMessage" => {
                            data_validation.show_input_message = string_to_bool(&string_value);
                        }
                        b"sqref" => {
                            data_validation.sqref = Some(string_value);
                        }
                        b"type" => {
                            data_validation.r#type = Some(string_value);
                        }
                        _ => {},
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        // Parse child elements
        let mut buf: Vec<u8> = Vec::new();
        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"formula1" => {
                    data_validation.formula1 = Some(Self::load_formula(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"formula2" => {
                    data_validation.formula2 = Some(Self::load_formula(reader)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"dataValidation" => break,
                Ok(Event::Eof) => bail!("unexpected end of file at `dataValidation`"),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        return Ok(data_validation);
    }

    fn load_formula(reader: &mut XmlReader<impl Read>) -> anyhow::Result<String> {
        let mut buf: Vec<u8> = Vec::new();
        loop {
            buf.clear();
            match reader.read_event_into(&mut buf) {
                Ok(Event::Text(e)) => {
                    return Ok(String::from_utf8(e.to_vec())?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"formula1" || 
                                         e.local_name().as_ref() == b"formula2" => {
                    return Ok(String::new());
                }
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.datavalidations?view=openxml-3.0.1
///
/// This element is a container for the data validation rules applied to cells in this worksheet.
///
/// Example:
/// ```
/// <dataValidations count="1">
///     <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1:A10">
///         <formula1>"Option1,Option2,Option3"</formula1>
///     </dataValidation>
/// </dataValidations>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxDataValidations {
    /// extLst (Future Feature Data Storage Area)	Not supported

    /// Child Elements
    /// dataValidation (Data Validation)
    pub data_validations: Vec<XlsxDataValidation>,

    /// Attributes
    /// count (Count)
    pub count: Option<u64>,
}

impl XlsxDataValidations {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut data_validations = Self {
            data_validations: vec![],
            count: None,
        };

        let mut buf: Vec<u8> = Vec::new();
        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dataValidation" => {
                    data_validations.data_validations.push(XlsxDataValidation::load(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"dataValidations" => break,
                Ok(Event::Eof) => bail!("unexpected end of file at `dataValidations`"),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        return Ok(data_validations);
    }
}
