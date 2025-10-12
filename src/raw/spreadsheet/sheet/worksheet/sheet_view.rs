use anyhow::bail;
use quick_xml::events::{BytesStart, Event};
use std::io::Read;

use crate::{excel::XmlReader, helper::{string_to_float, string_to_unsignedint, extract_val_attribute}};


/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sheetview?view=openxml-3.0.1
///
/// A single sheet view definition.
/// When more than one sheet view is defined in the file, it means that when opening the workbook, each sheet view corresponds to a separate window within the spreadsheet application, where each window is showing the particular sheet containing the same workbookViewId value, the last sheetView definition is loaded, and the others are discarded.
/// Example
/// ```
/// <sheetViews>
///   <sheetView tabSelected="1" workbookViewId="0">
///     <pane xSplit="2310" ySplit="2070" topLeftCell="C1" activePane="bottomRight"/>
///     <selection/>
///     <selection pane="bottomLeft" activeCell="A6" sqref="A6"/>
///     <selection pane="topRight" activeCell="C1" sqref="C1"/>
///     <selection pane="bottomRight" activeCell="E13" sqref="E13"/>
///   </sheetView>
/// </sheetViews>
/// ```
/// sheetView (Worksheet View)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxSheetView {
    // extLst (Future Feature Data Storage Area) Not supported

    // Child Elements
    // pane (View Pane)	§18.3.1.66
    pub pane: Option<XlsxPane>,
    // pivotSelection (PivotTable Selection)	§18.3.1.69
    // selection (Selection)
}

impl XlsxSheetView {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let mut sheet_view = Self {
            pane: None,
        };

        let mut buf = Vec::new();
        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref start_e)) if start_e.local_name().as_ref() == b"pane" => {
                    sheet_view.pane = Some(XlsxPane::load(start_e)?);
                    // Read to end of pane element
                    reader.read_to_end_into(start_e.to_end().to_owned().name(), &mut Vec::new())?;
                }
                Ok(Event::End(ref end_e)) if end_e.local_name().as_ref() == b"sheetView" => break,
                Ok(Event::Eof) => bail!("unexpected end of file at `sheetView`"),
                Err(err) => bail!(err.to_string()),
                _ => (),
            }
        }

        Ok(sheet_view)
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.pane?view=openxml-3.0.1
///
/// View Pane
/// This element specifies a view pane.
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxPane {
    // Attributes
    /// activePane (Active Pane)
    ///
    /// Specifies the active pane when in a frozen or split worksheet.
    /// Values are: topLeft, topRight, bottomLeft, bottomRight.
    pub active_pane: Option<String>,

    /// state (Pane State)
    ///
    /// Specifies the state of the pane.
    /// Values are: split, frozen, frozenSplit
    pub state: Option<String>,

    /// topLeftCell (Top Left Visible Cell)
    ///
    /// The cell that appears in the top left corner of the pane.
    pub top_left_cell: Option<String>,

    /// xSplit (Horizontal Split Position)
    ///
    /// Horizontal position of the split, in 1/20th of a point.
    pub x_split: Option<f64>,

    /// ySplit (Vertical Split Position)
    ///
    /// Vertical position of the split, in 1/20th of a point.
    pub y_split: Option<f64>,

    /// splitHorizontal (Horizontal Split Position (Pixels))
    ///
    /// Horizontal position of the split, in pixels.
    pub split_horizontal: Option<u64>,

    /// splitVertical (Vertical Split Position (Pixels))
    ///
    /// Vertical position of the split, in pixels.
    pub split_vertical: Option<u64>,

    /// thinHorizontal (Thin Horizontal Split Bar)
    ///
    /// Show a thin horizontal split bar.
    pub thin_horizontal: Option<bool>,

    /// thinVertical (Thin Vertical Split Bar)
    ///
    /// Show a thin vertical split bar.
    pub thin_vertical: Option<bool>,
}

impl XlsxPane {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let mut pane = Self {
            active_pane: None,
            state: None,
            top_left_cell: None,
            x_split: None,
            y_split: None,
            split_horizontal: None,
            split_vertical: None,
            thin_horizontal: None,
            thin_vertical: None,
        };

        // Iterate through all attributes
        for attr in e.attributes() {
            let attr = attr?;
            let key = attr.key.local_name().as_ref();
            let value = String::from_utf8(attr.value.to_vec())?;

            match key {
                b"activePane" => pane.active_pane = Some(value),
                b"state" => pane.state = Some(value),
                b"topLeftCell" => pane.top_left_cell = Some(value),
                b"xSplit" => {
                    if let Some(x) = string_to_float(&value) {
                        pane.x_split = Some(x);
                    }
                },
                b"ySplit" => {
                    if let Some(y) = string_to_float(&value) {
                        pane.y_split = Some(y);
                    }
                },
                b"splitHorizontal" => {
                    if let Some(horizontal) = string_to_unsignedint(&value) {
                        pane.split_horizontal = Some(horizontal);
                    }
                },
                b"splitVertical" => {
                    if let Some(vertical) = string_to_unsignedint(&value) {
                        pane.split_vertical = Some(vertical);
                    }
                },
                b"thinHorizontal" => {
                    if let Some(thin) = value.parse::<bool>().ok() {
                        pane.thin_horizontal = Some(thin);
                    }
                },
                b"thinVertical" => {
                    if let Some(thin) = value.parse::<bool>().ok() {
                        pane.thin_vertical = Some(thin);
                    }
                },
                _ => {}
            }
        }

        Ok(pane)
    }
}

/// Load sheet views from XML
pub(crate) fn load_sheet_views(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Vec<XlsxSheetView>> {
    let mut sheet_views = Vec::new();
    let mut buf = Vec::new();

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sheetView" => {
                sheet_views.push(XlsxSheetView::load(reader, e)?);
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sheetViews" => break,
            Ok(Event::Eof) => bail!("unexpected end of file at `sheetViews`"),
            Err(err) => bail!(err.to_string()),
            _ => (),
        }
    }

    Ok(sheet_views)
}