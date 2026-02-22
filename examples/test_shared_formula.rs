use excel_reader::excel::Excel;
use excel_reader::common_types::Coordinate;

fn main() -> anyhow::Result<()> {
    // Open the Excel file using raw string for path
    let mut excel = Excel::from_path(r"E:\share\tauri-excel\template.xlsx")?;
    
    // Get all sheets
    let sheets = excel.get_sheets()?;
    // Get the first worksheet
    let Some(first_sheet) = sheets.first() else {
        anyhow::bail!("No worksheet found");
    };
    // Get the complete worksheet object
    let worksheet = excel.get_worksheet(first_sheet)?;
    
    // Get cell D7 and print its formula
    let d7_coord = Coordinate::from_a1("D7".as_bytes()).ok_or(anyhow::anyhow!("Invalid coordinate D7"))?;
    let d7_cell = worksheet.get_cell(d7_coord)?;
    println!("D7 formula: {:?}", d7_cell.value);
    
    Ok(())
}