use excel_reader::excel::Excel;

fn main() -> anyhow::Result<()> {
    // Open a small Excel file for testing
    println!("Opening Excel file...");
    let mut excel = Excel::from_path(r"E:\share\tauri-excel\1.xlsx")?;
    
    // Get all sheets
    println!("Getting sheets...");
    let sheets = excel.get_sheets()?;
    
    // Get the first worksheet
    let Some(first_sheet) = sheets.first() else {
        anyhow::bail!("No worksheet found");
    };
    
    // Get the complete worksheet object
    println!("Getting worksheet...");
    let worksheet = excel.get_worksheet(first_sheet)?;
    
    // Get all cells
    println!("Getting all cells...");
    let cells = worksheet.get_cells()?;
    
    // Print the number of cells
    println!("Number of cells: {}", cells.len());
    
    // Print some sample cells
    println!("Sample cells:");
    for (i, cell) in cells.iter().take(5).enumerate() {
        println!("Cell {}: {:?}", i + 1, cell);
    }
    
    Ok(())
}