use excel_reader::excel::Excel;
use std::time::Instant;

fn main() -> anyhow::Result<()> {
    // Open the large Excel file
    println!("Opening Excel file...");
    let start = Instant::now();
    let mut excel = Excel::from_path(r"E:\share\tauri-excel\template_20260207_20260207.xlsx")?;
    println!("Excel file opened in {:?}", start.elapsed());
    
    // Get all sheets
    println!("Getting sheets...");
    let start = Instant::now();
    let sheets = excel.get_sheets()?;
    println!("Sheets retrieved in {:?}", start.elapsed());
    println!("Number of sheets: {}", sheets.len());
    
    // Get the first worksheet
    let Some(first_sheet) = sheets.first() else {
        anyhow::bail!("No worksheet found");
    };
    println!("First sheet name: {}", first_sheet.name);
    
    // Get the complete worksheet object
    println!("Getting worksheet...");
    let start = Instant::now();
    let worksheet = excel.get_worksheet(first_sheet)?;
    println!("Worksheet retrieved in {:?}", start.elapsed());
    
    // Get worksheet dimension
    if let Some(dimension) = worksheet.dimension {
        println!("Worksheet dimension: {:?}", dimension);
        let num_rows = dimension.end.row - dimension.start.row + 1;
        let num_cols = dimension.end.col - dimension.start.col + 1;
        println!("Number of rows: {}", num_rows);
        println!("Number of columns: {}", num_cols);
        println!("Estimated number of cells: {}", num_rows * num_cols);
    }
    
    // Get all cells
    println!("Getting all cells...");
    let start = Instant::now();
    let cells = worksheet.get_cells()?;
    println!("All cells retrieved in {:?}", start.elapsed());
    
    // Print the number of cells
    println!("Number of cells: {}", cells.len());
    println!("Test completed successfully!");
    
    Ok(())
}