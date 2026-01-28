use excel_reader::{excel::Excel};

/// Demo for data validation parsing
fn main() -> anyhow::Result<()> {
    let path = "examples/template.xlsx";

    // excel directly from path
    let mut excel = Excel::from_path(path)?;

    // get basic sheet info
    let sheets = excel.get_sheets()?;
    if sheets.is_empty() {
        println!("Excel contains no sheets");
        return Ok(());
    }

    // get worksheet detail
    let worksheet = excel.get_worksheet(&sheets[0].clone())?;
    println!("worksheet: {}", worksheet.name);

    // test data validation parsing
    if let Some(data_validations) = &worksheet.data_validations {
        println!("\ndata validations: ");
        for (index, dv) in data_validations.iter().enumerate() {
            println!("--------");
            println!("{}: DataValidation", index + 1);
            println!("type: {}", dv.r#type);
            println!("sqref: {}", dv.sqref);
            println!("allow_blank: {}", dv.allow_blank);
            println!("show_drop_down: {}", dv.show_drop_down);
            println!("show_error_message: {}", dv.show_error_message);
            println!("show_input_message: {}", dv.show_input_message);
            if let Some(formula1) = &dv.formula1 {
                println!("formula1: {}", formula1);
            }
            if let Some(formula2) = &dv.formula2 {
                println!("formula2: {}", formula2);
            }
            if let Some(operator) = &dv.operator {
                println!("operator: {}", operator);
            }
            if let Some(prompt_title) = &dv.prompt_title {
                println!("prompt_title: {}", prompt_title);
            }
            if let Some(prompt) = &dv.prompt {
                println!("prompt: {}", prompt);
            }
            if let Some(error_title) = &dv.error_title {
                println!("error_title: {}", error_title);
            }
            if let Some(error_message) = &dv.error_message {
                println!("error_message: {}", error_message);
            }
        }
        println!("--------");
    } else {
        println!("\nNo data validations found in worksheet");
    }

    Ok(())
}
