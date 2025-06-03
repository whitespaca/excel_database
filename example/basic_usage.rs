//! A simple example of how to use excel_database.
//! Run with: `cargo run --example basic_usage`
//!
//! Make sure you have an `example_data.xlsx` in the current directory and that it has
//! a header row on "Sheet1" before running this example.

use excel_database::{CellValue, ExcelDatabase, Row};
use std::collections::HashMap;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // 1) Create an ExcelDatabase instance for "example_data.xlsx".
    //    If no sheet name is provided, it defaults to "Sheet1".
    let mut db = ExcelDatabase::new("example_data.xlsx", None)?;

    // 2) SELECT example: Find all rows where "name" column equals "John Doe".
    let mut select_query: Row = HashMap::new();
    select_query.insert("name".to_string(), CellValue::Text("John Doe".to_string()));
    match db.select(Some(&select_query)) {
        Some(rows) => {
            println!("SELECT results:");
            for row in rows {
                println!("{:?}", row);
            }
        }
        None => {
            println!("No matching rows found for select query.");
        }
    }

    // 3) INSERT example: Add a new row with name="Jane Doe", age="30", city="New York".
    let mut new_row: Row = HashMap::new();
    new_row.insert("name".to_string(), CellValue::Text("Jane Doe".to_string()));
    new_row.insert("age".to_string(), CellValue::Text("30".to_string()));
    new_row.insert("city".to_string(), CellValue::Text("New York".to_string()));
    db.insert(new_row)?;
    println!("Inserted new row for Jane Doe.");

    // 4) UPDATE example: Update age to "31" for rows where name="Jane Doe".
    let mut update_query: Row = HashMap::new();
    update_query.insert("name".to_string(), CellValue::Text("Jane Doe".to_string()));
    let mut update_data: Row = HashMap::new();
    update_data.insert("age".to_string(), CellValue::Text("31".to_string()));
    db.update(&update_query, &update_data)?;
    println!("Updated Jane Doe's age to 31.");

    // 5) DELETE example: Delete any row where name="John Doe".
    let mut delete_query: Row = HashMap::new();
    delete_query.insert("name".to_string(), CellValue::Text("John Doe".to_string()));
    db.delete(&delete_query)?;
    println!("Deleted rows where name was John Doe.");

    // 6) get_column_value example: Retrieve the "city" value for the row where name="Jane Doe".
    if let Some(city_val) = db.get_column_value(
        "name",
        &CellValue::Text("Jane Doe".to_string()),
        "city",
    ) {
        println!("Jane Doe's city: {:?}", city_val);
    } else {
        println!("No city found for Jane Doe.");
    }

    // 7) add_sheet example: Create a new sheet named "Sheet2" with two initial rows.
    let mut row1: Row = HashMap::new();
    row1.insert("name".to_string(), CellValue::Text("Alice".to_string()));
    row1.insert("age".to_string(), CellValue::Text("25".to_string()));
    let mut row2: Row = HashMap::new();
    row2.insert("name".to_string(), CellValue::Text("Bob".to_string()));
    row2.insert("age".to_string(), CellValue::Text("30".to_string()));
    db.add_sheet("Sheet2", Some(vec![row1.clone(), row2.clone()]))?;
    println!("Added new sheet 'Sheet2' with initial data.");

    // 8) is_sheet_exists example: Check if "Sheet1" exists.
    if db.is_sheet_exists("Sheet1")? {
        println!("Sheet1 exists in the workbook.");
    } else {
        println!("Sheet1 does not exist.");
    }

    // 9) get_all_sheet_names example: List all sheet names.
    let sheet_names = db.get_all_sheet_names()?;
    println!("All sheet names in the file: {:?}", sheet_names);

    // 10) get_column_datas_number example: Count non-empty "name" column entries.
    let count = db.get_column_datas_number("name");
    println!("Number of non-empty 'name' cells: {}", count);

    // 11) add_column example: Add a new column "email" with default empty text.
    db.add_column("email", Some(CellValue::Text(String::new())))?;
    println!("Added 'email' column to all rows (default empty).");

    // 12) remove_column example: Remove the "age" column from all rows.
    db.remove_column("age")?;
    println!("Removed 'age' column from all rows.");

    Ok(())
}
