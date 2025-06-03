# excel_database (Rust)

This crate allows you to treat an Excel (`.xlsx`) file like a simple database, performing CRUD (Create, Read, Update, Delete) operations on its rows. It is a direct port of the TypeScript version ([whitespaca/excel-database](https://github.com/whitespaca/excel-database)) into Rust, using the `umya-spreadsheet` crate under the hood.

## Features

- **ExcelDatabase struct**  
  - `new(file_path: &str, sheet_name: Option<String>) -> Result<ExcelDatabase, ExcelDbError>`  
    - Loads data from the specified file and sheet (defaults to `"Sheet1"` if omitted).  
  - **CRUD operations**  
    - `select(query: Option<&Row>) -> Option<Vec<Row>>`  
    - `insert(new_row: Row) -> Result<(), ExcelDbError>`  
    - `update(query: &Row, update_data: &Row) -> Result<(), ExcelDbError>`  
    - `delete(query: &Row) -> Result<(), ExcelDbError>`  
  - **Column lookup**  
    - `get_column_value(search_column: &str, search_value: &CellValue, target_column: &str) -> Option<CellValue>`  
  - **Sheet management**  
    - `add_sheet(new_sheet_name: &str, initial_data: Option<Vec<Row>>) -> Result<(), ExcelDbError>`  
    - `is_sheet_exists(sheet_name: &str) -> Result<bool, ExcelDbError>`  
    - `get_all_sheet_names() -> Result<Vec<String>, ExcelDbError>`  
  - **Column statistics**  
    - `get_column_datas_number(column_name: &str) -> usize`  
  - **Column manipulation**  
    - `add_column(column_name: &str, default_value: Option<CellValue>) -> Result<(), ExcelDbError>`  
    - `remove_column(column_name: &str) -> Result<(), ExcelDbError>`

- **Error handling**  
  - `ExcelDbError` enum for various I/O, spreadsheet parsing/writing, or “sheet not found” errors.

- **Row and CellValue types**  
  - `Row = HashMap<String, CellValue>`  
  - `CellValue` currently supports only `Text(String)`, but you can extend it to support numbers, booleans, dates, etc.

## Installation

If you plan to publish on crates.io, you could install via:

```bash
cargo add excel_database
```

Otherwise, to use from a local path, add to your `Cargo.toml`:

```toml
[dependencies]
excel_database = { path = "/path/to/excel_database" }
```

## Usage Example

Take a look at `examples/basic_usage.rs`. To run it:

```bash
# From the root of this repository
cargo run --example basic_usage
```

Make sure you have an `.xlsx` file (e.g., `example_data.xlsx`) in the working directory, with a header row in `"Sheet1"`.

## License

This project is licensed under the MIT License. See [LICENSE] for details.
