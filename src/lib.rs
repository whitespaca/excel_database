//! # excel_database
//!
//! A library that lets you perform CRUD operations on an Excel file (`.xlsx`) as if it were a simple database.
//! Internally, it uses `umya-spreadsheet` to read from and write to XLSX files.

use std::collections::HashMap;
use std::path::Path;

use serde::{Deserialize, Serialize};
use thiserror::Error;
use umya_spreadsheet::{Cell, CellValue as UCellValue, reader, writer, Worksheet};

/// Represents a cell's value. Currently, only text is supported.
/// You can extend this enum to include Number(f64), Bool(bool), Date(String), etc.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub enum CellValue {
    /// Text-based cell
    Text(String),
}

impl From<UCellValue> for CellValue {
    fn from(raw: UCellValue) -> Self {
        // Convert any underlying value to a String, then wrap in CellValue::Text
        let s = raw.get_value().unwrap_or_default().to_string();
        CellValue::Text(s)
    }
}

impl Into<UCellValue> for CellValue {
    fn into(self) -> UCellValue {
        match self {
            CellValue::Text(s) => UCellValue::from(s),
        }
    }
}

/// A Row is a mapping from column name (String) to its cell value (CellValue).
pub type Row = HashMap<String, CellValue>;

/// Errors that can occur when working with an ExcelDatabase.
#[derive(Debug, Error)]
pub enum ExcelDbError {
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),
    #[error("Spreadsheet parsing/writing error: {0}")]
    SpreadsheetError(#[from] umya_spreadsheet::reader::XlsxError),
    #[error("Sheet \"{0}\" not found")]
    SheetNotFound(String),
    #[error("No headers found in sheet \"{0}\"")]
    NoHeaders(String),
}

/// An in-memory representation of an Excel sheet, providing CRUD-like operations.
pub struct ExcelDatabase {
    file_path: String,
    sheet_name: String,
    data: Vec<Row>,
}

impl ExcelDatabase {
    /// Create a new ExcelDatabase by loading data from the given file path and sheet name.
    ///
    /// # Arguments
    ///
    /// * `file_path` - Path to the `.xlsx` file (e.g., `"data.xlsx"`).
    /// * `sheet_name` - The sheet name to use; if `None`, defaults to `"Sheet1"`.
    ///
    /// # Errors
    ///
    /// Returns `ExcelDbError::SheetNotFound` if the sheet does not exist, or
    /// `ExcelDbError::NoHeaders` if the sheet is empty (no header row).
    pub fn new<P: AsRef<Path>>(
        file_path: P,
        sheet_name: Option<String>,
    ) -> Result<Self, ExcelDbError> {
        let path_str = file_path.as_ref().to_string_lossy().to_string();
        let sheet = sheet_name.unwrap_or_else(|| "Sheet1".to_string());
        let data = Self::load_data(&path_str, &sheet)?;
        Ok(Self {
            file_path: path_str,
            sheet_name: sheet,
            data,
        })
    }

    /// Load all rows from the given sheet into memory (`Vec<Row>`).
    ///
    /// The first row of the sheet is treated as the header (column names).
    ///
    /// # Errors
    ///
    /// - `SheetNotFound(sheet_name)` if the sheet is not found.
    /// - `NoHeaders(sheet_name)` if the sheet has no rows at all.
    fn load_data(file_path: &str, sheet_name: &str) -> Result<Vec<Row>, ExcelDbError> {
        // Open the workbook
        let book = Reader::new().load_workbook(Path::new(file_path))?;
        if !book.has_sheet(sheet_name) {
            return Err(ExcelDbError::SheetNotFound(sheet_name.to_string()));
        }
        let worksheet = book.get_sheet_by_name(sheet_name).unwrap();

        // Collect each row as Vec<CellValue>
        let mut rows: Vec<Vec<CellValue>> = Vec::new();
        for row in worksheet.get_row_iter() {
            let mut row_vals: Vec<CellValue> = Vec::new();
            for cell in row.get_cell_iter() {
                let cv: CellValue = cell.get_value().unwrap_or_default().clone().into();
                row_vals.push(cv);
            }
            rows.push(row_vals);
        }

        // If there are no rows, we cannot infer headers
        if rows.is_empty() {
            return Err(ExcelDbError::NoHeaders(sheet_name.to_string()));
        }

        // The first row is interpreted as header names
        let headers: Vec<String> = rows[0]
            .iter()
            .map(|cv| match cv {
                CellValue::Text(text) => text.clone(),
            })
            .collect();

        // Convert subsequent rows into Row maps
        let mut data: Vec<Row> = Vec::new();
        for row_vals in rows.into_iter().skip(1) {
            let mut row_map: Row = HashMap::new();
            for (col_idx, header) in headers.iter().enumerate() {
                let value = row_vals
                    .get(col_idx)
                    .cloned()
                    .unwrap_or(CellValue::Text(String::new()));
                row_map.insert(header.clone(), value);
            }
            data.push(row_map);
        }

        Ok(data)
    }

    /// Save the current in-memory `data` back into the Excel file, overwriting the sheet.
    ///
    /// # Errors
    ///
    /// - `SheetNotFound(sheet_name)` if the sheet cannot be found when writing.
    /// - I/O or spreadsheet errors if the underlying write fails.
    fn save_data(&self) -> Result<(), ExcelDbError> {
        let mut book = Reader::new().load_workbook(Path::new(&self.file_path))?;
        if !book.has_sheet(&self.sheet_name) {
            return Err(ExcelDbError::SheetNotFound(self.sheet_name.clone()));
        }

        // Remove the existing sheet and create a fresh one
        book.remove_sheet_by_name(&self.sheet_name);
        let mut new_ws = Worksheet::new();

        // If no rows are present, add an empty sheet
        if self.data.is_empty() {
            book.add_worksheet(&self.sheet_name, new_ws);
            Writer::new(&book).save_as(Path::new(&self.file_path))?;
            return Ok(());
        }

        // Use the keys from the first Row as headers
        let headers: Vec<String> = self.data[0].keys().cloned().collect();

        // Write header row (row index 1 in Excel)
        for (col_idx, header) in headers.iter().enumerate() {
            let cell = Cell::new((col_idx + 1) as u32, 1, UCellValue::from(header.clone()));
            new_ws.add_cell(cell);
        }

        // Write actual data rows starting at Excel row 2
        for (row_idx, row_map) in self.data.iter().enumerate() {
            let excel_row = (row_idx + 2) as u32; // +2 because Excel is 1-based, and row 1 is header
            for (col_idx, header) in headers.iter().enumerate() {
                let value = row_map
                    .get(header)
                    .cloned()
                    .unwrap_or(CellValue::Text(String::new()));
                let cell_value: UCellValue = value.into();
                let cell = Cell::new((col_idx + 1) as u32, excel_row, cell_value);
                new_ws.add_cell(cell);
            }
        }

        // Add the rebuilt sheet and save the file
        book.add_worksheet(&self.sheet_name, new_ws);
        Writer::new(&book).save_as(Path::new(&self.file_path))?;
        Ok(())
    }

    /// Reload the sheet data from disk, replacing the in-memory `data`.
    ///
    /// # Errors
    ///
    /// Propagates any errors from `load_data`.
    fn refresh_data(&mut self) -> Result<(), ExcelDbError> {
        self.data = Self::load_data(&self.file_path, &self.sheet_name)?;
        Ok(())
    }

    // -------------------------------------------------------
    // Public API: CRUD, lookups, sheet/column management
    // -------------------------------------------------------

    /// Return all rows that match EVERY key-value pair in `query`.
    ///
    /// If `query` is `None`, returns all rows. Returns `None` if no rows match.
    pub fn select(&self, query: Option<&Row>) -> Option<Vec<Row>> {
        let mut result: Vec<Row> = Vec::new();
        let q = query.unwrap_or(&HashMap::new());
        'outer: for row in self.data.iter() {
            for (column, wanted) in q.iter() {
                if let Some(cell_val) = row.get(column) {
                    if cell_val != wanted {
                        continue 'outer;
                    }
                } else {
                    continue 'outer;
                }
            }
            result.push(row.clone());
        }
        if result.is_empty() {
            None
        } else {
            Some(result)
        }
    }

    /// Find the first row where `search_column == search_value` and return that row's `target_column` value.
    pub fn get_column_value(
        &self,
        search_column: &str,
        search_value: &CellValue,
        target_column: &str,
    ) -> Option<CellValue> {
        for row in self.data.iter() {
            if let Some(val) = row.get(search_column) {
                if val == search_value {
                    return row.get(target_column).cloned();
                }
            }
        }
        None
    }

    /// Insert a new row into the in-memory data and immediately save to the Excel file.
    ///
    /// # Errors
    ///
    /// Propagates any error from `save_data`.
    pub fn insert(&mut self, new_row: Row) -> Result<(), ExcelDbError> {
        self.data.push(new_row);
        self.save_data()?;
        Ok(())
    }

    /// Update all rows matching `query` by merging in `update_data`, then save.
    ///
    /// # Errors
    ///
    /// Propagates any error from `save_data`.
    pub fn update(&mut self, query: &Row, update_data: &Row) -> Result<(), ExcelDbError> {
        for row in self.data.iter_mut() {
            let mut is_match = true;
            for (key, val) in query.iter() {
                if let Some(cell_val) = row.get(key) {
                    if cell_val != val {
                        is_match = false;
                        break;
                    }
                } else {
                    is_match = false;
                    break;
                }
            }
            if is_match {
                for (u_key, u_val) in update_data.iter() {
                    row.insert(u_key.clone(), u_val.clone());
                }
            }
        }
        self.save_data()?;
        Ok(())
    }

    /// Delete all rows that match `query`, then save back to the file.
    ///
    /// # Errors
    ///
    /// Propagates any error from `save_data`.
    pub fn delete(&mut self, query: &Row) -> Result<(), ExcelDbError> {
        self.data.retain(|row| {
            for (key, val) in query.iter() {
                if let Some(cell_val) = row.get(key) {
                    if cell_val != val {
                        return true; // keep this row
                    }
                } else {
                    return true; // keep if the column isn't present
                }
            }
            false // if all key-value pairs matched, drop this row
        });
        self.save_data()?;
        Ok(())
    }

    /// Add a new sheet with the given name. If `initial_data` is provided and non-empty,
    /// its first Row is used for headers, and subsequent Rows fill the sheet.
    ///
    /// # Errors
    ///
    /// - `SheetNotFound` if a sheet with that name already exists (to avoid overwriting).
    /// - Propagates I/O or spreadsheet errors if writing fails.
    pub fn add_sheet(
        &self,
        new_sheet_name: &str,
        initial_data: Option<Vec<Row>>,
    ) -> Result<(), ExcelDbError> {
        let mut book = Reader::new().load_workbook(Path::new(&self.file_path))?;
        if book.has_sheet(new_sheet_name) {
            return Err(ExcelDbError::SheetNotFound(new_sheet_name.to_string()));
        }
        let mut ws = Worksheet::new();

        if let Some(rows) = initial_data {
            if !rows.is_empty() {
                // Use keys from the first row as headers
                let headers: Vec<String> = rows[0].keys().cloned().collect();
                // Write header row
                for (col_idx, header) in headers.iter().enumerate() {
                    let cell = Cell::new((col_idx + 1) as u32, 1, UCellValue::from(header.clone()));
                    ws.add_cell(cell);
                }
                // Write the data rows
                for (row_idx, row_map) in rows.iter().enumerate() {
                    let excel_row = (row_idx + 2) as u32; // Start at row 2
                    for (col_idx, header) in headers.iter().enumerate() {
                        let value = row_map
                            .get(header)
                            .cloned()
                            .unwrap_or(CellValue::Text(String::new()));
                        let cell: Cell =
                            Cell::new((col_idx + 1) as u32, excel_row, UCellValue::from(match value {
                                CellValue::Text(s) => s,
                            }));
                        ws.add_cell(cell);
                    }
                }
            }
        }

        book.add_worksheet(new_sheet_name, ws);
        Writer::new(&book).save_as(Path::new(&self.file_path))?;
        Ok(())
    }

    /// Check if a sheet with `sheet_name` exists in the workbook.
    ///
    /// # Errors
    ///
    /// Propagates any I/O or spreadsheet parsing errors.
    pub fn is_sheet_exists(&self, sheet_name: &str) -> Result<bool, ExcelDbError> {
        let book = Reader::new().load_workbook(Path::new(&self.file_path))?;
        Ok(book.has_sheet(sheet_name))
    }

    /// Get a list of all sheet names in the Excel file.
    ///
    /// # Errors
    ///
    /// Propagates any I/O or spreadsheet parsing errors.
    pub fn get_all_sheet_names(&self) -> Result<Vec<String>, ExcelDbError> {
        let book = Reader::new().load_workbook(Path::new(&self.file_path))?;
        Ok(book.get_sheet_names().to_vec())
    }

    /// Count how many non-empty values exist in the specified column across all rows.
    pub fn get_column_datas_number(&self, column_name: &str) -> usize {
        self.data
            .iter()
            .filter(|row| {
                if let Some(CellValue::Text(s)) = row.get(column_name) {
                    !s.trim().is_empty()
                } else {
                    false
                }
            })
            .count()
    }

    /// Add a new column with the given default value (or empty string if `None`).
    /// Only rows that do not already have this column get the default.
    ///
    /// # Errors
    ///
    /// Propagates any I/O or spreadsheet errors from `save_data`.
    pub fn add_column(
        &mut self,
        column_name: &str,
        default_value: Option<CellValue>,
    ) -> Result<(), ExcelDbError> {
        let default_val = default_value.unwrap_or(CellValue::Text(String::new()));
        for row in self.data.iter_mut() {
            row.entry(column_name.to_string())
                .or_insert_with(|| default_val.clone());
        }
        self.save_data()?;
        Ok(())
    }

    /// Remove a column from every row, if it exists, and save back to the file.
    ///
    /// # Errors
    ///
    /// Propagates any I/O or spreadsheet errors from `save_data`.
    pub fn remove_column(&mut self, column_name: &str) -> Result<(), ExcelDbError> {
        for row in self.data.iter_mut() {
            row.remove(column_name);
        }
        self.save_data()?;
        Ok(())
    }
}