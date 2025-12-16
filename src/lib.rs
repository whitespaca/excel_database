use umya_spreadsheet::{
    // structs::Cell,
    // Worksheet,
    // reader::xlsx,
    writer::xlsx::write,
};
use std::collections::HashMap;
use serde_json::Value;
use std::path::PathBuf;
use calamine::{open_workbook_auto, Reader, Data};

pub type Row = HashMap<String, Value>;

pub struct ExcelDatabase {
    file_path: PathBuf,
    sheet_name: String,
    data: Vec<Row>,
}

impl ExcelDatabase {
    pub fn new<P: Into<PathBuf>>(file_path: P, sheet_name: Option<&str>) -> Result<Self, Box<dyn std::error::Error>> {
        let sheet = sheet_name.unwrap_or("Sheet1").to_string();
        let path = file_path.into();
        let data = Self::load_data_internal(&path, &sheet)?;

        Ok(Self {
            file_path: path,
            sheet_name: sheet,
            data,
        })
    }

    fn load_data_internal(
        file_path: &PathBuf,
        sheet_name: &str,
    ) -> Result<Vec<Row>, Box<dyn std::error::Error>> {
        let mut workbook = open_workbook_auto(file_path)?;
        let range = workbook
            .worksheet_range(sheet_name)
            .map_err(|_| "Sheet not found")?;

        let mut rows = range.rows();

        let headers: Vec<String> = rows
            .next()
            .ok_or("Empty sheet")?
            .iter()
            .map(|c| c.to_string())
            .collect();

        let mut result = Vec::new();

        for row in rows {
            let mut map = Row::new();

            for (header, cell) in headers.iter().zip(row.iter()) {
                map.insert(header.clone(), Self::data_type_to_json(cell));
            }

            result.push(map);
        }

        Ok(result)
    }

    fn data_type_to_json(cell: &Data) -> Value {
        match cell {
            Data::Int(v) => Value::from(*v),
            Data::Float(v) => Value::from(*v),
            Data::Bool(v) => Value::from(*v),
            Data::String(v) => Value::from(v.clone()),
            _ => Value::Null,
        }
    }

    pub fn save_data(&self) -> Result<(), Box<dyn std::error::Error>> {
        let mut book = umya_spreadsheet::new_file();
        let sheet = book.new_sheet(&self.sheet_name)?;

        if self.data.is_empty() {
            write(&book, &self.file_path)?;
            return Ok(());
        }

        // 헤더
        let headers: Vec<String> = self.data[0].keys().cloned().collect();

        for (col, header) in headers.iter().enumerate() {
            let coordinate = format!("{}1",(col + 1) as u32);
            sheet.get_cell_mut(coordinate).set_value(header);
        }

        // 데이터
        for (row_idx, row) in self.data.iter().enumerate() {
            for (col_idx, header) in headers.iter().enumerate() {
                let coordinate = format!("{}{}",
                    (col_idx + 1) as u32,
                    (row_idx + 2) as u32
                );

                let value = row
                    .get(header)
                    .map(|v| v.to_string())
                    .unwrap_or_default();

                sheet.get_cell_mut(coordinate).set_value(value);
            }
        }

        write(&book, &self.file_path)?;
        Ok(())
    }
}