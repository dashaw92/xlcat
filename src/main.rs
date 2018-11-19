extern crate calamine;

use std::env;
use std::io::{self, Write};
use calamine::{Reader, open_workbook_auto, DataType};

fn main() {
    let path = env::args().nth(1).expect("Missing path to Excel workbook.");
    let mut book = open_workbook_auto(path).expect("Not an Excel workbook.");
    let sheets = book.sheet_names().to_owned();
    let stdout = io::stdout();
    let mut lock = stdout.lock();

    //Iterate over all sheets in the workbook
    //Print each cell in every sheet.
    for sheet in sheets {
        if let Some(Ok(range)) = book.worksheet_range(&sheet) {
            writeln!(lock, "\nSheet: {}", sheet);
            for row in range.rows() {
                for cell in row.iter() {
                    let _ = match *cell {
                        DataType::Empty => write!(lock, "\t"),
                        DataType::Error(ref e) => write!(lock, "{:?}\t", e),
                        _ => write!(lock, "{}\t", cell),
                    };
                }
                writeln!(lock, "\n");
            }
        }
    }
}