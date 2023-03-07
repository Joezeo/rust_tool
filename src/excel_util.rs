#![allow(dead_code)]
use std::{collections::HashMap, path::Path};

use umya_spreadsheet::{Cell, Style, Color};

struct MatchedCell {
    val: String,
    col: u32,
    row: u32,
    matched: bool,
}

impl MatchedCell {
    pub fn new(cell: &Cell) -> Option<Self> {
        let val = cell.get_cell_value().get_value();
        if let Ok(_) = val.parse::<f64>() {
            Some(MatchedCell {
                val: val.to_string(),
                col: *cell.get_coordinate().get_col_num(),
                row: *cell.get_coordinate().get_row_num(),
                matched: false,
            })
        } else {
            None
        }
    }
}

pub fn read_and_compare_rows(path: &str) {
    let path = Path::new(path);
    let mut book = umya_spreadsheet::reader::xlsx::read(path).unwrap();
    let mut changed = false;
    let sheet_cnt = book.get_sheet_count();

    for i in 0..sheet_cnt {
        let mut map_1 = HashMap::new();
        let mut map_2 = HashMap::new();
        let mut need_mark = vec![];

        let sheet = book.get_sheet_mut(&i).unwrap();
        for row in sheet.get_row_dimensions() {
            let datas = sheet.get_collection_by_row(row.get_row_num());
            let cell_1 = *datas.get(0).unwrap();
            let cell_2 = *datas.get(1).unwrap();

            if let Some(cell_1) = MatchedCell::new(cell_1) {
                if cell_1.col == 1 {
                    let vec = map_1.entry(cell_1.val.clone()).or_insert(vec![]);
                    vec.push(cell_1);
                } else {
                    let vec = map_2.entry(cell_1.val.clone()).or_insert(vec![]);
                    vec.push(cell_1);
                }
            }

            if let Some(cell_2) = MatchedCell::new(cell_2) {
                if cell_2.col == 1 {
                    let vec = map_1.entry(cell_2.val.clone()).or_insert(vec![]);
                    vec.push(cell_2);
                } else {
                    let vec = map_2.entry(cell_2.val.clone()).or_insert(vec![]);
                    vec.push(cell_2);
                }
            }
        }

        map_1.iter_mut().for_each(|(val, cells_1)| {
            if let Some(cells_2) = map_2.get_mut(val) {
                if cells_1.len() > cells_2.len() {
                    mark_cell(cells_1, cells_2);
                } else {
                    mark_cell(cells_2, cells_1);
                }
            }
        });

        map_1.iter().for_each(|(_, cells)| {
            for cell in cells.iter() {
                if !cell.matched {
                    need_mark.push(cell);
                }
            }
        });

        map_2.iter().for_each(|(_, cells)| {
            for cell in cells.iter() {
                if !cell.matched {
                    need_mark.push(cell);
                }
            }
        });

        if need_mark.len() > 0 {
            changed = true;
            for cell in need_mark {
                let mut style = Style::default();
                style.set_background_color(Color::COLOR_YELLOW);
                sheet.set_style((cell.col, cell.row), style);
            }
        }
    }

    if changed {
        let _ = umya_spreadsheet::writer::xlsx::write(&book, path);
    }
}

fn mark_cell(long: &mut [MatchedCell], short: &mut [MatchedCell]) {
    for i in 0..short.len() {
        long.get_mut(i).unwrap().matched = true;
        short.get_mut(i).unwrap().matched = true;
    }
}
