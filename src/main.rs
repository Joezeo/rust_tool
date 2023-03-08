#![windows_subsystem = "windows"]
use ui::UI;

mod ui;
mod excel_util;

fn main() {
    let mut files = Box::new(vec![]);
    UI::start(&mut files);
}