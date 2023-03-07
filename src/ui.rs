use std::{
    mem::size_of,
    sync::atomic::{AtomicBool, AtomicIsize, Ordering, AtomicPtr},
    thread,
    time::Duration, ptr::null_mut,
};
use widestring::U16String;
use windows::{
    core::{HSTRING, PCWSTR, PWSTR},
    w,
    Win32::{
        Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, MAX_PATH, WPARAM},
        Graphics::Gdi::{InvalidateRect, UpdateWindow},
        System::LibraryLoader::GetModuleHandleW,
        UI::{
            Controls::Dialogs::{
                GetOpenFileNameW, OFN_ALLOWMULTISELECT, OFN_EXPLORER, OFN_FILEMUSTEXIST,
                OFN_NOCHANGEDIR, OFN_PATHMUSTEXIST, OPENFILENAMEW,
            },
            WindowsAndMessaging::{
                CreateWindowExW, DefWindowProcW, DispatchMessageW, GetMessageW, LoadCursorW,
                PeekMessageW, PostQuitMessage, RegisterClassW, ShowWindow, CS_HREDRAW, CS_VREDRAW,
                CW_USEDEFAULT, HMENU, IDC_ARROW, MSG, PM_NOREMOVE, SW_SHOW, WINDOW_EX_STYLE,
                WINDOW_STYLE, WM_COMMAND, WM_CREATE, WM_DESTROY, WM_MOVING, WNDCLASSW, WS_CHILD,
                WS_OVERLAPPEDWINDOW, WS_THICKFRAME, WS_VISIBLE, WM_ERASEBKGND, WM_PAINT,
            },
        },
    },
};

use crate::excel_util;

pub static mut RUNNING: AtomicBool = AtomicBool::new(true);
pub static mut HINS: AtomicIsize = AtomicIsize::new(0);
pub static mut FILES: AtomicPtr<Vec<String>> = AtomicPtr::new(null_mut());
pub struct UI;

impl UI {
    pub fn start(files: &mut Vec<String>) {
        unsafe {
            FILES.store(files, Ordering::SeqCst);
            let hins = GetModuleHandleW(None).unwrap();
            assert!(hins.0 != 0);
            HINS.store(hins.0, Ordering::SeqCst);

            let wcls = WNDCLASSW {
                hCursor: LoadCursorW(None, IDC_ARROW).unwrap(),
                hInstance: hins,
                lpszClassName: w!("TmuiMainClass"),

                style: CS_HREDRAW | CS_VREDRAW,
                lpfnWndProc: Some(wndproc),
                ..Default::default()
            };

            let atom = RegisterClassW(&wcls);
            assert!(atom != 0);

            let hwnd = CreateWindowExW(
                WINDOW_EX_STYLE::default(),
                w!("TmuiMainClass"),
                PCWSTR(HSTRING::from("工具").as_ptr()),
                WS_VISIBLE | WINDOW_STYLE(WS_OVERLAPPEDWINDOW.0 ^ WS_THICKFRAME.0),
                CW_USEDEFAULT,
                CW_USEDEFAULT,
                400,
                150,
                None,
                None,
                hins,
                None,
            );

            ShowWindow(hwnd, SW_SHOW);
            UpdateWindow(hwnd);

            while RUNNING.load(Ordering::SeqCst) {
                let mut message = MSG::default();
                if PeekMessageW(&mut message, hwnd, 0, 0, PM_NOREMOVE).into() {
                    if GetMessageW(&mut message, hwnd, 0, 0).into() {
                        DispatchMessageW(&message);
                    }
                }
                thread::sleep(Duration::from_nanos(1));
            }
        }
    }
}

extern "system" fn wndproc(window: HWND, message: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe {
        match message {
            WM_CREATE => {
                CreateWindowExW(
                    WINDOW_EX_STYLE::default(),
                    w!("button"),
                    w!("选择文件"),
                    WS_CHILD | WS_VISIBLE,
                    20,
                    20,
                    70,
                    30, 
                    window,
                    HMENU(100), 
                    HINSTANCE(HINS.load(Ordering::SeqCst)),
                    None,
                );
                CreateWindowExW(
                    WINDOW_EX_STYLE::default(),
                    w!("button"),
                    w!("执行"),
                    WS_CHILD | WS_VISIBLE,
                    20,
                    70,
                    70,
                    30, 
                    window,
                    HMENU(150), 
                    HINSTANCE(HINS.load(Ordering::SeqCst)),
                    None,
                );
                LRESULT(1)
            }
            WM_PAINT => {
                LRESULT(1)
            }
            WM_ERASEBKGND => {
                LRESULT(0)
            }
            WM_COMMAND => {
                match wparam.0 {
                    100 => {
                        choose_files(window);
                    }
                    150 => {
                        let files = FILES.load(Ordering::SeqCst).as_mut().unwrap();
                        for file in files.iter() {
                            excel_util::read_and_compare_rows(file);
                        }
                    }
                    _ => panic!()
                }
                LRESULT(1)
            }
            WM_DESTROY => {
                PostQuitMessage(0);
                RUNNING.store(false, Ordering::SeqCst);
                LRESULT(0)
            }
            WM_MOVING => {
                InvalidateRect(window, None, true);
                UpdateWindow(window);
                LRESULT(1)
            }
            _ => DefWindowProcW(window, message, wparam, lparam),
        }
    }
}

fn choose_files(hwnd: HWND) {
    let mut open_file_names = [0u16; MAX_PATH as usize * 80];
    let mut p = 0;

    let mut open = OPENFILENAMEW::default();
    open.hwndOwner = hwnd;
    open.lStructSize = size_of::<OPENFILENAMEW>() as u32;
    open.lpstrFile = PWSTR(&mut open_file_names as *mut u16);
    open.nMaxFile = open_file_names.len() as u32;
    open.nFilterIndex = 1;
    open.lpstrFileTitle = PWSTR::null();
    open.nMaxFileTitle = 0;
    open.Flags = OFN_EXPLORER
        | OFN_PATHMUSTEXIST
        | OFN_FILEMUSTEXIST
        | OFN_NOCHANGEDIR
        | OFN_ALLOWMULTISELECT;

    unsafe {
        if GetOpenFileNameW(&mut open).as_bool() {
            let offset = open.nFileOffset as usize;
            let mut path = U16String::from_vec(open_file_names[0..offset-1].to_vec()).to_string().unwrap();
            if !path.ends_with("\\") {
                path.push_str("\\")
            }
            let file_names = open_file_names[offset..].to_vec();

            let files = FILES.load(Ordering::SeqCst).as_mut().unwrap();
            files.clear();
            let mut last_c = 1;
            for (idx, c) in file_names.iter().enumerate() {
                if *c == 0 {
                    if last_c == 0 {
                        break;
                    }
                    let mut file = path.clone();
                    file.push_str(U16String::from_vec(file_names[p..idx].to_vec()).to_string().unwrap().as_str());
                    files.push(file);
                    p = idx + 1;
                }
                last_c = *c;
            }

            println!("{:?}", files);
        }
    }
}
