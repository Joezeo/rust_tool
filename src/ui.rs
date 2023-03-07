use std::{
    sync::atomic::{AtomicBool, Ordering, AtomicIsize},
    thread,
    time::Duration,
};
use windows::{
    core::{HSTRING, PCWSTR},
    w,
    Win32::{
        Foundation::{HWND, LPARAM, LRESULT, WPARAM, HINSTANCE},
        Graphics::Gdi::{InvalidateRect, UpdateWindow},
        System::LibraryLoader::GetModuleHandleW,
        UI::WindowsAndMessaging::{
            CreateWindowExW, DefWindowProcW, DispatchMessageW, GetMessageW, LoadCursorW,
            PeekMessageW, PostQuitMessage, RegisterClassW, ShowWindow, CS_HREDRAW, CS_VREDRAW,
            CW_USEDEFAULT, IDC_ARROW, MSG, PM_NOREMOVE, SW_SHOW, WINDOW_EX_STYLE, WM_CREATE,
            WNDCLASSW, WS_OVERLAPPEDWINDOW, WS_VISIBLE, WS_CHILD, HMENU, WM_DESTROY, WM_MOVING, WS_THICKFRAME, WINDOW_STYLE, WM_COMMAND,
        },
    },
};

pub static mut RUNNING: AtomicBool = AtomicBool::new(true);
pub static mut HINS: AtomicIsize = AtomicIsize::new(0);
pub struct UI;

impl UI {
    pub fn start() {
        unsafe {
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
                    30, //按钮在界面上出现的位置
                    window,
                    HMENU(100), //设置按钮的ID，随便设置一个数即可
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
                    30, //按钮在界面上出现的位置
                    window,
                    HMENU(150), //设置按钮的ID，随便设置一个数即可
                    HINSTANCE(HINS.load(Ordering::SeqCst)),
                    None,
                );
                LRESULT(1)
            },
            WM_COMMAND => {
                if wparam.0 == 100 {
                    println!("Button clicked");
                }
                LRESULT(1)
            }
            WM_DESTROY => {
                PostQuitMessage(0);
                RUNNING.store(false, Ordering::SeqCst);
                LRESULT(0)
            },
            WM_MOVING => {
                InvalidateRect(window, None, true);
                UpdateWindow(window);
                LRESULT(0)
            },
            _ => DefWindowProcW(window, message, wparam, lparam),
        }
    }
}
