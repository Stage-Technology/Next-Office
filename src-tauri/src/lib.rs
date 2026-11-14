mod file_types;

use std::collections::HashMap;
use std::path::{Path, PathBuf};

use percent_encoding::percent_decode_str;
use tauri::ipc::{InvokeBody, Request, Response};
use tauri::{AppHandle, Manager};

/// File passed on the command line, e.g. when a .nxslides file is double-clicked in Explorer.
#[tauri::command]
fn launch_file() -> Option<String> {
    std::env::args_os()
        .skip(1)
        .map(PathBuf::from)
        .find(|p| p.is_file())
        .map(|p| p.to_string_lossy().into_owned())
}

#[tauri::command]
fn read_file(path: String) -> Result<Response, String> {
    std::fs::read(&path)
        .map(Response::new)
        .map_err(|e| format!("Could not read {path}: {e}"))
}

/// Writes the raw request body to the path given in the (URI-encoded) `path` header.
/// Binary bodies avoid serialising large decks as JSON number arrays.
#[tauri::command]
fn write_file(request: Request<'_>) -> Result<(), String> {
    let InvokeBody::Raw(data) = request.body() else {
        return Err("Expected binary data".into());
    };
    let path = request
        .headers()
        .get("path")
        .and_then(|v| v.to_str().ok())
        .map(|v| percent_decode_str(v).decode_utf8_lossy().into_owned())
        .ok_or("Missing path")?;
    write_atomic(Path::new(&path), data).map_err(|e| format!("Could not save {path}: {e}"))
}

/// Write to a temp file first so a failed save never truncates the existing file.
fn write_atomic(path: &Path, data: &[u8]) -> std::io::Result<()> {
    let mut tmp = path.as_os_str().to_owned();
    tmp.push(".saving");
    std::fs::write(&tmp, data)
        .and_then(|_| std::fs::rename(&tmp, path))
        .inspect_err(|_| {
            let _ = std::fs::remove_file(&tmp);
        })
}

// ---- App state (autosave, recent files): one file per key in the app's data folder ----

fn state_dir(app: &AppHandle) -> Result<PathBuf, String> {
    app.path()
        .app_local_data_dir()
        .map(|d| d.join("state"))
        .map_err(|e| e.to_string())
}

fn valid_key(key: &str) -> bool {
    !key.is_empty() && key.len() <= 64 && key.chars().all(|c| c.is_ascii_alphanumeric() || c == ':' || c == '-')
}

// ':' isn't allowed in Windows file names.
fn key_file(key: &str) -> String {
    format!("{}.json", key.replace(':', "~"))
}

#[tauri::command]
fn load_state(app: AppHandle) -> HashMap<String, String> {
    let mut out = HashMap::new();
    let Ok(dir) = state_dir(&app) else { return out };
    let Ok(entries) = std::fs::read_dir(dir) else { return out };
    for entry in entries.flatten() {
        let name = entry.file_name().to_string_lossy().into_owned();
        let Some(stem) = name.strip_suffix(".json") else { continue };
        let key = stem.replace('~', ":");
        if valid_key(&key) {
            if let Ok(value) = std::fs::read_to_string(entry.path()) {
                out.insert(key, value);
            }
        }
    }
    out
}

/// Synchronous on purpose: commands without `async` run in order on the main thread,
/// so an older autosave can never overwrite a newer one.
#[tauri::command]
fn save_state(app: AppHandle, key: String, value: String) -> Result<(), String> {
    if !valid_key(&key) {
        return Err(format!("Invalid key {key}"));
    }
    let dir = state_dir(&app)?;
    std::fs::create_dir_all(&dir).map_err(|e| e.to_string())?;
    write_atomic(&dir.join(key_file(&key)), value.as_bytes()).map_err(|e| e.to_string())
}

// ---- Media store: pictures/video/audio named by content hash ("<32 hex>.<ext>") ----

fn media_path(app: &AppHandle) -> Result<PathBuf, String> {
    let dir = app.path().app_local_data_dir().map_err(|e| e.to_string())?.join("media");
    std::fs::create_dir_all(&dir).map_err(|e| e.to_string())?;
    Ok(dir)
}

fn valid_media_name(name: &str) -> bool {
    let Some((hash, ext)) = name.split_once('.') else { return false };
    hash.len() == 32
        && hash.chars().all(|c| c.is_ascii_hexdigit() && !c.is_ascii_uppercase())
        && (1..=5).contains(&ext.len())
        && ext.chars().all(|c| c.is_ascii_lowercase() || c.is_ascii_digit())
}

/// Folder the page loads media from through the asset protocol (scoped to it in tauri.conf.json).
#[tauri::command]
fn media_dir(app: AppHandle) -> Result<String, String> {
    Ok(media_path(&app)?.to_string_lossy().into_owned())
}

#[tauri::command]
fn media_has(app: AppHandle, name: String) -> Result<bool, String> {
    if !valid_media_name(&name) {
        return Err(format!("Invalid media name {name}"));
    }
    Ok(media_path(&app)?.join(name).is_file())
}

#[tauri::command]
fn media_put(app: AppHandle, request: Request<'_>) -> Result<(), String> {
    let InvokeBody::Raw(data) = request.body() else {
        return Err("Expected binary data".into());
    };
    let name = request.headers().get("name").and_then(|v| v.to_str().ok()).unwrap_or_default();
    if !valid_media_name(name) {
        return Err(format!("Invalid media name {name}"));
    }
    let path = media_path(&app)?.join(name);
    if path.is_file() {
        return Ok(()); // same name = same content
    }
    write_atomic(&path, data).map_err(|e| e.to_string())
}

// ---- Window theme: the native window and title bar follow the app's light/dark theme ----

/// Background behind the page (--surface-100), so resizing or a slow load never shows a white flash.
fn apply_window_theme(window: &tauri::WebviewWindow, dark: bool) {
    let (theme, bg) = if dark {
        (tauri::Theme::Dark, tauri::window::Color(0x14, 0x14, 0x1d, 255))
    } else {
        (tauri::Theme::Light, tauri::window::Color(0xf6, 0xf6, 0xfb, 255))
    };
    let _ = window.set_theme(Some(theme));
    let _ = window.set_background_color(Some(bg));
}

#[tauri::command]
fn set_window_theme(window: tauri::WebviewWindow, dark: bool) {
    apply_window_theme(&window, dark);
}

/// The theme saved by the app (src/theme.ts); "system" or nothing saved follows Windows.
fn saved_theme_is_dark(app: &AppHandle, window: &tauri::WebviewWindow) -> bool {
    let saved = state_dir(app)
        .ok()
        .and_then(|d| std::fs::read_to_string(d.join(key_file("next-office:theme"))).ok());
    match saved.as_deref().map(str::trim) {
        Some("dark") => true,
        Some("light") => false,
        _ => matches!(window.theme(), Ok(tauri::Theme::Dark)),
    }
}

/// Ask DWM for Windows 11 rounded corners on our frameless window.
#[cfg(windows)]
fn round_corners(window: &tauri::WebviewWindow) {
    use windows_sys::Win32::Graphics::Dwm::{DwmSetWindowAttribute, DWMWA_WINDOW_CORNER_PREFERENCE, DWMWCP_ROUND};
    if let Ok(hwnd) = window.hwnd() {
        let pref = DWMWCP_ROUND;
        // Fails harmlessly on Windows 10, which has no rounded corners.
        unsafe {
            DwmSetWindowAttribute(
                hwnd.0 as _,
                DWMWA_WINDOW_CORNER_PREFERENCE as u32,
                &pref as *const _ as *const _,
                std::mem::size_of_val(&pref) as u32,
            );
        }
    }
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .setup(|app| {
            if let Some(w) = app.get_webview_window("main") {
                #[cfg(windows)]
                round_corners(&w);
                let dark = saved_theme_is_dark(app.handle(), &w);
                apply_window_theme(&w, dark);
            }
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            launch_file,
            set_window_theme,
            read_file,
            write_file,
            load_state,
            save_state,
            media_dir,
            media_has,
            media_put,
            file_types::register_file_types,
            file_types::default_apps_status,
            file_types::open_default_apps_settings
        ])
        .run(tauri::generate_context!())
        .expect("error while running Next Office");
}
