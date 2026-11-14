//! Registering Next Office with Windows as a program that opens its file types, so it appears in
//! "Open with" and in Settings › Default apps. Windows (8 and later) lets only the user change a
//! default, so this never takes a file type over: it registers, reports which types already open in
//! Next Office, and opens Next Office's page in Settings where the user can make it the default.
//!
//! Everything is written under HKEY_CURRENT_USER (no admin rights), matching the per-user installer.

use serde::Deserialize;

/// Name in Settings › Default apps (the value under RegisteredApplications).
pub const APP_NAME: &str = "Next Office";
/// Prefix of our ProgIDs: "NextOffice.xlsx".
pub const PROGID_PREFIX: &str = "NextOffice";
const CAPABILITIES: &str = r"Software\NextOffice\Capabilities";

#[derive(Debug, Deserialize)]
pub struct FileType {
    pub ext: String,
    pub name: String,
}

/// Extensions are short and plain ("xlsx"): they become parts of registry paths.
pub fn valid_ext(ext: &str) -> bool {
    (1..=12).contains(&ext.len()) && ext.chars().all(|c| c.is_ascii_lowercase() || c.is_ascii_digit())
}

/// Registry keys a registration writes, so tests and the uninstaller agree with the app on the layout.
pub struct Layout<'a> {
    pub app_name: &'a str,
    pub progid_prefix: &'a str,
    pub capabilities: &'a str,
}

pub const LAYOUT: Layout<'static> = Layout { app_name: APP_NAME, progid_prefix: PROGID_PREFIX, capabilities: CAPABILITIES };

#[cfg(windows)]
mod win {
    use super::{valid_ext, FileType, Layout};
    use std::path::Path;
    use windows_sys::Win32::Foundation::ERROR_SUCCESS;
    use windows_sys::Win32::System::Registry::{
        RegDeleteTreeW, RegGetValueW, RegSetKeyValueW, HKEY, HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, REG_SZ, RRF_RT_REG_SZ,
    };
    use windows_sys::Win32::UI::Shell::{AssocQueryStringW, SHChangeNotify, ShellExecuteW, ASSOCSTR_EXECUTABLE, SHCNE_ASSOCCHANGED, SHCNF_IDLIST};
    use windows_sys::Win32::UI::WindowsAndMessaging::SW_SHOWNORMAL;

    fn wide(s: &str) -> Vec<u16> {
        s.encode_utf16().chain(std::iter::once(0)).collect()
    }

    /// Set a string value, creating the key if needed. `name` None is the key's (Default) value.
    fn set(key: &str, name: Option<&str>, value: &str) -> Result<(), String> {
        let k = wide(key);
        let n = name.map(wide);
        let v = wide(value);
        let rc = unsafe {
            RegSetKeyValueW(
                HKEY_CURRENT_USER,
                k.as_ptr(),
                n.as_ref().map_or(std::ptr::null(), |n| n.as_ptr()),
                REG_SZ,
                v.as_ptr().cast(),
                (v.len() * 2) as u32,
            )
        };
        if rc == ERROR_SUCCESS {
            Ok(())
        } else {
            Err(format!(r"Could not write HKCU\{key} (error {rc})"))
        }
    }

    fn delete_tree(key: &str) {
        let k = wide(key);
        unsafe { RegDeleteTreeW(HKEY_CURRENT_USER, k.as_ptr()) };
    }

    pub fn register(layout: &Layout, exe: &Path, description: &str, types: &[FileType]) -> Result<(), String> {
        let exe = exe.to_string_lossy();
        let exe_name = Path::new(exe.as_ref()).file_name().map(|n| n.to_string_lossy().into_owned()).unwrap_or_default();
        let open = format!("\"{exe}\" \"%1\"");
        let icon = format!("\"{exe}\",0");

        // The list of types is rewritten each time, so types dropped from the app disappear here too.
        delete_tree(&format!(r"{}\FileAssociations", layout.capabilities));
        set(layout.capabilities, Some("ApplicationName"), layout.app_name)?;
        set(layout.capabilities, Some("ApplicationDescription"), description)?;
        set(layout.capabilities, Some("ApplicationIcon"), &icon)?;
        let app_key = format!(r"Software\Classes\Applications\{exe_name}");
        set(&app_key, Some("FriendlyAppName"), layout.app_name)?;
        set(&format!(r"{app_key}\shell\open\command"), None, &open)?;

        for t in types.iter().filter(|t| valid_ext(&t.ext)) {
            let progid = format!("{}.{}", layout.progid_prefix, t.ext);
            let pk = format!(r"Software\Classes\{progid}");
            set(&pk, None, &t.name)?;
            set(&pk, Some("FriendlyTypeName"), &t.name)?;
            set(&format!(r"{pk}\DefaultIcon"), None, &icon)?;
            set(&format!(r"{pk}\shell\open\command"), None, &open)?;
            // Offered under "Open with" without changing the current default.
            set(&format!(r"Software\Classes\.{}\OpenWithProgids", t.ext), Some(&progid), "")?;
            set(&format!(r"{app_key}\SupportedTypes"), Some(&format!(".{}", t.ext)), "")?;
            set(&format!(r"{}\FileAssociations", layout.capabilities), Some(&format!(".{}", t.ext)), &progid)?;
        }
        set(r"Software\RegisteredApplications", Some(layout.app_name), layout.capabilities)?;
        unsafe { SHChangeNotify(SHCNE_ASSOCCHANGED as i32, SHCNF_IDLIST, std::ptr::null(), std::ptr::null()) };
        Ok(())
    }

    /// A string value (None: the key's default value), if present and not empty.
    fn get(root: HKEY, key: &str, name: Option<&str>) -> Option<String> {
        let k = wide(key);
        let n = name.map(wide);
        let mut buf = vec![0u16; 1024];
        let mut size = (buf.len() * 2) as u32;
        let rc = unsafe {
            RegGetValueW(
                root,
                k.as_ptr(),
                n.as_ref().map_or(std::ptr::null(), |n| n.as_ptr()),
                RRF_RT_REG_SZ,
                std::ptr::null_mut(),
                buf.as_mut_ptr().cast(),
                &mut size,
            )
        };
        if rc != ERROR_SUCCESS {
            return None;
        }
        let end = buf.iter().position(|&c| c == 0).unwrap_or(buf.len());
        let v = String::from_utf16_lossy(&buf[..end]);
        (!v.is_empty()).then_some(v)
    }

    /// The ProgID Windows uses for ".ext" on a double-click: the user's choice ("Always use this app",
    /// Settings › Default apps), else the type's registered default. None means Windows would ask.
    pub fn effective_progid(ext: &str) -> Option<String> {
        let fe = format!(r"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.{ext}");
        get(HKEY_CURRENT_USER, &format!(r"{fe}\UserChoiceLatest"), Some("ProgId"))
            .or_else(|| get(HKEY_CURRENT_USER, &format!(r"{fe}\UserChoice"), Some("ProgId")))
            .or_else(|| get(HKEY_CLASSES_ROOT, &format!(".{ext}"), None))
    }

    /// The program a ProgID (or ".ext") starts with its open command.
    pub fn handler(assoc: &str) -> Option<String> {
        let assoc = wide(assoc);
        let verb = wide("open");
        let mut buf = vec![0u16; 1024];
        let mut len = buf.len() as u32;
        let hr = unsafe { AssocQueryStringW(0, ASSOCSTR_EXECUTABLE, assoc.as_ptr(), verb.as_ptr(), buf.as_mut_ptr(), &mut len) };
        if hr != 0 {
            return None;
        }
        let end = buf.iter().position(|&c| c == 0).unwrap_or(buf.len());
        Some(String::from_utf16_lossy(&buf[..end]))
    }

    pub fn open_uri(uri: &str) -> Result<(), String> {
        let op = wide("open");
        let file = wide(uri);
        let h = unsafe { ShellExecuteW(std::ptr::null_mut(), op.as_ptr(), file.as_ptr(), std::ptr::null(), std::ptr::null(), SW_SHOWNORMAL) };
        // ShellExecute reports success with a value above 32.
        if h as isize > 32 {
            Ok(())
        } else {
            Err(format!("Could not open {uri}"))
        }
    }

    #[cfg(test)]
    pub fn remove(layout: &Layout, exe_name: &str, exts: &[&str]) {
        use windows_sys::Win32::System::Registry::RegDeleteKeyValueW;
        for ext in exts {
            delete_tree(&format!(r"Software\Classes\{}.{ext}", layout.progid_prefix));
            // Test-only extensions: the whole key goes.
            delete_tree(&format!(r"Software\Classes\.{ext}"));
        }
        delete_tree(layout.capabilities.trim_end_matches(r"\Capabilities"));
        delete_tree(&format!(r"Software\Classes\Applications\{exe_name}"));
        let k = wide(r"Software\RegisteredApplications");
        let v = wide(layout.app_name);
        unsafe { RegDeleteKeyValueW(HKEY_CURRENT_USER, k.as_ptr(), v.as_ptr()) };
    }

    #[cfg(test)]
    pub fn set_for_test(key: &str, value: &str) {
        set(key, None, value).unwrap();
    }

    #[cfg(test)]
    pub fn delete_for_test(key: &str) {
        delete_tree(key);
    }
}

/// Same file, ignoring case and the \\?\ prefix Rust may add.
fn same_path(a: &str, b: &str) -> bool {
    let norm = |p: &str| p.trim_start_matches(r"\\?\").replace('/', "\\").to_lowercase();
    norm(a) == norm(b)
}

fn current_exe() -> Result<std::path::PathBuf, String> {
    std::env::current_exe().map_err(|e| e.to_string())
}

/// Register (or refresh) Next Office for these file types. Runs at every start, so a moved or updated
/// app and newly supported types are picked up.
#[tauri::command]
pub fn register_file_types(types: Vec<FileType>) -> Result<(), String> {
    #[cfg(windows)]
    {
        win::register(&LAYOUT, &current_exe()?, "Documents, presentations and spreadsheets: opens Word, PowerPoint and Excel files.", &types)
    }
    #[cfg(not(windows))]
    {
        let _ = types;
        Ok(())
    }
}

/// Whether double-clicking a ".ext" file starts this copy of Next Office.
#[cfg(windows)]
pub fn opens_with_us(ext: &str, exe: &std::path::Path) -> bool {
    valid_ext(ext)
        && win::effective_progid(ext)
            .and_then(|progid| win::handler(&progid))
            .is_some_and(|h| same_path(&h, &exe.to_string_lossy()))
}

/// For each extension: does it open in this copy of Next Office now?
#[tauri::command]
pub fn default_apps_status(exts: Vec<String>) -> Result<Vec<bool>, String> {
    #[cfg(windows)]
    {
        let exe = current_exe()?;
        Ok(exts.iter().map(|e| opens_with_us(e, &exe)).collect())
    }
    #[cfg(not(windows))]
    {
        Ok(vec![false; exts.len()])
    }
}

/// Windows 11 opens Next Office's own page (with its "Set default" button); older versions open
/// the Default apps list.
#[tauri::command]
pub fn open_default_apps_settings() -> Result<(), String> {
    #[cfg(windows)]
    {
        let name: String = APP_NAME.chars().map(|c| if c == ' ' { "%20".to_string() } else { c.to_string() }).collect();
        win::open_uri(&format!("ms-settings:defaultapps?registeredAppUser={name}"))
    }
    #[cfg(not(windows))]
    {
        Err("Only available on Windows".into())
    }
}

#[cfg(all(test, windows))]
mod tests {
    use super::*;
    use std::process::Command;

    /// A separate name and ProgID prefix, so the test never touches the real Next Office registration.
    const TEST: Layout<'static> = Layout {
        app_name: "Next Office QA Test",
        progid_prefix: "NextOfficeQaTest",
        capabilities: r"Software\NextOfficeQaTest\Capabilities",
    };
    const EXTS: [&str; 2] = ["nxqaone", "nxqatwo"];

    fn reg_query(key: &str, value: Option<&str>) -> Option<String> {
        let mut cmd = Command::new("reg");
        cmd.args(["query", &format!(r"HKCU\{key}")]);
        match value {
            Some(v) => cmd.args(["/v", v]),
            None => cmd.arg("/ve"),
        };
        let out = cmd.output().ok()?;
        out.status.success().then(|| String::from_utf8_lossy(&out.stdout).into_owned())
    }

    #[test]
    fn registers_and_detects_defaults_without_taking_over() {
        let exe = std::env::current_exe().unwrap();
        let exe_name = exe.file_name().unwrap().to_string_lossy().into_owned();
        let types: Vec<FileType> = EXTS.iter().map(|e| FileType { ext: e.to_string(), name: format!("QA {e} file") }).collect();
        win::remove(&TEST, &exe_name, &EXTS);

        win::register(&TEST, &exe, "QA", &types).unwrap();

        // ProgID with the open command pointing at this exe
        let cmd = reg_query(r"Software\Classes\NextOfficeQaTest.nxqaone\shell\open\command", None).unwrap();
        assert!(cmd.contains(&exe.to_string_lossy().to_string()), "{cmd}");
        assert!(cmd.contains("\"%1\""), "{cmd}");
        assert!(reg_query(r"Software\Classes\NextOfficeQaTest.nxqaone", Some("FriendlyTypeName")).unwrap().contains("QA nxqaone file"));
        // Offered in "Open with", listed in Settings › Default apps
        assert!(reg_query(r"Software\Classes\.nxqaone\OpenWithProgids", Some("NextOfficeQaTest.nxqaone")).is_some());
        assert!(reg_query(r"Software\NextOfficeQaTest\Capabilities\FileAssociations", Some(".nxqatwo"))
            .unwrap()
            .contains("NextOfficeQaTest.nxqatwo"));
        assert!(reg_query(r"Software\RegisteredApplications", Some("Next Office QA Test"))
            .unwrap()
            .contains(r"Software\NextOfficeQaTest\Capabilities"));
        // Registering never makes it the default (Windows would still ask "How do you want to open this?")
        assert!(reg_query(r"Software\Classes\.nxqaone", None).map_or(true, |v| !v.contains("NextOfficeQaTest")));
        assert_eq!(win::effective_progid("nxqaone"), None);
        assert!(!opens_with_us("nxqaone", &exe));

        // Once the type does open with us (as after the user's choice), the check says so.
        win::set_for_test(r"Software\Classes\.nxqaone", "NextOfficeQaTest.nxqaone");
        assert_eq!(win::effective_progid("nxqaone").as_deref(), Some("NextOfficeQaTest.nxqaone"));
        assert!(opens_with_us("nxqaone", &exe));
        assert!(!opens_with_us("nxqatwo", &exe));
        // Real types keep whatever program the user chose (never this test build).
        assert!(!opens_with_us("txt", &exe));

        // A type dropped from the list leaves the Settings list on the next registration.
        win::register(&TEST, &exe, "QA", &types[..1]).unwrap();
        assert!(reg_query(r"Software\NextOfficeQaTest\Capabilities\FileAssociations", Some(".nxqatwo")).is_none());

        win::delete_for_test(r"Software\Classes\.nxqaone");
        win::remove(&TEST, &exe_name, &EXTS);
        assert!(reg_query(r"Software\RegisteredApplications", Some("Next Office QA Test")).is_none());
        assert!(reg_query(r"Software\Classes\NextOfficeQaTest.nxqaone", None).is_none());
    }

    #[test]
    fn checks_extensions_and_paths() {
        assert!(valid_ext("xlsx") && valid_ext("nxslides"));
        assert!(!valid_ext("") && !valid_ext("x\\y") && !valid_ext("XLSX") && !valid_ext(".xlsx"));
        assert!(same_path(r"\\?\C:\Apps\Next Office.exe", r"c:\apps\next office.EXE"));
        assert!(!same_path(r"C:\A\x.exe", r"C:\B\x.exe"));
    }
}
