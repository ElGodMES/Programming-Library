# Native OLE Drop Target — Dragging Emails from New Outlook

## Problem

Dragging an email **item** directly from New Outlook into the Tauri EML
Extractor did nothing — the JS `drop` event fired with an empty
`dataTransfer`, so no file was processed. Dragging a saved `.eml` from the
desktop worked fine.

## Root Cause

New Outlook for Windows is a WebView2/React application. When you drag an
email item out of it, the email is **not** offered as a ready file path.
Instead:

1. New Outlook writes the email as a real `.eml` to
   `%TEMP%\chrome_drag<id>_<timestamp>\` during the drag.
2. It exposes the file via **async drag-drop**
   (`IDataObjectAsyncCapability`), not synchronously.

Tauri/wry's built-in drop target (`wry/src/webview2/drag_drop.rs`) only
handles **synchronous** `CF_HDROP` — it calls `GetData(CF_HDROP)` directly in
`DragEnter`/`Drop`. New Outlook **lists** `CF_HDROP` in its format list
(`QueryGetData` returns `S_OK`), but `GetData` fails synchronously with
`DV_E_FORMATETC` because the file isn't materialised yet. This is the exact
behaviour a Delphi developer documented on StackOverflow
(<https://stackoverflow.com/questions/79352823/>).

### Why drag-to-email (inside Outlook) works but drag-to-Tauri doesn't

Dragging an email onto another email **inside Outlook** never leaves
Outlook's process — no cross-process file transfer is involved. Dragging to
an external app requires Windows OLE drag-drop with a real file path, which
New Outlook only provides asynchronously.

### Why COM automation isn't an alternative

New Outlook does **not** support COM automation (Microsoft confirmed COM
add-ins aren't supported). So a "grab the selected email via Outlook COM"
button is not possible on New Outlook. The native OLE drop target is the
only viable path.

## Solution

A custom Win32 `IDropTarget` implemented in Rust
(`src-tauri/src/ole_drop.rs`), registered on the WebView2 child HWND,
replacing WebView2's built-in drop target. This mirrors what `wry` does but
adds the async path that `wry` lacks.

### How it works

1. **`DragEnter`** — Always check `IDataObjectAsyncCapability` first
   (regardless of what `QueryGetData(CF_HDROP)` claims). If
   `GetAsyncMode() == TRUE`, use the async path. Accept with
   `DROPEFFECT_COPY`.

2. **`Drop`** —
   - **Async path** (New Outlook): Call `StartOperation`, then call
     `GetData(CF_HDROP)` **on the main thread** (which has an STA from
     `OleInitialize`) with a **message-pumping retry loop**. The STA message
     pump is critical — it allows the async data delivery to complete. A
     worker thread with MTA (`COINIT_MULTITHREADED`) fails because OLE
     drag-drop is STA-based and cross-apartment marshaling needs a message
     pump. After `GetData` succeeds, read + base64-encode each `.eml`, call
     `EndOperation`, and emit the `ole-drop` Tauri event.
   - **Sync path** (desktop files): `GetData(CF_HDROP)` directly, enumerate
     paths, read + base64-encode, emit `ole-drop`.

3. **Frontend** (`src/main.js`) — Listens for `ole-drop` and feeds the
   pre-built `{name, base64}` entries straight into the existing
   `run_eml_extractor` command. No new command, no base64 round-trip in JS.

### Key implementation details

- **STA message pump**: The retry loop calls `PeekMessageW(PM_REMOVE)` between
  attempts (up to 20 retries, 250ms apart ≈ 5s max). In practice,
  `GetData` succeeds on attempt 1 — the file is already written by the time
  `Drop` is called.

- **Registration race**: WebView2 may register its own OLE drop target
  asynchronously after the controller is created. `register()` calls
  `do_register()` immediately and again after a 2-second delay to win the
  race regardless of WebView2's init timing.

- **HWND targeting**: `EnumChildWindows` finds the WebView2 child HWND(s)
  and registers the drop target on each, same as `wry`.

- **COM object lifetime**: The `IDropTarget` COM objects are
  `std::mem::forget`ten so they live for the process lifetime
  (`RegisterDragDrop` holds its own ref, but forgetting guarantees no early
  `Release`). This avoids needing `Send`/`Sync` for Tauri managed state.

- **`windows-core` dependency**: The `#[implement(IDropTarget)]` macro
  requires `windows-core` as a direct dependency (same as `wry`). It was
  already a transitive dep; adding it explicitly is required for the macro
  to resolve `windows_core` in the crate root.

### What the data object offers (New Outlook)

Enumerated via `IEnumFORMATETC` during diagnostics:

| Format | Clipboard ID | tymed |
| --- | --- | --- |
| `DragContext` | registered (0xC0xx) | 4 (ISTORAGE) |
| `DragImageBits` | registered | 1 (HGLOBAL) |
| `chromium/x-renderer-taint` | registered | 1 (HGLOBAL) |
| `CF_HDROP` | 15 | 1 (HGLOBAL) |
| `Chromium Web Custom MIME Data Format` | registered | 1 (HGLOBAL) |

These are **Chromium's** internal drag formats (New Outlook is WebView2-based).
`CF_HDROP` is listed but only delivered asynchronously. The temp file is
written to `%TEMP%\chrome_drag<id>_<timestamp>\<filename>.eml`.

## Files Changed

| File | Change |
| --- | --- |
| `src-tauri/src/ole_drop.rs` | New — `IDropTarget` impl with async support |
| `src-tauri/src/lib.rs` | `setup` hook calling `ole_drop::register` |
| `src-tauri/Cargo.toml` | Added `windows` features + `windows-core` dep |
| `src/main.js` | `ole-drop`/`ole-drag-enter`/`ole-drag-leave` listeners, `eeSetNativeEntries` |

## Testing

Verified with New Outlook for Windows:

- Drag email item from Outlook → dropzone → `GetData succeeded on attempt 1`
- PDF extracted and renamed with PO number suffix
- Copied to the project's Acknowledgement & Confirmation folder (nested under
  Purchase Orders sub-folder, found via recursive search)

Desktop `.eml` drags also work via the same native path.

## Limitations

- The native drop target intercepts **all** external file drags over the
  window (not just the dropzone), so dragging a file over the app always
  shows a copy cursor. The `ole-drop` handler ignores drops when the EML
  overlay is closed, so nothing happens — just a cosmetic cursor change.
- Classic Outlook (COM-based) uses `CFSTR_FILEDESCRIPTOR`/
  `CFSTR_FILECONTENTS` virtual-file formats, not the async `CF_HDROP` path.
  Classic Outlook users should save as `.eml` first. (Classic Outlook is
  supported until 2029 per Microsoft.)

## Sample Code

The excerpts below are the minimum needed to reproduce the approach in a
Tauri (wry/WebView2) app on Windows. The full file is
`src-tauri/src/ole_drop.rs`.

### 1. `Cargo.toml` — Windows-only deps

```toml
[target.'cfg(windows)'.dependencies]
windows-core = "0.61"                       # required by #[implement]
windows = { version = "0.61", features = [
  "Win32_Foundation",
  "Win32_System_Com",
  "Win32_System_Ole",
  "Win32_System_SystemServices",
  "Win32_UI_Shell",
  "Win32_UI_WindowsAndMessaging",
] }
base64 = "0.22.1"
```

`windows-core` must be a **direct** dependency — the `#[implement(IDropTarget)]`
macro resolves `windows_core` in the crate root. It is already a transitive
dep of `windows`, but the macro won't find it unless it's explicit.

### 2. `lib.rs` — register from the `setup` hook

```rust
#[cfg(windows)]
mod ole_drop;

pub fn run() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![/* your commands */])
        .setup(|app| {
            #[cfg(windows)]
            ole_drop::register(app.handle());
            Ok(())
        })
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
```

### 3. The `IDropTarget` implementation

The struct holds the `AppHandle` (for emitting Tauri events) and per-drag
state in `UnsafeCell`s (COM calls arrive on the UI thread, so no `Send`/
`Sync` needed):

```rust
use windows::core::{implement, Interface, Ref, BOOL};
use windows::Win32::Foundation::{HWND, LPARAM, POINTL, S_OK};
use windows::Win32::System::Com::{IDataObject, DVASPECT_CONTENT, FORMATETC, TYMED_HGLOBAL};
use windows::Win32::System::Ole::{
    IDropTarget, IDropTarget_Impl, CF_HDROP, DROPEFFECT, DROPEFFECT_COPY, DROPEFFECT_NONE,
};
use windows::Win32::System::SystemServices::MODIFIERKEYS_FLAGS;
use windows::Win32::UI::Shell::{DragFinish, DragQueryFileW, HDROP, IDataObjectAsyncCapability};

#[implement(IDropTarget)]
pub struct OleDropTarget {
    app: tauri::AppHandle,
    cursor_effect: std::cell::UnsafeCell<DROPEFFECT>,
    enter_is_valid: std::cell::UnsafeCell<bool>,
    is_async:       std::cell::UnsafeCell<bool>,
    data_obj:       std::cell::UnsafeCell<Option<IDataObject>>,
}
```

#### `DragEnter` — check async capability FIRST

New Outlook lists `CF_HDROP` in its format list (`QueryGetData` returns
`S_OK`) but only delivers it asynchronously. So `QueryGetData` alone is a
lie — you must check `IDataObjectAsyncCapability::GetAsyncMode()` and prefer
the async path when it returns `TRUE`:

```rust
#[allow(non_snake_case)]
impl IDropTarget_Impl for OleDropTarget_Impl {
    fn DragEnter(
        &self,
        pDataObj: Ref<'_, IDataObject>,
        _grfKeyState: MODIFIERKEYS_FLAGS,
        _pt: &POINTL,
        pdwEffect: *mut DROPEFFECT,
    ) -> windows::core::Result<()> {
        let Some(data) = pDataObj.as_ref() else {
            unsafe { *pdwEffect = DROPEFFECT_NONE };
            return Ok(());
        };

        // Async check FIRST — New Outlook advertises CF_HDROP but won't
        // deliver it synchronously (GetData fails with DV_E_FORMATETC).
        let async_cap = match data.cast::<IDataObjectAsyncCapability>() {
            Ok(cap) => unsafe { cap.GetAsyncMode() }
                .map(|b| b.as_bool()).unwrap_or(false),
            Err(_) => false,
        };
        let sync = unsafe { has_sync_hdrop(data) };
        let is_async = async_cap;
        let valid = sync || is_async;

        unsafe {
            *self.is_async.get()       = is_async;
            *self.enter_is_valid.get() = valid;
            *self.data_obj.get()       = Some(data.clone());   // hold for Drop
        }

        if valid {
            unsafe {
                *pdwEffect = DROPEFFECT_COPY;
                *self.cursor_effect.get() = DROPEFFECT_COPY;
            }
            let _ = self.app.emit("ole-drag-enter", ());
        } else {
            unsafe { *pdwEffect = DROPEFFECT_NONE };
        }
        Ok(())
    }

    fn DragOver(&self, _: MODIFIERKEYS_FLAGS, _: &POINTL, pdwEffect: *mut DROPEFFECT)
        -> windows::core::Result<()>
    {
        unsafe { *pdwEffect = *self.cursor_effect.get() };
        Ok(())
    }

    fn DragLeave(&self) -> windows::core::Result<()> {
        if unsafe { *self.enter_is_valid.get() } {
            let _ = self.app.emit("ole-drag-leave", ());
            unsafe { *self.enter_is_valid.get() = false; *self.data_obj.get() = None; }
        }
        Ok(())
    }
    // Drop — see next section
}
```

#### `Drop` — async path with STA message pump

The critical insight: `GetData` must run on the **main thread** (which has
an STA from `OleInitialize`), and you must **pump STA messages** between
retries. A worker thread with `COINIT_MULTITHREADED` fails because OLE
drag-drop is STA-based and cross-apartment marshaling needs a message pump.

```rust
    fn Drop(
        &self,
        pDataObj: Ref<'_, IDataObject>,
        _: MODIFIERKEYS_FLAGS,
        _: &POINTL,
        _pdwEffect: *mut DROPEFFECT,
    ) -> windows::core::Result<()> {
        let valid    = unsafe { *self.enter_is_valid.get() };
        let is_async = unsafe { *self.is_async.get() };
        if !valid { return Ok(()); }

        let data = pDataObj.as_ref().map(|d| d.clone())
            .or_else(|| unsafe { (*self.data_obj.get()).clone() });
        let Some(data) = data else { return Ok(()); };

        if is_async {
            let cap = data.cast::<IDataObjectAsyncCapability>().ok();
            if let Some(cap) = cap {
                let _ = unsafe { cap.StartOperation(None) };
                // Main-thread GetData with STA message pump (see below).
                let payload = unsafe { enumerate_hdrop_with_pump(&data) };
                let _ = unsafe { cap.EndOperation(S_OK, None, DROPEFFECT_COPY.0) };
                let _ = self.app.emit("ole-drop", payload);
            }
            unsafe { *self.enter_is_valid.get() = false; *self.data_obj.get() = None; }
            return Ok(());
        }

        // Sync path (desktop files): GetData is fast, do it inline.
        let payload = unsafe { enumerate_hdrop(&data) }
            .map(|(hdrop, paths)| {
                let p = paths_to_payload(&paths);
                unsafe { DragFinish(hdrop) };
                p
            })
            .unwrap_or_else(|| serde_json::json!({ "entries": [], "names": [] }));
        let _ = self.app.emit("ole-drop", payload);
        Ok(())
    }
```

#### The message-pumping retry loop

This is the heart of the async fix. Between `GetData` attempts, pump STA
messages for ~250ms so the async data delivery can complete. In practice
`GetData` succeeds on attempt 1 — the `.eml` is already written by the time
`Drop` is called:

```rust
use windows::Win32::UI::WindowsAndMessaging::{PeekMessageW, PM_REMOVE, MSG};

unsafe fn enumerate_hdrop_with_pump(data_obj: &IDataObject) -> serde_json::Value {
    let fmt = FORMATETC {
        cfFormat: CF_HDROP.0,
        ptd: std::ptr::null_mut(),
        dwAspect: DVASPECT_CONTENT.0,
        lindex: -1,
        tymed: TYMED_HGLOBAL.0 as u32,
    };

    for _attempt in 1..=20 {                       // 20 × 250ms ≈ 5s max
        match data_obj.GetData(&fmt) {
            Ok(medium) => {
                let hdrop = HDROP(medium.u.hGlobal.0 as _);
                let count = DragQueryFileW(hdrop, 0xFFFFFFFF, None);
                let mut paths = Vec::with_capacity(count as usize);
                for i in 0..count {
                    let n = DragQueryFileW(hdrop, i, None) as usize;
                    let mut buf = vec![0u16; n + 1];
                    DragQueryFileW(hdrop, i, Some(&mut buf));
                    let s = std::ffi::OsString::from_wide(&buf[..n])
                        .to_string_lossy().to_string();
                    paths.push(std::path::PathBuf::from(s));
                }
                let payload = paths_to_payload(&paths);
                DragFinish(hdrop);
                return payload;
            }
            Err(_) => {
                // Pump STA messages for 250ms — lets async delivery finish.
                let deadline = std::time::Instant::now()
                    + std::time::Duration::from_millis(250);
                while std::time::Instant::now() < deadline {
                    let mut msg = MSG::default();
                    let _ = PeekMessageW(&mut msg, None, 0, 0, PM_REMOVE);
                    std::thread::sleep(std::time::Duration::from_millis(10));
                }
            }
        }
    }
    serde_json::json!({ "entries": [], "names": [] })
}
```

#### Building the payload

Read each `.eml` and base64-encode it so the existing extractor command can
consume it without a new Tauri command or a JS-side base64 round-trip:

```rust
use base64::{engine::general_purpose, Engine as _};

fn paths_to_payload(paths: &[std::path::PathBuf]) -> serde_json::Value {
    let entries: Vec<_> = paths.iter()
        .filter(|p| p.extension().and_then(|e| e.to_str())
            .map(|e| e.eq_ignore_ascii_case("eml")).unwrap_or(false))
        .filter_map(|p| {
            let bytes = std::fs::read(p).ok()?;
            let name = p.file_name().and_then(|n| n.to_str())
                .unwrap_or("input.eml").to_string();
            Some(serde_json::json!({
                "name": name,
                "base64": general_purpose::STANDARD.encode(&bytes),
            }))
        })
        .collect();
    let names: Vec<String> = paths.iter()
        .filter_map(|p| p.file_name().and_then(|n| n.to_str()).map(String::from))
        .collect();
    serde_json::json!({ "entries": entries, "names": names })
}
```

### 4. Registration — win the WebView2 race

WebView2 registers its own OLE drop target asynchronously after the
controller is created. Register immediately **and** re-register after a
short delay so you win the race regardless of WebView2's init timing. Find
the WebView2 child HWND via `EnumChildWindows` (same as `wry`):

```rust
use windows::Win32::System::Ole::{OleInitialize, RegisterDragDrop, RevokeDragDrop};
use windows::Win32::Foundation::DRAGDROP_E_INVALIDHWND;
use windows::Win32::UI::WindowsAndMessaging::EnumChildWindows;
use tauri::Manager;

pub fn register(app: &tauri::AppHandle) {
    do_register(app);
    let app = app.clone();
    std::thread::spawn(move || {
        std::thread::sleep(std::time::Duration::from_secs(2));
        let app2 = app.clone();
        let _ = app.run_on_main_thread(move || { do_register(&app2); });
    });
}

fn do_register(app: &tauri::AppHandle) {
    let _ = unsafe { OleInitialize(None) };        // idempotent (STA)

    let Some(win) = app.get_webview_window("main") else { return; };
    let Ok(hwnd) = win.hwnd() else { return; };

    let app = app.clone();
    let mut targets: Vec<IDropTarget> = Vec::new();

    let mut inject = |child: HWND| -> bool {
        let target: IDropTarget = OleDropTarget::new(app.clone()).into();
        let revoked = unsafe { RevokeDragDrop(child) }
            != Err(DRAGDROP_E_INVALIDHWND.into());
        if revoked && unsafe { RegisterDragDrop(child, &target) }.is_ok() {
            targets.push(target);
        }
        true
    };
    let mut trait_obj: &mut dyn FnMut(HWND) -> bool = &mut inject;
    let ptr = unsafe { std::mem::transmute(&mut trait_obj) };
    unsafe extern "system" fn cb(hwnd: HWND, lp: LPARAM) -> BOOL {
        let f = &mut *(lp.0 as *mut &mut dyn FnMut(HWND) -> bool);
        f(hwnd).into()
    }
    let _ = unsafe { EnumChildWindows(Some(hwnd), Some(cb), LPARAM(ptr as _)) };

    // Keep the COM objects alive for the process lifetime. RegisterDragDrop
    // holds its own ref, but forgetting our handles guarantees no early
    // Release (avoids needing Send+Sync for Tauri managed state).
    std::mem::forget(targets);
}
```

### 5. Frontend — listen for the `ole-drop` event

The JS `drop` handler is unchanged (it still handles desktop `.eml` files
via `dataTransfer`). A separate Tauri event listener handles the native
OLE drops that WebView2 can't see:

```js
const _oleListen = window.__TAURI__?.event?.listen;

if (_oleListen) {
  _oleListen("ole-drag-enter", () => {
    if (!document.querySelector("#eml-extractor-overlay")?.classList.contains("hidden"))
      document.querySelector("#ee-dropzone")?.classList.add("drag-over");
  });
  _oleListen("ole-drag-leave", () => {
    document.querySelector("#ee-dropzone")?.classList.remove("drag-over");
  });
  _oleListen("ole-drop", (e) => {
    document.querySelector("#ee-dropzone")?.classList.remove("drag-over");
    const entries = e?.payload?.entries || [];   // [{name, base64}, ...]
    const names   = e?.payload?.names   || [];
    if (document.querySelector("#eml-extractor-overlay")?.classList.contains("hidden"))
      return;                                     // ignore drops when overlay closed
    if (entries.length) {
      // Feed straight into the existing run_eml_extractor flow —
      // no new command, no base64 round-trip in JS.
      eeSetNativeEntries(entries);
    } else if (names.length) {
      document.querySelector("#ee-dropzone-text").textContent =
        `No .eml files — dropped: ${names.join(", ")}.`;
    }
  });
}
```

### 6. `Cargo.toml` feature flags (recap)

The `Win32_UI_Shell` feature is what provides `IDataObjectAsyncCapability`.
Without it the cast in `DragEnter` won't compile. The full feature list
needed is shown in step 1.
