# Programming Library

Field notes and playbooks from real Windows desktop work — Tauri/WebView2, OLE drag-drop, and MS Access report conversion. Each writeup is a self-contained `.md` covering problem, root cause, solution, and sample code.

## Writeups

### [Native OLE Drop Target — Dragging Emails from New Outlook](OLE-DROP-OUTLOOK.md)

A custom Win32 `IDropTarget` in Rust that lets a Tauri (wry/WebView2) app accept email drags directly from **New Outlook for Windows**. New Outlook is WebView2-based and only delivers `CF_HDROP` asynchronously via `IDataObjectAsyncCapability` — wry's built-in drop target calls `GetData` synchronously and fails. This impl adds the STA message-pump retry loop that makes the async path work.

- Rust + the `windows` crate, `#[implement(IDropTarget)]`
- Registers on the WebView2 child HWND, replacing the built-in drop target
- Full sample code: `Cargo.toml` deps, `lib.rs` setup hook, the `IDropTarget` impl

### [MS Access → Rust Report Conversion](ACCESS-REPORT-CONVERSION.md)

A playbook for converting MS Access reports into a Tauri/Rust report viewer. Report definitions are exported from Access via VBA (`ExtractSingleReport` reads the live `Printer` object + section/control properties into a `.json` sidecar; the body comes from `SaveAsText` as a sanitized `.bas` via the [Version Control for Microsoft Access](https://github.com/joyfullservice/msaccess-vcs-addin) add-in, or a raw `.rpt` legacy fallback). Covers the converter, Tauri commands, and JS renderer.

## License

MIT — see [LICENSE](LICENSE).
