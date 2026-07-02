# MS Access → Rust Report Conversion

A playbook + reference for converting MS Access reports into a Tauri/Rust
report viewer.  Report definitions are exported from Access via VBA: a custom
`ExtractSingleReport` function (run from inside the `.accdb`) opens each report
in hidden Design View and reads the live `Printer` object + section/control
properties directly into a clean `.json` sidecar.  The report body itself comes
from `Application.SaveAsText` — either sanitized to a `.bas` by the **Version
Control for Microsoft Access** add-in (if installed) or as a raw `.rpt`
(UTF-16LE) legacy fallback.  This doc captures everything learned while
converting a complex Access report (with an embedded sub-report) so future
conversions are faster.

> **Credits:** The [Version Control for Microsoft Access](https://github.com/joyfullservice/msaccess-vcs-addin)
> add-in by JoyfullService was invaluable for producing clean, diff-friendly
> `.bas` exports of the report bodies. The VBA property extractor, the Rust
> converter, the Tauri commands, and the JS renderer are all our own work.

## Architecture (3 pieces)

| Piece | Role |
| ------- | ----- |
| **Converter** (Rust) | Parse `.bas` (or legacy `.rpt`) text → serde `ReportDef` (sections, controls, geometry, page setup, sub-report links). Pure, no I/O. |
| **Commands** (Rust/Tauri) | `convert_report(name)` reads `{name}.bas` + `{name}.json`, falls back to `Report_<name>.rpt`, and parses. `get_report_data(jobNum)` fetches live data and aliases it to the report's field names. |
| **Renderer** (JS) | Consumes `ReportDef` + data → absolutely-positioned HTML, evaluates Access expressions. |

Source `.bas` + `.json` files live in the Access report export directory
(override with an env var).  Legacy `Report_<name>.rpt` files in the same
folder still work as a fallback.  Their record-source **queries** live
alongside as `Query_<name>.sql` — **read these**, they are the source of
truth for the data.

The `convert_report` function tries formats in priority order:

1. `{name}.bas` (VCS add-in sanitized `SaveAsText`) + `{name}.json` (VBA extractor sidecar)
2. `{name}.json` (VBA extractor — print settings only, no body)
3. `Report_{name}.rpt` (legacy raw `SaveAsText` fallback)

### Output model (`ReportDef`)

The converter produces a serde-serializable struct that the JS renderer
consumes as JSON:

```rust
#[derive(Debug, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ReportDef {
    pub name: String,
    pub caption: String,
    pub record_source: String,
    pub width_px: f64,
    /// Group/sort fields, outermost first (e.g. ["Category", "Item"]).
    pub group_fields: Vec<String>,
    /// 0=none, 1=whole group, 2=first detail with header.
    pub grp_keep_together: i32,
    pub page: PageSetup,
    pub sections: Vec<Section>,
}

#[derive(Debug, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PageSetup {
    pub orientation: String,    // "portrait" | "landscape"
    pub paper_size: String,     // "Letter" | "Legal" | "A4" | ...
    pub paper_w_in: f64,
    pub paper_h_in: f64,
    pub margin_left: f64,       // px (twips/15)
    pub margin_top: f64,
    pub margin_right: f64,
    pub margin_bottom: f64,
}

#[derive(Debug, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Section {
    /// reportHeader | pageHeader | groupHeader | detail | groupFooter | ...
    pub kind: String,
    pub name: String,
    pub height_px: f64,
    pub can_grow: bool,
    pub can_shrink: bool,
    pub keep_together: bool,
    pub controls: Vec<Control>,
}

#[derive(Debug, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Control {
    pub kind: String,           // label | textBox | line | checkBox | subReport
    pub name: String,
    pub left: f64,              // px (twips/15)
    pub top: f64,
    pub width: f64,
    pub height: f64,
    pub text: Option<String>,   // static label caption
    pub field: Option<String>,  // bound data field
    pub expr: Option<String>,   // Access expression (starts with '=')
    pub font_name: Option<String>,
    pub font_size: Option<f64>,
    pub bold: bool,
    pub align: Option<String>,  // left | center | right
    pub fore_color: Option<String>,
    pub back_color: Option<String>,
    pub transparent: bool,
    pub visible: bool,
    pub can_grow: bool,
    pub running_sum: i32,       // 0=none, 1=cumulative per group
    pub source_object: Option<String>, // sub-report name
    pub link_child: Option<String>,
    pub link_master: Option<String>,
}
```

---

## 0. The Access-side export (VBA)

Before the Rust converter can parse anything, the report definitions must be
exported from Access. There are two export paths, both run from inside the
Access database via VBA.

### Path A — `SaveAsText` (produces `.rpt` or `.bas`)

`Application.SaveAsText` writes the raw report definition. The **Version Control
for Microsoft Access** add-in wraps this to produce a sanitized `.bas` (UTF-8,
print settings stripped); a direct call produces a raw `.rpt` (UTF-16LE). The
`.json` sidecar is produced separately by Path B (the VBA extractor), not by
the VCS add-in.

```vba
' Export every report in the database via SaveAsText.
' Produces Report_<name>.rpt (raw) or <name>.bas (VCS add-in sanitized).
Public Sub ExportAllReports()
    Dim accObj As AccessObject
    Dim exportFolder As String
    exportFolder = "C:\Path\To\Report\Exports\"

    For Each accObj In CurrentProject.AllReports
        Application.SaveAsText acReport, accObj.Name, _
            exportFolder & "Report_" & accObj.Name & ".rpt"
    Next accObj
End Sub
```

### Path B — VBA property extractor (produces `.json` sidecar)

This opens each report in hidden Design View and reads the live `Printer`
object + section/control properties directly, writing a clean JSON sidecar.
This is more reliable than parsing the `PrtMip`/`PrtDevMode` binary blobs in
the `.rpt` file because it reads the active printer settings.

```vba
' Extract printer settings, group levels, sections, and controls from a
' report's Design View into a JSON sidecar file.
Function ExtractSingleReport(strReportName As String, strFolderPath As String) As Boolean
    Dim rpt As Report, ctl As Control, prt As Printer, sec As Section
    Dim strFilePath As String, fileNum As Integer, i As Integer

    strFilePath = strFolderPath & "\" & strReportName & ".json"
    fileNum = FreeFile()

    ' Open in hidden Design View to read live properties
    DoCmd.OpenReport strReportName, acViewDesign, , , acHidden
    Set rpt = Reports(strReportName)
    Set prt = rpt.Printer

    Open strFilePath For Output As #fileNum
    Print #fileNum, "{"
    Print #fileNum, "  ""ReportName"": """ & strReportName & ""","
    Print #fileNum, "  ""RecordSource"": """ & EscapeJSON(rpt.RecordSource) & ""","
    Print #fileNum, "  ""WidthPx"": " & TwipsToPx(rpt.Width) & ","

    ' --- Printer settings (read from the live Printer object) ---
    Dim orientationStr As String
    If prt.Orientation = acPRORPortrait Then orientationStr = "Portrait" Else orientationStr = "Landscape"
    Print #fileNum, "  ""PrinterSettings"": {"
    Print #fileNum, "    ""DeviceName"": """ & EscapeJSON(prt.DeviceName) & ""","
    Print #fileNum, "    ""Orientation"": """ & orientationStr & ""","
    Print #fileNum, "    ""PaperSize"": """ & DecodePaperSize(prt.PaperSize) & ""","
    Print #fileNum, "    ""Copies"": " & prt.Copies & ","
    Print #fileNum, "    ""ColorMode"": """ & DecodeColorMode(prt.ColorMode) & ""","
    Print #fileNum, "    ""Duplex"": """ & DecodeDuplex(prt.Duplex) & ""","
    Print #fileNum, "    ""TopMarginPx"": " & TwipsToPx(prt.TopMargin) & ","
    Print #fileNum, "    ""BottomMarginPx"": " & TwipsToPx(prt.BottomMargin) & ","
    Print #fileNum, "    ""LeftMarginPx"": " & TwipsToPx(prt.LeftMargin) & ","
    Print #fileNum, "    ""RightMarginPx"": " & TwipsToPx(prt.RightMargin)
    Print #fileNum, "  },"

    ' --- Group levels (sort/group fields) ---
    Print #fileNum, "  ""GroupLevels"": ["
    For i = 0 To 9
        On Error Resume Next
        Dim gl As GroupLevel
        Set gl = rpt.GroupLevel(i)
        If Err.Number = 0 Then
            If i > 0 Then Print #fileNum, "    ,"
            Print #fileNum, "    {"
            Print #fileNum, "      ""ControlSource"": """ & EscapeJSON(gl.ControlSource) & ""","
            Print #fileNum, "      ""SortOrder"": " & gl.SortOrder & ","
            Print #fileNum, "      ""GroupHeader"": " & LCase(CStr(gl.GroupHeader)) & ","
            Print #fileNum, "      ""GroupFooter"": " & LCase(CStr(gl.GroupFooter)) & ","
            Print #fileNum, "      ""KeepTogether"": " & gl.KeepTogether
            Print #fileNum, "    }"
        End If
        Err.Clear
        On Error GoTo 0
    Next i
    Print #fileNum, "  ],"

    ' --- Controls (geometry, fonts, colors, bindings) ---
    Print #fileNum, "  ""Controls"": ["
    Dim isFirst As Boolean: isFirst = True
    For Each ctl In rpt.Controls
        If Not isFirst Then Print #fileNum, "    ,"
        isFirst = False
        Print #fileNum, "    {"
        Print #fileNum, "      ""Name"": """ & EscapeJSON(ctl.Name) & ""","
        Print #fileNum, "      ""Type"": " & ctl.ControlType & ","
        Print #fileNum, "      ""SectionCode"": " & ctl.Section & ","
        Print #fileNum, "      ""TopPx"": " & TwipsToPx(ctl.Top) & ","
        Print #fileNum, "      ""LeftPx"": " & TwipsToPx(ctl.Left) & ","
        Print #fileNum, "      ""WidthPx"": " & TwipsToPx(ctl.Width) & ","
        Print #fileNum, "      ""HeightPx"": " & TwipsToPx(ctl.Height) & ","
        Print #fileNum, "      ""Visible"": " & LCase(CStr(ctl.Visible)) & ","

        ' Data binding
        If ctl.ControlType = acTextBox Then
            Print #fileNum, "      ""ControlSource"": """ & EscapeJSON(ctl.ControlSource) & ""","
            Print #fileNum, "      ""Format"": """ & EscapeJSON(ctl.Format) & ""","
            Print #fileNum, "      ""RunningSum"": " & ctl.RunningSum & ","
        ElseIf ctl.ControlType = acLabel Then
            Print #fileNum, "      ""Caption"": """ & EscapeJSON(ctl.Caption) & ""","
        ElseIf ctl.ControlType = acSubform Then
            Print #fileNum, "      ""SourceObject"": """ & EscapeJSON(ctl.SourceObject) & ""","
            Print #fileNum, "      ""LinkMasterFields"": """ & EscapeJSON(ctl.LinkMasterFields) & ""","
            Print #fileNum, "      ""LinkChildFields"": """ & EscapeJSON(ctl.LinkChildFields) & ""","
        End If

        ' Fonts
        If ctl.ControlType = acTextBox Or ctl.ControlType = acLabel Then
            Print #fileNum, "      ""FontName"": """ & EscapeJSON(ctl.FontName) & ""","
            Print #fileNum, "      ""FontSizePx"": " & PtToPx(ctl.FontSize) & ","
            Print #fileNum, "      ""FontWeight"": " & ctl.FontWeight & ","
            Print #fileNum, "      ""TextAlign"": " & ctl.TextAlign & ","
        End If

        ' Colors (OLE BGR → hex)
        Print #fileNum, "      ""ForeColorHex"": """ & ColorToHex(ctl.ForeColor) & ""","
        Print #fileNum, "      ""BackColorHex"": """ & ColorToHex(ctl.BackColor) & """"
        Print #fileNum, "    }"
    Next ctl
    Print #fileNum, "  ]"
    Print #fileNum, "}"

    Close #fileNum
    DoCmd.Close acReport, strReportName, acSaveNo
    ExtractSingleReport = True
End Function
```

### Helper functions (VBA)

```vba
' Twips → px (1440 twips/in ÷ 96 px/in = 15)
Function TwipsToPx(ByVal twipsVal As Long) As Long
    TwipsToPx = Round(twipsVal / 15)
End Function

' Points → px (96 dpi / 72 pt/in ≈ 1.333)
Function PtToPx(ByVal ptVal As Integer) As Long
    PtToPx = Round(ptVal * 1.3333)
End Function

' OLE BGR integer → CSS #RRGGBB. Negative = Windows system color.
Function ColorToHex(ByVal colorVal As Long) As String
    Dim R As Integer, G As Integer, B As Integer
    If colorVal < 0 Then
        Select Case colorVal
            Case -2147483633: ColorToHex = "#FFFFFF"
            Case -2147483630: ColorToHex = "#000000"
            Case -2147483624: ColorToHex = "#F0F0F0"
            Case Else:        ColorToHex = "#000000"
        End Select
        Exit Function
    End If
    R = colorVal Mod 256
    G = (colorVal \ 256) Mod 256
    B = (colorVal \ 65536) Mod 256
    ColorToHex = "#" & Right("0" & Hex(R), 2) & Right("0" & Hex(G), 2) & Right("0" & Hex(B), 2)
End Function

' Paper size code → name
Function DecodePaperSize(SizeCode As Integer) As String
    Select Case SizeCode
        Case 1: DecodePaperSize = "Letter"
        Case 3: DecodePaperSize = "Tabloid"
        Case 5: DecodePaperSize = "Legal"
        Case 8: DecodePaperSize = "A3"
        Case 9: DecodePaperSize = "A4"
        Case Else: DecodePaperSize = "Other (" & SizeCode & ")"
    End Select
End Function

' JSON string escaping
Function EscapeJSON(ByVal strVal As String) As String
    If IsNull(strVal) Then EscapeJSON = "": Exit Function
    Dim res As String: res = strVal
    res = Replace(res, "\", "\\")
    res = Replace(res, """", "\""")
    res = Replace(res, vbCrLf, "\n")
    res = Replace(res, vbTab, "\t")
    EscapeJSON = res
End Function
```

### Reading physical paper dimensions from the spooler

The `Printer.PaperSize` enum gives a code, but the actual physical dimensions
can vary by printer driver. For accurate page layout, query the spooler:

```vba
' Windows spooler API — reads physical paper dimensions from the printer driver.
Private Declare PtrSafe Function DeviceCapabilities Lib "winspool.drv" _
    Alias "DeviceCapabilitiesA" ( _
    ByVal pDevice As String, ByVal pPort As String, _
    ByVal iCapability As Integer, ByVal pOutput As String, _
    ByVal pDevMode As Long) As Long
Private Const DC_PAPERWIDTH As Integer = 3
Private Const DC_PAPERHEIGHT As Integer = 4

' Returns paper width in twips (0.1mm units from spooler → twips).
' Falls back to 0 if the API fails; caller then uses the PaperSize lookup table.
Function GetPaperWidthTwips(ByRef prt As Printer) As Long
    On Error Resume Next
    Dim mm10 As Long
    mm10 = DeviceCapabilities(prt.DeviceName, prt.Port, DC_PAPERWIDTH, vbNullString, 0)
    If Err.Number <> 0 Or mm10 <= 0 Then
        GetPaperWidthTwips = 0
    Else
        GetPaperWidthTwips = Round(mm10 * 1440 / 254) ' 0.1mm → twips
    End If
End Function
```

### Full pipeline

```text
Access .accdb
    │
    ├── ExportAllReports (VBA)  ──► Report_<name>.rpt  (SaveAsText, UTF-16LE)
    │                               or <name>.bas       (VCS add-in, UTF-8)
    │
    ├── ExtractSingleReport     ──► <name>.json         (printer settings + props)
    │   (VBA, Design View)
    │
    └── Query SQL export        ──► Query_<name>.sql    (record source)
        (VBA, QueryDef.SQL)
                │
                ▼
    Rust converter
        parse_body() → Node tree → ReportDef struct
        decode_page() / decode_page_from_json() → PageSetup
                │
                ▼
    JSON { def: ReportDef, data: rows[] }
                │
                ▼
    JS renderer
        renderReport() → absolutely-positioned HTML
        evalAccessExpr() → Access expression evaluation
                │
                ▼
    WebView2 / browser
        Paginated print-preview with @page CSS matching Access page setup
```

---

## 1. The `SaveAsText` export format

### `.bas` (preferred — VCS add-in export)

- **Encoding: UTF-8.** The VCS add-in sanitizes the raw `SaveAsText` output:
  print settings are stripped out and VBA code is split to a separate `.cls`
  file.  The result is a clean, diff-friendly `.bas` file.
- Same nested `Begin <Kind> … End` tree and `Key =value` property lines as the
  legacy `.rpt` export.

### `.rpt` (legacy fallback — raw `SaveAsText`)

- **Encoding: UTF-16LE with BOM `FF FE`.** Decode as UTF-16, not UTF-8.
  Reading as UTF-8 yields an empty parse (every other byte is `\0`).

```rust
pub fn decode(bytes: &[u8]) -> String {
    if bytes.len() >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE {
        let u16s: Vec<u16> = bytes[2..]
            .chunks_exact(2)
            .map(|c| u16::from_le_bytes([c[0], c[1]]))
            .collect();
        String::from_utf16_lossy(&u16s)
    } else if bytes.len() >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF {
        // UTF-16BE (rare)
        let u16s: Vec<u16> = bytes[2..]
            .chunks_exact(2)
            .map(|c| u16::from_be_bytes([c[0], c[1]]))
            .collect();
        String::from_utf16_lossy(&u16s)
    } else {
        // .bas files are UTF-8
        String::from_utf8_lossy(bytes).into_owned()
    }
}
```

### Shared structure (both formats)

Both `.bas` and `.rpt` are a nested `Begin <Kind> … End` tree of `Key =value`
property lines. Note the space before `=` and no space after for numbers
(`Left =4965`).

- Top: `Version`/`VersionRequired`/`Checksum`, then `Begin Report … End`.
- Inside `Begin Report`: report-level props (+ binary blobs in `.rpt` only) +
  **one bare `Begin`** that contains the body.
- The bare `Begin` body, in order:
  1. **Default-format prototype controls** (`Begin Label`, `Begin TextBox`,
     `Begin Rectangle`, `Begin Image`, `Begin Subform`, `Begin UnboundObjectFrame`…)
     with no `Name`/geometry — these are style defaults, **SKIP them** (the parser
     skips control-typed nodes at the top level).
  2. **`Begin BreakLevel`** blocks = group/sort levels; their `ControlSource` is
     the group field (e.g. `Heading`, `Category`).
  3. **Section** blocks (see §2).

The parser is a simple recursive tree builder:

```rust
struct Node {
    kind: String,
    props: Vec<(String, String)>,
    children: Vec<Node>,
}

/// Parse a block body (lines after `Begin`) until the matching `End`.
fn parse_body(lines: &[&str], i: &mut usize) -> (Vec<(String, String)>, Vec<Node>) {
    let mut props = Vec::new();
    let mut children = Vec::new();
    while *i < lines.len() {
        let line = lines[*i].trim().to_string();
        *i += 1;
        if line == "End" { break; }
        if line.is_empty() { continue; }
        if line == "Begin" || line.starts_with("Begin ") {
            let kind = line["Begin".len()..].trim().to_string();
            let (p, c) = parse_body(lines, i);
            children.push(Node { kind, props: p, children: c });
        } else if let Some((k, v)) = split_prop(&line) {
            let vt = v.trim();
            if vt == "Begin" {
                // Binary blob (hex lines) — keep PrtMip/PrtDevMode, discard rest.
                let hex = consume_blob(lines, i);
                let keep = matches!(k.as_str(), "PrtMip" | "PrtDevMode");
                props.push((k, if keep { hex } else { String::new() }));
            } else if vt.starts_with('"') {
                // Quoted string — may continue on following "…" lines
                let mut inner = strip_quotes(vt).to_string();
                while *i < lines.len() {
                    let peek = lines[*i].trim();
                    if peek.starts_with('"') && peek.ends_with('"') {
                        inner.push_str(strip_quotes(peek));
                        *i += 1;
                    } else { break; }
                }
                props.push((k, unescape_inner(&inner)));
            } else {
                props.push((k, vt.to_string()));
            }
        }
    }
    (props, children)
}
```

### Binary blobs (legacy `.rpt` only)

In raw `.rpt` exports you will see `Key = Begin … End` blocks containing `0x…`
hex lines. Most are discardable (`NameMap`, `GUID`, `RecSrcDt`, `PrtDevNames`,
`PrtDevModeW`). **Keep `PrtMip` and `PrtDevMode`** — they hold printer setup
(see §4). The report name is **not** reliably decodable from `NameMap`; pass it
in from the filename.

`.bas` files produced by the VCS add-in **do not contain these binary blobs**;
page setup is supplied by the companion `.json` sidecar instead.

---

## 2. Sections

`SaveAsText` keyword → logical kind:

| Keyword | Kind |
| --------- | ------ |
| `FormHeader` | `reportHeader` |
| `FormFooter` | `reportFooter` |
| `PageHeader` | `pageHeader` |
| `PageFooter` | `pageFooter` |
| `BreakHeader` | `groupHeader` |
| `BreakFooter` | `groupFooter` |
| `Section` | `detail` |

Each section node = props (`Height`, `Name`, `OnFormat`, `GUID`) **plus a nested
bare `Begin`** whose children are the controls. Collect controls **recursively**
(a control can have an attached child label).

---

## 3. Controls & properties

| Property | Notes |
| ---------- | ------- |
| `Left/Top/Width/Height` | **TWIPS**. px = twips / **15** (1440 twips/in ÷ 96 px/in). `Top`/`Left` omitted ⇒ 0. **CRITICAL**: Controls often omit `Height` and only provide `LayoutCachedHeight`. Parser must fall back to `LayoutCachedHeight`, `LayoutCachedLeft`, `LayoutCachedTop`, `LayoutCachedWidth`. |
| `Caption` | static label text. |
| `ControlSource` | starts with `=` ⇒ **expression**; else a **bound field**. |
| `FontName` / `FontSize` | size is **points** → px × 4/3 (≈1.333). |
| `FontWeight` | ≥ 700 ⇒ bold. `FontItalic`/`FontUnderline` = 1. |
| `TextAlign` | 1=left, 2=center, 3=right, 0=general. |
| `ForeColor`/`BackColor` | OLE **BGR** int → `#RRGGBB` (R=`n&0xFF`, G=`(n>>8)&0xFF`, B=`(n>>16)&0xFF`). **Negative = system color** — must map common codes: `-2147483610` → `#F0F0F0`, others via `GetSysColor` approximation. |
| `BackStyle` | 0 ⇒ transparent. |
| `BorderStyle`/`BorderLineStyle` | ≠ 0 ⇒ border. |
| **`Visible = NotDefault`** (or `0`) | **HIDDEN — skip the control.** Access hides internal sort/group fields. If you don't skip them they render as garbage columns. |
| `SourceObject` | sub-report/Subform target, `"Report.<name>"` (strip `Report.`). |
| `LinkChildFields` / `LinkMasterFields` | `;`-separated join keys (see §6). |
| `CanGrow` / `CanShrink` | Section-level: section height is `min-height` (not fixed). Control-level: text controls inside get `overflow:visible;white-space:normal`. |
| `KeepTogether` | Section property; true means don't split this section across pages. |
| `AlternateBackColor` | Section property; apply alternating background to detail/sub-report rows. |
| `RunningSum` | Control property; `1` = cumulative sum per group. Must accumulate across rows. |

Control kinds: `Label`, `TextBox`, `Line`, `Rectangle`, `Image`, `CheckBox`,
`ToggleButton`, `OptionButton`, `Subform`/`SubReport`. (`ToggleButton`/`OptionButton`
⇒ map to `checkBox` kind. `Subform` with a `SourceObject` ⇒ treat as `subReport`.)

### Unit + color conversions

```rust
const TWIPS_PER_PX: f64 = 15.0; // 1440 twips/in ÷ 96 px/in

/// Read a twips property, falling back to LayoutCached<key> if missing.
fn px(props: &[(String, String)], key: &str) -> f64 {
    pf64(props, key)
        .or_else(|| pf64(props, &format!("LayoutCached{}", key)))
        .map(|v| v / TWIPS_PER_PX)
        .unwrap_or(0.0)
}

/// OLE BGR integer → CSS #RRGGBB. Negative = Windows system color.
fn ole_color(n: i64) -> String {
    if n < 0 {
        // Map common system colors; others approximate as light gray.
        return match n {
            -2147483610 => "#F0F0F0".into(), // btnFace
            -2147483633 => "#C0C0C0".into(), // btnShadow
            _ => "#F0F0F0".into(),
        };
    }
    let r = (n & 0xFF) as u8;
    let g = ((n >> 8) & 0xFF) as u8;
    let b = ((n >> 16) & 0xFF) as u8;
    format!("#{:02X}{:02X}{:02X}", r, g, b)
}
```

---

## 4. Printer / page setup

### `.json` sidecar (preferred — VBA extractor)

When a `{name}.json` file exists next to `{name}.bas`, the converter reads page
setup from clean JSON instead of hex blobs:

```json
{
  "orientation": "landscape",
  "paperSize": "legal",
  "margins": {
    "left": 0.2,
    "top": 0.2,
    "right": 0.2,
    "bottom": 0.2
  }
}
```

- `orientation`: `"portrait"` or `"landscape"`.
- `paperSize`: `"letter"`, `"legal"`, `"tabloid"`, `"a3"`, `"a4"`, `"a5"`.
- `margins`: all values are in **inches** (converted to twips internally:
  `inches × 1440`).

### Legacy binary blobs (`PrtMip` + `PrtDevMode` in `.rpt`)

If no `.json` sidecar is found, the converter falls back to the binary blobs
inside a legacy `.rpt` file:

- **`PrtMip`**: first four little-endian `int32` = **left, top, right, bottom
  margins in twips**.
- **`PrtDevMode`** (Windows `DEVMODE`): bytes 0–31 = `dmDeviceName` (printer name).
  - offset **44** (LE u16) = `dmOrientation` — **1 = portrait, 2 = landscape**.
  - offset **46** (LE u16) = `dmPaperSize` — **1=Letter, 5=Legal, 3=Tabloid,
    8=A3, 9=A4, 11=A5**.
- One report decoded to **Legal, Landscape, 0.2″ margins** — and the
  report `Width` (19440 twips = 13.5″) confirmed it needs Legal landscape.

```rust
fn paper_info(code: u16) -> (&'static str, f64, f64) {
    match code {
        1  => ("Letter",  8.5,  11.0),
        5  => ("Legal",   8.5,  14.0),
        3  => ("Tabloid", 11.0, 17.0),
        8  => ("A3",      11.69, 16.54),
        9  => ("A4",      8.27, 11.69),
        11 => ("A5",      5.83,  8.27),
        _  => ("Custom",  0.0,   0.0),
    }
}

fn decode_page(props: &[(String, String)]) -> PageSetup {
    let mut page = PageSetup::default();
    // PrtMip: four LE int32s = left/top/right/bottom margins (twips).
    if let Some(hex) = pget(props, "PrtMip") {
        let b = hex_to_bytes(hex);
        if let Some(v) = le_i32(&b, 0)  { page.margin_left   = v as f64 / TWIPS_PER_PX; }
        if let Some(v) = le_i32(&b, 4)  { page.margin_top    = v as f64 / TWIPS_PER_PX; }
        if let Some(v) = le_i32(&b, 8)  { page.margin_right  = v as f64 / TWIPS_PER_PX; }
        if let Some(v) = le_i32(&b, 12) { page.margin_bottom = v as f64 / TWIPS_PER_PX; }
    }
    // PrtDevMode (DEVMODE): orientation @44, paper size @46.
    if let Some(hex) = pget(props, "PrtDevMode") {
        let b = hex_to_bytes(hex);
        page.orientation = match le_u16(&b, 44) {
            Some(2) => "landscape".into(),
            _       => "portrait".into(),
        };
        if let Some(code) = le_u16(&b, 46) {
            let (name, w, h) = paper_info(code);
            page.paper_size  = name.to_string();
            page.paper_w_in  = w;
            page.paper_h_in  = h;
        }
    }
    page
}
```

---

## 5. Expression evaluation

`ControlSource` expressions must be evaluated. Supported subset:
`[Field]` / `[Table].[Field]`, string/number literals, `IIf`, comparisons
(`<= >= <> < > =`), `And`/`Or`/`Not`, `&` (concat), `+ - * /`, functions
(`Sum`→arg, `Val`, `Nz`, `IsNull`, `Date`, `Time`), `Null`/`True`/`False`.

### Field-resolution rule (important!)

- `[A].[B]` (**two** brackets joined by `.`) → resolve to the **last segment**
  (`B`) → maps to a data field.
- `[A.B]` (**single** bracket containing a dot) → resolve to the **full** name
  (`A.B`) → usually **absent** in supplied data → `null`.

This distinction is what makes ambiguous expressions resolve correctly, e.g.:

```text
=IIf([Supplier]<="004"," ", Sum([orders].[Quantity]) * [items.Quantity])
```

`[orders].[Quantity]` → `Quantity` (data) ; `[items.Quantity]` →
absent → `null`. With **lenient arithmetic (null = 1 in `*`/`/`, 0 in `+`/`-`)**,
QTY = `Quantity * 1 = Quantity`. ✔

Other rules:

- Comparison: string compare if **either** operand is a string, else numeric
  (so `[Code]<=10999` is numeric, `[Supplier]<="004"` is string).
- **`IIf(null, a, b)` must return `null`** (not `b`). Access treats a null
  condition as null, and the renderer then shows blank. If the evaluator falls
  through to the false branch, hidden columns leak values.
- `IIf([Code]<=10999, Null, [Code])` is a common pattern to hide low codes —
  evaluates to the code for real products, blank for low codes.

### Tokenizer + parser (JavaScript)

The evaluator is a recursive-descent parser with precedence:
`Or → And → comparison → & (concat) → +/- → */ → unary → primary`.

```javascript
function exprTokenize(s) {
  const toks = []; let i = 0; const n = s.length;
  while (i < n) {
    const c = s[i];
    if (c === " " || c === "\t") { i++; continue; }
    if (c === "[") {
      // [Field] or [Table].[Field] — join segments with "."
      const segs = [];
      while (i < n && s[i] === "[") {
        const end = s.indexOf("]", i);
        if (end < 0) { segs.push(s.slice(i + 1)); i = n; break; }
        segs.push(s.slice(i + 1, end)); i = end + 1;
        if (s[i] === "." && s[i + 1] === "[") { i++; continue; }
        break;
      }
      toks.push({ t: "field", v: segs.join(".") });
      continue;
    }
    if (c === '"') { /* parse string literal */ ... }
    if (/[0-9]/.test(c)) { /* parse number */ ... }
    if (/[A-Za-z_]/.test(c)) { /* parse identifier (IIf, Sum, And, etc.) */ ... }
    const two = s.substr(i, 2);
    if (two === "<=" || two === ">=" || two === "<>") { toks.push({ t: "op", v: two }); i += 2; continue; }
    if ("<>=&+-*/".includes(c)) { toks.push({ t: "op", v: c }); i++; continue; }
    if (c === "(") { toks.push({ t: "lp" }); i++; continue; }
    if (c === ")") { toks.push({ t: "rp" }); i++; continue; }
    if (c === ",") { toks.push({ t: "comma" }); i++; continue; }
    i++;
  }
  return toks;
}
```

Field resolution with the last-segment fallback:

```javascript
function exParsePrimary(p, row, groupRows) {
  const t = p.toks[p.i];
  if (t.t === "field") {
    p.i++;
    let v = row ? row[t.v] : null;
    // [A.B] not found → try last segment "B"
    if (v === undefined && t.v.includes(".")) {
      v = row ? row[t.v.split(".").pop()] : null;
    }
    // Case-insensitive fallback
    if (v === undefined && row) {
      const lowerKey = t.v.toLowerCase();
      v = Object.keys(row).find(k => k.toLowerCase() === lowerKey)?.let(row[_]);
    }
    return v === undefined ? null : v;
  }
  if (t.t === "id" && p.toks[p.i + 1]?.t === "lp") {
    // Function call: IIf, Sum, Nz, Val, IsNull, Left, Right, Mid, Format, ...
    p.i++; // skip function name
    p.i++; // skip "("
    const args = parseArgs(p, row, groupRows);
    if (t.v.toLowerCase() === "iif") {
      if (args[0] == null) return null;     // null condition → null (NOT false branch)
      return exTruthy(args[0]) ? args[1] : args[2];
    }
    if (t.v.toLowerCase() === "nz")
      return args[0] == null ? (args[1] ?? 0) : args[0];
    if (t.v.toLowerCase() === "isnull")
      return args[0] == null;
    // ... Sum, Val, Left, Right, Mid, Format, Date, etc.
  }
  // numbers, strings, parens, Not, unary minus ...
}
```

Lenient arithmetic (null acts as identity):

```javascript
const exNum = (x, d) =>
  x == null ? d : (typeof x === "number" ? x : (parseFloat(x) || d));

// In exParseMul: null * x = x (d=1), null + x = x (d=0)
left = op === "*"
  ? exNum(left, 1) * exNum(r, 1)   // null → 1 (identity for *)
  : exNum(left, 1) / exNum(r, 1);
```

---

## 6. Data binding — the hard part

The report's `RecordSource` is an **Access query**, not an API. Steps:

1. **Read the query** `Query_<recordsource>.sql`. It is authoritative
   for joins, filters, grouping, and key relationships.
2. **Replicate it as a Node endpoint** (or map to an existing one). Match the
   query's `WHERE`/`HAVING`/`GROUP BY` exactly.
3. **Alias field names** to the report's `ControlSource` names if your API
   differs from the Access query's field names.
4. **Inject project-level fields the expressions read per-row.** If an
   expression like `IIf([Metric],…)` reads a project-level field per row,
   inject it into every row, or the condition falls to the wrong branch
   and renders blank.

### Sub-reports

- `SourceObject` = child report; convert it too (`convert_report` is generic).
- `LinkMasterFields` = child key from the parent. Group the child data by
  that key; render the matching slice under each parent group.
- **Detail may be empty** — Access often puts the per-record fields in an inner
  **group header**. `renderSubreport` picks the section with the **most bound
  fields** as the row template.

### Common pitfalls

- **Wrong join key.** Always trace the *query's* join path to find the correct
  foreign key. Using a stale or wrong key links items to the wrong group.
- **Over-aggregation drops rows.** If the real query groups per line item,
  collapsing by product loses line items and the fields that columns depend on.
- **Replicate the query's filters.** If the Access query excludes placeholder
  rows (e.g. supplier `0000`), omitting that filter leaks junk rows.
- **Aggregate expressions need pre-computation.** `Sum([Quantity])` in a group
  footer cannot be evaluated row-by-row. Pre-compute per-group aggregates and
  inject them as `__computed.<controlName>` on the data object before rendering.
- **Blank last page** — report footer on its own page. Merge footer onto last
  content page if it fits; drop empty trailing pages.

---

## 7. Rendering model

- `ReportDef` → stack sections vertically. **Each section div must be
  `position: relative`** so a control's absolute `left/top` is *section-local*.
  Without it, every control collapses to the page's top-left (the #1 layout bug).
- **Paginated (print-preview).** The flow is split into blocks — per group: an
  *intro* block (`groupHeader` + `detail` rows + clipped `groupFooter`) kept
  together, then one block per sub-report row so sub-report content splits
  across pages — plus a `reportFooter` block. Blocks are packed into pages by
  their section heights; each page renders the `pageHeader` (top zone), its
  blocks (middle zone), and the `pageFooter` (bottom zone) with `[Page]`/`[Pages]`
  set to that page's number / total.
- **Clip the group footer to the sub-report control's `top`** before flowing the
  sub-report rows, so they sit directly under their column headers instead of
  after the footer's full (mostly empty) height.
- Page: width = paper size in px (landscape swaps W/H); margins as padding.
- **Print CSS is critical**: the `@page` rule (e.g. `@page { size: legal
  landscape; margin: 0; }`) must match the decoded page setup (`.json` sidecar
  or legacy `PrtDevMode`) or the browser will print with default size/margins,
  completely mismatching Access. Use `overflow: visible` on the modal in
  `@media print` and `break-inside: avoid` on page containers.
- **Pagination rules**: `GrpKeepTogether=1` means the entire group (intro + all
  sub-report rows) starts on a new page if it can't fit on the current page.
  `keepTogether` on a section means don't split that section across pages.
- Conversions: twips→px `/15`; points→px `×4/3`.
- `CheckBox` / `ToggleButton` / `OptionButton` → render checkmark from bound boolean.
- `Line` → a bordered div (horizontal if `width ≥ height`).

### Control → HTML

Each control becomes an absolutely-positioned `<div>` inside a
`position: relative` section:

```javascript
function reportCtrlHtml(ctrl, data, sectionCanGrow, groupRows = null) {
  const pt = (p) => (p * 4) / 3; // points → px @96dpi

  if (ctrl.kind === "line") {
    const horiz = ctrl.width >= ctrl.height;
    const color = ctrl.borderColor || "#333";
    const w = ctrl.borderWidth ? Math.max(1, Math.round(ctrl.borderWidth)) : 1;
    const st = `left:${ctrl.left}px;top:${ctrl.top}px;` +
      (horiz
        ? `width:${ctrl.width}px;border-top:${w}px solid ${color};`
        : `height:${ctrl.height}px;border-left:${w}px solid ${color};`);
    return `<div class="rpt-ctrl" style="${st}"></div>`;
  }

  const st = [
    `left:${ctrl.left}px`, `top:${ctrl.top}px`, `width:${ctrl.width}px`,
  ];
  const grow = ctrl.canGrow || sectionCanGrow;
  if (ctrl.height > 0 && !grow)      st.push(`height:${ctrl.height}px`);
  else if (ctrl.height > 0)          st.push(`min-height:${ctrl.height}px`);
  if (ctrl.visible === false)        st.push("display:none");
  if (ctrl.fontName)                 st.push(`font-family:'${ctrl.fontName}',sans-serif`);
  if (ctrl.fontSize)                 st.push(`font-size:${pt(ctrl.fontSize).toFixed(1)}px`);
  if (ctrl.bold)                     st.push("font-weight:700");
  if (ctrl.italic)                   st.push("font-style:italic");
  if (ctrl.underline)                st.push("text-decoration:underline");
  if (ctrl.align)                    st.push(`text-align:${ctrl.align}`);
  if (ctrl.foreColor)                st.push(`color:${ctrl.foreColor}`);
  if (ctrl.backColor && !ctrl.transparent)
                                     st.push(`background:${ctrl.backColor}`);

  // Resolve the value: static label > bound field > expression
  let raw = "";
  if (ctrl.text != null)       raw = ctrl.text;
  else if (ctrl.field)         raw = data?.[ctrl.field] ?? "";
  else if (ctrl.expr) {
    // Check pre-computed aggregates first, then evaluate
    if (data?.__computed?.[ctrl.name] != null)
      raw = String(data.__computed[ctrl.name]);
    else
      raw = evalAccessExpr(ctrl.expr.replace(/^=/, ""), data, groupRows);
  }
  const display = ctrl.format ? formatValue(raw, ctrl.format) : raw;

  return `<div class="rpt-ctrl" style="${st.join(";")}">${escHtml(display)}</div>`;
}
```

### Sub-report row rendering

The sub-report row template is the section with the most bound fields.
Each row gets its own `position: relative` div so controls position
section-local:

```javascript
function renderSubreport(subDef, rows) {
  // Pick the section with the most bound fields as the row template
  let subTmpl = null, best = -1;
  for (const s of subDef.sections) {
    const n = s.controls.filter(c => c.field).length;
    if (n > best) { best = n; subTmpl = s; }
  }
  const rowH = Math.max(subTmpl.heightPx, 14);
  const altColor = subTmpl.alternateBackColor;

  return rows.map((row, idx) => {
    let inner = "";
    for (const c of subTmpl.controls)
      inner += reportCtrlHtml(c, row, subTmpl.canGrow, [row]);
    const bg = altColor && idx % 2 === 1 ? `background:${altColor};` : "";
    const h = subTmpl.canGrow ? `min-height:${rowH}px` : `height:${rowH}px`;
    return `<div class="rpt-section" style="${h};width:100%;${bg}">${inner}</div>`;
  }).join("");
}
```

### Print CSS

```css
@media print {
  @page { size: legal landscape; margin: 0; }
  .report-modal { overflow: visible; }
  .rpt-page    { break-inside: avoid; overflow: hidden; }
}
```

---

## 8. Playbook for a NEW report

1. **Export the report** from Access. Two paths:
   - **VBA extractor** (preferred): `ExtractSingleReport` opens the report in
     hidden Design View and writes `{name}.json` with printer settings, group
     levels, sections, and controls read directly from the live objects.
   - **`SaveAsText`** for the report body: `{name}.bas` (VCS add-in sanitized,
     UTF-8) or `Report_<name>.rpt` (raw, UTF-16LE legacy fallback).
   - The converter consumes `{name}.bas`/`.rpt` for the body + `{name}.json`
     for page setup. Both can coexist — the `.json` sidecar wins for page setup.
2. Run `convert_report("<name>")` and inspect the `ReportDef` JSON. Confirm
   sections, group fields, page setup, control bindings.
3. Read `Query_<recordSource>.sql`. Identify joins, filters, keys.
4. Add/choose a data endpoint that returns those rows; alias field names to the
   `ControlSource` names; inject any project-level fields used in expressions.
5. For each sub-report: convert it, find its link key, group child data by it.
6. Render; compare side-by-side with the Access print. Check: hidden columns
   gone, expressions populate, groups line up, page size/margins.

> **Legacy fallback:** If the VCS add-in is not available, a raw `.rpt`
> `SaveAsText` export still works.  Place `Report_<name>.rpt` in the report
> directory; `convert_report` will fall back to it automatically.

---

## 9. Known limitations / TODO

- **Pagination uses deterministic heights.** Blocks pack by their section *design*
  heights (twips → px), so page breaks/count are close to but not identical to
  Access (which grows sections to fit wrapped content). Descriptions long enough
  to wrap a second line can clip at the page edge (`.rpt-page` is
  `overflow:hidden`); most fit one line. `CanGrow` sections now use `min-height`
  which helps but does not recalculate page breaks after text wraps.
- **Sub-report internal grouping/sorting** isn't fully replicated; rows are
  ordered/consolidated per the SQL `ORDER BY`/`GROUP BY` and rendered flat
  per group, but multi-level sort hierarchies within a sub-report may differ.
- Expression evaluator is a subset — `Date()`/`Time()` use the browser locale (not
  the control's `Format`); no `DLookup`, domain aggregates, multi-arg date math.
- **Aggregate functions** (`Sum`, `Avg`, etc.) in expressions are only supported
  via pre-computation (`__computed`). The evaluator itself does not walk datasets.
