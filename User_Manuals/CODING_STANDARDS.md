# VBA Macro Projects — Coding Standards & Style Guide

This document covers every convention used across the macro projects in this
repository (GenerateBPAFiles, WireProjectRework, AppLoader,
VendorNormalizationEngine, banking_details_macroRework).
New contributors and Claude agents starting a fresh thread should read this
before writing or modifying any code.

---

## Table of Contents

0. [Living Document Process](#0-living-document-process)
1. [Project Layout](#1-project-layout)
2. [Module Naming & Numbering](#2-module-naming--numbering)
3. [Module Header Block](#3-module-header-block)
4. [Section Dividers](#4-section-dividers)
5. [Option Explicit](#5-option-explicit)
6. [Constants — Config Module Only](#6-constants--config-module-only)
7. [Global Run State Variables](#7-global-run-state-variables)
8. [ReadConfig Pattern](#8-readconfig-pattern)
9. [AppLock Pattern](#9-applock-pattern)
10. [Naming Conventions](#10-naming-conventions)
11. [Variable Declarations](#11-variable-declarations)
12. [Error Handling](#12-error-handling)
13. [Performance — Array Processing](#13-performance--array-processing)
14. [Private Helper Naming](#14-private-helper-naming)
15. [Orchestrator Pattern](#15-orchestrator-pattern)
16. [Pre-Run Validation Pattern](#16-pre-run-validation-pattern)
17. [Post-Run Summary Pattern](#17-post-run-summary-pattern)
18. [User Defined Types (UDTs)](#18-user-defined-types-udts)
19. [Progress UI Pattern](#19-progress-ui-pattern)
20. [Logging Pattern](#20-logging-pattern)
21. [Sheet Utilities](#21-sheet-utilities)
22. [User-Facing Messages](#22-user-facing-messages)
23. [Forms (UserForms)](#23-forms-userforms)
24. [Deprecated Constants](#24-deprecated-constants)
25. [What NOT to Do](#25-what-not-to-do)
26. [Known Inconsistencies in Earlier Projects](#26-known-inconsistencies-in-earlier-projects)
27. [Common VBA Errors — Causes and Fixes](#27-common-vba-errors--causes-and-fixes)

---

## 0. Living Document Process

This document is the single source of truth for coding conventions across all
macro projects in this repo.  It must be updated whenever a new pattern is
adopted, a new error is discovered and solved, or a new project is registered.

### When to update

| Trigger | What to do |
| --- | --- |
| New project started | Add acronym to §2 registered table AND to `CLAUDE.md` |
| New coding pattern adopted across ≥ 2 projects | Add to relevant section; add row to Quick Reference Card |
| New VBA error found and solved | Add subsection under §27; add row to §25; add to Quick Reference Card |
| Old pattern officially retired | Add row to §26 Known Inconsistencies; mark as `(Deprecated)` in code |
| Existing section changed significantly | Add row to the Changelog below |

### How to update

1. Edit the relevant section in this file directly.
2. Add or update the corresponding row in the **Quick Reference Card** at the end.
3. If a new project acronym is registered, update §2 in this file, `CLAUDE.md`, **and** `MEMORY.md`.
4. Add a row to the **Changelog** table below (newest entry at top).
5. Commit with a message like: `docs: update CODING_STANDARDS §N — <short description>`.

### Changelog

| Date | Section(s) | Change |
| --- | --- | --- |
| 2026-03-15 | §2, CLAUDE.md | Registered TFT acronym — Template Filler Tool |
| 2026-03-14 | §0, CLAUDE.md | Added living document process; created CLAUDE.md for auto-load enforcement |
| 2026-03-14 | §27 | Added §27a reserved words, §27b Else without If, §27c Error 9 |
| 2026-03-14 | §2, §25, §26 | Mandated 3-letter acronym prefix; renamed 38 module files across 4 projects |
| 2026-03-14 | §7, §16, §17, §18, §24 | Added Global Run State, Pre-Run Validation, Post-Run Summary, UDTs, Deprecated Constants sections |
| 2026-03-14 | §9, §23 | AppLock updated to include DisplayAlerts; UserForm comment style documented |
| 2026-03-14 | §1–§27 | Initial document created from audit of VNE, APL, BDR, BPA, WPR projects |

---

## 1. Project Layout

Each macro project lives in its own folder inside the repo root.

```text
/ProjectName/
    ACR1_Config.bas           <- ACR = 3-letter project acronym
    ACR2_Helpers.bas
    ACR3_...bas
    ...
    ACRN_Orchestrator.bas     <- always the highest number
    frmACRWizard_Code.bas     <- UserForm code exported as .bas
    frmACRProgress_Code.bas
```

- Modules are exported from the .xlsm and committed as plain `.bas` files.
- UserForm code is exported as a `.bas` file with the suffix `_Code.bas`.
- The `.xlsm` workbook itself is **not** committed (binary, too large).
- Never mix modules from two different projects in the same folder.

---

## 2. Module Naming & Numbering

Every module filename — and its internal `Attribute VB_Name` — must begin
with the project's **unique 3-letter acronym** followed by the module number
and an underscore.  This prevents file-collision when multiple projects are
downloaded to the same desktop folder and makes it immediately obvious which
macro a file belongs to.

### Registered project acronyms

| Acronym | Project folder | Full project name |
| --- | --- | --- |
| `VNE` | `VendorNormalizationEngine/` | Vendor Normalization Engine |
| `WPR` | `WireProjectRework/` | Wire Payment Review |
| `APL` | `AppLoader/` | AppLoader Formatter |
| `BDR` | `banking_details_macroRework/` | Banking Details Rework |
| `BPA` | `GenerateBPAFilesRework/` | Generate BPA Files |
| `TFT` | `TemplateFillerTool/` | Template Filler Tool |

When starting a new project, choose a 3-letter acronym that:

- Is not already in the table above.
- Is clearly related to the project name (not random).
- Contains only uppercase letters.

Add the new acronym to this table before writing any module.

### Naming pattern

```text
<ACR><N>_<Description>.bas        regular module
frm<ACR><Description>_Code.bas   UserForm code export
```

Examples: `WPR1_Config.bas`, `APL6_Orchestrator.bas`, `frmAPLWizard_Code.bas`

### Numbering rules

- **N1** — Config (constants + global run-state vars only, no logic)
- **N2** — Helpers (shared utilities, no business logic)
- **N3** — Setup / Reader (workbook setup + config reader)
- **N4+** — Step modules in execution order
- **Second-to-last** — ProgressUI wrappers
- **Last number** — Orchestrator (all public entry points)

The number makes the dependency graph obvious at a glance:
higher-numbered modules may call lower-numbered ones, never the reverse.

---

## 3. Module Header Block

Every module starts with the `Attribute` line (auto-generated on export),
`Option Explicit`, and then a fixed-width comment box:

```vba
Attribute VB_Name = "VNE4_OracleReduce"
Option Explicit

' ============================================================
' VNE4_ORACLEREDUCE  |  Vendor Normalization Engine  -  Step 2
'
' One paragraph: what this module does and why.
' State what the output sheet / data structure looks like.
' State any performance notes that affect how the code is written.
'
' Entry point:  RunOracleReduce()   (called by VNE7_Orchestrator)
' Dependencies: VNE1_Config, VNE2_Helpers, VNE3_Setup
' ============================================================
```

Rules:

- Width: exactly 60 `=` characters. Each line starts with `'` followed by a space.
- Module name in ALLCAPS followed by `|` and the project name.
- List the **Entry point** and **Dependencies** so a reader knows what calls
  this and what it calls.
- No `@author`, no dates — git log is the authoritative history.

### Change Log (optional, for heavily-evolved modules)

If a module has gone through significant breaking changes that are not obvious
from the code, a Change Log block can be added **below** the main header box.
Keep it brief — one line per notable change, newest at top:

```vba
' CHANGE LOG
' 2026-03-12  Log sheet redesigned: 1 row per file (was 1 row per cell).
'             Output_File and Output_Folder hyperlink columns added.
' 2026-03-11  Added GENERATE workflow constants and routing.
```

This is supplementary — `git log` is still the canonical history.
Do **not** add a Change Log to a new module that has not yet changed.

---

## 4. Section Dividers

Logical sections within a module use a shorter divider:

```vba
' -----------------------------------------------------------
' 1. STRING UTILITIES
' -----------------------------------------------------------
```

or for procedure groups:

```vba
' -----------------------------------------------------------
' PUBLIC ENTRY POINT
' -----------------------------------------------------------
```

- Width: exactly 59 `-` characters. Each line starts with `'` followed by a space.
- Section title in ALLCAPS.
- Leave one blank line above and below the divider block.
- Number sections sequentially when a module has more than three of them.

> **Inconsistency in earlier projects:** AppLoader and banking_details_macroRework
> use `'--...` (no leading space, shorter length) or `'====...` (= signs).
> Always use the `-` dash style shown above in new code.

---

## 5. Option Explicit

`Option Explicit` is **mandatory** in every module.  No exceptions.
Undeclared variables are the single biggest source of silent bugs in VBA.

---

## 6. Constants — Config Module Only

All constants (sheet names, cell addresses, column headers, threshold values,
color codes) live **exclusively** in the Config module (`Mod1_Config` or
`VNE1_Config`).

```vba
' Good — in VNE1_Config.bas
Public Const VNE_SH_RUN             As String = "Run"
Public Const VNE_DEF_THRESH_CLUSTER As Double = 0.85
Public Const VNE_HEADER_COLOR       As Long   = 14344935

' Bad — hardcoded in a processing module
ws.Name = "Run"                ' never
If score > 0.85 Then           ' never
.Interior.Color = 14344935     ' never
```

Runtime user values (file paths, column names) are read from the Config sheet
via `ReadConfig()` (see §8), not from constants.

### What belongs in Config

| Allowed | Not allowed |
| --- | --- |
| `Public Const` declarations | Any `Sub` or `Function` |
| `Public` global run-state variables (see §7) | Business logic |
| `Public Type` UDT definitions (see §18) | Sheet manipulation |
| `Private Const` that are local to Config itself | Calls to Helpers or other modules |

> **Inconsistency in earlier projects:** banking_details_macroRework puts
> functions (`GetDefaultOutputColumns()`, `GetUserSpecifiedOutputColumns()`,
> `RequiredSourceColumns()`) inside Mod1_Config. These should be in Mod2_Helpers.
> Config is a declarations-only file.

---

## 7. Global Run State Variables

Variables that hold the current run's parameters (user ID, output path, file
queue, abort flag) are declared `Public` in the Config module, clearly grouped
and labeled:

```vba
' ----- Global Run State -----
' Set by the wizard before RunAll() is called.
Public g_ProcessorID   As String   ' e.g. "1234"
Public g_RunDate       As String   ' DD-MMM-YY format
Public g_Initials      As String   ' e.g. "DK"
Public g_OutputPath    As String   ' full folder path for output files
Public g_AbortProcess  As Boolean
Public g_EnableLogging As Boolean

' ----- Global File Queue -----
Public g_TemplateNames() As String
Public g_FilePaths()     As String
Public g_FileCount       As Long
```

Rules:

- Always use the `g_` prefix (see §10 Naming Conventions).
- Group related globals under a labeled comment block.
- Never declare globals in step or helper modules — Config is the only
  valid home.
- Reset all globals to their zero/empty state at the start of each run
  (don't rely on leftover values from the previous run).

> **Inconsistency in earlier projects:** GenerateBPAFilesRework declares
> `Public AbortProcess As Boolean` and `Public EnableLogging As Boolean`
> in Mod1_Helpers — both without a `g_` prefix and in the wrong module.
> New code must use `g_` prefix and declare in Config.

---

## 8. ReadConfig Pattern

`ReadConfig()` lives in the Setup module and returns a
`Scripting.Dictionary`.  Every step module calls it at the top of its
entry point — never cache config across calls.

```vba
Public Sub RunSomeStep()
    Dim cfg As Object: Set cfg = ReadConfig()
    If cfg Is Nothing Then Exit Sub   ' ReadConfig already showed an error

    Dim filePath As String: filePath = cfg("VendorPath")
    ' ...
End Sub
```

Key rules:

- `ReadConfig()` returns `Nothing` on validation failure — callers **must**
  check for `Nothing` immediately and `Exit Sub` if so.
- Never pass individual config values as parameters through the call stack;
  re-call `ReadConfig()` or pass the dictionary object.
- Dictionary keys use PascalCase short names, not the cell address constants
  (e.g. `cfg("VendorPath")` not `cfg("C4")`).

### Dictionary factory

Always create Scripting.Dictionary objects via a factory helper so the
compare mode is set consistently:

```vba
' In Mod2_Helpers:
Public Function NewDict(Optional ByVal textCompare As Boolean = True) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    If textCompare Then d.CompareMode = 1   ' vbTextCompare
    Set NewDict = d
End Function

' Usage:
Dim lookup As Object: Set lookup = NewDict()
```

Using `textCompare = True` (the default) means dictionary key lookups are
case-insensitive, which matches how column header names are compared.

---

## 9. AppLock Pattern

Any procedure that does heavy work must bracket it with `AppLock`:

```vba
Private Sub AppLock(ByVal isLock As Boolean)
    Application.ScreenUpdating = Not isLock
    Application.EnableEvents   = Not isLock
    Application.DisplayAlerts  = Not isLock
    If isLock Then
        Application.Calculation = xlCalculationManual
    Else
        Application.Calculation = xlCalculationAutomatic
        Application.StatusBar   = False
    End If
End Sub
```

Usage:

```vba
AppLock True
' ... all heavy work ...
AppLock False
```

Rules:

- `AppLock False` must be called in **all** exit paths (normal, error, abort).
- Use `GoTo CleanUp` labels when multiple exit paths exist to guarantee unlock.
- Never call `AppLock True` twice in a row (it doesn't nest).
- The Orchestrator calls `AppLock`; individual step modules do **not** — they
  are always called from the Orchestrator.
  Exception: step modules have their own lock when called standalone via
  their individual `RunStepN_X()` button entry point.

> **Note vs earlier versions:** `Application.DisplayAlerts = False` is now
> part of the standard AppLock. Earlier projects (WireProjectRework, VNE)
> omitted it and added it ad-hoc. Include it in every AppLock going forward.

---

## 10. Naming Conventions

| Item | Convention | Example |
| --- | --- | --- |
| Public Sub / Function | PascalCase | `RunOracleReduce`, `ReadConfig` |
| Private Sub / Function | PascalCase | `OracleReduce_Core` |
| Private helper (utility) | `H_` prefix + PascalCase | `H_StripWord`, `H_Min3` |
| Module-level variable | `m_` prefix + PascalCase | `m_AbortFlag` |
| Global variable | `g_` prefix + PascalCase | `g_VNE_Abort`, `g_VNE_Silent` |
| Local variable | camelCase | `lastRow`, `vendorName`, `cfg` |
| Boolean parameter | `isXxx` or `hasXxx` | `isLock`, `hasHeader` |
| Output (ByRef) parameter | `outXxx` | `outDataStartRow`, `outA` |
| Constants | `PREFIX_UPPER_SNAKE` | `VNE_SH_RUN`, `VNE_DEF_THRESH_CLUSTER` |
| Sheet name constants | `PREFIX_SH_NAME` | `VNE_SH_LOG`, `VNE_SH_CONFIG` |
| Cell address constants | `PREFIX_CELL_NAME` | `VNE_CELL_VENDOR_PATH` |
| UDT name | PascalCase, noun | `ProcessResult`, `TemplateConfig` |
| UDT field name | PascalCase | `result.RowsProcessed`, `cfg.TemplateType` |

### Canonical utility function names

Some utility functions were written multiple times across early projects with
different names. Use these canonical names in all new code:

| Canonical name | Old names to retire |
| --- | --- |
| `NzStr(v)` | `NzTrim`, `GF_NzStr`, `GF_NzStr` |
| `CleanVendorName(s)` | `GF_CleanCompanyName` |
| `Similarity(a, b)` | `GF_SimilarityScore` |
| `Levenshtein(s, t)` | `GF_Levenshtein` |
| `FindHeaderCol(ws, hdr)` | `GF_FindHeaderColumn`, `FindHeaderInRow` (when row 1) |
| `GetOrCreateSheet(wb, name)` | `EnsureSheet`, `GF_EnsureSheet` |

> **Do not use the `GF_` prefix** on any new utility function. That prefix
> came from an era when all helpers lived in one mega-module (GenerateBPA).
> Today each project has a dedicated Helpers module so the prefix adds no
> value and creates inconsistency.

---

## 11. Variable Declarations

**Single-line declaration + initialisation** when the initial value is
immediately known:

```vba
' Good
Dim wb  As Workbook: Set wb  = ThisWorkbook
Dim cfg As Object:   Set cfg = ReadConfig()
Dim r   As Long:     r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

' Bad — split across two lines for no reason
Dim wb As Workbook
Set wb = ThisWorkbook
```

**Multi-line blocks** when declaring several related variables before any
assignment:

```vba
Dim lastRow    As Long
Dim vendorNum  As String
Dim vendorName As String
Dim score      As Double
```

Rules:

- Always declare the most specific type (`Long` not `Integer`,
  `String` not `Variant` when you know it's a string).
- Use `As Object` for late-bound COM objects (Scripting.Dictionary,
  FileSystemObject) because early-binding requires references that may
  not be set on all machines.
- Declare loop counters (`i`, `j`) with `As Long`, never `As Integer`
  (Integer overflows at 32 767 — arrays are always larger).
- Do **not** declare variables inside a loop body (VBA re-declares on every
  iteration — move all Dim statements to the top of the procedure).

---

## 12. Error Handling

### In utility functions (must not crash the caller)

```vba
Public Function NzStr(ByVal v As Variant) As String
    On Error GoTo EH
    If IsError(v) Or IsNull(v) Then NzStr = "" Else NzStr = CStr(v)
    Exit Function
EH:
    NzStr = ""
End Function
```

Small utility functions may use `EH:` as the label (short and obvious in a
short function). The `CleanUp:` label is reserved for entry-point procedures.

### Suppressing expected errors (sheet lookup, object access)

```vba
On Error Resume Next
Set ws = wb.Worksheets(sheetName)
On Error GoTo 0
' Always re-enable error handling immediately after the guarded line.
```

**Never** leave `On Error Resume Next` active across more than 2–3 lines.
Re-enable with `On Error GoTo 0` as soon as the guarded operation is done.

### In step/entry-point procedures — always use `CleanUp:`

Use a `CleanUp` label to guarantee resource cleanup on every exit path:

```vba
Public Sub RunSomething()
    Dim openedWb As Workbook
    On Error GoTo CleanUp

    ' ... work ...
    GoTo CleanUp   ' normal exit — falls through to cleanup

CleanUp:
    If Not openedWb Is Nothing Then openedWb.Close SaveChanges:=False
    AppLock False
    If Err.Number <> 0 Then
        MsgBox "Error in RunSomething: " & Err.Description, vbCritical, "ProjectName"
    End If
End Sub
```

Rules:

- **Always name the label `CleanUp:`** in entry-point procedures.
  Do not use `EH:` or `ErrorHandler:` in entry-points — those names
  suggest error-only paths, but `CleanUp:` is always reached.
- The `CleanUp:` block calls `AppLock False` unconditionally.
- `Err.Number <> 0` check shows the error message only when an error
  actually occurred.

> **Inconsistency in earlier projects:** GenerateBPAFilesRework uses `ErrorHandler:`,
> some AppLoader code uses `EH:`. Standardize on `CleanUp:` for all new entry points.

---

## 13. Performance — Array Processing

For any loop over more than ~500 rows, read data into a Variant array first,
process in memory, then write the result back in a single operation.

```vba
' Good — one read, in-memory loop, one write
Dim data As Variant
data = ws.Range("A1:C" & lastRow).Value   ' one hit to the worksheet
Dim i As Long
For i = 1 To UBound(data, 1)
    ' work on data(i, 1), data(i, 2), data(i, 3)
Next i
ws.Range("E1").Resize(outRow - 1, 3).Value = outArr   ' one write

' Bad — cell-by-cell loop (100x slower)
For i = 1 To lastRow
    ws.Cells(i, 5).Value = ws.Cells(i, 1).Value & " " & ws.Cells(i, 2).Value
Next i
```

### Header lookup — always use Application.Match

```vba
' Good — single worksheet call, fast on any size header row
Public Function FindHeaderCol(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim idx As Variant
    idx = Application.Match(headerText, ws.Rows(1), 0)
    If IsError(idx) Then FindHeaderCol = 0 Else FindHeaderCol = CLng(idx)
End Function

' Bad — For loop over every cell in row 1 (slow on wide sheets)
Dim c As Long
For c = 1 To lastCol
    If UCase$(ws.Cells(1, c).Value) = UCase$(headerText) Then ...
Next c
```

> **Inconsistency in earlier projects:** banking_details_macroRework and
> GenerateBPAFilesRework use a For loop for header scanning. VNE and new
> projects use `Application.Match`. Always use `Application.Match` — it is
> orders of magnitude faster on wide sheets and requires less code.

Additional rules:

- Use `Scripting.Dictionary` for deduplication and lookups (O(1) per key,
  not a nested loop).
- Use `Application.StatusBar` + `DoEvents` every ~5 000 rows so Excel does
  not appear frozen, but don't call `DoEvents` inside tight inner loops.
- Define a `WRITE_CHUNK_SIZE` constant (e.g. `5000`) when you need to write
  large arrays in batches to avoid memory exhaustion.

---

## 14. Private Helper Naming

Helper functions that are only used inside one module are prefixed `H_`:

```vba
Private Function H_StripWord(ByVal s As String, ByVal w As String) As String
Private Sub     H_StripCommonSuffixTokens(...)
Private Function H_Min3(...)
Private Sub     H_RunRow(...)     ' setup helpers
```

The `H_` prefix signals "not part of the public API; refactor freely".
It also groups related private helpers together in the VBE procedure list.

Module-local private constants are the **one exception** where constants are
allowed outside Config — only when the constant is truly private to one
module and has no use elsewhere:

```vba
' Fine — private constant, only relevant to this one module
Private Const ALLOW_DERIVED_PRICE_FROM_BREAK As Boolean = False
```

---

## 15. Orchestrator Pattern

Every project has exactly one Orchestrator module (always the last-numbered
module).  It owns:

- All `Public Sub` entry points that users assign to buttons.
- The `AppLock True / AppLock False` brackets.
- The `ConfirmRun()` dialog before destructive operations.
- The `GoRunSheet` / navigation back to the Run sheet.
- The `RunAll` and `RunSelected` sequences.
- Step-guard validation ("Step 3 requires Step 1 output sheet").

Step modules (`VNE4_`, `VNE5_`, etc.) expose **one public entry point each**
(`RunOracleReduce`, `RunNameClustering`, etc.) and do nothing else publicly.
They do not call `AppLock`, show MsgBox dialogs (except errors), or navigate
sheets — that is the Orchestrator's job.

```vba
' Good — Orchestrator controls the frame
Public Sub RunStep2_OracleReduce()
    If Not ConfirmRun("Step 2", "...") Then Exit Sub
    AppLock True
    VNEProgressStart "Step 2", 1
    VNEProgressBeginStep 1, "Oracle Reduce"
    VNE4_OracleReduce.RunOracleReduce     ' step module does pure work
    VNEProgressStepComplete
    VNEProgressFinish
    AppLock False
    GoRunSheet
End Sub
```

---

## 16. Pre-Run Validation Pattern

Before any heavy run (RunAll or multi-step), validate all required global
state in a dedicated private function:

```vba
Private Function ValidateRunState() As Boolean
    If g_FileCount = 0 Then
        MsgBox "No files added. Please add at least one file.", vbExclamation, "AppLoader"
        ValidateRunState = False: Exit Function
    End If
    If Trim$(g_ProcessorID) = "" Then
        MsgBox "Processor ID is required.", vbExclamation, "AppLoader"
        ValidateRunState = False: Exit Function
    End If
    If Trim$(g_OutputPath) = "" Then
        MsgBox "Output path is required.", vbExclamation, "AppLoader"
        ValidateRunState = False: Exit Function
    End If
    ValidateRunState = True
End Function

' Usage in RunAll():
Public Sub RunAll()
    If Not ValidateRunState() Then Exit Sub
    AppLock True
    ' ...
End Sub
```

Rules:

- One validation function per Orchestrator — do not scatter validation
  checks inline throughout RunAll.
- Each check shows a specific, actionable message before returning False.
- Returns False on the first failure — do not try to report all errors at once.
- For the VNE-style projects (no global queue, config is read from a sheet),
  config validation is done inside `ReadConfig()` itself rather than a
  separate `ValidateRunState()`.

---

## 17. Post-Run Summary Pattern

Show a single summary MsgBox at the end of RunAll — never mid-run pop-ups
for each file. Keep a private `ShowSummary()` function:

```vba
Private Sub ShowSummary(ByRef results() As ProcessResult, _
                        ByVal total As Long, _
                        ByVal processed As Long, _
                        ByVal failed As Long)
    Dim msg As String
    msg = "Processing complete." & vbCrLf & vbCrLf
    msg = msg & "  Total:         " & total     & vbCrLf
    msg = msg & "  Processed OK:  " & processed & vbCrLf
    msg = msg & "  Failed:        " & failed    & vbCrLf

    If failed > 0 Then
        msg = msg & vbCrLf & "FAILED FILES:" & vbCrLf
        Dim i As Long
        For i = 0 To total - 1
            If Not results(i).Success Then
                msg = msg & "  - " & results(i).SourceFile & vbCrLf
                msg = msg & "    " & results(i).ErrorMessage & vbCrLf
            End If
        Next i
        msg = msg & vbCrLf & "See the Log sheet for details."
    End If

    msg = msg & vbCrLf & "Output saved to:" & vbCrLf & "  " & g_OutputPath
    MsgBox msg, IIf(failed > 0, vbExclamation, vbInformation), "ProjectName - Run Summary"
End Sub
```

Rules:

- Summary format: totals first, then failures list, then output location.
- Always direct the user to the Log sheet when failures exist.
- The `IIf(failed > 0, vbExclamation, vbInformation)` pattern picks the
  right icon automatically.

---

## 18. User Defined Types (UDTs)

Use `Type...End Type` structures to group related per-item result data rather
than passing 5+ separate ByRef parameters:

```vba
' In Mod1_Config (or at the top of the Orchestrator if only used there):
Public Type ProcessResult
    SourceFile    As String
    TemplateName  As String
    Success       As Boolean
    RowsProcessed As Long
    ChangesApplied As Long
    NoMatchCount  As Long
    ErrorMessage  As String
    OutputFile    As String
End Type

' Usage:
Dim result As ProcessResult
result = ProcessFile(filePath, templateName)
If result.Success Then processed = processed + 1
```

Rules:

- Declare UDTs in Config if they are used by multiple modules.
- Declare UDTs at the top of the Orchestrator if used only there.
- UDT names are PascalCase nouns (`ProcessResult`, `TemplateConfig`).
- UDT fields are PascalCase (`result.RowsProcessed`, not `result.rows_processed`).
- Never use a `Variant` array or a Scripting.Dictionary to pass a structured
  result when a UDT would be cleaner.

---

## 19. Progress UI Pattern

Progress reporting uses a dedicated ProgressUI module (always second-to-last
numbered) that wraps a `frmProgress` UserForm.

Public API (four procedures):

```vba
VNEProgressStart    "Title", totalSteps    ' show form, initialise
VNEProgressBeginStep stepNum, "Label"      ' update step label
VNEProgressStepComplete                    ' advance bar
VNEProgressFinish                          ' hide form
```

Rules:

- The progress form is **modal** during each step. Use `SafeDoEvents` (a
  wrapper that suppresses re-entrant calls) — never raw `DoEvents` inside a
  tight loop.
- `g_VNE_Abort` (or equivalent global) is set by the form's Cancel button.
  Step modules check it only if they have a natural checkpoint (e.g. after
  each cluster). They do not poll on every row.
- When running a single step standalone the bar shows 1 step.
  When running multi-step (`RunAll`, `RunSelected`) the bar shows N steps.
- `g_VNE_Silent = True` suppresses per-step completion MsgBoxes during
  a multi-step run; the Orchestrator shows one summary MsgBox at the end.

---

## 20. Logging Pattern

Every project has a Log sheet. Use the shared `SafeLog` helper:

```vba
SafeLog ThisWorkbook, "VNE4_OracleReduce", "Processed 412 319 rows -> 8 744 unique"
```

### Log granularity — file-level not row-level

Write **one log row per processed file**, not one row per cell change.
The AppLoader pattern is the standard to follow:

| Column | Content |
| --- | --- |
| Timestamp | `Now` |
| Processor_ID / Source | Who/what ran |
| Source_File | Filename only (not full path) |
| Template / Step | Which template or step was active |
| Rows_Processed | Count of data rows touched |
| Changes_Applied | Count of cells actually changed |
| Status | `OK` or `FAILED` |
| Notes | Plain-English success summary or failure hint for the user |

> **Lesson from earlier projects:** GenerateBPAFilesRework logged one row
> per cell action, producing thousands of log rows from just two files.
> This made the Log sheet unusable as a run record. Always log at the
> file/step level.

Additional rules:

- `SafeLog` wraps its writes in `On Error Resume Next` so a missing Log
  sheet never crashes a step.
- Log the **start** (input counts) and **end** (output counts) of each
  major step. Don't log individual row decisions (use `Debug.Print` for that
  during development, then remove before shipping).
- For failure rows, the Notes column should contain a user-facing tip
  (what to fix), not a raw VBA error message.

---

## 21. Sheet Utilities

Always use the shared helper functions from the Helpers module — never write
inline sheet-access code:

| Helper | Purpose |
| --- | --- |
| `GetOrCreateSheet(wb, name)` | Return existing sheet or create at end |
| `GetSheetIfExists(wb, name)` | Return sheet or `Nothing` (no error) |
| `FindHeaderCol(ws, header)` | Column index via `Application.Match` in row 1 |
| `FindHeaderColInRow(ws, hdr, row)` | Same but in a specific row number |
| `ClearAndPrepareSheet(ws)` | Clear content + outlines + tables |
| `StyleHeader(rng)` | Apply standard bold + color to header row |
| `ApplyAutoFilter(ws)` | Enable AutoFilter safely (won't toggle off) |
| `NewDict([textCompare])` | Create Scripting.Dictionary with consistent compare mode |
| `EnsureFolderExists(path)` | Create folder tree (uses FSO, handles intermediate dirs) |
| `PathJoin(folder, file)` | Combine path + filename, handles trailing separator |

Sheet names used in code must always come from a Config constant, never a
string literal:

```vba
' Good
Dim ws As Worksheet: Set ws = GetOrCreateSheet(wb, VNE_SH_LOG)

' Bad
Dim ws As Worksheet: Set ws = wb.Worksheets("Log")
```

---

## 22. User-Facing Messages

### Confirmation dialogs (`ConfirmRun`)

Use the private `ConfirmRun` helper in the Orchestrator for any destructive
operation (sheet will be overwritten, file will be modified):

```vba
If Not ConfirmRun("Step 1: Name Clustering", _
    "This will read the vendor name source and group similar names." & vbCrLf & _
    "The '" & VNE_DEF_SH_CLUSTERS & "' sheet will be overwritten.") Then Exit Sub
```

### Error messages

```vba
MsgBox "Oracle Vendor Number Column must be filled in.", vbExclamation, "VNE - Step 2"
```

Rules:

- Always provide a `title` (third arg) in the format `"ProjectName - Context"`.
- Use `vbExclamation` for configuration/validation errors.
- Use `vbCritical` for unexpected runtime errors.
- Use `vbInformation` for success summaries.
- Multi-line success messages list each output sheet/file on its own line,
  indented with two spaces.

### User-facing error hints in the Log

When writing a failure to the Log sheet, translate the raw VBA error into
an actionable tip. Pattern from AppLoader `BuildFailureHint()`:

```vba
' Check most-specific patterns first, fall through to generic.
If InStr(UCase$(errMsg), "PERMISSION DENIED") > 0 Then
    BuildFailureHint = errMsg & " | TIP: The file is locked by another application."
    Exit Function
End If
' ... more specific patterns ...
BuildFailureHint = errMsg & " | TIP: Check the Log sheet and contact your administrator."
```

### Status bar updates

```vba
Application.StatusBar = "Oracle Reduce: loading " & lastRow & " rows..."
DoEvents
```

Use plain English, include the module name, and include a count when relevant.
Always reset with `Application.StatusBar = False` (done by `AppLock False`).

---

## 23. Forms (UserForms)

- Form code is exported as `frmXxx_Code.bas` — the actual `.frm` / `.frx`
  files are not committed (binary).
- Forms follow a **Wizard** pattern: one form walks the user through all
  configuration fields, then writes them to the Config sheet and calls the
  Orchestrator.
- Forms do **not** call step modules directly — they call the Orchestrator
  (`RunSelected`, `QuickRerun`, etc.).
- Form controls use the prefix `txt` (TextBox), `cmb` (ComboBox),
  `chk` (CheckBox), `btn` (CommandButton), `lbl` (Label).
- The `btnCancel` / `btnClose` handler always calls `Unload Me`.

### Documenting UserForm layout in comments

When writing layout comments (e.g. at the top of a form's code module to
describe its controls), list each control's properties **in the same order
they appear top-to-bottom in the VBE Properties window** (alphabetical tab).
This lets you read values straight off the screen without hunting for them.

Group position and size onto a single line (`Left Top Width Height`) since
they are always set together.

**Preferred style:**

```vba
' FRAME 1  -  Vendor Names File  (Step 1 input)
' -------------------------------------------------------
'   fraVendor       Frame
'       Caption = "Step 1  -  Vendor Names File"
'       Left=10  Top=38  Width=544  Height=100
'
'   Inside fraVendor (coordinates are relative to the frame):
'
'   lblVFile        Label
'       Caption = "File:"
'       Left=8  Top=24  Width=36  Height=14
'
'   txtVFile        TextBox
'       Left=48  Top=22  Width=380  Height=18
'
'   btnVFile        CommandButton
'       Caption = "..."
'       Left=432  Top=22  Width=24  Height=18
```

**Not preferred** — verbose label-per-line format makes it slow to fill in
and hard to scan:

```vba
'  #2  FRAME — fraPath
'      Control type : Frame
'      (Name)       : fraPath
'      Caption      : Save Location
'      Left         : 10
'      Top          : 36
'      Width        : 454
'      Height       : 72
```

Rules:

- One header line per control: `ControlName    ControlType` (name left-aligned,
  type separated by enough spaces to visually align across the block).
- List only the properties you actually set — skip defaults you are leaving
  unchanged (e.g. `Enabled = True`, `Visible = True`).
- Position line is always last: `Left=N  Top=N  Width=N  Height=N`.
- For controls inside a container (Frame, MultiPage), add an indent and note
  that coordinates are relative to the parent.
- The VBE Properties window **Alphabetic** tab is the reference order. When
  in doubt, open the Properties panel and read top-to-bottom.

---

## 24. Deprecated Constants

When a constant is superseded but kept for backward compatibility, mark it
clearly and do not use it in any new code:

```vba
' (Deprecated - superseded by CRIT_COL_VENDOR_NUM; kept for reference only)
Public Const CRIT_HDR_VENDOR_NUM As String = "VENDOR_NUMBER"
```

Rules:

- Never remove a deprecated constant until you have confirmed no existing
  workbook formulas or cell references use it.
- The `(Deprecated)` comment must appear on the same line as the `Const`.
- New features never reference deprecated constants.

---

## 25. What NOT to Do

| Prohibited | Reason |
| --- | --- |
| Hardcode sheet names as string literals | Config constants make renames a one-line change |
| Cell-by-cell loops over large data | 100-1000x slower than array reads |
| `On Error Resume Next` left active | Masks bugs; always close with `On Error GoTo 0` |
| `Integer` for loop counters or row indices | Overflows at 32 767; use `Long` |
| Business logic in the Config module | It's a declarations-only file |
| Functions or Subs in the Config module | Only `Const`, `Public` vars, and `Type` allowed |
| Business logic in the Helpers module | Helpers are pure utilities with no domain knowledge |
| Step modules calling `AppLock` or `MsgBox` | That belongs in the Orchestrator |
| Storing config in module-level variables | Re-read from `ReadConfig()` each run |
| `GoTo` for normal flow control | Only use `GoTo` for `CleanUp` / error-exit labels |
| Magic numbers inline (`If score > 0.82`) | Define a named constant in Config |
| Comments that restate the code (`' increment i`) | Comment *why*, not *what* |
| `For` loop to scan row 1 for a header | Use `Application.Match` — much faster |
| `GF_` prefix on new utility functions | That prefix is retired; no prefix needed |
| Globals without `g_` prefix | Breaks naming convention; hard to spot in code |
| Globals declared outside Config module | Config is the only valid home for globals |
| Per-row/per-cell log entries | Log at file or step level; cell-level logs explode in size |
| Dim inside a loop body | VBA hoists all Dims anyway; put them at the top of the procedure |
| Variable named `asc`, `name`, `left`, `row`, `date`, etc. | Shadows VBA built-ins silently — see §27a reserved word table |
| Single-line `If...Then` mixed with block `ElseIf` | Causes "Else without If" compile error — always use block form |
| Array index without bounds check | Causes error 9 — check `UBound` and use `GetSheetIfExists` |

---

## 26. Known Inconsistencies in Earlier Projects

These are patterns found in older projects in this repo that do **not** match
the standards above. When reading old code, you will encounter these.
When writing new code or extending an old project, fix these forward.

| Project | Inconsistency | Standard to apply |
| --- | --- | --- |
| banking_details_macroRework | `GetDefaultOutputColumns()` function lives in Mod1_Config | Move functions to Mod2_Helpers |
| banking_details_macroRework | Section dividers use `'========================` (= signs) | Use `' ---` dash style |
| GenerateBPAFilesRework | `Public AbortProcess As Boolean` without `g_` prefix | Always prefix with `g_` |
| GenerateBPAFilesRework | Globals declared in Mod1_Helpers | Declare globals in Config |
| GenerateBPAFilesRework | All utility functions prefixed `GF_` | Drop `GF_` prefix in new code |
| GenerateBPAFilesRework | `ErrorHandler:` label in entry points | Use `CleanUp:` |
| GenerateBPAFilesRework | `FindHeaderCol` uses a For loop | Use `Application.Match` |
| AppLoader | AppLock did not include `DisplayAlerts` in early versions | Include `DisplayAlerts` in every AppLock |
| AppLoader / BPA early | Log wrote one row per cell change | Log one row per file/step |
| WireProjectRework | No `Application.DisplayAlerts` in AppLock | Include in all new projects |
| All pre-VNE projects | `NzStr` / `NzTrim` / `GF_NzStr` — three names for same function | Use `NzStr` |

---

## 27. Common VBA Errors — Causes and Fixes

These three errors appeared repeatedly during development across this repo.
Understanding the root cause prevents hours of debugging.

---

### 27a. Variable Names That Clash with VBA Reserved Words

**Symptom:** Compile error, unexpected behaviour, or a built-in function stops
working after you declare a variable with the same name.

**Root cause:** VBA lets you shadow built-in functions and keywords with your
own variable names. The local name wins, breaking every call to the built-in
in that scope.

```vba
' BAD — "asc" shadows the Asc() built-in string function
Dim asc As Long
asc = 65
Debug.Print Asc("A")   ' now broken — refers to the Long variable, not the function

' BAD — "name" shadows the Name statement (renames files)
Dim name As String

' BAD — "left" shadows the Left() string function AND the .Left property
Dim left As Long
```

**Fix:** Add a meaningful prefix or suffix that makes the intent clear and
guarantees no collision:

```vba
' Good — prefix with the role or type
Dim ascVal    As Long     ' ASCII value
Dim vendName  As String   ' vendor name
Dim leftPos   As Long     ' left position / column index
Dim rowIdx    As Long     ' row index (not "row" — clashes with Range.Row)
Dim colIdx    As Long     ' column index (not "col" is fine, but "column" is not)
```

**VBA names to avoid as local variables** — these shadow built-ins or
keywords in ways that are hard to spot:

| Dangerous name | Why | Safe alternative |
| --- | --- | --- |
| `asc` | Shadows `Asc()` | `ascVal`, `charCode` |
| `name` | Shadows `Name` statement | `sheetName`, `vendName`, `fileName` |
| `left` | Shadows `Left()` and `.Left` | `leftPos`, `colStart` |
| `right` | Shadows `Right()` and `.Right` | `rightPos`, `colEnd` |
| `mid` | Shadows `Mid()` | `midStr`, `midVal` |
| `date` | Shadows `Date` function/type | `runDate`, `startDate` |
| `time` | Shadows `Time` function | `runTime`, `elapsed` |
| `error` | Shadows `Error()` | `errMsg`, `errNum` |
| `string` | Shadows `String()` fill function | `strVal`, `rawText` |
| `array` | Shadows `Array()` constructor | `dataArr`, `resultArr` |
| `index` | Shadows `.Index` property | `rowIdx`, `colIdx`, `matchIdx` |
| `count` | Shadows `.Count` property | `rowCount`, `fileCount` |
| `value` | Shadows `.Value` property | `cellVal`, `rawVal` |
| `row` | Shadows `.Row` property | `rowNum`, `rowIdx` |
| `column` | Shadows `.Column` property | `colNum`, `colIdx` |
| `type` | Shadows `Type` UDT keyword | `fileType`, `templateType` |
| `end` | Reserved word | `endRow`, `endCol` |

**Rule of thumb:** If a name appears highlighted in the VBE editor in a
different colour, it is already claimed by VBA. Rename it immediately.

---

### 27b. "Else without If" / "ElseIf without If" Compile Error

**Symptom:** VBA compile error: `Else without If` or `ElseIf without If`,
pointing at a line that does have a matching `If` above it.

**Root cause:** VBA has two forms of `If`. They cannot be mixed:

```vba
' SINGLE-LINE form — entire statement on one line, no End If
If condition Then DoSomething

' BLOCK form — spans multiple lines, requires End If
If condition Then
    DoSomething
End If
```

If you write a single-line `If...Then statement` and then put `ElseIf` or
`Else` on the NEXT line, VBA reads the first line as a complete statement
and the `ElseIf` is orphaned — hence "Else without If".

```vba
' BAD — mixes single-line form with block ElseIf
If cidl < cid2 Then H_CmpCluster = -1        ' VBA sees this as COMPLETE
ElseIf cidl > cid2 Then H_CmpCluster = 1     ' Compile error: ElseIf without If

' GOOD — pure block form, all branches visible, no ambiguity
If cidl < cid2 Then
    H_CmpCluster = -1
ElseIf cidl > cid2 Then
    H_CmpCluster = 1
ElseIf cnt1 > cnt2 Then
    H_CmpCluster = -1
ElseIf cnt1 < cnt2 Then
    H_CmpCluster = 1
Else
    H_CmpCluster = 0
End If
```

**Rules:**

- **Never use the single-line `If...Then statement` form** when you need
  `ElseIf` or `Else`. Always switch to the block form.
- The single-line form is only acceptable for truly trivial one-liners with
  no branches: `If ws Is Nothing Then Exit Sub`
- When in doubt, always use the block form. It is never wrong.
- The `Else: H_CmpCluster = 0` colon syntax is also fragile — avoid it.
  Write `Else` on its own line, followed by the statement on the next line.

---

### 27c. Runtime Error '9' — Subscript Out of Range

**Symptom:** Runtime error 9: `Subscript out of range`, usually pointing at
an array index, a `Worksheets("name")` call, or a `Dictionary` lookup.

**Root cause:** You are asking for an element that does not exist at that
index. Four common causes:

#### Cause 1 — Array not yet initialized

```vba
' BAD — arr() is declared but never ReDim'd; UBound crashes with error 9
Dim arr() As String
Debug.Print UBound(arr)   ' error 9

' GOOD — check before accessing
If g_FileCount = 0 Then Exit Sub
Debug.Print UBound(g_FilePaths)
```

#### Cause 2 — Wrong array base (0 vs 1)

```vba
' BAD — data from ws.Range().Value is 1-based; accessing index 0 crashes
Dim data As Variant
data = ws.Range("A1:C10").Value
Debug.Print data(0, 1)    ' error 9 — first row is data(1, 1)

' GOOD — always start at 1 for worksheet range arrays
For i = 1 To UBound(data, 1)
    Debug.Print data(i, 1)
Next i
```

#### Cause 3 — Sheet name does not exist

```vba
' BAD — crashes if the sheet was renamed or not yet created
Set ws = ThisWorkbook.Worksheets("VN_Clusters")   ' error 9 if missing

' GOOD — use GetSheetIfExists and check
Set ws = GetSheetIfExists(ThisWorkbook, VNE_SH_CLUSTERS)
If ws Is Nothing Then
    MsgBox "Run Step 1 first.", vbExclamation, "VNE"
    Exit Sub
End If
```

#### Cause 4 — Reading past the end of an array or Split result

```vba
' BAD — Split("A,B", ",") gives indices 0 and 1; accessing index 2 crashes
Dim parts() As String
parts = Split(someText, ",")
Debug.Print parts(2)    ' error 9 if fewer than 3 tokens

' GOOD — always check UBound before indexing
If UBound(parts) >= 2 Then Debug.Print parts(2)
```

**General defence pattern** — wrap array access in a bounds check helper:

```vba
' In Mod2_Helpers — safe array read that returns "" instead of crashing
Public Function SafeArr(ByRef arr() As String, ByVal idx As Long) As String
    On Error Resume Next
    SafeArr = arr(idx)
    On Error GoTo 0
End Function
```

**Quick checklist when you see error 9:**

1. Is the array declared with `()` but never `ReDim`'d?
2. Are you using index `0` on a worksheet-range array (which is 1-based)?
3. Does the sheet / workbook name actually exist at that moment?
4. Did `Split()` return fewer tokens than you expected?
5. Is a `For...Next` loop going one index past the end?

---

## Quick Reference Card

```text
Every module      Option Explicit + 60-char header box (leading space: ' ===...)
Numbering         1=Config, 2=Helpers, 3=Setup, 4..N-2=Steps, N-1=ProgressUI, N=Orchestrator
Config module     Const + Public globals (g_ prefix) + Type UDTs only. No functions.
Globals           g_ prefix, declared in Config, reset at start of each run.
Runtime config    ReadConfig() -> Dictionary -> check Nothing -> Exit Sub
AppLock           ScreenUpdating + EnableEvents + DisplayAlerts + Calculation. Orchestrator only.
Array reads       Read range -> Variant array -> process -> write once
Header lookup     Application.Match, not a For loop
Private helpers   H_ prefix. No GF_ prefix.
Error handling    On Error Resume Next for 1-2 lines only; GoTo CleanUp for entry points
Logging           1 row per file/step. SafeLog wb, "ModuleName", "message with counts".
Sheet access      GetOrCreateSheet / GetSheetIfExists — never direct Worksheets("name")
Messages          vbExclamation=config error  vbCritical=runtime error  vbInformation=success
UDTs              PascalCase noun. Declare in Config if shared, in Orchestrator if local.
Validation        ValidateRunState() private function before RunAll(). One check per guard.
Summary           ShowSummary() private function at end of RunAll(). One MsgBox, not many.
Reserved words    Avoid: asc name left right mid date time error string array index count value row column type end
If/ElseIf         Always block form (If...Then / ElseIf...Then / Else / End If). Never mix single-line with block.
Error 9           Check UBound before array access. Use GetSheetIfExists. 1-based for ws.Range().Value arrays.
```

---

*This is a living document — see §0 for update instructions and changelog.*
