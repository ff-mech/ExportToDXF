# ExportToDXF

> **A SolidWorks VBA macro for batch-exporting sheet metal parts to DXF — built for production use.**

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Requirements](#requirements)
- [Repository Contents](#repository-contents)
- [Setup](#setup)
- [Usage](#usage)
- [Output](#output)
- [Processing Logic](#processing-logic)
- [Validation & Decision Rules](#validation--decision-rules)
- [Log Reference](#log-reference)
- [Troubleshooting](#troubleshooting)
- [Extending the Macro](#extending-the-macro)
- [Changelog](#changelog)
- [License](#license)

---

## Overview

`ExportToDXF.bas` is a SolidWorks VBA macro that scans a source folder for `.sldprt` part files, analyzes each one, and batch-exports valid sheet metal models to DXF format in a specified destination folder.

It is designed to be reliable and practical for real production environments — handling large file sets, applying rigorous validation before each export, and producing a structured log file with timing data and a full summary.

---

## Features

| Capability | Details |
|---|---|
| **Batch processing** | Processes all `.sldprt` files in a folder in one run |
| **Sheet metal detection** | Automatically identifies and skips non-sheet-metal parts |
| **Flat-pattern handling** | Detects bend features and unsuppresses the Flat-Pattern feature before export |
| **Rollback correction** | Detects and auto-corrects parts with a rolled-back feature tree |
| **Feature validation** | Checks for feature errors after flattening before attempting export |
| **Silent operation** | Opens and processes files without interrupting the SolidWorks UI |
| **Structured logging** | Writes a timestamped log grouped by failures, skips, and successes |
| **Per-file timing** | Records elapsed time for every part and reports average and total run time |
| **Clean document tracking** | Tracks documents opened during processing and closes extras automatically |

---

## Requirements

- **SolidWorks** — Any version compatible with the SolidWorks VBA API (`SldWorks`, `ModelDoc2`, `PartDoc`, `ExportToDWG2`)
- **Macro execution enabled** — Macro security must permit running unsigned macros in your SolidWorks installation
- **Pre-existing folders** — Both the source and destination folders must already exist before running the macro
- **Read access** to the source part folder
- **Write access** to the destination folder

---

## Repository Contents

```
ExportToDXF/
├── ExportToDXF.bas      # Main SolidWorks VBA macro (entry point: main)
├── ExportToDXF.swp      # SolidWorks macro project file
└── README.md            # This file
```

---

## Setup

1. Open **SolidWorks**.
2. Navigate to **Tools → Macro → New** (or **Edit** to load an existing file).
3. Import `ExportToDXF.bas` into the macro editor, or open `ExportToDXF.swp` directly.
4. Confirm that SolidWorks object library references are available in the VBA IDE (**Tools → References**).
5. Save the macro to your standard macro library location for easy access.

---

## Usage

1. Launch the macro by running the `main` subroutine.
2. When prompted, enter the **source folder path** containing the `.sldprt` files to process.
3. When prompted, enter the **destination folder path** where exported `.dxf` files will be saved.
4. The macro will process each part silently. Wait for the completion dialog to appear.
5. Review the summary in the dialog, then open the generated log file for full details:

```
<destination folder>\DXF_Export_Log.txt
```

> **Note:** The macro processes only the top-level files in the source folder. Subfolders are not traversed. Both the source and destination folders must exist before running.

---

## Output

### DXF Files

- One `.dxf` file is created per successfully exported part.
- The output file name matches the source part's base name (without extension).
- Files are written directly to the destination folder.

### Log File

A structured plain-text log is written to `<destination>\DXF_Export_Log.txt` after all parts are processed.

**Log structure:**

```
DXF EXPORT LOG
Started: <timestamp>
Source Folder: <path>
Destination Folder: <path>
======================================================================

ERRORS / FAILURES FIRST
----------------------------------------------------------------------
FILE: part_a.sldprt
  FAIL - No Flat-Pattern feature found.
  Elapsed: 00:00:03

SKIPS / WARNINGS
----------------------------------------------------------------------
FILE: assembly_bracket.sldprt
  SKIP - Not a sheet metal part.
  Elapsed: 00:00:01

SUCCESSFUL EXPORTS
----------------------------------------------------------------------
FILE: panel_left.sldprt
  INFO - Rebuild succeeded.
  INFO - Rollback bar already at end.
  INFO - Flattening: Flat-Pattern1
  INFO - Flat pattern validation passed.
  OK - Exported to: C:\DXF\panel_left.dxf
  Elapsed: 00:00:05

SUMMARY
----------------------------------------------------------------------
Processed:            12
Exported:             9
Failed:               2
Skipped:              1
Average Time / Part:  00:00:04
Total Time:           00:00:48
======================================================================
```

---

## Processing Logic

The macro applies the following sequence to each `.sldprt` file:

```
┌─────────────────────────────────────────────┐
│  Open part silently (OpenDoc6)              │
│  ↓                                          │
│  Initial rebuild (ForceRebuild3)            │
│  ↓                                          │
│  Analyze feature tree:                      │
│    • Sheet metal present?  → No  → SKIP    │
│    • Flat-Pattern feature?  → No  → FAIL   │
│    • Bend features present?                 │
│    • Rollback bar at end?                   │
│    • Multiple configurations?               │
│  ↓                                          │
│  Correct rollback (if detected)             │
│  ↓                                          │
│  Unsuppress Flat-Pattern (if bends exist)   │
│  ↓                                          │
│  Rebuild + validate feature error state     │
│  ↓                                          │
│  Export DXF via ExportToDWG2               │
│  ↓                                          │
│  Log result + elapsed time                  │
│  ↓                                          │
│  Close document + clean up extras           │
└─────────────────────────────────────────────┘
```

**Outcome classification:**

| Result | Meaning |
|---|---|
| `Exported` | DXF written successfully |
| `Failed` | A required step failed (open, rebuild, flatten, validate, or export) |
| `Skipped` | Part is not a sheet metal part |

---

## Validation & Decision Rules

- **Non-sheet-metal parts** — Detected by absence of a `SheetMetal` feature; classified as `Skipped`.
- **Missing Flat-Pattern feature** — Sheet metal part with no `FlatPattern` feature in the tree; classified as `Failed`.
- **Rollback bar not at end** — Macro detects rolled-back features, moves the rollback bar to end, and rebuilds before continuing.
- **Multiple non-flat configurations** — Logged as a warning but processing continues.
- **Feature error after flattening** — A non-warning error code on the Flat-Pattern feature after unsuppression is treated as `Failed`.
- **Flatten fallback** — If `Select2` fails, the macro falls back to `SetSuppression2` on the active configuration.

---

## Log Reference

| Message | Meaning |
|---|---|
| `FAIL - Could not open file` | `OpenDoc6` returned `Nothing`; part could not be opened |
| `FAIL - Initial rebuild failed` | `ForceRebuild3` raised an exception on first rebuild |
| `SKIP - Not a sheet metal part` | No `SheetMetal` feature found in the part's feature tree |
| `FAIL - No Flat-Pattern feature found` | Sheet metal part is missing a `FlatPattern` feature |
| `WARN - Rollback state detected` | One or more features are rolled back; macro attempted roll-forward |
| `FAIL - Could not flatten part` | `EditUnsuppress2` and `SetSuppression2` both failed |
| `FAIL - Rebuild failed after flatten` | `ForceRebuild3` raised an exception after flattening |
| `FAIL - Feature tree error(s) after flatten` | Flat-Pattern feature reports a non-warning error code |
| `FAIL - ExportToDWG2 returned False` | The export API call completed but reported failure |
| `OK - Exported to: <path>` | DXF was written successfully |

---

## Troubleshooting

**Many parts failing to open**
- Verify the source folder path is correct and accessible from the machine running SolidWorks.
- Confirm macro security settings allow execution (**Tools → Options → System Options → General**).

**Parts failing at flatten or validate**
- Open the part manually in SolidWorks and inspect the Flat-Pattern feature for existing errors.
- Ensure the part does not have a suppressed or missing Flat-Pattern feature.

**Export returning False**
- Check that the destination folder exists and the macro has write permission.
- Verify the part opens and flattens cleanly when done manually.

**Parts opening with unexpected documents**
- The macro tracks all documents open before processing each part and closes extras after. If references load automatically, they will be closed unless they were already open before the run started.

---

## Extending the Macro

The following changes are safe starting points if you need to customize behavior:

- **Recursive folder traversal** — Replace the `Dir(sourceFolder & "*.sldprt")` loop with a recursive folder-walking approach.
- **File name filtering** — Add prefix/suffix checks to `fileName` before calling `ProcessPart`.
- **Configurable export options** — Replace the hardcoded `smOptions = 71` constant with a user-configurable variable or INI file read.
- **CSV or JSON output** — Add a secondary output loop after the main processing loop to write structured data alongside `DXF_Export_Log.txt`.
- **Continue on non-critical warnings** — Adjust the `CheckFeatureErrors` function to ignore specific error codes rather than treating all non-warning codes as failures.

> When modifying, preserve the `prExported` / `prFailed` / `prSkipped` result classification so that log grouping and summary counts remain consistent.

---

## Changelog

Initial Release
---

## License

This macro is provided as-is for internal use. Add your organization's license terms here if distributing within your team or externally.
