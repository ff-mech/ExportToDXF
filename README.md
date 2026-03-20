# ExportToDXF

## Overview

`ExportToDXF.bas` is a SolidWorks macro that scans a source folder for part files, analyzes each part, and exports valid sheet metal models to DXF in a destination folder.

The macro is designed to be practical for production use:

- Handles many files in one run
- Applies validation before export
- Separates results into Exported, Failed, and Skipped
- Produces a readable text log with timing and summary metrics

## What This Macro Does

For each `.sldprt` file, the macro will:

1. Open the part silently.
2. Run an initial rebuild.
3. Detect whether it is a sheet metal part.
4. Detect rollback state and attempt roll-forward if needed.
5. Find and flatten the Flat-Pattern feature when bend features are present.
6. Rebuild and validate feature errors after flattening.
7. Export to DXF using `ExportToDWG2` with predefined sheet metal options.
8. Log the result with elapsed time.

## What It Does Not Do

- It does not recurse into subfolders (current behavior is top-level files only).
- It does not create destination folders automatically; source and destination must already exist.
- It does not save check/failed preview images (those image features were intentionally removed in this revision).

## Requirements

- SolidWorks installed (version compatible with the VBA API used)
- Macro execution enabled in SolidWorks
- Access to `ExportToDXF.bas`
- Read access to source part folder
- Write access to destination folder

## Files in This Repository

- `ExportToDXF.bas`: Main SolidWorks VBA macro.
- `README.md`: Documentation and usage instructions.
- `ExportToDXF.swp`: SolidWorks macro project file (when available in your environment).

## Setup in SolidWorks

1. Open SolidWorks.
2. Go to **Tools > Macro > New** or **Tools > Macro > Edit**.
3. Load `ExportToDXF.bas` (or open the `.swp` if preferred).
4. Ensure references to SolidWorks object libraries are available.
5. Save the macro in your standard macro location.

## How to Run

1. Start the macro (`main`).
2. When prompted, paste the source folder path that contains `.sldprt` files.
3. Paste the destination folder path for DXF output.
4. Wait for processing to complete.
5. Review the completion message and open the generated log:
	- `DXF_Export_Log.txt` in the destination folder.

## Output

### Exported Files

- One `.dxf` per successfully exported source part.
- Output file name matches the source part base name.

### Log File

The macro writes a structured log at:

- `<destination>\\DXF_Export_Log.txt`

Log sections include:

- Errors / Failures first
- Skips / Warnings
- Successful exports
- Summary counts and timing

Summary metrics include:

- Total processed
- Exported count
- Failed count
- Skipped count
- Average time per part
- Total run time

## Validation and Decision Rules

The macro classifies each file into one of three outcomes:

- **Exported**: DXF export succeeded.
- **Failed**: Required processing step failed (open, rebuild, flatten, validation, or export).
- **Skipped**: Part is not sheet metal.

Additional rules:

- Multiple non-flat configurations trigger a warning.
- Missing Flat-Pattern feature is treated as failure.
- Rollback bar not at end triggers automatic roll-forward attempt.
- Flat pattern error code (non-warning) after flattening is treated as failure.

## Internal Processing Flow

High-level sequence per part:

1. Open file with `OpenDoc6` (silent).
2. Rebuild using `ForceRebuild3`.
3. Analyze features for:
	- Sheet metal presence
	- Flat-Pattern feature
	- Bend-related features
	- Rollback state
	- Configuration count
4. If rolled back, move rollback bar to end and rebuild.
5. If bends exist, unsuppress flat pattern and rebuild.
6. Validate flat pattern feature error status.
7. Export DXF via `ExportToDWG2`.
8. Close processed documents and clean up extras opened during processing.

## Troubleshooting

If you see many failures, check the following first:

- Source folder path is correct and accessible.
- Destination folder exists and is writable.
- Parts open manually in SolidWorks without major errors.
- Flat-Pattern feature exists for expected sheet metal parts.
- Macro security settings allow execution.

Common log messages and meanings:

- `FAIL - Could not open file`: SolidWorks could not open the part.
- `SKIP - Not a sheet metal part`: Non-sheet-metal file detected.
- `FAIL - No Flat-Pattern feature found`: Sheet metal model missing required flat pattern.
- `WARN - Rollback state detected`: Macro found rollback and attempted correction.
- `FAIL - ExportToDWG2 returned False`: Export API call failed.

## Safe Customization Ideas

If you plan to extend this macro, these are usually safe places to start:

- Add recursive folder traversal.
- Add filtering by file naming conventions.
- Make sheet metal export options configurable.
- Add CSV/JSON summary output in addition to text log.
- Add optional "continue on non-critical warning" behavior.

When modifying, keep the existing result classification (Exported/Failed/Skipped) so reporting remains consistent.

## Notes on Revision

Initial Release

## License / Usage

If you use this internally, consider adding your team/company license terms here.

---

If helpful, this README can also be expanded with a "Developer Notes" section documenting each VBA function and expected SolidWorks API behavior in more depth.
