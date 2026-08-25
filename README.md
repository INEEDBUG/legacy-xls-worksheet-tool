# Legacy XLS Worksheet Tool

Version 1.2.0 is a completely offline Windows PowerShell tool for inspecting
and removing legacy BIFF/OLE `.xls` worksheet-protection records from files
you are authorized to process.

This repository contains the real PowerShell/WinForms implementation and the
scripts used to build the portable package. It intentionally contains no
workbooks, processed copies, logs, credentials, personal paths, or binary
release archive.

## Scope and safety boundary

The tool handles legacy `.xls` worksheet, workbook-structure, window, and
object protection records by making an equal-length, minimal BIFF edit in a
new output copy. The original file is never overwritten, and the output is
verified after writing.

Use it only for files that you own or are explicitly authorized to process.

The tool does **not**:

- recover or bypass a true file-open password/encryption;
- support `.xlsx`, `.xlsm`, or `.xlsb` files;
- recover the original text behind a legacy 16-bit worksheet verifier;
- modify the source workbook in place;
- require Excel, Python, Node.js, an installer, a network connection, or a
  third-party runtime.

If a file is encrypted before Excel can display its contents, the tool reports
and skips it. Worksheet protection is an edit restriction, not file-open
encryption.

## Use the GUI

On a Windows machine with built-in Windows PowerShell 5.1:

1. Keep all files in this repository together.
2. Double-click `启动Excel解锁工具.vbs`.
3. Add one or more `.xls` files, or choose a folder.
4. Review the detected status and select **一键解除保护**.

The GUI supports multi-file selection, folder recursion, drag-and-drop, and
automatic unique output names such as `report_已解除保护.xls`. The VBS entry
point starts the GUI without leaving a console window open.

## Use the CLI

```powershell
ExcelUnlocker.cmd doctor
ExcelUnlocker.cmd inspect <file.xls>
ExcelUnlocker.cmd unlock <file.xls>
ExcelUnlocker.cmd unlock <folder> --recursive --out <output-folder>
ExcelUnlocker.cmd unlock <folder> --recursive --json
```

`doctor` performs an environment/capability check. `inspect` only reads the
input. `unlock` writes a new copy and then inspects that copy. `--json` emits a
stable machine-readable result. Exit code `0` means success or a safe skip,
`1` means at least one input failed, and `2` means invalid arguments.

## Repository layout

- `ExcelUnlocker.ps1` — OLE Compound File and BIFF parser plus safe writer.
- `ExcelUnlocker-GUI.ps1` — Windows Forms front end.
- `启动Excel解锁工具.vbs` — hidden GUI launcher.
- `ExcelUnlocker.cmd` — command-line entry point.
- `tests/Test-ExcelUnlocker.ps1` — self-contained regression test. It builds a
  synthetic protected OLE/BIFF workbook in a temporary directory at runtime;
  no `.xls` fixture is stored in the repository.
- `build.ps1` — creates a fresh v1.2.0 source package and SHA-256 manifest.

## Test

Run from this repository on Windows PowerShell 5.1:

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-ExcelUnlocker.ps1
```

The test checks PowerShell parsing, `doctor`, synthetic BIFF inspection,
non-destructive unlock, output re-inspection, and original-file preservation.
Its temporary synthetic files are removed in a `finally` block.

## Build the portable package

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\build.ps1
```

The command writes an ignored `dist\Excel工作表离线解锁工具_v1.2.0.zip` and a
SHA-256 manifest. The public repository keeps the source and build recipe,
not the generated binary archive.

