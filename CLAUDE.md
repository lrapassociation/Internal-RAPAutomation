# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

`FileDownload` is a Windows-only WPF desktop app (.NET 6, `net6.0-windows`, x64) for LRAP internal RAP (request/approval) automation. It has two workflows:

1. **Document download** (`MainWindow`): look up an employment by NID, list its documents, and download them from Azure Blob Storage to a chosen folder.
2. **RAP approval** (`ApprovalWindow`, opened via `btnApprovalProcess_Click`): read data from a RAP Excel workbook, check whether the request has already been recorded in the downstream workbooks, then insert the data into those workbooks and into SQL Server.

## Build / run

- Build: `dotnet build FileDownload.sln` (or build in Visual Studio). Run: `dotnet run --project FileDownload.csproj`.
- There are no tests and no linter configured.
- Building needs the machine-specific environment below. Without it the build or run will fail.
  - Two non-NuGet DLLs are referenced by relative `HintPath` outside the repo (in the Dropbox tree): `AzureStorageExplorerApp.dll` and `TechArch.dll`.
  - The `Microsoft.Office.Interop.Excel` COM reference needs Excel installed (`EmbedInteropTypes=False`).
- At startup, `MainWindow` reads a hardcoded Dropbox config path (`...\Request Group Document Download\Config\Configuration.txt`). It uses `TechArch` (`Service.getConfigVar`, `Security.revealConnectionString`) to obtain the SQL connection string, which is passed into `ApprovalWindow`.
- `bin/`, `obj/` and the `*_wpftmp.csproj` files are build artifacts, so ignore them.

## Architecture

- **Two Excel stacks are in use.** Excel COM interop (`Microsoft.Office.Interop.Excel`) and ClosedXML (plus `DocumentFormat.OpenXml`). Check which one a file uses before editing.
  - `GetExcelLists.cs`: read the ID lists out of the tracking workbooks. These are the RAP Tracker, Ex Gratia, Payment Spreadsheet and Weekly RAP Approvals lists. `ExtractWorksheet` returns a sheet as a `MemoryStream`.
  - `GetRAPExcelData.cs`: parse a RAP workbook into the models `RapData`, `RapCalcTableData`, `PaymentSheetData`, `RAPTrackerData` and `ExGratiaClosingData`. Header-to-property mapping goes through `Normalize` and `BuildHeaderMap`, which match by normalized header text, so changing a header name in the workbooks breaks the mapping.
  - `InsertExcelData.cs`: write those models into the destination workbooks (historical data, payment spreadsheet, RAP tracker, ex gratia, quarter).
  - `InsertSQL.cs`: write the Payment, Ex Gratia and RAP Tracker data to SQL Server via `System.Data.SqlClient`.
  - `RapExcelHelper.cs`: pull ranges for DataGrid display (including a transposed variant) and `ExportRangeAsPdf`.
- **Approval flow:** `ApprovalWindowHelperFunctions.CheckHistory` compares the RAP request number against the four ID lists to build `GridRow`s. `ApprovalWindow` shows these and then triggers the Excel and SQL inserts. The shared model classes (`RapData`, `RapHistoryCheck`, `GridRow`, `RapFile`, `RapCalcTableData`) live in `ApprovalWindowHelperFunctions.cs`, not in separate files.
- The Excel files and shares are on Dropbox paths, so file locks and sync state can affect runs.
