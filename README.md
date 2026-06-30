# ProcessDocumentFunction

[![.NET](https://img.shields.io/badge/.NET-10.0-512BD4?logo=dotnet)](https://dotnet.microsoft.com/)
[![Azure Functions](https://img.shields.io/badge/Azure_Functions-v4-0062AD?logo=azurefunctions)](https://learn.microsoft.com/azure/azure-functions/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE.txt)

An HTTP-triggered Azure Function that validates and annotates `.xlsx` files. It renames administrative sheets, detects duplicates, flags incorrect titles, and highlights excessive time entries — returning a color-coded workbook.

## Features

- **Sheet renaming** — renames legacy admin sheet names to clean ones (e.g. `AdministrativeActivitiesDirecto` → `Directors`)
- **Vacation/Illness validation** — detects duplicate entries (red) and incorrect titles (orange) on the `VacationIllness` sheet
- **Other sheet validation** — flags excessive daily hours (>10h, orange), duplicate rows (red), and vacation/illness titles on non-Vacation sheets (red)
- **Formatting-only mode** — skip data validation and only apply column widths + auto-filters
- **OpenTelemetry** — built-in observability via Azure Monitor Exporter

## Color Coding

| Color  | Hex       | Meaning                                          |
|--------|-----------|--------------------------------------------------|
| Red    | `#FF0000` | Duplicate rows / incorrect titles on wrong sheet |
| Orange | `#F4B084` | Incorrect Vacation/Illness title / excessive hours (>10h) |

## Tech Stack

- **.NET 10** — isolated worker model
- **Azure Functions v4** — HTTP trigger
- **DocumentFormat.OpenXml** — workbook manipulation
- **OpenTelemetry + Azure Monitor** — observability

## Prerequisites

- [.NET 10 SDK](https://dotnet.microsoft.com/download)
- [Azure Functions Core Tools](https://learn.microsoft.com/azure/azure-functions/functions-run-local)
- [Azurite](https://learn.microsoft.com/azure/storage/common/storage-use-azurite) (for local storage emulation)

## Local Setup

```bash
# Clone the repo
git clone <repo-url>
cd "Azure Functions"

# Restore dependencies
dotnet restore

# Start Azurite (requires Azurite installed)
azurite

# In a separate terminal, start the function
func start
```

Or use the helper script:

```powershell
# Run with formatting only
.\run.ps1 -FormattingOnly

# Run with full validation
.\run.ps1
```

## Usage

Send an HTTP POST request with an `.xlsx` file to the function endpoint:

```powershell
$url = "http://localhost:7071/api/ProcessDocument"
$form = @{
    file = Get-Item -Path "path\to\file.xlsx"
    isFormattingOnly = "false"
}
Invoke-RestMethod -Uri $url -Method Post -Form $form -OutFile "result.xlsx"
```

### Parameters

| Field             | Type    | Required | Description                          |
|-------------------|---------|----------|--------------------------------------|
| `file`            | File    | Yes      | The `.xlsx` file to process          |
| `isFormattingOnly`| Boolean | No       | Skip validation, format only (default: `false`) |

### Response

Returns the processed `.xlsx` file with a `Content-Disposition` header containing the original filename.

## Configuration

| Setting                                   | Purpose                          |
|-------------------------------------------|----------------------------------|
| `AzureWebJobsStorage`                     | Storage connection (use Azurite for local) |
| `APPLICATIONINSIGHTS_CONNECTION_STRING`   | Optional — enables OpenTelemetry |

## Project Structure

```
Azure Functions/
├── Models/
│   ├── Constants/Excel/   # Sheet/column/config constants
│   └── Excel/             # Data transfer objects (ReportLogOutData, VacationData)
├── Services/
│   ├── Excel/             # Low-level OpenXML operations (read, validate, update)
│   └── ExcelService.cs    # Static helpers (cell parsing, header detection)
├── Workflows/
│   ├── Excel/             # Business logic (VacationIllness, OtherSheets processors)
│   └── ExcelWorkflow.cs   # Orchestrator — ties everything together
├── ProcessDocument.cs     # HTTP trigger entry point
├── Program.cs             # Host builder / DI registration
├── host.json              # Azure Functions host config
├── local.settings.json    # Local env settings (gitignored)
└── run.ps1                # Helper script to send a file to the function
```

## Notes

- The class `VacationIllnesProcess` contains a known typo (kept for consistency)

## License

[MIT](LICENSE.txt) — Copyright (c) 2026 Maksym Kurkudiuk
