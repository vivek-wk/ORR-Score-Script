# ORR Score and Checkpoint Report

PowerShell utility for retrieving ORR evaluations and checkpoint details from the Pegasus production ORR API.

## Prerequisites

- Windows PowerShell 5.1 or PowerShell 7+
- Azure CLI (`az`)
- PowerShell `Az.Accounts` module
- Access to the Pegasus production Azure account/subscription
- Access to the ORR portal
- Microsoft Edge or Google Chrome for PDF export
- Network access to the Pegasus ORR API

Install the Azure CLI from `https://learn.microsoft.com/cli/azure/install-azure-cli`.

Install the PowerShell module:

```powershell
Install-Module Az.Accounts -Scope CurrentUser
```

## Sign In

Use the Pegasus production account. Do not put passwords or tokens in this README or in command-line arguments.

### Azure CLI

```powershell
az login
az account list --output table
az account set --subscription "<Pegasus production subscription name or ID>"
az account show --output table
```

Confirm that the selected tenant, subscription, and signed-in user are the expected Pegasus production values.

### PowerShell Az

The script uses `Get-AzAccessToken`, so also establish the PowerShell Az context if required:

```powershell
Connect-AzAccount
Get-AzContext
Set-AzContext -Subscription "<Pegasus production subscription name or ID>"
```

For a specific tenant:

```powershell
Connect-AzAccount -Tenant "<Pegasus production tenant ID or domain>"
```

### ORR portal browser session

Sign in to the ORR portal in Edge or Chrome using the same Pegasus production account and keep that browser session active:

```text
https://orr.pegasus.wkcloud.io
```

The browser session opens ORR links from the report. API authentication comes from the PowerShell Az context.

## Location

Keep the files together in:

```text
C:\orr
```

Main files:

- `Get-OrrScore.ps1` - report script
- `Start-OrrEvaluation.ps1` - authenticated evaluation trigger client
- `Run-Orr.ps1` - single interactive entry point for evaluations or score review
- `Run-OrrScore.cmd` - Windows launcher
- `evaluation.txt` - reference/sample output
- `.gitignore` - ignores generated reports, ZIP packages, and backup scripts

Generated HTML and PDF files are saved under:

```text
C:\orr\reports
```

The reports folder is created automatically when an export is requested.

## Run Interactively

From PowerShell:

```powershell
Set-Location C:\orr
.\Get-OrrScore.ps1
```

Or double-click:

```text
C:\orr\Run-OrrScore.cmd
```

The launcher first asks whether to run evaluations or review recent scores.
Reviewing delegates to `Get-OrrScore.ps1`, including its existing output,
environment, division, lifecycle, and result menus. Running evaluations uses
only the BIT ID choice, then queues jobs through
`POST /api/v4.0/orr/jobs/{BIT_ID}`. Runs ask for a department / division,
one or more environments, and whether to include only IPM-managed and only
active applications. The defaults are Production, Yes, and Yes. One job is
queued for each application/environment combination. The API returns a status
URL for each job. Before any evaluation starts, the launcher lists all
selected application/environment combinations and asks for confirmation.

Interactive menus support arrow-key navigation. Use `Space` to select or deselect environments, `Enter` to accept, and `Esc` to restore the default.

The default output mode is Table. Available modes are Table, Detailed, and Summary.

## Run with Parameters

```powershell
.\Get-OrrScore.ps1 -BitId 041800002285 -OutputMode Detailed
```

```powershell
.\Get-OrrScore.ps1 -BitId 041800002285 -OutputMode Table
```

```powershell
.\Get-OrrScore.ps1 -AllMatching -Environment Production -Division DXG -OutputMode Table
```

```powershell
.\Get-OrrScore.ps1 -AllMatching -Environment "Production,Staging" -OutputMode Summary
```

For non-failing checkpoints:

```powershell
.\Get-OrrScore.ps1 -BitId 041800002285 -Failing:$false -OutputMode Table
```

## Trigger an Evaluation

`Get-OrrScore.ps1` only reads evaluations. The unified `Run-Orr.ps1` launcher
uses the confirmed ORR job endpoint and sends the selected environment name:

```text
POST /api/v4.0/orr/jobs/{BIT_ID}
Body: { "environments": ["Production"] }
```

## Export HTML or PDF

The script asks whether to export and then lets you select `HTML page` or `PDF` with arrow keys.

Select the export format directly:

```powershell
.\Get-OrrScore.ps1 -BitId 041800002285 -OutputMode Detailed -ExportFormat HTML
.\Get-OrrScore.ps1 -BitId 041800002285 -OutputMode Detailed -ExportFormat PDF
```

Specify a custom output path:

```powershell
.\Get-OrrScore.ps1 -BitId 041800002285 -OutputMode Summary -ExportFormat PDF -PdfPath "C:\orr\reports\my-report.pdf"
```

Table mode exports a table. Detailed and Summary exports contain document-style content. Detailed exports include the complete checkpoint information shown in the CLI.

## Filters

- `-OutputMode Table|Detailed|Summary`
- `-BitId <letters and numbers>`
- `-AllMatching`
- `-Environment <environment>`
- `-Division <division>`
- `-IpmOnly $true|$false`
- `-EvaluatedOnly $true|$false`
- `-ActiveOnly $true|$false`
- `-Live $true|$false`
- `-Preview $true|$false`
- `-Exception $true|$false`
- `-Maintenance $true|$false`
- `-Failing $true|$false`
- `-ExportFormat HTML|PDF`
- `-PdfPath <path>`

## Overview

The script:

- Retrieves the latest ORR evaluation for each application.
- Supports one BIT ID or multiple applications.
- Filters by environment, division, lifecycle, and result.
- Groups multiple checkpoints under one application.
- Shows evaluation timestamps in UTC.
- Adds checkpoint links using evaluation ID, checkpoint ID, and checkpoint version.
- Exports HTML or PDF reports to `C:\orr\reports`.
- Applies lifecycle colors for Live, Preview, Exception, and Maintenance.
- Displays the actual BIT ID in remarks.

## Troubleshooting

### Script is not digitally signed / Execution Policy error

If PowerShell blocks script execution with `PSSecurityException` or `is not digitally signed`:

- **Option 1 (Current Session Only - Recommended):**
  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
  ```

- **Option 2 (Current User):**
  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```

- **Option 3 (Run via CMD Launcher):**
  Run `.\Run-OrrScore.cmd` or launch PowerShell with bypass:
  ```powershell
  powershell -ExecutionPolicy Bypass -File .\Run-Orr.ps1
  ```

- **Option 4 (Unblock Files):**
  ```powershell
  Unblock-File .\*.ps1
  ```

### `Az.Accounts is not installed`

```powershell
Install-Module Az.Accounts -Scope CurrentUser
Import-Module Az.Accounts
```

### Access token or permission error

```powershell
az account show
Get-AzContext
```

Sign in again with the Pegasus production account if either context is incorrect.

### PDF export fails

Confirm Microsoft Edge or Google Chrome is installed. PDF export uses the browser in headless print-to-PDF mode.

### GitHub push permission denied

The GitHub account used for authentication must have write access to the repository. A local commit does not guarantee permission to push.
