<#
.SYNOPSIS
    Sends an .xlsx file to the ProcessDocument Azure Function and saves the result.
.PARAMETER FunctionUrl
    URL of the Azure Function (default: localhost for local dev).
.PARAMETER InputFile
    Path to a specific .xlsx file. If omitted, picks the first file from 'Input' folder.
.PARAMETER InputFolder
    Folder to pick a file from when -InputFile is not specified (default: ./Input).
.PARAMETER OutputFolder
    Folder to save the result (default: ./Outputs).
.PARAMETER FormattingOnly
    If set, only apply formatting without data validation.
.PARAMETER FunctionKey
    Optional function key (appended as ?code=...). Prefer environment variable FUNCTIONS_KEY.
#>

param(
    [string]$FunctionUrl = "http://localhost:7071/api/ProcessDocument",
    [string]$InputFile,
    [string]$InputFolder = (Join-Path -Path $PSScriptRoot -ChildPath "Input"),
    [string]$OutputFolder = (Join-Path -Path $PSScriptRoot -ChildPath "Outputs"),
    [switch]$FormattingOnly,
    [string]$FunctionKey = $env:FUNCTIONS_KEY
)

# --- Resolve input file ---
if (-not $InputFile) {
    if (-not (Test-Path -Path $InputFolder)) {
        throw "Input folder not found: $InputFolder"
    }
    $InputFile = Get-ChildItem -Path $InputFolder -File | Select-Object -First 1 -ExpandProperty FullName
    if (-not $InputFile) {
        throw "No files found in $InputFolder"
    }
}

if (-not (Test-Path -Path $InputFile)) {
    throw "File not found: $InputFile"
}

Write-Host "Processing: $InputFile"

# --- Ensure output folder exists ---
$null = New-Item -Path $OutputFolder -ItemType Directory -Force

# --- Build URL with optional key ---
$url = $FunctionUrl
if ($FunctionKey) {
    $separator = if ($url.Contains('?')) { '&' } else { '?' }
    $url = "$url${separator}code=$FunctionKey"
}

# --- Send request ---
$fileStream = [System.IO.File]::OpenRead($InputFile)
$client = [System.Net.Http.HttpClient]::new()
try {
    $content = [System.Net.Http.MultipartFormDataContent]::new()
    $content.Add([System.Net.Http.StreamContent]::new($fileStream), "file", [System.IO.Path]::GetFileName($InputFile))
    $content.Add([System.Net.Http.StringContent]::new($FormattingOnly.IsPresent.ToString().ToLower()), "isFormattingOnly")

    $response = $client.PostAsync($url, $content).GetAwaiter().GetResult()
    $response.EnsureSuccessStatusCode()

    $responseBytes = $response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult()
    $cdHeader = $response.Content.Headers.ContentDisposition.ToString()
    $originalName = if ($cdHeader -match 'filename="(.+?)"') { $matches[1] } else { (Get-Item $InputFile).Name }
    $resultName = $originalName -replace '_result', ''

    $outputPath = Join-Path -Path $OutputFolder -ChildPath $resultName
    [System.IO.File]::WriteAllBytes($outputPath, $responseBytes)

    Write-Host "Result saved to: $outputPath"
}
catch {
    Write-Error "Request failed: $($_.Exception.Message)"
    exit 1
}
finally {
    $fileStream.Dispose()
    $client.Dispose()
    if ($response) { $response.Dispose() }
}
