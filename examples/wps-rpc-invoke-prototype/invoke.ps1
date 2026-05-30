$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$PdfPath = $env:OFFICE_TOOLS_SAMPLE_PDF
if (-not $PdfPath) {
  throw "Set OFFICE_TOOLS_SAMPLE_PDF to a local PDF path before running this prototype."
}

$WpsExe = (Get-Process wps -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty Path)
if (-not $WpsExe) {
  $WpsExe = Join-Path $env:LOCALAPPDATA "Kingsoft\WPS Office\office6\wps.exe"
}

$Framework = Join-Path (Split-Path $WpsExe -Parent) "addons\kappessframework\kappessframework.dll"
$Fso = New-Object -ComObject Scripting.FileSystemObject

$WpsExeShort = $Fso.GetFile($WpsExe).ShortPath
$FrameworkShort = $Fso.GetFile($Framework).ShortPath
$PdfShort = $Fso.GetFile($PdfPath).ShortPath

$Html = Get-Content (Join-Path $Root "invoke.html") -Raw
$Html = $Html -replace "window.codexBridge.ping\(\);", @"
window.codexBridgeContext = {
  wpsExe: "$($WpsExeShort -replace '\\', '\\')",
  appFramework: "$($FrameworkShort -replace '\\', '\\')",
  pdfPath: "$($PdfShort -replace '\\', '\\')"
};
window.codexBridge.ping();
"@

$Generated = Join-Path $Root "invoke.generated.html"
Set-Content -Path $Generated -Value $Html -Encoding UTF8

Write-Host "Generated $Generated"
Write-Host "Open http://127.0.0.1:48990/invoke.generated.html"
