$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$PdfPath = Get-ChildItem -LiteralPath "C:\Code\test\pdfreadtest" -Filter "*.pdf" |
  Where-Object { $_.Name -like "*2023.10.26*" } |
  Select-Object -First 1 -ExpandProperty FullName
if (-not $PdfPath) {
  throw "Test PDF was not found under C:\Code\test\pdfreadtest"
}
$WpsExe = (Get-Process wps -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty Path)
if (-not $WpsExe) {
  $WpsExe = "C:\Users\十九\AppData\Local\Kingsoft\WPS Office\12.1.0.26375\office6\wps.exe"
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
