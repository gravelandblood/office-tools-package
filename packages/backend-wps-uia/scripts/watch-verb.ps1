param(
  [string]$VerbPattern = "",
  [string]$InputPdf = ""
)

$ErrorActionPreference = "Stop"

function New-Text {
  param([int[]]$CodePoints, [string]$Suffix = "", [string]$Prefix = "")
  return $Prefix + (($CodePoints | ForEach-Object { [char]$_ }) -join "") + $Suffix
}

if (-not $VerbPattern) {
  $VerbPattern = New-Text -CodePoints @(0x8F6C, 0x4E3A) -Suffix "Word"
}

if (-not $InputPdf) {
  $InputPdf = $env:OFFICE_TOOLS_SAMPLE_PDF
}

if (-not $InputPdf) {
  throw "Pass -InputPdf or set OFFICE_TOOLS_SAMPLE_PDF to a local PDF path."
}

$pdf = (Resolve-Path -LiteralPath $InputPdf).Path
$before = @{}
Get-CimInstance Win32_Process | ForEach-Object { $before[[int]$_.ProcessId] = $true }

$shell = New-Object -ComObject Shell.Application
$folder = $shell.Namespace((Split-Path $pdf -Parent))
$item = $folder.ParseName((Split-Path $pdf -Leaf))
$verbs = @($item.Verbs())
$matches = @($verbs | Where-Object { ($_.Name -replace "&", "") -like "*$VerbPattern*" })

if ($matches.Count -eq 0) {
  throw "No verb matching $VerbPattern"
}

$verb = $matches[-1]
Write-Host "Invoking verb: $($verb.Name)"
$verb.DoIt()

$seen = @{}
$deadline = (Get-Date).AddSeconds(12)
while ((Get-Date) -lt $deadline) {
  Get-CimInstance Win32_Process |
    Where-Object { -not $before.ContainsKey([int]$_.ProcessId) } |
    ForEach-Object {
      if (-not $seen.ContainsKey([int]$_.ProcessId)) {
        $seen[[int]$_.ProcessId] = $true
        [pscustomobject]@{
          Time = (Get-Date).ToString("HH:mm:ss.fff")
          Pid = $_.ProcessId
          Parent = $_.ParentProcessId
          Name = $_.Name
          CommandLine = $_.CommandLine
        } | ConvertTo-Json -Compress
      }
    }
  Start-Sleep -Milliseconds 300
}
