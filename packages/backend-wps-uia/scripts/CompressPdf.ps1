param(
  [Parameter(Mandatory = $true)]
  [string]$InputPdf,

  [Parameter(Mandatory = $true)]
  [string]$OutputPath,

  [ValidateSet("high", "standard", "medium", "low")]
  [string]$Level = "standard",

  [int]$TimeoutSeconds = 180,

  [string]$WpsExe = "",

  [string]$AppFramework = "",

  [switch]$Overwrite,

  [switch]$NoCleanup,

  [int]$CleanupSeconds = 20
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

function ConvertTo-JsonLine($Object) {
  $Object | ConvertTo-Json -Compress -Depth 6
}

function New-Text {
  param([int[]]$CodePoints, [string]$Suffix = "", [string]$Prefix = "")
  return $Prefix + (($CodePoints | ForEach-Object { [char]$_ }) -join "") + $Suffix
}

function ConvertTo-CommandLineArgument {
  param([Parameter(Mandatory = $true)][string]$Value)

  if ($Value -notmatch '[\s"]') {
    return $Value
  }

  return '"' + ($Value -replace '"', '\"') + '"'
}

function Find-WpsExe {
  $running = Get-Process wps -ErrorAction SilentlyContinue |
    Where-Object { $_.Path } |
    Select-Object -First 1 -ExpandProperty Path
  if ($running) {
    return $running
  }

  $root = Join-Path $env:LOCALAPPDATA "Kingsoft\WPS Office"
  if (Test-Path -LiteralPath $root) {
    $candidate = Get-ChildItem -LiteralPath $root -Filter wps.exe -Recurse -ErrorAction SilentlyContinue |
      Where-Object { $_.FullName -like "*\office6\wps.exe" } |
      Sort-Object LastWriteTime -Descending |
      Select-Object -First 1 -ExpandProperty FullName
    if ($candidate) {
      return $candidate
    }
  }

  throw "Unable to locate wps.exe. Pass -WpsExe explicitly."
}

function Find-AppFramework {
  param([Parameter(Mandatory = $true)][string]$ResolvedWpsExe)

  $framework = Join-Path (Split-Path $ResolvedWpsExe -Parent) "addons\kappessframework\kappessframework.dll"
  if (Test-Path -LiteralPath $framework) {
    return $framework
  }

  throw "Unable to locate kappessframework.dll. Pass -AppFramework explicitly."
}

function Get-ProcessSnapshot {
  $snapshot = @{}
  Get-Process wps,et,wpp,wpspdf,wpscloudsvr,promecefpluginhost -ErrorAction SilentlyContinue | ForEach-Object {
    $snapshot[[int]$_.Id] = $true
  }
  return ,$snapshot
}

function Stop-ProcessIfRunning {
  param([Parameter(Mandatory = $true)][int]$ProcessId)

  $proc = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
  if ($proc) {
    Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue
    return $true
  }

  return $false
}

function Close-NewWpsArtifacts {
  param(
    [hashtable]$BeforeProcesses,
    [int]$TimeoutSeconds = 20
  )

  $events = @()
  $deadline = (Get-Date).AddSeconds($TimeoutSeconds)

  while ((Get-Date) -lt $deadline) {
    $closedAny = $false
    foreach ($proc in @(Get-Process wps,et,wpp,wpspdf,promecefpluginhost -ErrorAction SilentlyContinue)) {
      $isNew = -not $BeforeProcesses.ContainsKey([int]$proc.Id)
      if ($isNew) {
        if (Stop-ProcessIfRunning -ProcessId $proc.Id) {
          $closedAny = $true
          $events += [pscustomobject]@{
            action = "stop-new-process"
            processId = $proc.Id
            processName = $proc.ProcessName
            title = [string]$proc.MainWindowTitle
            newProcess = $true
          }
        }
      }
    }

    $remaining = @(Get-Process wps,et,wpp,wpspdf,promecefpluginhost -ErrorAction SilentlyContinue | Where-Object {
      -not $BeforeProcesses.ContainsKey([int]$_.Id)
    })

    if ($remaining.Count -eq 0) {
      break
    }

    if (-not $closedAny) {
      Start-Sleep -Milliseconds 700
    }
  }

  return @($events)
}

function Close-UiaCompressWindows {
  param(
    [hashtable]$BeforeProcesses,
    [int]$TimeoutSeconds = 10
  )

  $events = @()
  $compressTitle = "PDF" + (New-Text -CodePoints @(0x538B, 0x7F29))

  Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes -ErrorAction SilentlyContinue
  $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
  while ((Get-Date) -lt $deadline) {
    $closedAny = $false
    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $windows = $root.FindAll([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition)

    foreach ($win in $windows) {
      $name = ""
      $windowProcessId = 0
      try { $name = $win.Current.Name } catch {}
      try { $windowProcessId = $win.Current.ProcessId } catch {}

      if ($name -ne $compressTitle) {
        continue
      }

      $pattern = $null
      if ($win.TryGetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern, [ref]$pattern)) {
        try {
          $pattern.Close()
          Start-Sleep -Milliseconds 800
          $closedAny = $true
          $events += [pscustomobject]@{
            action = "uia-close-window"
            processId = $windowProcessId
            title = $name
          }
        } catch {}
      }

      $proc = Get-Process -Id $windowProcessId -ErrorAction SilentlyContinue
      $isNew = $windowProcessId -and (-not $BeforeProcesses.ContainsKey([int]$windowProcessId))
      if ($isNew -and $proc -and $proc.MainWindowTitle -eq $compressTitle) {
        if (Stop-ProcessIfRunning -ProcessId $windowProcessId) {
          $closedAny = $true
          $events += [pscustomobject]@{
            action = "stop-new-process"
            processId = $windowProcessId
            processName = $proc.ProcessName
            title = [string]$proc.MainWindowTitle
            newProcess = $true
          }
        }
      }
    }

    if (-not $closedAny) {
      break
    }

    Start-Sleep -Milliseconds 700
  }

  return @($events)
}

function Start-WpsPdfCompress {
  param([Parameter(Mandatory = $true)][string]$Path)

  $resolvedWpsExe = $WpsExe
  if (-not $resolvedWpsExe) {
    $resolvedWpsExe = Find-WpsExe
  }
  $resolvedWpsExe = (Resolve-Path -LiteralPath $resolvedWpsExe).Path

  $resolvedFramework = $AppFramework
  if (-not $resolvedFramework) {
    $resolvedFramework = Find-AppFramework -ResolvedWpsExe $resolvedWpsExe
  }
  $resolvedFramework = (Resolve-Path -LiteralPath $resolvedFramework).Path

  $args = @(
    "Run",
    "/InstanceId=batchcompress",
    $resolvedFramework,
    "/appId=batchcompress",
    "/appname=PDF Compress",
    "/size=960&670",
    "/src=office_tools_package",
    "/switchskin=0",
    "/file=$Path"
  )

  $argumentLine = ($args | ForEach-Object { ConvertTo-CommandLineArgument -Value $_ }) -join " "
  Start-Process -FilePath $resolvedWpsExe -ArgumentList $argumentLine | Out-Null

  return [pscustomobject]@{
    wpsExe = $resolvedWpsExe
    appFramework = $resolvedFramework
    arguments = $args
    argumentLine = $argumentLine
  }
}

function Invoke-CompressStartButton {
  param([string]$Level, [int]$TimeoutSeconds = 45)

  $code = @'
using System;
using System.Windows.Automation;
using System.Threading;

public class WpsPdfCompressAutomation {
  static string FromCodes(int[] codes) {
    char[] chars = new char[codes.Length];
    for (int i = 0; i < codes.Length; i++) chars[i] = (char)codes[i];
    return new string(chars);
  }

  static AutomationElement FindByName(AutomationElement root, string name) {
    return root.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, name));
  }

  static bool TryInvoke(AutomationElement element) {
    if (element == null) return false;
    object pattern;
    if (element.TryGetCurrentPattern(InvokePattern.Pattern, out pattern)) {
      ((InvokePattern)pattern).Invoke();
      return true;
    }
    if (element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out pattern)) {
      ((SelectionItemPattern)pattern).Select();
      return true;
    }
    return false;
  }

  public static int Invoke(string levelName, int timeoutSeconds) {
    string windowTitle = "PDF" + FromCodes(new int[] { 0x538B, 0x7F29 });
    string startText = FromCodes(new int[] { 0x5F00, 0x59CB, 0x538B, 0x7F29 });
    DateTime deadline = DateTime.Now.AddSeconds(timeoutSeconds);
    while (DateTime.Now < deadline) {
      var root = AutomationElement.RootElement;
      var win = root.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, windowTitle));
      if (win != null) {
        if (!String.IsNullOrWhiteSpace(levelName)) {
          var level = FindByName(win, levelName);
          TryInvoke(level);
          Thread.Sleep(300);
        }

        var start = FindByName(win, startText);
        if (TryInvoke(start)) return 0;
        return 3;
      }
      Thread.Sleep(500);
    }
    return 2;
  }
}
'@

  Add-Type -TypeDefinition $code -ReferencedAssemblies UIAutomationClient,UIAutomationTypes -ErrorAction SilentlyContinue

  $levelName = switch ($Level) {
    "high" { New-Text -CodePoints @(0x9AD8, 0x54C1, 0x8D28) }
    "standard" { New-Text -CodePoints @(0x6807, 0x51C6, 0x54C1, 0x8D28) }
    "medium" { New-Text -CodePoints @(0x4E2D, 0x7B49, 0x54C1, 0x8D28) }
    "low" { New-Text -CodePoints @(0x4F4E, 0x54C1, 0x8D28) }
    default { New-Text -CodePoints @(0x6807, 0x51C6, 0x54C1, 0x8D28) }
  }

  $result = [WpsPdfCompressAutomation]::Invoke($levelName, $TimeoutSeconds)
  if ($result -ne 0) {
    throw "Unable to invoke WPS PDF compress start button. Code: $result"
  }
}

function Wait-FileReady {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [int]$TimeoutSeconds = 180,
    [datetime]$NotBefore = [datetime]::MinValue
  )

  $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
  $lastLength = -1
  $stableCount = 0

  while ((Get-Date) -lt $deadline) {
    if (Test-Path -LiteralPath $Path) {
      $item = Get-Item -LiteralPath $Path
      if ($item.LastWriteTime -lt $NotBefore) {
        Start-Sleep -Seconds 1
        continue
      }

      if ($item.Length -eq $lastLength -and $item.Length -gt 0) {
        $stableCount += 1
      } else {
        $stableCount = 0
        $lastLength = $item.Length
      }

      if ($stableCount -ge 2) {
        return $item
      }
    }
    Start-Sleep -Seconds 1
  }

  throw "Timed out waiting for output file: $Path"
}

$inputPath = (Resolve-Path -LiteralPath $InputPdf).Path
if ([IO.Path]::GetExtension($inputPath).ToLowerInvariant() -ne ".pdf") {
  throw "Input must be a PDF: $inputPath"
}

$finalOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
if ([IO.Path]::GetExtension($finalOutputPath).ToLowerInvariant() -ne ".pdf") {
  throw "OutputPath must end with .pdf: $finalOutputPath"
}

if ((Test-Path -LiteralPath $finalOutputPath) -and -not $Overwrite) {
  throw "Output already exists. Pass -Overwrite to replace it: $finalOutputPath"
}

$compressedPrefix = New-Text -CodePoints @(0xFF08, 0x5DF2, 0x538B, 0x7F29, 0xFF09)
$expectedOutputPath = Join-Path (Split-Path $inputPath -Parent) ($compressedPrefix + [IO.Path]::GetFileName($inputPath))
if ((Test-Path -LiteralPath $expectedOutputPath) -and -not $Overwrite) {
  throw "Default WPS output already exists. Pass -Overwrite or remove it first: $expectedOutputPath"
}
if ((Test-Path -LiteralPath $expectedOutputPath) -and $Overwrite) {
  Remove-Item -LiteralPath $expectedOutputPath -Force
}

$beforeProcesses = Get-ProcessSnapshot
$startedAt = Get-Date
$cleanupEvents = @()
$runner = $null
$output = $null
try {
  $runner = Start-WpsPdfCompress -Path $inputPath
  Invoke-CompressStartButton -Level $Level
  $output = Wait-FileReady -Path $expectedOutputPath -TimeoutSeconds $TimeoutSeconds -NotBefore $startedAt
} finally {
  if (-not $NoCleanup) {
    $cleanupEvents = @(
      Close-UiaCompressWindows -BeforeProcesses $beforeProcesses -TimeoutSeconds ([Math]::Min($CleanupSeconds, 10))
      Close-NewWpsArtifacts -BeforeProcesses $beforeProcesses -TimeoutSeconds $CleanupSeconds
    )
  }
}

$parent = Split-Path $finalOutputPath -Parent
if ($parent -and -not (Test-Path -LiteralPath $parent)) {
  New-Item -ItemType Directory -Path $parent | Out-Null
}
Move-Item -LiteralPath $output.FullName -Destination $finalOutputPath -Force:$Overwrite
$output = Get-Item -LiteralPath $finalOutputPath
$inputItem = Get-Item -LiteralPath $inputPath

ConvertTo-JsonLine ([pscustomobject]@{
  ok = $true
  backend = "wps-uia"
  command = "pdf.compress"
  input = $inputPath
  output = $output.FullName
  level = $Level
  inputLength = $inputItem.Length
  length = $output.Length
  savedBytes = $inputItem.Length - $output.Length
  lastWriteTime = $output.LastWriteTime.ToString("o")
  runner = $runner
  cleanup = $cleanupEvents
})
