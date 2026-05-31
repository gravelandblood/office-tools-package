param(
  [Parameter(Mandatory = $true)]
  [string]$InputPdf,

  [string]$OutputPath = "",

  [int]$TimeoutSeconds = 180,

  [switch]$NoClick,

  [ValidateSet("shell", "native", "cloud")]
  [string]$LaunchMode = "shell",

  [string]$WpsExe = "",

  [string]$CloudExe = "",

  [string]$AppFramework = "",

  [string]$InstanceId = "kpdf2wordv2",

  [string]$AppId = "kpdf2wordv2",

  [string]$AppName = "WPS PDF Convert",

  [string]$WindowSize = "960&670",

  [string]$Src = "office_tools_package",

  [string]$CloudSrc = "office_tools_package",

  [string]$CloudAppParams = "",

  [int]$SwitchSkin = 0,

  [string]$Action = "ConvertToWord",

  [string[]]$RunnerParam = @(),

  [string[]]$RunnerArg = @(),

  [string[]]$CloudArg = @(),

  [string[]]$PreferredVerb = @(),

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

function Invoke-ShellVerb {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [string[]]$PreferredVerbs = @()
  )

  if ($PreferredVerbs.Count -eq 0) {
    $PreferredVerbs = @(
      (New-Text -CodePoints @(0x8F6C, 0x4E3A) -Suffix "Word"),
      (New-Text -CodePoints @(0x8F6C, 0x6362, 0x4E3A) -Suffix "Word"),
      (New-Text -CodePoints @(0x8F6C) -Prefix "PDF" -Suffix "office")
    )
  }

  $resolved = (Resolve-Path -LiteralPath $Path).Path
  $shell = New-Object -ComObject Shell.Application
  $folder = $shell.Namespace((Split-Path $resolved -Parent))
  if (-not $folder) {
    throw "Unable to open shell namespace for $resolved"
  }

  $item = $folder.ParseName((Split-Path $resolved -Leaf))
  if (-not $item) {
    throw "Unable to parse shell item for $resolved"
  }

  $verbs = @($item.Verbs())
  foreach ($preferred in $PreferredVerbs) {
    $match = @($verbs | Where-Object { (($_.Name -replace "&", "").Trim()) -eq $preferred }) | Select-Object -First 1
    if ($match) {
      $match.DoIt()
      return $preferred
    }
  }

  $available = ($verbs | ForEach-Object { ($_.Name -replace "&", "").Trim() } | Where-Object { $_ }) -join ", "
  throw "No supported WPS PDF conversion shell verb found. Available verbs: $available"
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

function Find-CloudExe {
  param([string]$ResolvedWpsExe = "")

  if ($ResolvedWpsExe) {
    $candidate = Join-Path (Split-Path $ResolvedWpsExe -Parent) "wpscloudsvr.exe"
    if (Test-Path -LiteralPath $candidate) {
      return $candidate
    }
  }

  $running = Get-Process wpscloudsvr -ErrorAction SilentlyContinue |
    Where-Object { $_.Path } |
    Select-Object -First 1 -ExpandProperty Path
  if ($running) {
    return $running
  }

  $root = Join-Path $env:LOCALAPPDATA "Kingsoft\WPS Office"
  if (Test-Path -LiteralPath $root) {
    $candidate = Get-ChildItem -LiteralPath $root -Filter wpscloudsvr.exe -Recurse -ErrorAction SilentlyContinue |
      Sort-Object LastWriteTime -Descending |
      Select-Object -First 1 -ExpandProperty FullName
    if ($candidate) {
      return $candidate
    }
  }

  throw "Unable to locate wpscloudsvr.exe. Pass -CloudExe explicitly."
}

function ConvertTo-RunnerArgument {
  param([Parameter(Mandatory = $true)][string]$Param)

  if ($Param.StartsWith("/")) {
    return $Param
  }

  $parts = $Param.Split("=", 2)
  if ($parts.Count -ne 2 -or -not $parts[0]) {
    throw "RunnerParam must be key=value or /key=value: $Param"
  }

  return "/" + $parts[0] + "=" + $parts[1]
}

function ConvertTo-CommandLineArgument {
  param([Parameter(Mandatory = $true)][string]$Value)

  if ($Value -notmatch '[\s"]') {
    return $Value
  }

  return '"' + ($Value -replace '"', '\"') + '"'
}

function Start-WpsPdfConverter {
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
    "/InstanceId=$InstanceId",
    $resolvedFramework,
    "/appId=$AppId",
    "/appname=$AppName",
    "/size=$WindowSize",
    "/src=$Src",
    "/switchskin=$SwitchSkin",
    "/action=$Action",
    "/file=$Path"
  )

  foreach ($param in $RunnerParam) {
    $args += ConvertTo-RunnerArgument -Param $param
  }

  foreach ($arg in $RunnerArg) {
    $args += $arg
  }

  $argumentLine = ($args | ForEach-Object { ConvertTo-CommandLineArgument -Value $_ }) -join " "
  Start-Process -FilePath $resolvedWpsExe -ArgumentList $argumentLine | Out-Null

  return [pscustomobject]@{
    wpsExe = $resolvedWpsExe
    appFramework = $resolvedFramework
    arguments = $args
    argumentLine = $argumentLine
  }
}

function Start-WpsCloudPdfConverter {
  param([Parameter(Mandatory = $true)][string]$Path)

  $resolvedWpsExe = $WpsExe
  if (-not $resolvedWpsExe) {
    $resolvedWpsExe = Find-WpsExe
  }
  $resolvedWpsExe = (Resolve-Path -LiteralPath $resolvedWpsExe).Path

  $resolvedCloudExe = $CloudExe
  if (-not $resolvedCloudExe) {
    $resolvedCloudExe = Find-CloudExe -ResolvedWpsExe $resolvedWpsExe
  }
  $resolvedCloudExe = (Resolve-Path -LiteralPath $resolvedCloudExe).Path

  $args = @(
    "Run",
    "/InstanceId:WpsCloudSvr",
    "/app_id=$AppId",
    "/src=$CloudSrc"
  )

  if ($CloudAppParams) {
    $args += "/app_params=$CloudAppParams"
  }

  foreach ($arg in $CloudArg) {
    $args += $arg
  }

  $argumentLine = ($args | ForEach-Object { ConvertTo-CommandLineArgument -Value $_ }) -join " "
  Start-Process -FilePath $resolvedCloudExe -ArgumentList $argumentLine | Out-Null

  return [pscustomobject]@{
    cloudExe = $resolvedCloudExe
    wpsExe = $resolvedWpsExe
    arguments = $args
    argumentLine = $argumentLine
    note = "cloud mode exposes the observed wpscloudsvr entrypoint; app_params must be supplied by the caller when needed"
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

function Get-ProcessSnapshot {
  $snapshot = @{}
  Get-Process wps,et,wpp,wpspdf,wpscloudsvr -ErrorAction SilentlyContinue | ForEach-Object {
    $snapshot[[int]$_.Id] = $true
  }
  return ,$snapshot
}

function Close-MainWindowByProcessId {
  param([Parameter(Mandatory = $true)][int]$ProcessId)

  $proc = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
  if (-not $proc) {
    return $false
  }

  $hadWindow = [bool]$proc.MainWindowHandle
  if ($hadWindow) {
    [void]$proc.CloseMainWindow()
    Start-Sleep -Milliseconds 800
  }

  $proc = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
  if ($proc -and $proc.MainWindowHandle) {
    Stop-Process -Id $ProcessId -Force
  }

  return $hadWindow
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

function Close-WpsArtifacts {
  param(
    [hashtable]$BeforeProcesses,
    [string]$InputPath,
    [string]$OutputPath,
    [int]$TimeoutSeconds = 20
  )

  $events = New-Object System.Collections.Generic.List[object]
  $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
  $inputName = [IO.Path]::GetFileNameWithoutExtension($InputPath)
  $outputName = [IO.Path]::GetFileNameWithoutExtension($OutputPath)

  while ((Get-Date) -lt $deadline) {
    $closedAny = $false
    $processes = @(Get-Process wps,et,wpp,wpspdf -ErrorAction SilentlyContinue)

    foreach ($proc in $processes) {
      $isNew = -not $BeforeProcesses.ContainsKey([int]$proc.Id)
      $title = $proc.MainWindowTitle
      $titleMatches = $title -and (
        $title -like "*$inputName*" -or
        $title -like "*$outputName*" -or
        $title -like "*WPS PDF*" -or
        $title -eq "工作簿1 - WPS Office" -or
        $title -eq "文档1 - WPS Office"
      )

      if ($isNew) {
        $stopped = Stop-ProcessIfRunning -ProcessId $proc.Id
        if ($stopped) {
          $closedAny = $true
          $events.Add([pscustomobject]@{
            action = "stop-new-process"
            processId = $proc.Id
            processName = $proc.ProcessName
            title = $title
            newProcess = $true
          })
        }
      } elseif ($titleMatches) {
        $closed = Close-MainWindowByProcessId -ProcessId $proc.Id
        if ($closed) {
          $closedAny = $true
          $events.Add([pscustomobject]@{
            action = "close-window"
            processId = $proc.Id
            processName = $proc.ProcessName
            title = $title
            newProcess = $isNew
          })
        }
      }
    }

    if (-not $closedAny) {
      Start-Sleep -Milliseconds 700
    }

    $remaining = @(Get-Process wps,et,wpp,wpspdf -ErrorAction SilentlyContinue | Where-Object {
      $isNew = -not $BeforeProcesses.ContainsKey([int]$_.Id)
      $title = $_.MainWindowTitle
      $titleMatches = $title -and (
        $title -like "*$inputName*" -or
        $title -like "*$outputName*" -or
        $title -like "*WPS PDF*" -or
        $title -eq "工作簿1 - WPS Office" -or
        $title -eq "文档1 - WPS Office"
      )
      $titleMatches
    })

    if ($remaining.Count -eq 0) {
      break
    }
  }

  return ,@($events)
}

function Close-UiaWindowArtifacts {
  param(
    [string]$InputPath,
    [string]$OutputPath,
    [int]$TimeoutSeconds = 10
  )

  $events = New-Object System.Collections.Generic.List[object]
  $inputName = [IO.Path]::GetFileNameWithoutExtension($InputPath)
  $outputName = [IO.Path]::GetFileNameWithoutExtension($OutputPath)
  $convertTitle = "WPS PDF" + (New-Text -CodePoints @(0x8F6C, 0x6362))
  $doneTitle = New-Text -CodePoints @(0x64CD, 0x4F5C, 0x5B8C, 0x6210)

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

      if (-not $name) {
        continue
      }

      $matches = (
        $name -like "*$inputName*" -or
        $name -like "*$outputName*" -or
        $name -like "*$convertTitle*" -or
        $name -like "*$doneTitle*"
      )

      if ($matches) {
        $pattern = $null
        if ($win.TryGetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern, [ref]$pattern)) {
          try {
            $pattern.Close()
            Start-Sleep -Milliseconds 800
            $proc = Get-Process -Id $windowProcessId -ErrorAction SilentlyContinue
            if ($proc -and $proc.MainWindowTitle -and (
              $proc.MainWindowTitle -like "*$inputName*" -or
              $proc.MainWindowTitle -like "*$outputName*" -or
              $proc.MainWindowTitle -like "*$convertTitle*"
            )) {
              Stop-Process -Id $windowProcessId -Force
            }
            $closedAny = $true
            $events.Add([pscustomobject]@{
              action = "uia-close-window"
              processId = $windowProcessId
              title = $name
            })
          } catch {}
        }
      }
    }

    if (-not $closedAny) {
      break
    }

    Start-Sleep -Milliseconds 700
  }

  return ,@($events)
}

function Test-WpsArtifactTitle {
  param(
    [string]$Title,
    [string]$InputName,
    [string]$OutputName
  )

  if (-not $Title) {
    return $false
  }

  $blankWorkbookTitle = (New-Text -CodePoints @(0x5DE5, 0x4F5C, 0x7C3F) -Suffix "1 - WPS Office")
  $blankDocumentTitle = (New-Text -CodePoints @(0x6587, 0x6863) -Suffix "1 - WPS Office")
  return [bool](
    $Title -like "*$InputName*" -or
    $Title -like "*$OutputName*" -or
    $Title -like "*WPS PDF*" -or
    $Title -eq "WPS Office - WPS Office" -or
    $Title -eq $blankWorkbookTitle -or
    $Title -eq $blankDocumentTitle
  )
}

function Close-WpsArtifactsSafe {
  param(
    [hashtable]$BeforeProcesses,
    [string]$InputPath,
    [string]$OutputPath,
    [int]$TimeoutSeconds = 20
  )

  $events = @()
  $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
  $inputName = [IO.Path]::GetFileNameWithoutExtension($InputPath)
  $outputName = [IO.Path]::GetFileNameWithoutExtension($OutputPath)

  while ((Get-Date) -lt $deadline) {
    $closedAny = $false
    foreach ($proc in @(Get-Process wps,et,wpp,wpspdf -ErrorAction SilentlyContinue)) {
      $isNew = -not $BeforeProcesses.ContainsKey([int]$proc.Id)
      $title = [string]$proc.MainWindowTitle
      $titleMatches = Test-WpsArtifactTitle -Title $title -InputName $inputName -OutputName $outputName

      if ($isNew) {
        if (Stop-ProcessIfRunning -ProcessId $proc.Id) {
          $closedAny = $true
          $events += [pscustomobject]@{
            action = "stop-new-process"
            processId = $proc.Id
            processName = $proc.ProcessName
            title = $title
            newProcess = $true
          }
        }
      } elseif ($titleMatches) {
        if (Close-MainWindowByProcessId -ProcessId $proc.Id) {
          $closedAny = $true
          $events += [pscustomobject]@{
            action = "close-window"
            processId = $proc.Id
            processName = $proc.ProcessName
            title = $title
            newProcess = $false
          }
        }
      }
    }

    $remaining = @(Get-Process wps,et,wpp,wpspdf -ErrorAction SilentlyContinue | Where-Object {
      Test-WpsArtifactTitle -Title ([string]$_.MainWindowTitle) -InputName $inputName -OutputName $outputName
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

function Close-UiaWindowArtifactsSafe {
  param(
    [string]$InputPath,
    [string]$OutputPath,
    [int]$TimeoutSeconds = 10
  )

  $events = @()
  $inputName = [IO.Path]::GetFileNameWithoutExtension($InputPath)
  $outputName = [IO.Path]::GetFileNameWithoutExtension($OutputPath)
  $convertTitle = "WPS PDF" + (New-Text -CodePoints @(0x8F6C, 0x6362))
  $doneTitle = New-Text -CodePoints @(0x64CD, 0x4F5C, 0x5B8C, 0x6210)

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

      $matches = (
        $name -like "*$inputName*" -or
        $name -like "*$outputName*" -or
        $name -like "*$convertTitle*" -or
        $name -like "*$doneTitle*"
      )

      if ($matches -and $windowProcessId) {
        if (Stop-ProcessIfRunning -ProcessId $windowProcessId) {
          $closedAny = $true
          $events += [pscustomobject]@{
            action = "uia-stop-process"
            processId = $windowProcessId
            title = $name
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

function Invoke-WpsPdfStartButton {
  $code = @'
using System;
using System.Windows.Automation;
using System.Threading;

public class WpsPdfStartButton {
  static string FromCodes(int[] codes) {
    char[] chars = new char[codes.Length];
    for (int i = 0; i < codes.Length; i++) chars[i] = (char)codes[i];
    return new string(chars);
  }

  static AutomationElement FindByName(AutomationElement root, string name) {
    return root.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, name));
  }

  public static int Invoke(int timeoutSeconds) {
    string windowTitle = "WPS PDF" + FromCodes(new int[] { 0x8F6C, 0x6362 });
    string startButtonText = FromCodes(new int[] { 0x5F00, 0x59CB, 0x8F6C, 0x6362 });
    DateTime deadline = DateTime.Now.AddSeconds(timeoutSeconds);
    while (DateTime.Now < deadline) {
      var root = AutomationElement.RootElement;
      var win = root.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, windowTitle));
      if (win != null) {
        var text = FindByName(win, startButtonText);
        if (text != null) {
          AutomationElement cur = text;
          AutomationElement btn = null;
          for (int i = 0; i < 8 && cur != null; i++) {
            try {
              if (cur.Current.ControlType == ControlType.Button) {
                btn = cur;
                break;
              }
            } catch {}
            cur = TreeWalker.ControlViewWalker.GetParent(cur);
          }
          if (btn != null) {
            object pattern;
            if (btn.TryGetCurrentPattern(InvokePattern.Pattern, out pattern)) {
              ((InvokePattern)pattern).Invoke();
              return 0;
            }
            return 3;
          }
        }
      }
      Thread.Sleep(500);
    }
    return 2;
  }
}
'@

  Add-Type -TypeDefinition $code -ReferencedAssemblies UIAutomationClient,UIAutomationTypes -ErrorAction SilentlyContinue
  $result = [WpsPdfStartButton]::Invoke(30)
  if ($result -ne 0) {
    throw "Unable to invoke WPS PDF start button. Code: $result"
  }
}

$inputPath = (Resolve-Path -LiteralPath $InputPdf).Path
if ([IO.Path]::GetExtension($inputPath).ToLowerInvariant() -ne ".pdf") {
  throw "Input must be a PDF: $inputPath"
}

$beforeProcesses = Get-ProcessSnapshot
$finalOutputPath = ""
if ($OutputPath) {
  $finalOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
}

function Get-ExpectedOutputExtension {
  param(
    [string]$Action,
    [string]$FinalOutputPath
  )

  if ($FinalOutputPath) {
    $extension = [IO.Path]::GetExtension($FinalOutputPath).ToLowerInvariant()
    if ($extension) {
      return $extension
    }
  }

  switch ($Action.ToLowerInvariant()) {
    "converttoexcel" { return ".xlsx" }
    "converttopowerpoint" { return ".pptx" }
    default { return ".docx" }
  }
}

$expectedExtension = Get-ExpectedOutputExtension -Action $Action -FinalOutputPath $finalOutputPath
$expectedOutputPath = [IO.Path]::ChangeExtension($inputPath, $expectedExtension)
if (-not $finalOutputPath) {
  $finalOutputPath = $expectedOutputPath
}

$startedAt = Get-Date
if ((Test-Path -LiteralPath $finalOutputPath) -and -not $Overwrite) {
  throw "Output already exists. Pass -Overwrite to replace it: $finalOutputPath"
}

if ((Test-Path -LiteralPath $expectedOutputPath) -and $finalOutputPath -ne $expectedOutputPath) {
  throw "Default WPS output already exists. Move or remove it before using -OutputPath: $expectedOutputPath"
}

if ((Test-Path -LiteralPath $expectedOutputPath) -and -not $Overwrite) {
  throw "Default WPS output already exists. Pass -Overwrite or remove it first: $expectedOutputPath"
}

if ((Test-Path -LiteralPath $expectedOutputPath) -and $Overwrite -and $finalOutputPath -eq $expectedOutputPath) {
  Remove-Item -LiteralPath $expectedOutputPath -Force
}

$verb = $null
$runner = $null
if ($LaunchMode -eq "shell") {
  $verb = Invoke-ShellVerb -Path $inputPath -PreferredVerbs $PreferredVerb
} elseif ($LaunchMode -eq "native") {
  $runner = Start-WpsPdfConverter -Path $inputPath
} else {
  $runner = Start-WpsCloudPdfConverter -Path $inputPath
}

if (-not $NoClick) {
  Invoke-WpsPdfStartButton
}

$output = Wait-FileReady -Path $expectedOutputPath -TimeoutSeconds $TimeoutSeconds -NotBefore $startedAt

$cleanupEvents = @()
if ($finalOutputPath -ne $expectedOutputPath -and -not $NoCleanup) {
  $cleanupEvents = @(
    Close-WpsArtifactsSafe -BeforeProcesses $beforeProcesses -InputPath $inputPath -OutputPath $output.FullName -TimeoutSeconds $CleanupSeconds
    Close-UiaWindowArtifactsSafe -InputPath $inputPath -OutputPath $output.FullName -TimeoutSeconds ([Math]::Min($CleanupSeconds, 10))
  )
}

if ($finalOutputPath -ne $expectedOutputPath) {
  $parent = Split-Path $finalOutputPath -Parent
  if ($parent -and -not (Test-Path -LiteralPath $parent)) {
    New-Item -ItemType Directory -Path $parent | Out-Null
  }
  Move-Item -LiteralPath $output.FullName -Destination $finalOutputPath -Force:$Overwrite
  $output = Get-Item -LiteralPath $finalOutputPath
}

if ($finalOutputPath -eq $expectedOutputPath -and -not $NoCleanup) {
  $cleanupEvents = @(
    Close-WpsArtifactsSafe -BeforeProcesses $beforeProcesses -InputPath $inputPath -OutputPath $output.FullName -TimeoutSeconds $CleanupSeconds
    Close-UiaWindowArtifactsSafe -InputPath $inputPath -OutputPath $output.FullName -TimeoutSeconds ([Math]::Min($CleanupSeconds, 10))
  )
}

ConvertTo-JsonLine ([pscustomobject]@{
  ok = $true
  backend = "wps-uia"
  launchMode = $LaunchMode
  action = $Action
  input = $inputPath
  output = $output.FullName
  shellVerb = $verb
  runner = $runner
  length = $output.Length
  lastWriteTime = $output.LastWriteTime.ToString("o")
  cleanup = $cleanupEvents
})
