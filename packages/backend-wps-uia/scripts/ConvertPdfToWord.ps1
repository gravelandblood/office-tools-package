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

  [switch]$Overwrite
)

$ErrorActionPreference = "Stop"

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

$expectedOutputPath = [IO.Path]::ChangeExtension($inputPath, ".docx")
$finalOutputPath = $expectedOutputPath
if ($OutputPath) {
  $finalOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
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

if ($finalOutputPath -ne $expectedOutputPath) {
  $parent = Split-Path $finalOutputPath -Parent
  if ($parent -and -not (Test-Path -LiteralPath $parent)) {
    New-Item -ItemType Directory -Path $parent | Out-Null
  }
  Move-Item -LiteralPath $output.FullName -Destination $finalOutputPath -Force:$Overwrite
  $output = Get-Item -LiteralPath $finalOutputPath
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
})
