param(
  [Parameter(Mandatory = $true)]
  [string]$InputPdf,

  [int]$TimeoutSeconds = 180,

  [switch]$NoClick
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

function Wait-FileReady {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [int]$TimeoutSeconds = 180
  )

  $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
  $lastLength = -1
  $stableCount = 0

  while ((Get-Date) -lt $deadline) {
    if (Test-Path -LiteralPath $Path) {
      $item = Get-Item -LiteralPath $Path
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

$outputPath = [IO.Path]::ChangeExtension($inputPath, ".docx")
$verb = Invoke-ShellVerb -Path $inputPath

if (-not $NoClick) {
  Invoke-WpsPdfStartButton
}

$output = Wait-FileReady -Path $outputPath -TimeoutSeconds $TimeoutSeconds

ConvertTo-JsonLine ([pscustomobject]@{
  ok = $true
  input = $inputPath
  output = $output.FullName
  shellVerb = $verb
  length = $output.Length
  lastWriteTime = $output.LastWriteTime.ToString("o")
})
