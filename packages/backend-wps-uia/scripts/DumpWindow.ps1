param(
  [string]$Title = "",

  [int]$ProcessId = 0,

  [int]$MaxDepth = 6,

  [int]$TimeoutSeconds = 5,

  [switch]$All
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

function ConvertTo-JsonLine($Object) {
  $Object | ConvertTo-Json -Compress -Depth 16
}

function Convert-ControlTypeName {
  param($ControlType)
  if (-not $ControlType) {
    return ""
  }
  return ([string]$ControlType.ProgrammaticName) -replace "^ControlType\.", ""
}

function Convert-UiaElement {
  param(
    [Parameter(Mandatory = $true)]
    $Element,
    [int]$Depth = 0,
    [int]$MaxDepth = 6
  )

  $children = @()
  if ($Depth -lt $MaxDepth) {
    try {
      $children = @($Element.FindAll(
        [System.Windows.Automation.TreeScope]::Children,
        [System.Windows.Automation.Condition]::TrueCondition
      ) | ForEach-Object {
        Convert-UiaElement -Element $_ -Depth ($Depth + 1) -MaxDepth $MaxDepth
      })
    } catch {
      $children = @()
    }
  }

  $patterns = @()
  foreach ($patternInfo in @(
    [System.Windows.Automation.InvokePattern]::Pattern,
    [System.Windows.Automation.SelectionItemPattern]::Pattern,
    [System.Windows.Automation.ValuePattern]::Pattern,
    [System.Windows.Automation.TogglePattern]::Pattern,
    [System.Windows.Automation.WindowPattern]::Pattern
  )) {
    try {
      $pattern = $null
      if ($Element.TryGetCurrentPattern($patternInfo, [ref]$pattern)) {
        $patterns += $patternInfo.ProgrammaticName
      }
    } catch {}
  }

  [pscustomobject]@{
    name = [string]$Element.Current.Name
    automationId = [string]$Element.Current.AutomationId
    className = [string]$Element.Current.ClassName
    controlType = Convert-ControlTypeName $Element.Current.ControlType
    processId = [int]$Element.Current.ProcessId
    isEnabled = [bool]$Element.Current.IsEnabled
    isOffscreen = [bool]$Element.Current.IsOffscreen
    patterns = @($patterns)
    children = @($children)
  }
}

Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes

$deadline = (Get-Date).AddSeconds($TimeoutSeconds)
$matches = @()
while ((Get-Date) -lt $deadline) {
  $root = [System.Windows.Automation.AutomationElement]::RootElement
  $windows = @($root.FindAll(
    [System.Windows.Automation.TreeScope]::Children,
    [System.Windows.Automation.Condition]::TrueCondition
  ))

  $matches = @($windows | Where-Object {
    if ($All) {
      return $true
    }
    $ok = $true
    if ($Title) {
      $ok = $ok -and ([string]$_.Current.Name -like "*$Title*")
    }
    if ($ProcessId) {
      $ok = $ok -and ([int]$_.Current.ProcessId -eq $ProcessId)
    }
    return $ok
  })

  if ($matches.Count -gt 0 -or $All) {
    break
  }
  Start-Sleep -Milliseconds 300
}

ConvertTo-JsonLine ([pscustomobject]@{
  ok = $true
  backend = "wps-uia"
  title = $Title
  processId = $ProcessId
  maxDepth = $MaxDepth
  windows = @($matches | ForEach-Object {
    Convert-UiaElement -Element $_ -MaxDepth $MaxDepth
  })
})
