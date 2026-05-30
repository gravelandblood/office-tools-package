$ErrorActionPreference = "Stop"

$code = @'
using System;
using System.Windows.Automation;

class ClickStartConvert {
  static string FromCodes(int[] codes) {
    char[] chars = new char[codes.Length];
    for (int i = 0; i < codes.Length; i++) chars[i] = (char)codes[i];
    return new string(chars);
  }

  static AutomationElement FindByName(AutomationElement root, string name) {
    return root.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, name));
  }

  static void Main() {
    string windowTitle = "WPS PDF" + FromCodes(new int[] { 0x8F6C, 0x6362 });
    string startButtonText = FromCodes(new int[] { 0x5F00, 0x59CB, 0x8F6C, 0x6362 });
    var root = AutomationElement.RootElement;
    var win = root.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, windowTitle));
    if (win == null) {
      Console.WriteLine("window not found");
      Environment.Exit(2);
      return;
    }

    var text = FindByName(win, startButtonText);
    if (text == null) {
      Console.WriteLine("start text not found");
      Environment.Exit(3);
      return;
    }

    AutomationElement cur = text;
    AutomationElement btn = null;
    for (int i = 0; i < 6 && cur != null; i++) {
      try {
        if (cur.Current.ControlType == ControlType.Button) {
          btn = cur;
          break;
        }
      } catch {}
      cur = TreeWalker.ControlViewWalker.GetParent(cur);
    }

    if (btn == null) {
      Console.WriteLine("button parent not found");
      Environment.Exit(4);
      return;
    }

    object pattern;
    if (btn.TryGetCurrentPattern(InvokePattern.Pattern, out pattern)) {
      ((InvokePattern)pattern).Invoke();
      Console.WriteLine("invoked");
    } else {
      Console.WriteLine("invoke pattern not available");
      Environment.Exit(5);
    }
  }
}
'@

$exe = Join-Path $PSScriptRoot "ClickStartConvert.exe"
Add-Type -TypeDefinition $code `
  -ReferencedAssemblies UIAutomationClient,UIAutomationTypes `
  -OutputAssembly $exe `
  -OutputType ConsoleApplication

& $exe
