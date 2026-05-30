$ErrorActionPreference = "Stop"

$code = @'
using System;
using System.Windows.Automation;

class DumpWpsPdfUi {
  static string FromCodes(int[] codes) {
    char[] chars = new char[codes.Length];
    for (int i = 0; i < codes.Length; i++) chars[i] = (char)codes[i];
    return new string(chars);
  }

  static void Dump(AutomationElement el, int depth, int maxDepth) {
    if (el == null || depth > maxDepth) return;
    string indent = new string(' ', depth * 2);
    string name = "";
    string cls = "";
    string aid = "";
    string type = "";
    try { name = el.Current.Name; } catch {}
    try { cls = el.Current.ClassName; } catch {}
    try { aid = el.Current.AutomationId; } catch {}
    try { type = el.Current.ControlType.ProgrammaticName; } catch {}
    if (!string.IsNullOrWhiteSpace(name) || !string.IsNullOrWhiteSpace(cls) || !string.IsNullOrWhiteSpace(aid)) {
      Console.WriteLine(String.Format("{0}{1} name=[{2}] class=[{3}] id=[{4}]", indent, type, name, cls, aid));
    }
    AutomationElementCollection children = null;
    try { children = el.FindAll(TreeScope.Children, Condition.TrueCondition); } catch {}
    if (children == null) return;
    foreach (AutomationElement child in children) Dump(child, depth + 1, maxDepth);
  }

  static void Main() {
    string windowTitle = "WPS PDF" + FromCodes(new int[] { 0x8F6C, 0x6362 });
    var root = AutomationElement.RootElement;
    var win = root.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, windowTitle));
    if (win == null) {
      Console.WriteLine("WPS PDF conversion window not found");
      Environment.Exit(2);
      return;
    }
    Dump(win, 0, 8);
  }
}
'@

$exe = Join-Path $PSScriptRoot "DumpWpsPdfUi.exe"
Add-Type -TypeDefinition $code `
  -ReferencedAssemblies UIAutomationClient,UIAutomationTypes `
  -OutputAssembly $exe `
  -OutputType ConsoleApplication

& $exe
