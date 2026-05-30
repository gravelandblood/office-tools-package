$ErrorActionPreference = "Stop"

param(
  [Parameter(Mandatory = $true)]
  [string]$Path,
  [int]$Take = 30
)

$code = @'
using System;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

class ExtractDocxText {
  static void Main(string[] args) {
    string path = args[0];
    int take = Int32.Parse(args[1]);
    using (var zip = ZipFile.OpenRead(path)) {
      var entry = zip.GetEntry("word/document.xml");
      using (var stream = entry.Open()) {
        var doc = XDocument.Load(stream);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var paras = doc.Descendants(w + "p")
          .Select(p => string.Concat(p.Descendants(w + "t").Select(t => (string)t)))
          .Where(s => !string.IsNullOrWhiteSpace(s))
          .Take(take)
          .ToList();
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        Console.WriteLine("Paragraphs=" + paras.Count);
        foreach (var p in paras) Console.WriteLine(p);
      }
    }
  }
}
'@

$exe = Join-Path $PSScriptRoot "ExtractDocxText.exe"
Add-Type -TypeDefinition $code `
  -ReferencedAssemblies System.IO.Compression,System.IO.Compression.FileSystem,System.Xml.Linq,System.Xml `
  -OutputAssembly $exe `
  -OutputType ConsoleApplication

& $exe $Path $Take
