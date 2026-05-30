export const capabilities = {
  backend: "officecli",
  status: "planned-adapter",
  preferredWhen: [
    "headless OOXML inspection",
    "headless OOXML mutation",
    "deterministic JSON output",
    "CI or MCP execution"
  ],
  commands: [
    "docx.inspect",
    "docx.query",
    "docx.text",
    "docx.replace",
    "xlsx.inspect",
    "xlsx.query",
    "xlsx.set",
    "pptx.inspect",
    "pptx.query",
    "pptx.addShape",
    "ooxml.validate",
    "ooxml.applyPlan"
  ]
};
