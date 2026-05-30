export const capabilities = {
  backend: "wps-jsapi",
  status: "planned-adapter",
  preferredWhen: [
    "active WPS document context",
    "selection or cursor operations",
    "add-in/sidebar workflows",
    "web-to-WPS callbacks",
    "controlled OAAssist.ShellExecute launch"
  ],
  commands: [
    "wpsJsapi.ping",
    "wps.activeDocument",
    "wps.selection",
    "wps.open",
    "wps.saveAs",
    "wps.insertText",
    "wps.replace",
    "et.activeWorkbook",
    "et.range",
    "wpp.activePresentation",
    "wpp.slides",
    "wpp.shapes",
    "wps.launch"
  ]
};
