(function () {
  function parseParam(param) {
    if (!param) return {};
    if (typeof param === "string") {
      try {
        return JSON.parse(param);
      } catch (error) {
        return { raw: param, parseError: String(error) };
      }
    }
    return param;
  }

  function ok(data) {
    return {
      ok: true,
      at: new Date().toISOString(),
      data: data || {}
    };
  }

  function fail(error, data) {
    return {
      ok: false,
      at: new Date().toISOString(),
      error: String(error && error.message ? error.message : error),
      data: data || {}
    };
  }

  function getApplicationInfo() {
    var info = {
      hasWindowApplication: typeof window !== "undefined" && !!window.Application,
      hasGlobalWps: typeof wps !== "undefined" && !!wps
    };

    try {
      if (window.Application) {
        info.version = window.Application.Version || "";
        info.name = window.Application.Name || "";
        info.activeDocumentName = window.Application.ActiveDocument ? window.Application.ActiveDocument.Name : "";
        info.documentsCount = window.Application.Documents ? window.Application.Documents.Count : null;
      }
    } catch (error) {
      info.applicationError = String(error);
    }

    try {
      if (typeof wps !== "undefined" && wps && wps.OAAssist) {
        info.hasOAAssist = true;
        info.hasShellExecute = typeof wps.OAAssist.ShellExecute === "function";
        info.hasWebNotify = typeof wps.OAAssist.WebNotify === "function";
      } else {
        info.hasOAAssist = false;
      }
    } catch (error) {
      info.oaAssistError = String(error);
    }

    return info;
  }

  function quoteArg(value) {
    return '"' + String(value).replace(/"/g, '\\"') + '"';
  }

  window.CodexBridge_Ping = function (param) {
    return ok({
      param: parseParam(param),
      app: getApplicationInfo(),
      location: window.location ? window.location.href : ""
    });
  };

  window.CodexBridge_GetActiveDocument = function (param) {
    try {
      var args = parseParam(param);
      var doc = window.Application && window.Application.ActiveDocument;
      if (!doc) {
        return ok({ param: args, active: false });
      }
      return ok({
        param: args,
        active: true,
        name: doc.Name || "",
        path: doc.Path || "",
        fullName: doc.FullName || ""
      });
    } catch (error) {
      return fail(error);
    }
  };

  window.CodexBridge_OpenDocument = function (param) {
    var args = parseParam(param);
    try {
      if (!args.filePath) {
        return fail("filePath is required", args);
      }
      var doc = window.Application.Documents.OpenFromUrl(args.filePath);
      return ok({
        opened: true,
        name: doc && doc.Name ? doc.Name : "",
        filePath: args.filePath
      });
    } catch (error) {
      return fail(error, args);
    }
  };

  window.CodexBridge_LaunchPdfConvert = function (param) {
    var args = parseParam(param);
    try {
      if (!args.filePath) {
        return fail("filePath is required", args);
      }

      if (!args.wpsExe || !args.appFramework) {
        return fail("wpsExe and appFramework are required", args);
      }

      var wpsExe = args.wpsExe;
      var appFramework = args.appFramework;
      var action = args.action || "ConvertToWord";
      var command = [
        quoteArg(wpsExe),
        "Run",
        "/InstanceId=kpdf2wordv2",
        quoteArg(appFramework),
        "/appId=kpdf2wordv2",
        quoteArg("/appname=WPS PDF Convert"),
        "/size=960&670",
        "/src=codex_bridge",
        "/switchskin=0",
        "/action=" + action,
        "/file=" + quoteArg(args.filePath)
      ].join(" ");

      if (typeof wps === "undefined" || !wps || !wps.OAAssist || typeof wps.OAAssist.ShellExecute !== "function") {
        return fail("wps.OAAssist.ShellExecute is not available", { command: command });
      }

      wps.OAAssist.ShellExecute(command);
      return ok({ launched: true, command: command });
    } catch (error) {
      return fail(error, args);
    }
  };

  window.OnAddinLoad = function () {};
  window.OnAction = function (control) {
    if (control && control.Id === "btnCodexPing") {
      alert(JSON.stringify(window.CodexBridge_Ping({ from: "ribbon" })));
    }
  };
})();
