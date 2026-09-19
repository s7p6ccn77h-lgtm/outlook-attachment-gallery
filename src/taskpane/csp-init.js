// Diagnostic only, used by csp-test.html (the separate "Attaché CSP test" add-in).
// Applies one of several candidate Content Security Policies (?p=1..4, ?p=0 for none) and shows,
// on screen, what Outlook's environment tries to load that the policy blocks.
(function () {
  var BASE = "default-src 'self'; object-src 'none'; base-uri 'none'; form-action 'none'; ";
  var MS = "https://appsforoffice.microsoft.com https://ajax.aspnetcdn.com"; // Mac office.js loads MicrosoftAjax.js from the latter
  var POLICIES = {
    0: { text: "", note: "No policy (baseline)" },
    1: {
      note: "Strict: this site + Microsoft (incl. Ajax CDN)",
      text: BASE +
        "script-src 'self' " + MS + "; style-src 'self'; img-src 'self' data:; " +
        "connect-src 'self' " + MS + " https://*.microsoft.com https://*.office.com https://*.office.net https://*.outlook.com https://*.live.com; " +
        "frame-src 'self' " + MS + "; worker-src 'self'",
    },
    2: {
      note: "P1 + any https for frames/connections/images, inline styles",
      text: BASE +
        "script-src 'self' " + MS + "; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob: https:; " +
        "connect-src 'self' https:; frame-src 'self' https:; worker-src 'self'",
    },
    3: {
      note: "P2 + inline scripts",
      text: BASE +
        "script-src 'self' 'unsafe-inline' " + MS + "; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob: https:; " +
        "connect-src 'self' https:; frame-src 'self' https:; worker-src 'self'",
    },
    4: {
      note: "P3 + eval and blob scripts/workers",
      text: BASE +
        "script-src 'self' 'unsafe-inline' 'unsafe-eval' blob: " + MS + "; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob: https:; " +
        "connect-src 'self' https: wss:; frame-src 'self' https:; worker-src 'self' blob:",
    },
  };

  var params = new URLSearchParams(location.search);
  var level = parseInt(params.get("p"), 10);
  if (!(level in POLICIES)) level = 1;

  var violations = [];
  var seen = {};
  document.addEventListener("securitypolicyviolation", function (e) {
    var key = e.effectiveDirective + "|" + e.blockedURI;
    if (seen[key]) return;
    seen[key] = true;
    violations.push({
      directive: e.effectiveDirective,
      blocked: e.blockedURI || "(inline)",
      sample: (e.sample || "").slice(0, 60),
      where: (e.sourceFile || "").replace(/^https?:\/\/[^/]+\//, "") + (e.lineNumber ? ":" + e.lineNumber : ""),
    });
  });

  if (POLICIES[level].text) {
    var meta = document.createElement("meta");
    meta.httpEquiv = "Content-Security-Policy";
    meta.content = POLICIES[level].text;
    document.head.appendChild(meta);
  }

  var started = Date.now();
  var ready = null;
  var hooked = false;

  function build() {
    var panel = document.createElement("div");
    panel.id = "cspPanel";
    Object.assign(panel.style, {
      position: "fixed", left: "0", right: "0", bottom: "0", maxHeight: "60%", overflow: "auto",
      background: "#FFF8E1", borderTop: "2px solid #E6A700", padding: "6px 8px",
      font: "11px/1.35 -apple-system, Segoe UI, sans-serif", color: "#3b2f00", zIndex: "99999",
    });
    document.body.appendChild(panel);
    setInterval(function () { render(panel); }, 500);
    render(panel);
  }

  function button(label, target) {
    var b = document.createElement("button");
    b.textContent = label;
    Object.assign(b.style, { marginRight: "6px", padding: "2px 8px", fontSize: "11px" });
    b.addEventListener("click", function () { location.search = "?p=" + target; });
    return b;
  }

  function line(parent, text, bold) {
    var d = document.createElement("div");
    d.textContent = text;
    if (bold) d.style.fontWeight = "700";
    parent.appendChild(d);
  }

  function render(panel) {
    if (!hooked && window.Office && typeof Office.onReady === "function") {
      hooked = true;
      Office.onReady(function (info) { ready = { ms: Date.now() - started, host: String(info && info.host) }; });
    }
    var waited = Math.round((Date.now() - started) / 1000);
    panel.textContent = "";
    line(panel, "CSP TEST - policy P" + level + ": " + POLICIES[level].note, true);
    line(panel, "office.js loaded: " + (window.Office ? "yes" : "NO"));
    line(panel, "Office ready: " + (ready ? "YES after " + ready.ms + " ms (host " + ready.host + ")" : waited < 12 ? "waiting… " + waited + "s" : "NO after " + waited + "s"));
    var count = document.getElementById("attachmentCount");
    line(panel, "Attachments shown: " + (count ? count.textContent : "?"));
    line(panel, "Blocked by policy (" + violations.length + "):");
    if (!violations.length) line(panel, "  none");
    violations.slice(0, 14).forEach(function (v) {
      line(panel, "  [" + v.directive + "] " + v.blocked + (v.sample ? "  \"" + v.sample + "\"" : "") + (v.where ? "  @" + v.where : ""));
    });
    var nav = document.createElement("div");
    nav.style.marginTop = "4px";
    if (level > 0) nav.appendChild(button("< P" + (level - 1), level - 1));
    if (level < 4) nav.appendChild(button("P" + (level + 1) + " >", level + 1));
    panel.appendChild(nav);
  }

  if (document.body) build();
  else document.addEventListener("DOMContentLoaded", build);
})();
