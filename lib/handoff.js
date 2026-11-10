/* Reports handoff — ส่งต่อข้อมูลระหว่างผลัด (offline code + QR + link)
   PREFIX-base64url(deflate(JSON{v,form,draft,shift,extra})) */
(function (global) {
  "use strict";

  var VER = 1;
  var FORMS = {
    DLA: { path: "delay/", label: "รายงานเที่ยวบินล่าช้า" },
    RF:  { path: "return-flight/", label: "รายงาน Return Flight" },
    RT:  { path: "return-taxi/", label: "รายงาน Return Taxi" },
    AOG: { path: "aog/", label: "รายงาน Aircraft on Ground (AOG)" },
    DIV: { path: "divert/", label: "รายงาน Divert" }
  };

  function utf8(s) {
    return new TextEncoder().encode(s);
  }
  function fromUtf8(u8) {
    return new TextDecoder().decode(u8);
  }
  function b64urlEncode(u8) {
    var s = "";
    for (var i = 0; i < u8.length; i++) s += String.fromCharCode(u8[i]);
    return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  }
  function b64urlDecode(str) {
    var s = String(str || "").replace(/-/g, "+").replace(/_/g, "/");
    while (s.length % 4) s += "=";
    var bin = atob(s);
    var u8 = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    return u8;
  }

  function deflateBytes(u8) {
    if (typeof pako !== "undefined" && pako.deflate) return pako.deflate(u8);
    throw new Error("ไม่มีไลบรารีบีบอัด (pako)");
  }
  function inflateBytes(u8) {
    if (typeof pako !== "undefined" && pako.inflate) return pako.inflate(u8);
    throw new Error("ไม่มีไลบรารีบีบอัด (pako)");
  }

  function encodePayload(obj) {
    var json = JSON.stringify(obj);
    var raw = utf8(json);
    var z = deflateBytes(raw);
    return { codeBody: b64urlEncode(z), rawBytes: raw.length, zipBytes: z.length, json: json };
  }
  function decodePayload(body) {
    var z = b64urlDecode(body);
    var raw = inflateBytes(z);
    return JSON.parse(fromUtf8(raw));
  }

  function parseCode(text) {
    var s = String(text || "").trim();
    // allow full URL with #h= or ?h=
    var hm = s.match(/[?#]h=([A-Za-z0-9\-_=]+)/);
    if (hm) s = hm[1];
    // allow pasted PREFIX-body or just body with known prefix embedded
    var m = s.match(/^([A-Z]{2,4})-(.+)$/i);
    if (!m) return { ok: false, error: "รหัสไม่ถูกต้อง — ต้องขึ้นต้นด้วย PREFIX- เช่น DIV-…" };
    var prefix = m[1].toUpperCase();
    var body = m[2].replace(/\s+/g, "");
    if (!FORMS[prefix]) return { ok: false, error: "ไม่รู้จักรหัส " + prefix + "- — ตรวจว่าคัดลอกครบหรือยัง" };
    try {
      var payload = decodePayload(body);
      if (!payload || payload.v !== VER) {
        return { ok: false, error: "เวอร์ชันรหัสไม่รองรับ (ต้องการ v" + VER + ")" };
      }
      if (payload.form && String(payload.form).toUpperCase() !== prefix) {
        // tolerate mismatch if form field differs — prefer prefix
      }
      payload.form = prefix;
      return { ok: true, prefix: prefix, payload: payload, meta: FORMS[prefix] };
    } catch (e) {
      return { ok: false, error: "ถอดรหัสไม่ได้ — รหัสเสียหรือคัดลอกไม่ครบ" };
    }
  }

  function makeCode(prefix, state) {
    var p = String(prefix || "").toUpperCase();
    if (!FORMS[p]) throw new Error("unknown prefix");
    var obj = {
      v: VER,
      form: p,
      draft: state.draft == null ? null : state.draft,
      shift: state.shift || null,
      extra: state.extra || null
    };
    var enc = encodePayload(obj);
    return {
      code: p + "-" + enc.codeBody,
      rawBytes: enc.rawBytes,
      zipBytes: enc.zipBytes,
      json: enc.json
    };
  }

  function siteOrigin() {
    // Prefer absolute URL of current page's site root (…/reports/ or preview root)
    var u = new URL(location.href);
    // strip form subpath: find /delay/ /divert/ etc
    var path = u.pathname;
    var forms = Object.keys(FORMS).map(function (k) { return FORMS[k].path.replace(/\/$/, ""); });
    // locate "site root" as directory containing form folders OR current folder's parent chain
    // For GitHub Pages: …/reports/delay/ → root …/reports/
    // For preview server: /delay/ → /
    var root = u.origin + "/";
    var idx = -1;
    ["delay", "divert", "return-flight", "return-taxi", "aog", "changelog"].forEach(function (seg) {
      var i = path.indexOf("/" + seg + "/");
      if (i < 0 && path.endsWith("/" + seg)) i = path.lastIndexOf("/" + seg);
      if (i >= 0 && (idx < 0 || i < idx)) idx = i;
    });
    if (idx > 0) root = u.origin + path.slice(0, idx + 1);
    else if (idx === 0) root = u.origin + "/";
    else {
      // fallback: dirname of current path without trailing form
      root = u.origin + path.replace(/\/[^/]*\/?$/, "/");
      if (!/[a-z]/i.test(root.replace(u.origin, ""))) root = u.origin + "/";
    }
    return root;
  }

  function formUrl(prefix, code) {
    var meta = FORMS[prefix];
    if (!meta) return "";
    var root = siteOrigin();
    var url = root + meta.path + "#h=" + encodeURIComponent(code);
    return url;
  }

  function qrDataUrl(text, opts) {
    opts = opts || {};
    if (typeof qrcode !== "function") throw new Error("ไม่มีไลบรารี QR");
    // Prefer ECC L for capacity; type 0 = auto
    var ecc = opts.ecc || "L";
    var qr = qrcode(0, ecc);
    qr.addData(text);
    qr.make();
    var cell = opts.cell || 3;
    var margin = opts.margin != null ? opts.margin : 2;
    return qr.createDataURL(cell, margin);
  }

  function qrCapacityHint(len) {
    // Approx max alphanumeric for QR version 40 ECC L ~ 2953; binary ~ 2953 bytes
    // Our codes are base64url ≈ alphanumeric. Practical phone scan ~ 800–1500 chars.
    if (len <= 800) return { ok: true, level: "ดี", note: "" };
    if (len <= 1500) return { ok: true, level: "พอใช้", note: "คิวอาร์ค่อนข้างแน่น — สแกนใกล้ ๆ หรือใช้โค้ด/ลิงก์แทนได้" };
    if (len <= 2500) return { ok: true, level: "ใหญ่", note: "คิวอาร์ใหญ่มาก อาจสแกนยากบนมือถือบางรุ่น — แนะนำส่งโค้ดหรือลิงก์" };
    return { ok: false, level: "ใหญ่เกิน", note: "ข้อมูลใหญ่เกินกว่าจะทำคิวอาร์ที่สแกนได้แน่นอน — ใช้คัดลอกโค้ดหรือลิงก์แทน" };
  }

  function ensureCss() {
    if (document.getElementById("handoffCss")) return;
    var css = document.createElement("style");
    css.id = "handoffCss";
    css.textContent =
      ".topbar{display:flex;align-items:center;justify-content:space-between;gap:8px;flex-wrap:nowrap}" +
      ".topbar .back{flex:0 1 auto;min-width:0}" +
      "a.ho-pill,button.ho-pill{display:inline-flex;align-items:center;justify-content:center;gap:6px;min-height:44px;padding:0 14px;border-radius:22px;background:var(--card,#fff);color:var(--accent,#3a5a8c);font-size:15px;font-weight:600;text-decoration:none;box-shadow:0 1px 3px var(--shadow-c,rgba(0,0,0,.08));border:1px solid var(--line,#e2e0ea);cursor:pointer;font-family:inherit;white-space:nowrap;flex:0 0 auto;margin-left:auto}" +
      "a.ho-pill:active,button.ho-pill:active{background:var(--accent-soft,#e8eef8)}" +
      ".ho-sheet{position:fixed;inset:0;z-index:85;background:rgba(20,24,40,.55);display:flex;align-items:flex-end;justify-content:center;padding:12px}" +
      ".ho-sheet.hidden{display:none!important}" +
      ".ho-card{background:var(--card,#fff);color:var(--ink,#2b2f45);border-radius:18px 18px 12px 12px;width:min(520px,100%);padding:14px 16px 16px;border:1px solid var(--line,#e2e0ea);box-shadow:0 -4px 24px rgba(0,0,0,.18);max-height:min(88vh,760px);overflow:auto;box-sizing:border-box}" +
      ".ho-card h3{margin:0 0 10px;font-size:18px;color:var(--accent,#3a5a8c)}" +
      ".ho-tabs{display:flex;background:var(--seg-bg,#eceaf3);border-radius:12px;padding:4px;gap:4px;border:1px solid var(--line,#e2e0ea);margin-bottom:12px}" +
      ".ho-tabs button{flex:1;white-space:nowrap;font-size:15px;padding:10px 6px;border-radius:8px;background:transparent;color:var(--muted,#7a7f93);border:0;cursor:pointer;font-family:inherit;font-weight:600}" +
      ".ho-tabs button.on{background:var(--seg-on,#fff);color:var(--seg-on-ink,#2b2f45);box-shadow:inset 0 0 0 1px var(--seg-on-line,#d8d4e4),0 1px 3px rgba(58,48,110,.18)}" +
      ".ho-pane.hidden{display:none!important}" +
      ".ho-lab{font-size:13px;color:var(--muted2,#7a7f93);margin:10px 0 4px}" +
      ".ho-code{font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;font-size:12px;line-height:1.35;word-break:break-all;background:var(--chip,#eceaf3);border-radius:10px;padding:10px;border:1px solid var(--line,#e2e0ea);max-height:96px;overflow:auto}" +
      ".ho-row{display:flex;gap:8px;margin-top:8px;flex-wrap:wrap}" +
      ".ho-row button{flex:1;min-width:100px;font-size:15px;padding:10px;border:0;border-radius:12px;cursor:pointer;font-family:inherit;font-weight:600}" +
      ".ho-primary{background:var(--line-green,#15834a);color:#fff}" +
      ".ho-secondary{background:var(--btn,#cddcf3);color:var(--btn-ink,#1f3a66);border:1px solid var(--accent-line,#9eb6d9)!important}" +
      ".ho-muted{background:var(--chip,#eceaf3);color:var(--ink,#2b2f45)}" +
      ".ho-qr{display:flex;justify-content:center;margin:12px 0;padding:10px;background:#fff;border-radius:12px;border:1px solid var(--line,#e2e0ea)}" +
      ".ho-qr img{width:min(260px,70vw);height:auto;image-rendering:pixelated}" +
      ".ho-hint{font-size:13px;color:var(--muted2,#7a7f93);margin:6px 0 0}" +
      ".ho-warn{font-size:13px;color:var(--warn-ink,#b4470f);background:var(--warn-bg,#fff3e6);border:1px solid var(--warn-line,#f1c08a);border-radius:10px;padding:8px 10px;margin:8px 0}" +
      ".ho-input{width:100%;box-sizing:border-box;font-size:14px;padding:10px;border-radius:10px;border:1px solid var(--field-line,#cfd3e0);font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;min-height:88px;resize:vertical;background:var(--card,#fff);color:var(--ink,#2b2f45)}" +
      ".ho-err{color:#b00020;font-size:13px;margin-top:6px;font-weight:600}" +
      ".ho-err.hidden{display:none!important}" +
      "@media (max-width:420px){a.ho-pill,button.ho-pill{padding:0 12px;font-size:14px}}";
    document.head.appendChild(css);
  }

  function el(tag, attrs, html) {
    var n = document.createElement(tag);
    if (attrs) Object.keys(attrs).forEach(function (k) {
      if (k === "className") n.className = attrs[k];
      else if (k === "text") n.textContent = attrs[k];
      else n.setAttribute(k, attrs[k]);
    });
    if (html != null) n.innerHTML = html;
    return n;
  }

  function toast(msg) {
    var t = document.getElementById("toast");
    if (t) {
      t.textContent = msg;
      t.classList.add("show");
      setTimeout(function () { t.classList.remove("show"); }, 1800);
      return;
    }
    alert(msg);
  }

  async function copyText(txt) {
    try {
      if (navigator.clipboard && window.isSecureContext) {
        await navigator.clipboard.writeText(txt);
        toast("คัดลอกแล้ว ✓");
        return;
      }
    } catch (e) {}
    var ta = document.createElement("textarea");
    ta.value = txt;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand("copy");
    ta.remove();
    toast("คัดลอกแล้ว ✓");
  }

  function closeSheet(node) {
    if (node && node.parentNode) node.parentNode.removeChild(node);
  }

  function fillCreatePane(pane, cfg, made) {
    while (pane.firstChild) pane.removeChild(pane.firstChild);
    var url = formUrl(cfg.prefix, made.code);
    var hint = qrCapacityHint(made.code.length);
    pane.appendChild(el("div", { className: "ho-hint", text: "รหัสขึ้นต้น " + cfg.prefix + "- · บีบอัดแล้ว " + made.zipBytes + " ไบต์ (ดิบ " + made.rawBytes + ")" }));

    pane.appendChild(el("div", { className: "ho-lab", text: "รหัสส่งต่อ" }));
    pane.appendChild(el("div", { className: "ho-code", text: made.code }));
    var row1 = el("div", { className: "ho-row" });
    var bCopy = el("button", { type: "button", className: "ho-primary", text: "📋 คัดลอกรหัส" });
    bCopy.onclick = function () { copyText(made.code); };
    row1.appendChild(bCopy);
    pane.appendChild(row1);

    pane.appendChild(el("div", { className: "ho-lab", text: "ลิงก์เปิดฟอร์มพร้อมรหัส" }));
    pane.appendChild(el("div", { className: "ho-code", text: url }));
    var row2 = el("div", { className: "ho-row" });
    var bLink = el("button", { type: "button", className: "ho-secondary", text: "🔗 คัดลอกลิงก์" });
    bLink.onclick = function () { copyText(url); };
    row2.appendChild(bLink);
    if (navigator.share) {
      var bShare = el("button", { type: "button", className: "ho-primary", text: "แชร์" });
      bShare.onclick = function () {
        navigator.share({ title: "ส่งต่อข้อมูลรายงาน", text: made.code, url: url }).catch(function (e) {
          if (e && e.name !== "AbortError") copyText(url);
        });
      };
      row2.appendChild(bShare);
    }
    pane.appendChild(row2);

    pane.appendChild(el("div", { className: "ho-lab", text: "คิวอาร์โค้ด (สแกนแล้วไปหน้าฟอร์มพร้อมรหัส)" }));
    var qrWrap = el("div", { className: "ho-qr" });
    try {
      if (!hint.ok) {
        qrWrap.appendChild(el("div", { className: "ho-warn", text: hint.note }));
      } else {
        var img = el("img", { alt: "QR ส่งต่อข้อมูล" });
        img.src = qrDataUrl(url, { ecc: "L", cell: 3 });
        qrWrap.appendChild(img);
        if (hint.note) pane.appendChild(el("div", { className: "ho-hint", text: hint.note }));
      }
    } catch (e) {
      qrWrap.appendChild(el("div", { className: "ho-err", text: "สร้างคิวอาร์ไม่ได้: " + (e && e.message ? e.message : e) }));
    }
    pane.appendChild(qrWrap);
  }

  function fillPastePane(pane, cfg, prefill) {
    while (pane.firstChild) pane.removeChild(pane.firstChild);
    pane.appendChild(el("div", { className: "ho-hint", text: "วางรหัส " + cfg.prefix + "-… หรือลิงก์ที่มี #h=" }));
    var ta = el("textarea", { className: "ho-input", placeholder: cfg.prefix + "-…" });
    if (prefill) ta.value = prefill;
    pane.appendChild(ta);
    var err = el("div", { className: "ho-err hidden" });
    pane.appendChild(err);
    var row = el("div", { className: "ho-row" });
    var bGo = el("button", { type: "button", className: "ho-primary", text: "นำเข้าข้อมูล" });
    bGo.onclick = function () {
      err.classList.add("hidden");
      var r = applyImport(cfg, ta.value, { fromUi: true });
      if (r === "ok") {
        var sheet = pane.closest(".ho-sheet");
        closeSheet(sheet);
      } else if (r && r.error) {
        err.textContent = r.error;
        err.classList.remove("hidden");
      }
    };
    row.appendChild(bGo);
    pane.appendChild(row);
    setTimeout(function () { ta.focus(); }, 50);
  }

  function openHandoffSheet(cfg, opts) {
    opts = opts || {};
    ensureCss();
    var existing = document.querySelector(".ho-sheet");
    if (existing) closeSheet(existing);

    var sheet = el("div", { className: "ho-sheet", role: "dialog", "aria-modal": "true" });
    var card = el("div", { className: "ho-card" });
    card.appendChild(el("h3", { text: "ส่งต่อข้อมูล" }));

    var tabs = el("div", { className: "ho-tabs", role: "tablist" });
    var tCreate = el("button", { type: "button", className: "on", text: "สร้างโค้ด", role: "tab", "aria-selected": "true", "data-tab": "create" });
    var tPaste = el("button", { type: "button", text: "ใส่โค้ด", role: "tab", "aria-selected": "false", "data-tab": "paste" });
    tabs.appendChild(tCreate);
    tabs.appendChild(tPaste);
    card.appendChild(tabs);

    var paneCreate = el("div", { className: "ho-pane", id: "hoPaneCreate", role: "tabpanel" });
    var panePaste = el("div", { className: "ho-pane hidden", id: "hoPanePaste", role: "tabpanel" });
    card.appendChild(paneCreate);
    card.appendChild(panePaste);

    function showTab(name) {
      var create = name === "create";
      tCreate.classList.toggle("on", create);
      tPaste.classList.toggle("on", !create);
      tCreate.setAttribute("aria-selected", create ? "true" : "false");
      tPaste.setAttribute("aria-selected", create ? "false" : "true");
      paneCreate.classList.toggle("hidden", !create);
      panePaste.classList.toggle("hidden", create);
      if (create && !paneCreate.dataset.ready) {
        try {
          if (typeof cfg.saveBefore === "function") cfg.saveBefore();
          var made = makeCode(cfg.prefix, cfg.getState());
          fillCreatePane(paneCreate, cfg, made);
          paneCreate.dataset.ready = "1";
        } catch (e) {
          paneCreate.appendChild(el("div", { className: "ho-err", text: "สร้างรหัสไม่ได้: " + (e && e.message ? e.message : e) }));
          paneCreate.dataset.ready = "1";
        }
      }
      if (!create && !panePaste.dataset.ready) {
        fillPastePane(panePaste, cfg, opts.prefill || "");
        panePaste.dataset.ready = "1";
      }
    }

    tCreate.onclick = function () { showTab("create"); };
    tPaste.onclick = function () { showTab("paste"); };

    var rowClose = el("div", { className: "ho-row" });
    var bClose = el("button", { type: "button", className: "ho-muted", text: "ปิด" });
    bClose.onclick = function () { closeSheet(sheet); };
    rowClose.appendChild(bClose);
    card.appendChild(rowClose);

    sheet.appendChild(card);
    sheet.addEventListener("click", function (e) { if (e.target === sheet) closeSheet(sheet); });
    document.body.appendChild(sheet);
    showTab(opts.tab === "paste" ? "paste" : "create");
    return sheet;
  }

  function clearHash() {
    try {
      if (location.hash && /(^|[&#])h=/.test(location.hash)) {
        history.replaceState(null, "", location.pathname + location.search);
      }
      var u = new URL(location.href);
      if (u.searchParams.has("h")) {
        u.searchParams.delete("h");
        history.replaceState(null, "", u.pathname + (u.search || "") + (u.hash || ""));
      }
    } catch (e) {}
  }

  function readHashCode() {
    try {
      var u = new URL(location.href);
      if (u.searchParams.get("h")) return u.searchParams.get("h");
      var m = (location.hash || "").match(/(?:^|[&#])h=([^&]+)/);
      if (m) return decodeURIComponent(m[1]);
    } catch (e) {}
    return "";
  }

  function applyImport(cfg, raw, opts) {
    opts = opts || {};
    var parsed = parseCode(raw);
    if (!parsed.ok) return { error: parsed.error };

    var want = String(cfg.prefix).toUpperCase();
    var got = parsed.prefix;
    var compatible = got === want ||
      (FORMS[want] && FORMS[want].also && FORMS[want].also.indexOf(got) >= 0) ||
      (FORMS[got] && FORMS[got].also && FORMS[got].also.indexOf(want) >= 0);

    if (!compatible) {
      var meta = FORMS[got];
      if (!meta) return { error: "ไม่รู้จักรหัส " + got + "-" };
      var codeOnly = String(raw).trim();
      var hm = codeOnly.match(/[?#]h=([A-Za-z0-9\-_=]+)/);
      if (hm) codeOnly = decodeURIComponent(hm[1]);
      var mm = codeOnly.match(/^([A-Z]{2,4})-(.+)$/i);
      var fullCode = mm ? (mm[1].toUpperCase() + "-" + mm[2].replace(/\s+/g, "")) : (got + "-" + codeOnly.replace(/\s+/g, ""));
      var goUrl = formUrl(got, fullCode);
      var msg = "รหัสนี้เป็นของ「" + meta.label + "」ไม่ใช่ฟอร์มนี้\n\nต้องการเปิดหน้านั้นเพื่อนำเข้าหรือไม่?";
      if (confirm(msg)) location.href = goUrl;
      return { error: "รหัสเป็นของฟอร์มอื่น (" + got + "-)" };
    }

    if (typeof cfg.isEmpty === "function" && !cfg.isEmpty()) {
      if (!confirm("ฟอร์มมีข้อมูลอยู่แล้ว — ต้องการทับด้วยข้อมูลที่รับต่อหรือไม่?")) {
        return { error: "ยกเลิกการนำเข้า" };
      }
    }

    try {
      cfg.applyState(parsed.payload);
      clearHash();
      toast("รับข้อมูลต่อแล้ว ✓");
      return "ok";
    } catch (e) {
      return { error: "นำเข้าไม่สำเร็จ: " + (e && e.message ? e.message : e) };
    }
  }

  function mount(cfg) {
    if (!cfg || !cfg.prefix || !FORMS[cfg.prefix]) throw new Error("handoff: bad config");
    ensureCss();

    var top = document.querySelector(".topbar");
    if (top && !document.getElementById("handoffPill")) {
      var pill = el("button", { type: "button", className: "ho-pill", id: "handoffPill", text: "🔁 ส่งต่อข้อมูล", title: "ส่งต่อ / รับข้อมูลระหว่างผลัด" });
      pill.onclick = function () { openHandoffSheet(cfg, { tab: "create" }); };
      top.appendChild(pill);
    }

    // Auto hash import — confirm only; no panel unless error
    var pending = readHashCode();
    if (pending) {
      setTimeout(function () {
        var r = applyImport(cfg, pending, { fromHash: true });
        if (r && r.error && r.error.indexOf("ฟอร์มอื่น") < 0 && r.error.indexOf("ยกเลิก") < 0) {
          openHandoffSheet(cfg, {
            tab: "paste",
            prefill: pending.indexOf("-") >= 0 ? pending : (cfg.prefix + "-" + pending)
          });
        }
      }, 200);
    }

    return {
      makeCode: function () { return makeCode(cfg.prefix, cfg.getState()); },
      parseCode: parseCode,
      applyImport: function (raw) { return applyImport(cfg, raw, {}); },
      formUrl: function (code) { return formUrl(cfg.prefix, code); },
      openSheet: function (opts) { return openHandoffSheet(cfg, opts || {}); }
    };
  }

  global.ReportsHandoff = {
    VER: VER,
    FORMS: FORMS,
    mount: mount,
    makeCode: makeCode,
    parseCode: parseCode,
    formUrl: formUrl,
    qrDataUrl: qrDataUrl,
    qrCapacityHint: qrCapacityHint,
    siteOrigin: siteOrigin,
    encodePayload: encodePayload,
    decodePayload: decodePayload
  };
})(window);
