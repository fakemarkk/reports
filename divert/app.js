/* Divert Flight preview — offline single-file companion */
(function () {
  "use strict";

  const LS_KEY = "reports_divert_draft_v1";

  let _tplABuf = null, _tplBBuf = null;
  async function templateBuf(kind) {
    if (kind === "A") {
      if (!_tplABuf) {
        const res = await fetch("template-a-form.xlsx", { cache: "force-cache" });
        if (!res.ok) throw new Error("โหลด template A ไม่ได้ (" + res.status + ")");
        _tplABuf = await res.arrayBuffer();
      }
      return _tplABuf.slice(0);
    }
    if (!_tplBBuf) {
      const res = await fetch("template-b-table.xlsx", { cache: "force-cache" });
      if (!res.ok) throw new Error("โหลด template B ไม่ได้ (" + res.status + ")");
      _tplBBuf = await res.arrayBuffer();
    }
    return _tplBBuf.slice(0);
  }

  const M = ["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."];
  const WTH = ["อาทิตย์","จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์","เสาร์"];
  const FIXED_TO = "ผอก.ฝปข.ทสภ. และผู้บริหาร ฝปข.ทสภ.";
  const FIXED_RECIPIENT = "ศูนย์รักษาความปลอดภัย ทสภ. (SCC)";
  const FIXED_FROM = "งพบ.สบบ.ฝปข.ทสภ. (โทร. 24103)";
  const FIXED_SUBJECT = "แจ้งข้อมูลเที่ยวบิน Divert";
  const FORM_UNIT = "งพบ.สบบ.ฝปข.ทสภ.";

  const AIRPORTS = {
    BKK: "BANGKOK", DMK: "BANGKOK", UTP: "U-TAPAO", CNX: "CHIANG MAI", HKT: "PHUKET",
    KBV: "KRABI", HDY: "HAT YAI", CEI: "CHIANG RAI", UTH: "UDON THANI", KKC: "KHON KAEN",
    KUL: "KUALA LUMPUR", CKG: "CHONGQING", HKG: "HONG KONG", SIN: "SINGAPORE", DAC: "DHAKA",
    HAN: "HANOI", DOH: "DOHA", BOM: "MUMBAI", KMG: "KUNMING", PVG: "SHANGHAI", CAN: "GUANGZHOU",
    SGN: "HO CHI MINH", PNH: "PHNOM PENH", RGN: "YANGON", NRT: "TOKYO", ICN: "SEOUL",
    DXB: "DUBAI", AUH: "ABU DHABI", SYD: "SYDNEY", MEL: "MELBOURNE", DEL: "DELHI",
    MAA: "CHENNAI", CCU: "KOLKATA", TPE: "TAIPEI", MNL: "MANILA", JKT: "JAKARTA",
    CGK: "JAKARTA", PEN: "PENANG", LGK: "LANGKAWI", USM: "KOH SAMUI", URT: "SURAT THANI"
  };

  const STATUS = {
    unknown: "ยังไม่ทราบเวลาที่จะกลับมาลง ทสภ.",
    eta: "เวลาที่คาดว่าจะกลับมาลงที่ ทสภ.",
    actual: "กลับมาลงที่ ทสภ. เวลา"
  };

  const $ = (id) => document.getElementById(id);

  function showToast(msg) {
    const t = $("toast");
    t.textContent = msg;
    t.classList.add("show");
    setTimeout(() => t.classList.remove("show"), 1600);
  }

  function cleanTimeStr(s) {
    if (!s) return "";
    let x = String(s).replace(/[:,]/g, ".").replace(/[^0-9.]/g, "").slice(0, 5);
    if (!x) return "";
    let h, m;
    if (x.includes(".")) [h, m] = x.split(".");
    else {
      const d = x.replace(/\D/g, "");
      if (d.length <= 2) { h = d; m = "00"; }
      else { h = d.slice(0, d.length - 2); m = d.slice(-2); }
    }
    h = (h || "0").padStart(2, "0").slice(-2);
    m = (m || "0").padEnd(2, "0").slice(0, 2);
    return h + "." + m;
  }

  function up(s) { return String(s || "").trim().toUpperCase(); }

  function fmtFlight(s) {
    let x = up(s).replace(/\s+/g, " ");
    const m = x.match(/^([A-Z0-9]{2})\s*(\d{1,4}[A-Z]?)$/);
    if (m) return m[1] + " " + m[2];
    return x;
  }

  function fmtReg(s) {
    let x = up(s).replace(/\s+/g, "");
    if (/^HS[A-Z0-9]{3}$/.test(x)) return "HS-" + x.slice(2);
    if (/^B\d{4}$/.test(x)) return "B-" + x.slice(1);
    return x;
  }

  function airportLabel(name, code) {
    const n = up(name), c = up(code);
    if (n && c) {
      if (n.includes("(" + c + ")")) return n;
      return n + " (" + c + ")";
    }
    return n || c || "";
  }

  function divertAirportNameOnly(name, code) {
    let n = up(name).replace(/\s*\([A-Z0-9]{3}\)\s*$/, "").trim();
    if (!n && code) n = AIRPORTS[up(code)] || up(code);
    return n;
  }

  function titleish(s) {
    return String(s || "").toLowerCase().replace(/\b[a-z]/g, c => c.toUpperCase());
  }

  function blankFlight() {
    return {
      airline: "", flight: "", sibt: "",
      originName: "", originCode: "",
      destName: "BANGKOK", destCode: "BKK",
      reg: "", divertName: "U-TAPAO", divertCode: "UTP",
      status: "unknown", etaTime: "", actualTime: ""
    };
  }

  function defaultState() {
    return {
      kind: "initial",
      cause: "ติดสภาพอากาศที่ ทสภ.",
      formDate: "",
      formTime: "",
      formSigner: "",
      formShift: "3",
      name: "",
      unit: "งพบ.สบบ.ฝปข.ทสภ.",
      shift: "3",
      flights: [blankFlight()]
    };
  }

  let state = defaultState();

  function save() {
    try {
      const o = Object.assign({}, state);
      // never persist locked form header fields or shift dates
      delete o.recipient; delete o.fromLine; delete o.subject;
      delete o.dateRange; delete o.period; delete o.d1; delete o.toLine; delete o.causeCustom;
      localStorage.setItem(LS_KEY, JSON.stringify(o));
    } catch (e) {}
  }

  function load() {
    try {
      const raw = localStorage.getItem(LS_KEY);
      if (!raw) return;
      const o = JSON.parse(raw);
      if (!o || typeof o !== "object") return;
      const d = defaultState();
      state = Object.assign(d, o);
      // ignore stale locked fields
      delete state.recipient; delete state.fromLine; delete state.subject;
      delete state.dateRange; delete state.period; delete state.d1; delete state.toLine; delete state.causeCustom;
      if (!Array.isArray(state.flights) || !state.flights.length) state.flights = [blankFlight()];
      state.flights = state.flights.map(f => Object.assign(blankFlight(), f || {}));
      if (!state.cause) state.cause = "ติดสภาพอากาศที่ ทสภ.";
      // formDate: store ISO; convert legacy Thai text drafts if possible
      state.formDate = parseFormDateToIso(state.formDate);
      if (!state.formShift) state.formShift = "3";
      if (state.formSigner == null) state.formSigner = "";
    } catch (e) {}
  }

  function todayThaiParts(t) {
    const fmt = new Intl.DateTimeFormat("en-GB", {
      timeZone: "Asia/Bangkok", year: "numeric", month: "2-digit", day: "2-digit",
      hour: "2-digit", minute: "2-digit", hour12: false
    });
    const parts = Object.fromEntries(fmt.formatToParts(new Date(t == null ? Date.now() : t)).map(p => [p.type, p.value]));
    const y = Number(parts.year), m = Number(parts.month), d = Number(parts.day);
    return { y, m, d, time: parts.hour + "." + parts.minute, hour: Number(parts.hour), be2: String(y + 543).slice(-2) };
  }

  // Divert text samples: no space before BE year (14-15 ต.ค.69 / 4 พ.ย.69)
  function th(d) {
    return d.getDate() + " " + M[d.getMonth()] + String(d.getFullYear() + 543).slice(-2);
  }
  function thSpaced(d) {
    return d.getDate() + " " + M[d.getMonth()] + " " + String(d.getFullYear() + 543).slice(-2);
  }

  function datesFromInputs() {
    const s = ($("d1").value || "").trim();
    if (!s) return { range: "", day1: "", day2: "", iso: "" };
    const a = new Date(s + "T00:00:00");
    const b = new Date(a); b.setDate(a.getDate() + 1);
    const night = ($("period").value || "").startsWith("17");
    let range;
    if (!night) range = th(a);
    else if (a.getMonth() === b.getMonth()) {
      range = a.getDate() + "-" + b.getDate() + " " + M[b.getMonth()] + String(b.getFullYear() + 543).slice(-2);
    } else {
      range = th(a) + " - " + th(b);
    }
    return { range, day1: th(a), day2: th(b), iso: s, night };
  }

  function expandYear(dateRange) {
    // 14-15 ต.ค.69 -> 14-15 ต.ค. 2569
    return String(dateRange || "").replace(/([ก-๙.]+\s*)(\d{2})(?!\d)/, function (_, a, y) {
      return a.replace(/\s*$/, " ") + "25" + y;
    });
  }

  function causeText() {
    return (state.cause || "").trim() || "ติดสภาพอากาศที่ ทสภ.";
  }

  function positionLine() {
    return "หน.พนง.งานกะ " + state.unit + " (ผลัด " + state.shift + ")";
  }

  function formPositionLine() {
    return "หน.พนง.งานกะชุดที่.   " + (state.formShift || "3") + "   " + FORM_UNIT;
  }

  function todayIso() {
    const t = todayThaiParts();
    return t.y + "-" + String(t.m).padStart(2, "0") + "-" + String(t.d).padStart(2, "0");
  }

  /** Thai short form for Excel/PNG, e.g. "14 ต.ค. 69" */
  function formatFormDateThai(iso) {
    const s = String(iso || "").trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return "";
    const a = new Date(s + "T00:00:00");
    if (isNaN(a.getTime())) return "";
    return thSpaced(a);
  }

  /** Accept ISO or legacy Thai text from old drafts ("14 ต.ค. 69"). */
  function parseFormDateToIso(raw) {
    const s = String(raw || "").trim();
    if (!s) return "";
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    for (let mi = 0; mi < M.length; mi++) {
      const mon = M[mi].replace(/\./g, "\\.");
      const re = new RegExp("^(\\d{1,2})\\s+" + mon + "\\s*(\\d{2}|\\d{4})$");
      const m = s.match(re);
      if (!m) continue;
      const day = Number(m[1]);
      if (day < 1 || day > 31) continue;
      let be = Number(m[2]);
      if (m[2].length === 2) be = 2500 + be;
      const ce = be - 543;
      if (ce < 2000 || ce > 2100) continue;
      return ce + "-" + String(mi + 1).padStart(2, "0") + "-" + String(day).padStart(2, "0");
    }
    return "";
  }

  function autoFormDate() {
    return todayIso();
  }

  function paintDateLabel() {
    const s = $("d1").value;
    if (s) {
      const a = new Date(s + "T00:00:00");
      $("dateLabel").textContent = thSpaced(a);
      $("dateSub").textContent = "วัน" + WTH[a.getDay()];
    } else {
      $("dateLabel").textContent = "แตะเพื่อเลือกวันที่";
      $("dateSub").textContent = "";
    }
  }

  function paintFormDateLabel() {
    const s = $("formDate").value;
    if (s) {
      const a = new Date(s + "T00:00:00");
      $("formDateLabel").textContent = thSpaced(a);
      $("formDateSub").textContent = "วัน" + WTH[a.getDay()];
    } else {
      $("formDateLabel").textContent = "แตะเพื่อเลือกวันที่";
      $("formDateSub").textContent = "";
    }
  }

  function buildText() {
    const D = datesFromInputs();
    const kindLabel = state.kind === "initial" ? "รายงานข้อมูลเบื้องต้น" : "รายงานข้อมูลเพิ่มเติม";
    const cause = causeText();
    const lines = [];
    lines.push("เรียน " + FIXED_TO);
    lines.push("วันที่ " + (D.range || "-"));
    lines.push("ระหว่างเวลา " + ($("period").value || state.period || "17.00-08.00") + " น.");
    lines.push("");
    lines.push(kindLabel);
    lines.push("1. เที่ยวบิน Divert Flight เนื่องจาก " + cause);
    lines.push("");

    state.flights.forEach((f, i) => {
      const n = i + 2;
      const airline = up(f.airline);
      const flight = fmtFlight(f.flight);
      const sibt = cleanTimeStr(f.sibt);
      const oName = up(f.originName), oCode = up(f.originCode);
      const dName = up(f.destName) || "BANGKOK", dCode = up(f.destCode) || "BKK";
      const reg = fmtReg(f.reg);
      const divLabel = airportLabel(f.divertName, f.divertCode);
      const route = [oName, dName].filter(Boolean).join(" - ") +
        ([oCode, dCode].filter(Boolean).length ? " (" + [oCode, dCode].filter(Boolean).join("-") + ")" : "");

      lines.push(n + ". สายการบิน " + airline + " เที่ยวบิน " + flight);
      lines.push("   เวลาตามตารางการบิน (SIBT) " + (sibt || "-") + " น.");
      lines.push("   เส้นทางบิน " + route + (reg ? " ทะเบียนอากาศยาน " + reg : ""));
      lines.push("   " + n + ".1 ไปลงที่สนามบิน " + divLabel);
      if (f.status === "unknown") lines.push("   " + n + ".2 " + STATUS.unknown);
      else if (f.status === "eta") lines.push("   " + n + ".2 " + STATUS.eta + " " + (cleanTimeStr(f.etaTime) || "..") + " น.");
      else if (f.status === "actual") lines.push("   " + n + ".2 " + STATUS.actual + " " + (cleanTimeStr(f.actualTime) || "..") + " น.");
      if (i < state.flights.length - 1) lines.push("");
    });

    lines.push("");
    lines.push("จึงเรียนมาเพื่อทราบ");
    lines.push((state.name || "").trim() || "(ชื่อผู้รายงาน)");
    lines.push(positionLine());
    return lines.join("\n");
  }

  function esc(s) {
    return String(s || "").replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
  }
  function xmlEsc(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  function lookupCity(codeInput, nameInput) {
    const code = up(codeInput.value);
    codeInput.value = code;
    if (!code) return;
    const city = AIRPORTS[code];
    if (city && !nameInput.value.trim()) nameInput.value = city;
  }

  function renderFlights() {
    const box = $("flights");
    box.innerHTML = "";
    state.flights.forEach((f, idx) => {
      const card = document.createElement("div");
      card.className = "flight " + (idx % 2 === 0 ? "odd" : "even");
      card.dataset.i = String(idx);
      card.innerHTML =
        '<div class="flight-head"><strong>เที่ยวบินที่ ' + (idx + 1) + '</strong>' +
        '<button type="button" class="mini danger" data-act="del"' + (state.flights.length === 1 ? " disabled" : "") + '>ลบ</button></div>' +
        '<div class="grid">' +
        '<label class="full">สายการบิน<input data-k="airline" class="up" value="' + esc(f.airline) + '" placeholder="THAI AIRWAYS"></label>' +
        '<label>เที่ยวบิน<input data-k="flight" class="up" value="' + esc(f.flight) + '" placeholder="TG 242"></label>' +
        '<label>SIBT<input data-k="sibt" class="t" value="' + esc(f.sibt) + '" placeholder="11.00" inputmode="decimal"></label>' +
        '<label>ต้นทาง (เมือง)<input data-k="originName" class="up" value="' + esc(f.originName) + '" placeholder="KRABI"></label>' +
        '<label>รหัสต้นทาง<input data-k="originCode" class="up code" value="' + esc(f.originCode) + '" placeholder="KBV" maxlength="4"></label>' +
        '<label>ปลายทาง (เมือง)<input data-k="destName" class="up" value="' + esc(f.destName) + '" placeholder="BANGKOK"></label>' +
        '<label>รหัสปลายทาง<input data-k="destCode" class="up code" value="' + esc(f.destCode) + '" placeholder="BKK" maxlength="4"></label>' +
        '<label>ทะเบียนอากาศยาน<input data-k="reg" class="up" value="' + esc(f.reg) + '" placeholder="HS-TXA"></label>' +
        '<label>สนามบิน Divert<input data-k="divertName" class="up" value="' + esc(f.divertName) + '" placeholder="U-TAPAO"></label>' +
        '<label>รหัส Divert<input data-k="divertCode" class="up code" value="' + esc(f.divertCode) + '" placeholder="UTP" maxlength="4"></label>' +
        '<label class="full">สถานะกลับมาลง ทสภ.<select data-k="status">' +
        '<option value="unknown"' + (f.status === "unknown" ? " selected" : "") + '>1) ยังไม่ทราบเวลาที่จะกลับมาลง ทสภ.</option>' +
        '<option value="eta"' + (f.status === "eta" ? " selected" : "") + '>2) เวลาที่คาดว่าจะกลับมาลงที่ ทสภ. … น.</option>' +
        '<option value="actual"' + (f.status === "actual" ? " selected" : "") + '>3) กลับมาลงที่ ทสภ. เวลา … น.</option>' +
        '</select></label>' +
        '<label class="etaWrap' + (f.status === "eta" ? "" : " hidden") + '">เวลาคาดการณ์<input data-k="etaTime" class="t" value="' + esc(f.etaTime) + '" placeholder="18.30" inputmode="decimal"></label>' +
        '<label class="actWrap' + (f.status === "actual" ? "" : " hidden") + '">เวลาลงจริง<input data-k="actualTime" class="t" value="' + esc(f.actualTime) + '" placeholder="12.51" inputmode="decimal"></label>' +
        "</div>";
      box.appendChild(card);
    });
  }

  function syncFormFromState() {
    $("kind").value = state.kind;
    document.querySelectorAll("#kindSeg button").forEach(b => b.classList.toggle("on", b.dataset.v === state.kind));
    $("toLine").value = FIXED_TO;
    $("recipient").value = FIXED_RECIPIENT;
    $("fromLine").value = FIXED_FROM;
    $("subject").value = FIXED_SUBJECT;
    document.querySelectorAll("#periodSeg button").forEach(b => b.classList.toggle("on", b.dataset.v === $("period").value));
    $("cause").value = state.cause || "ติดสภาพอากาศที่ ทสภ.";
    $("formDate").value = state.formDate || "";
    paintFormDateLabel();
    $("formTime").value = state.formTime;
    $("formSigner").value = state.formSigner || "";
    $("formShift").value = state.formShift || "3";
    $("name").value = state.name;
    $("shift").value = state.shift;
    $("unit").value = state.unit;
    document.querySelectorAll("#unitSeg button").forEach(b => b.classList.toggle("on", b.dataset.v === state.unit));
    $("posPreview").textContent = positionLine();
    $("formPosPreview").textContent = "จะลงในฟอร์มเป็น: " + formPositionLine();
    paintDateLabel();
    renderFlights();
    refreshOut();
  }

  function readFormBasics() {
    state.kind = $("kind").value;
    state.cause = $("cause").value;
    state.formDate = $("formDate").value;
    state.formTime = $("formTime").value;
    state.formSigner = $("formSigner").value;
    state.formShift = $("formShift").value;
    state.name = $("name").value;
    state.shift = $("shift").value;
    state.unit = $("unit").value;
    $("toLine").value = FIXED_TO;
    $("recipient").value = FIXED_RECIPIENT;
    $("fromLine").value = FIXED_FROM;
    $("subject").value = FIXED_SUBJECT;
  }

  function readFlightCard(card) {
    const i = Number(card.dataset.i);
    const f = state.flights[i];
    if (!f) return;
    card.querySelectorAll("[data-k]").forEach(el => {
      let v = el.value;
      if (el.classList && el.classList.contains("up")) {
        const u = String(v || "").toUpperCase();
        if (el.value !== u) el.value = u;
        v = u.trim();
      }
      f[el.dataset.k] = v;
    });
  }

  function refreshOut() {
    const D = datesFromInputs();
    $("dateHint").textContent = D.range ? ("จะแสดงในรายงานเป็น: วันที่ " + D.range + " · ระหว่างเวลา " + $("period").value + " น.") : "";
    paintDateLabel();
    document.querySelectorAll("#periodSeg button").forEach(b => b.classList.toggle("on", b.dataset.v === $("period").value));
    document.querySelectorAll("#kindSeg button").forEach(b => b.classList.toggle("on", b.dataset.v === state.kind));
    document.querySelectorAll("#unitSeg button").forEach(b => b.classList.toggle("on", b.dataset.v === state.unit));
    $("posPreview").textContent = positionLine();
    $("formPosPreview").textContent = "จะลงในฟอร์มเป็น: " + formPositionLine();
    $("out").textContent = buildText();
  }

  function onAnyChange() {
    readFormBasics();
    document.querySelectorAll("#flights .flight").forEach(readFlightCard);
    document.querySelectorAll("#flights .flight").forEach(card => {
      const st = card.querySelector('[data-k="status"]').value;
      card.querySelector(".etaWrap").classList.toggle("hidden", st !== "eta");
      card.querySelector(".actWrap").classList.toggle("hidden", st !== "actual");
    });
    save();
    refreshOut();
  }

  // ---- shift auto like return-flight ----
  function shiftNow(t) {
    const b = new Date((t == null ? Date.now() : t) + 7 * 36e5);
    const h = b.getUTCHours();
    if (h < 8) b.setUTCDate(b.getUTCDate() - 1);
    return {
      d1: b.getUTCFullYear() + "-" + String(b.getUTCMonth() + 1).padStart(2, "0") + "-" + String(b.getUTCDate()).padStart(2, "0"),
      day: h >= 8 && h < 17
    };
  }
  const AUTO = {
    day: "08.00-17.00", night: "17.00-08.00", sk: "divert_shift", away: 36e5, hiddenAt: null,
    key() { const s = shiftNow(); return s.d1 + (s.day ? "D" : "N"); },
    set() {
      const s = shiftNow(), p = s.day ? this.day : this.night;
      const ch = $("d1").value !== s.d1 || $("period").value !== p;
      $("d1").value = s.d1; $("period").value = p; return ch;
    },
    read() { try { return JSON.parse(sessionStorage.getItem(this.sk) || "null"); } catch (e) { return null; } },
    touch() {
      setTimeout(() => {
        try { sessionStorage.setItem(this.sk, JSON.stringify({ k: this.key(), d1: $("d1").value, period: $("period").value })); } catch (e) {}
      }, 0);
    },
    manual() {
      const m = this.read();
      if (m && m.k === this.key() && (m.period === this.day || m.period === this.night)) return m;
      try { sessionStorage.removeItem(this.sk); } catch (e) {}
      return null;
    },
    apply() {
      const m = this.manual();
      if (m) { $("d1").value = m.d1; $("period").value = m.period; }
      else this.set();
    },
    hide() { if (this.hiddenAt === null) this.hiddenAt = Date.now(); },
    back() {
      const t = this.hiddenAt; this.hiddenAt = null;
      if (t === null || Date.now() - t <= this.away) return;
      const a = $("d1").value, b = $("period").value;
      this.apply();
      if ($("d1").value !== a || $("period").value !== b) refreshOut();
    },
    init() {
      this.apply();
      ["input", "change"].forEach(e => $("d1").addEventListener(e, () => { this.touch(); onAnyChange(); }));
      document.querySelectorAll("#periodSeg button").forEach(b => b.addEventListener("click", () => this.touch()));
      addEventListener("pagehide", () => this.hide());
      addEventListener("pageshow", e => { if (e && e.persisted) this.back(); });
      document.addEventListener("visibilitychange", () => {
        if (document.visibilityState === "hidden") this.hide(); else this.back();
      });
    }
  };

  function clearForm() {
    if (!confirm("ล้างฟอร์ม Divert ทั้งหมด? ข้อมูลที่ค้างไว้จะหาย")) return;
    state = defaultState();
    const t = todayThaiParts();
    state.formDate = autoFormDate();
    state.formTime = t.time;
    try { localStorage.removeItem(LS_KEY); } catch (e) {}
    AUTO.set();
    syncFormFromState();
    save();
    showToast("ล้างฟอร์มแล้ว");
  }

  async function copyText() {
    const txt = $("out").textContent;
    const done = () => showToast("คัดลอกแล้ว ✓");
    try {
      if (navigator.clipboard && window.isSecureContext) { await navigator.clipboard.writeText(txt); done(); return; }
    } catch (e) {}
    const ta = document.createElement("textarea");
    ta.value = txt; document.body.appendChild(ta); ta.select();
    document.execCommand("copy"); ta.remove(); done();
  }

  async function shareText() {
    const txt = $("out").textContent;
    if (navigator.share) {
      try { await navigator.share({ text: txt, title: "Divert Flight" }); return; }
      catch (e) { if (e && e.name === "AbortError") return; }
    }
    copyText();
  }

  // ---------- Excel via JSZip ----------

  function setCell(sheet, ref, opts) {
    const re = new RegExp('<c r="' + ref + '"([^>]*?)(\\/>|>([\\s\\S]*?)</c>)');
    const m = sheet.match(re);
    if (!m) return sheet;
    const attrs = m[1] || "";
    let sAttr = "";
    if (opts && opts.s != null && opts.s !== "") sAttr = ' s="' + opts.s + '"';
    else {
      const sm = attrs.match(/\ss="(\d+)"/);
      sAttr = sm ? ' s="' + sm[1] + '"' : "";
    }
    let repl;
    if (opts.empty) repl = '<c r="' + ref + '"' + sAttr + "/>";
    else if (opts.num != null && opts.num !== "") repl = '<c r="' + ref + '"' + sAttr + "><v>" + opts.num + "</v></c>";
    else {
      const t = xmlEsc(opts.str == null ? "" : opts.str);
      const space = /^\s|\s$/.test(String(opts.str || "")) ? ' xml:space="preserve"' : "";
      repl = '<c r="' + ref + '"' + sAttr + ' t="inlineStr"><is><t' + space + ">" + t + "</t></is></c>";
    }
    return sheet.replace(re, () => repl);
  }

  function downloadBlob(blob, filename) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
  }

  function flightRowsForExcel() {
    return state.flights.map((f, i) => {
      const oName = up(f.originName), dName = up(f.destName) || "BANGKOK";
      const route = [oName, dName].filter(Boolean).join("-");
      const airline = up(f.airline);
      return {
        seq: i + 1,
        airline: airline,
        airlineUp: airline,
        flight: fmtFlight(f.flight),
        flightCompact: fmtFlight(f.flight).replace(/\s+/g, ""),
        sibt: cleanTimeStr(f.sibt),
        routeTitle: route,   // UPPERCASE (was Title Case)
        routeUpper: route,
        divert: divertAirportNameOnly(f.divertName, f.divertCode), // already UPPER
        actual: f.status === "actual" ? cleanTimeStr(f.actualTime) : ""
      };
    });
  }

  function formMeta() {
    const D = datesFromInputs();
    const iso = parseFormDateToIso(state.formDate) || autoFormDate();
    const formDate = formatFormDateThai(iso) || formatFormDateThai(autoFormDate());
    const formTime = cleanTimeStr(state.formTime) || todayThaiParts().time;
    const cause = causeText();
    const rows = flightRowsForExcel();
    const n = rows.length;
    return {
      D, formDate, formTime, cause,
      causeLine: "สาเหตุเกิดจาก" + cause.replace(/^เนื่องจาก\s*/, ""),
      dateLine: "ในวันที่      " + expandYear(D.range || "-"),
      period: $("period").value || "17.00-08.00",
      countLine: "มีเที่ยวบิน DIVERT FLIGHT จำนวน   " + n + "   เที่ยวบิน ",
      signer: (state.formSigner || "").trim() || "XXXXXX",
      signLine: "ลงชื่อ   " + ((state.formSigner || "").trim() || "XXXXXX"),
      posLine: formPositionLine(),
      rows
    };
  }

  /** Insert extra flight rows after template row 22; shift signature/footer/merges down. */
  function expandExcelASheet(sheet, nFlights) {
    const extra = Math.max(0, nFlights - 10);
    const locs = { extra, sigRow: 24 + extra, posRow: 25 + extra, foot1: 45 + extra, foot2: 46 + extra, firstFlight: 13 };
    if (extra === 0) return { sheet, locs };

    const row13 = sheet.match(/<row r="13"[^>]*>[\s\S]*?<\/row>/);
    if (!row13) throw new Error("template flight row 13 missing");
    const tpl = row13[0];

    function shiftRowXml(full, r, attrs, body) {
      r = Number(r);
      if (r < 23) return full;
      const nr = r + extra;
      body = body.replace(/\br="([A-Z]+)\d+"/g, function (_, c) { return 'r="' + c + nr + '"'; });
      return '<row r="' + nr + '"' + attrs + ">" + body + "</row>";
    }
    sheet = sheet.replace(/<row r="(\d+)"([^>]*)>([\s\S]*?)<\/row>/g, shiftRowXml);

    sheet = sheet.replace(/<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\/>/g, function (full, c1, r1, c2, r2) {
      r1 = Number(r1); r2 = Number(r2);
      if (r1 < 23) return full;
      return '<mergeCell ref="' + c1 + (r1 + extra) + ":" + c2 + (r2 + extra) + '"/>';
    });

    let insertXml = "";
    let mergeXml = "";
    for (let i = 0; i < extra; i++) {
      const r = 23 + i;
      let row = tpl
        .replace(/<row r="13"/, '<row r="' + r + '"')
        .replace(/\br="([A-Z]+)13"/g, function (_, c) { return 'r="' + c + r + '"'; });
      row = row.replace(/<c r="A\d+"[^>]*\/>|<c r="A\d+"[^>]*>[\s\S]*?<\/c>/,
        '<c r="A' + r + '" s="10"><v>' + (11 + i) + "</v></c>");
      insertXml += row;
      mergeXml += '<mergeCell ref="B' + r + ":G" + r + '"/>'
        + '<mergeCell ref="H' + r + ":I" + r + '"/>'
        + '<mergeCell ref="K' + r + ":L" + r + '"/>';
    }
    sheet = sheet.replace(/(<row r="22"[^>]*>[\s\S]*?<\/row>)/, "$1" + insertXml);

    sheet = sheet.replace(/<mergeCells([^>]*) count="(\d+)"([^>]*)>/, function (full, a, count, b) {
      const n = Number(count) + extra * 3;
      return "<mergeCells" + a + ' count="' + n + '"' + b + ">" + mergeXml;
    });
    // If mergeCells has no attrs between, still ok. Ensure merges inserted after opening tag.
    if (mergeXml && !sheet.includes('ref="B23:G23"') && extra > 0) {
      // fallback if count attr pattern differed
      sheet = sheet.replace(/<mergeCells[^>]*>/, function (m) { return m + mergeXml; });
    }

    sheet = sheet.replace(/<dimension ref="A1:([A-Z]+)(\d+)"/, function (full, col, r) {
      return '<dimension ref="A1:' + col + (Number(r) + extra) + '"';
    });
    return { sheet, locs };
  }

  // Style ids from template row 13 (center+borders+font) — apply to every flight row
  const FLIGHT_STYLE = { A: "10", B: "27", H: "27", J: "11", K: "27", M: "12", N: "13" };
  // Empty numbered rows in template use A s="13"; keep seq readable: use A="10" like row 13 for consistency

  function normalizeFlightRowStyles(sheet, fromRow, toRow) {
    const S = FLIGHT_STYLE;
    const map = { A: S.A, B: S.B, H: S.H, J: S.J, K: S.K, M: S.M, N: S.N };
    for (let r = fromRow; r <= toRow; r++) {
      Object.keys(map).forEach(function (col) {
        const ref = col + r;
        const re = new RegExp('<c r="' + ref + '"([^>]*?)(\\/>|>([\\s\\S]*?)</c>)');
        sheet = sheet.replace(re, function (full, attrs, rest) {
          let a = (attrs || "").replace(/\ss="\d+"/, "");
          return '<c r="' + ref + '" s="' + map[col] + '"' + a + rest;
        });
      });
    }
    return sheet;
  }

  function fillFlightRow(sheet, r, row, seq) {
    const S = FLIGHT_STYLE;
    if (!row) {
      sheet = setCell(sheet, "A" + r, { num: seq, s: S.A });
      sheet = setCell(sheet, "B" + r, { empty: true, s: S.B });
      sheet = setCell(sheet, "H" + r, { empty: true, s: S.H });
      sheet = setCell(sheet, "J" + r, { empty: true, s: S.J });
      sheet = setCell(sheet, "K" + r, { empty: true, s: S.K });
      sheet = setCell(sheet, "M" + r, { empty: true, s: S.M });
      sheet = setCell(sheet, "N" + r, { empty: true, s: S.N });
      return sheet;
    }
    sheet = setCell(sheet, "A" + r, { num: row.seq, s: S.A });
    sheet = setCell(sheet, "B" + r, { str: row.airline, s: S.B });
    sheet = setCell(sheet, "H" + r, { str: row.flightCompact || row.flight, s: S.H });
    if (row.sibt) sheet = setCell(sheet, "J" + r, { str: row.sibt, s: S.J });
    else sheet = setCell(sheet, "J" + r, { empty: true, s: S.J });
    sheet = setCell(sheet, "K" + r, { str: row.routeTitle, s: S.K });
    sheet = setCell(sheet, "M" + r, { str: row.divert, s: S.M });
    if (row.actual) sheet = setCell(sheet, "N" + r, { str: row.actual, s: S.N });
    else sheet = setCell(sheet, "N" + r, { empty: true, s: S.N });
    return sheet;
  }


  const DOC_CODE = "VTBS-FM-AOD-CA-003";
  const DOC_REV = "Rev00 Effective 01Oct26";
  // A4 portrait pt; pageMargins top/bottom inches; pageSetup.scale 91%
  const A4_PT = 841.89;
  const MARGIN_IN = { t: 0.27, b: 0.17 };
  const PAGE_SCALE_PCT = 0.91;

  function sumRowHeightsPt(sheet, r1, r2) {
    // default row height 15pt if unspecified; template uses ~18.75-21.75 for body
    const hMap = {};
    sheet.replace(/<row r="(\d+)"([^>]*)>/g, function (full, r, attrs) {
      const rr = Number(r);
      if (rr < r1 || rr > r2) return full;
      const hm = attrs.match(/ht="([0-9.]+)"/);
      hMap[rr] = hm ? Number(hm[1]) : 15;
      return full;
    });
    let sum = 0;
    for (let r = r1; r <= r2; r++) sum += (hMap[r] != null ? hMap[r] : 15);
    return sum;
  }

  function compactExcelASpacers(sheet, L) {
    // After signature: 2 small blanks + several spacer rows + doc-code rows at sheet bottom.
    // pageSetup uses fitToHeight=1 so the whole sheet (incl. codes) stays on 1 page with codes at bottom.
    const pos = L.posRow;
    const rowMap = {};
    sheet.replace(/<row r="(\d+)"[^>]*>[\s\S]*?<\/row>/g, function (full, r) { rowMap[Number(r)] = full; return full; });
    if (!rowMap[pos]) return sheet;
    const blankSrc = rowMap[pos + 1] || rowMap[23] || rowMap[pos];
    const footSrc = rowMap[L.foot1] || rowMap[45] || blankSrc;
    function renum(full, nr, ht) {
      let out = full
        .replace(/<row r="\d+"/, '<row r="' + nr + '"')
        .replace(/\br="([A-Z]+)\d+"/g, function (_, c) { return 'r="' + c + nr + '"'; });
      if (ht != null) {
        if (/ht="[0-9.]+"/.test(out)) out = out.replace(/ht="[0-9.]+"/, 'ht="' + ht + '"');
        else out = out.replace(/<row r="\d+"/, '<row r="' + nr + '" ht="' + ht + '" customHeight="1"');
        if (!/customHeight=/.test(out)) out = out.replace(/<row r="\d+"([^>]*)/, '<row r="' + nr + '"$1 customHeight="1"');
      }
      return out;
    }
    function clearRowText(full) {
      return full
        .replace(/\st="s"/g, "")
        .replace(/\st="inlineStr"/g, "")
        .replace(/<v>[\s\S]*?<\/v>/g, "")
        .replace(/<is>[\s\S]*?<\/is>/g, "");
    }
    const keep = [];
    Object.keys(rowMap).map(Number).sort(function (a, b) { return a - b; }).forEach(function (r) {
      if (r <= pos) keep.push(rowMap[r]);
    });
    let rNext = pos + 1;
    keep.push(renum(clearRowText(blankSrc), rNext++, 10));
    keep.push(renum(clearRowText(blankSrc), rNext++, 10));
    // Enough spacer so codes sit at sheet bottom; fitToPage keeps 1 page (up to ~23 flights).
    const used = sumRowHeightsPt(sheet, 1, pos) + 20;
    const printablePt = A4_PT - (MARGIN_IN.t + MARGIN_IN.b) * 72;
    // Aim sheet height ≈ printable/scale so after fit the codes hug the bottom.
    let need = Math.max(40, printablePt / PAGE_SCALE_PCT - used - 27);
    if (need > 320) need = 320;
    const chunk = 20;
    const nSp = Math.max(2, Math.round(need / chunk));
    const each = need / nSp;
    for (let i = 0; i < nSp; i++) {
      keep.push(renum(clearRowText(blankSrc), rNext++, Number(each.toFixed(2))));
    }
    const f1 = rNext++, f2 = rNext++;
    const footHt = 13.5;
    keep.push(renum(clearRowText(footSrc), f1, footHt));
    keep.push(renum(clearRowText(footSrc), f2, footHt));
    L.foot1 = f1; L.foot2 = f2;
    sheet = sheet.replace(/<sheetData>[\s\S]*?<\/sheetData>/, "<sheetData>" + keep.join("") + "</sheetData>");
    const merges = [];
    (sheet.match(/<mergeCell ref="[^"]+"\/>/g) || []).forEach(function (mm) {
      const g = mm.match(/ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/);
      if (!g) return;
      if (Number(g[4]) < L.sigRow - 1) merges.push(mm);
    });
    merges.push('<mergeCell ref="L' + L.sigRow + ":N" + L.sigRow + '"/>');
    merges.push('<mergeCell ref="L' + L.posRow + ":N" + L.posRow + '"/>');
    merges.push('<mergeCell ref="M' + f1 + ":N" + f1 + '"/>');
    merges.push('<mergeCell ref="M' + f2 + ":N" + f2 + '"/>');
    sheet = sheet.replace(/<mergeCells[^>]*>[\s\S]*?<\/mergeCells>/, '<mergeCells count="' + merges.length + '">' + merges.join("") + "</mergeCells>");
    sheet = sheet.replace(/<dimension ref="A1:([A-Z]+)\d+"/, '<dimension ref="A1:$1' + f2 + '"');
    sheet = sheet.replace(/<pageMargins[^/]*\/>/,
      '<pageMargins left="0.21" right="0" top="0.27" bottom="0.17" header="0.3" footer="0.21"/>');
    sheet = sheet.replace(/<headerFooter[\s\S]*?<\/headerFooter>/g, "");
    // Fit sheet to 1 page tall so spacer+codes stay on page 1 at the bottom
    if (/<sheetPr[\s\S]*?\/>/.test(sheet) || /<sheetPr[\s\S]*?<\/sheetPr>/.test(sheet)) {
      sheet = sheet.replace(/<sheetPr[\s\S]*?(?:\/>|<\/sheetPr>)/, '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>');
    } else {
      sheet = sheet.replace(/<sheetData/, '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><sheetData');
    }
    sheet = sheet.replace(/<pageSetup[^/]*\/>/,
      '<pageSetup paperSize="9" fitToWidth="1" fitToHeight="1" orientation="portrait" r:id="rId1"/>');
    return sheet;
  }

  async function buildExcelABlob() {
    if (typeof JSZip === "undefined") throw new Error("JSZip missing");
    const zip = await JSZip.loadAsync(await templateBuf("A"));
    let sheet = await zip.file("xl/worksheets/sheet1.xml").async("string");
    const meta = formMeta();
    const n = meta.rows.length;
    const expanded = expandExcelASheet(sheet, n);
    sheet = expanded.sheet;
    const L = expanded.locs;
    const flightSlots = Math.max(10, n);
    sheet = normalizeFlightRowStyles(sheet, 13, 12 + flightSlots);

    sheet = setCell(sheet, "B3", { str: FIXED_RECIPIENT });
    sheet = setCell(sheet, "B4", { str: meta.formDate });
    sheet = setCell(sheet, "D4", { str: meta.formTime });
    sheet = setCell(sheet, "B5", { str: FIXED_FROM });
    sheet = setCell(sheet, "B6", { str: FIXED_SUBJECT + " " });
    sheet = setCell(sheet, "B7", { str: meta.causeLine + " " });
    sheet = setCell(sheet, "B8", { str: meta.dateLine });
    sheet = setCell(sheet, "G8", { str: meta.period });
    sheet = setCell(sheet, "A10", { str: meta.countLine });

    for (let i = 0; i < flightSlots; i++) {
      sheet = fillFlightRow(sheet, 13 + i, meta.rows[i], i + 1);
    }

    sheet = setCell(sheet, "L" + L.sigRow, { str: meta.signLine });
    sheet = setCell(sheet, "L" + L.posRow, { str: meta.posLine });
    // Compact huge empty spacer rows between signature and footer so print can fit one page
    // while keeping a small gap (2 rows) + footer text.
    sheet = compactExcelASpacers(sheet, L);
    sheet = setCell(sheet, "M" + L.foot1, { str: DOC_CODE });
    sheet = setCell(sheet, "M" + L.foot2, { str: DOC_REV });
    zip.file("xl/worksheets/sheet1.xml", sheet);
    const blob = await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    return { blob, meta, locs: L };
  }


  /** Required fields for Excel/PNG (screen markers + confirm). */
  function missingExportFields(needSigner) {
    readFormBasics();
    document.querySelectorAll("#flights .flight").forEach(readFlightCard);
    const out = [];
    const d1 = ($("d1").value || "").trim();
    if (!d1) out.push({ el: $("d1"), label: "วันที่เริ่มผลัด" });
    const fd = ($("formDate").value || "").trim();
    if (!fd) out.push({ el: $("formDate"), label: "วันที่บนแบบฟอร์ม" });
    const ft = ($("formTime").value || "").trim();
    if (!ft) out.push({ el: $("formTime"), label: "เวลาบนแบบฟอร์ม" });
    if (needSigner) {
      const sg = ($("formSigner").value || "").trim();
      if (!sg) out.push({ el: $("formSigner"), label: "ชื่อผู้ลงชื่อบนแบบฟอร์ม" });
    }
    const cards = document.querySelectorAll("#flights .flight");
    cards.forEach(function (card, i) {
      const n = "เที่ยวบินที่ " + (i + 1) + ": ";
      function need(k, label) {
        const el = card.querySelector('[data-k="' + k + '"]');
        if (!el || !(el.value || "").trim()) out.push({ el: el, label: n + label });
      }
      need("airline", "สายการบิน");
      need("flight", "เที่ยวบิน");
      need("sibt", "SIBT");
      need("originName", "ต้นทาง");
      need("destName", "ปลายทาง");
      need("divertName", "สนามบิน Divert");
    });
    return out;
  }

  function paintMiss(list) {
    document.querySelectorAll("input.miss,select.miss,textarea.miss").forEach(function (el) {
      el.classList.remove("miss");
      el.removeAttribute("aria-invalid");
    });
    document.querySelectorAll(".mh").forEach(function (hn) { hn.classList.add("hidden"); });
    const set = new Set(list.map(function (m) { return m.el; }).filter(Boolean));
    set.forEach(function (el) {
      el.classList.add("miss");
      el.setAttribute("aria-invalid", "true");
      const parent = el.parentNode;
      if (!parent) return;
      let hn = parent.querySelector(":scope > .mh");
      if (!hn) {
        hn = document.createElement("span");
        hn.className = "mh";
        hn.textContent = "⚠️ ยังไม่ได้กรอก";
        el.after(hn);
      }
      hn.classList.remove("hidden");
    });
  }

  function okToMakeExport(needSigner) {
    const miss = missingExportFields(!!needSigner);
    paintMiss(miss);
    const n = miss.length;
    return !n || confirm("ยังขาดข้อมูล " + n + " ช่อง ต้องการสร้างไฟล์ต่อหรือไม่?");
  }

  async function makeExcelA() {
    if (!okToMakeExport(true)) return;
    const { blob } = await buildExcelABlob();
    downloadBlob(blob, "divert-form.xlsx");
    showToast("ดาวน์โหลดแบบฟอร์มแล้ว");
  }

  async function makeExcelB() {
    if (!okToMakeExport(false)) return;
    if (typeof JSZip === "undefined") throw new Error("JSZip missing");
    const zip = await JSZip.loadAsync(await templateBuf("B"));
    let sheet = await zip.file("xl/worksheets/sheet1.xml").async("string");
    const rows = flightRowsForExcel();
    sheet = setCell(sheet, "A1", { str: "มีเที่ยวบิน DIVERT FLIGHT จำนวน   " + rows.length + "   เที่ยวบิน " });
    sheet = sheet.replace(/<row r="([4-9]|\d{2,})"[^>]*>[\s\S]*?<\/row>/g, "");
    let dataXml = "";
    rows.forEach((row, i) => {
      const r = 4 + i;
      const s = i === 0 ? 4 : 5;
      const aCell = i === 1
        ? '<c r="A' + r + '" s="3" t="inlineStr"><is><t>' + row.seq + "</t></is></c>"
        : '<c r="A' + r + '" s="' + s + '"><v>' + row.seq + "</v></c>";
      dataXml += '<row r="' + r + '" spans="1:6" ht="24.95" customHeight="1" x14ac:dyDescent="0.35">' +
        aCell +
        '<c r="B' + r + '" s="' + s + '" t="inlineStr"><is><t>' + xmlEsc(row.airlineUp) + "</t></is></c>" +
        '<c r="C' + r + '" s="' + s + '" t="inlineStr"><is><t>' + xmlEsc(row.flight) + "</t></is></c>" +
        '<c r="D' + r + '" s="' + s + '" t="inlineStr"><is><t>' + xmlEsc((row.sibt || "") + (row.sibt ? " น." : "")) + "</t></is></c>" +
        '<c r="E' + r + '" s="' + s + '" t="inlineStr"><is><t>' + xmlEsc(row.routeUpper) + "</t></is></c>" +
        '<c r="F' + r + '" s="' + s + '" t="inlineStr"><is><t>' + xmlEsc(row.divert) + "</t></is></c>" +
        "</row>";
    });
    sheet = sheet.replace(/<sheetData>[\s\S]*?<\/sheetData>/, (block) => {
      const head = [];
      for (let r = 1; r <= 3; r++) {
        const m = block.match(new RegExp('<row r="' + r + '"[\\s\\S]*?<\\/row>'));
        if (m) head.push(m[0]);
      }
      return "<sheetData>" + head.join("") + dataXml + "</sheetData>";
    });
    const last = 3 + Math.max(rows.length, 1);
    sheet = sheet.replace(/<dimension ref="[^"]*"/, '<dimension ref="A1:F' + last + '"');
    zip.file("xl/worksheets/sheet1.xml", sheet);
    const blob = await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    downloadBlob(blob, "divert-table.xlsx");
    showToast("ดาวน์โหลดตารางแล้ว");
  }

  // ---------- PNG form: faithful Template A sheet grid ----------
  function pngFileName() {
    const D = datesFromInputs();
    if (D.iso) return "Divert_" + D.iso.replace(/-/g, "") + ".png";
    return "Divert.png";
  }

  // Exact A..N widths from sheet1.xml <cols> (openpyxl misreads ranged cols)
  const XL_COLS = [4.625, 8.625, 5.625, 5.625, 3.625, 3.625, 4.625, 4.625, 4.625, 10.625, 12.625, 12.625, 11.25, 12.125];
  const XL_MDW = 7; // default font max digit width (px)
  // Excel width→px: trunc((width*mdw+5)/mdw*256)/256*mdw  (approx; use integer form)
  function xlWidthToPx(w) {
    return Math.floor(((256 * w + Math.floor(128 / XL_MDW)) / 256) * XL_MDW);
  }
  // Prefer the +5 form which matches LO print column span more closely
  function xlWidthToPxPad(w) {
    return Math.floor((w * XL_MDW + 5) / XL_MDW * 256) / 256 * XL_MDW;
  }
  const XL_COL_PX = XL_COLS.map(xlWidthToPxPad); // ~804
  const XL_SHEET_W = XL_COL_PX.reduce(function (a, b) { return a + b; }, 0);
  const A4_W = 794, A4_H = 1123; // CSS px @96dpi
  const MARGIN = { l: 0.21 * 96, r: 0, t: 0.27 * 96, b: 0.17 * 96 };
  const PRINT_W = A4_W - MARGIN.l - MARGIN.r;
  const PAGE_SCALE = 0.91; // pageSetup.scale
  // Uniform scale: pageSetup 91% (sheet already fits printable width at 100%)
  const SHEET_SCALE = PAGE_SCALE;
  const PT = 96 / 72;

  function rowHeightPx(pt) { return pt * PT; }

  function buildFormHTML(meta) {
    const n = meta.rows.length;
    const flightSlots = Math.max(10, n);
    // Compacted print layout (same as Excel export after compactExcelASpacers):
    // rows 1..12 header, 13..(12+slots) flights, spacer, sig, pos, blank, blank, foot, foot
    const flightsStart = 13;
    const flightsEnd = 12 + flightSlots;
    const spacerRow = flightsEnd + 1;
    const sigRow = spacerRow + 1;
    const posRow = sigRow + 1;
    const blank1 = posRow + 1;
    const blank2 = blank1 + 1;
    const totalRows = blank2;

    const heights = {};
    for (let r = 1; r <= totalRows; r++) heights[r] = 21;
    heights[2] = 19.5;
    heights[spacerRow] = 19.5;
    heights[posRow] = 21.75;
    // Keep a couple of post-signature spacer nuances from template (rows 27/28 → blank1/blank2)
    heights[blank1] = 20.25;
    heights[blank2] = 18.75;

    const colLine = XL_COL_PX.map(function (x) { return x.toFixed(2) + "px"; }).join(" ");
    const rowLine = [];
    let sheetH = 0;
    for (let r = 1; r <= totalRows; r++) {
      const px = rowHeightPx(heights[r]);
      rowLine.push(px + "px");
      sheetH += px;
    }
    const scaledW = XL_SHEET_W * SHEET_SCALE;
    const scaledH = sheetH * SHEET_SCALE;

    function cell(r1, c1, r2, c2, text, cls) {
      return '<div class="xlc ' + (cls || "") + '" style="grid-area:' + r1 + "/" + c1 + "/" + (r2 + 1) + "/" + (c2 + 1) + '">' +
        (text == null || text === "" ? "" : text) + "</div>";
    }

    const parts = [];
    parts.push(cell(1, 1, 1, 14, esc("แบบฟอร์มการจัดส่งข้อมูลให้ผู้บริหาร"), "xl-title"));
    parts.push(cell(3, 1, 3, 1, esc("ถึง"), "xl-lab"));
    parts.push(cell(3, 2, 3, 14, esc(FIXED_RECIPIENT), "xl-val"));
    parts.push(cell(4, 1, 4, 1, esc("วันที่"), "xl-lab"));
    parts.push(cell(4, 2, 4, 2, esc(meta.formDate), "xl-val xl-left"));
    parts.push(cell(4, 3, 4, 3, esc("เวลา"), "xl-lab xl-center"));
    parts.push(cell(4, 4, 4, 4, esc(meta.formTime), "xl-val xl-center"));
    parts.push(cell(4, 5, 4, 5, esc("น."), "xl-val xl-center"));
    parts.push(cell(5, 1, 5, 1, esc("จาก"), "xl-lab"));
    parts.push(cell(5, 2, 5, 5, esc(FIXED_FROM), "xl-val"));
    parts.push(cell(6, 1, 6, 1, esc("เรื่อง"), "xl-lab"));
    parts.push(cell(6, 2, 6, 4, esc(FIXED_SUBJECT), "xl-val"));
    parts.push(cell(7, 2, 7, 14, esc(meta.causeLine), "xl-val"));
    parts.push(cell(8, 2, 8, 4, esc(meta.dateLine), "xl-val xl-left"));
    parts.push(cell(8, 5, 8, 6, esc("ช่วงเวลา"), "xl-lab xl-center"));
    parts.push(cell(8, 7, 8, 8, esc(meta.period), "xl-val xl-center"));
    parts.push(cell(8, 9, 8, 9, esc("น."), "xl-val xl-center"));
    parts.push(cell(10, 1, 10, 14, esc(meta.countLine), "xl-count"));
    // Header 11-12 (matches merges A11:A12, B11:G12, H11:I12, J11, J12, K11:L12, M11:M12, N11, N12)
    parts.push(cell(11, 1, 12, 1, esc("ลำดับ"), "xl-th xl-t xl-l"));
    parts.push(cell(11, 2, 12, 7, esc("สายการบิน"), "xl-th xl-t"));
    parts.push(cell(11, 8, 12, 9, esc("เที่ยวบิน"), "xl-th xl-t"));
    parts.push(cell(11, 10, 12, 10, "เวลาลงตาม<br>ตารางการบิน", "xl-th xl-t xl-nh"));
    parts.push(cell(11, 11, 12, 12, esc("เส้นทางบิน"), "xl-th xl-t"));
    parts.push(cell(11, 13, 12, 13, esc("ทำการบินไปลง"), "xl-th xl-t"));
    parts.push(cell(11, 14, 12, 14, "เวลาที่ทำการบิน<br>ลง ทสภ.", "xl-th xl-t xl-nh"));
    for (let i = 0; i < flightSlots; i++) {
      const r = flightsStart + i;
      const row = meta.rows[i];
      parts.push(cell(r, 1, r, 1, String(i + 1), "xl-td xl-center xl-seq xl-l"));
      parts.push(cell(r, 2, r, 7, esc(row ? row.airline : ""), "xl-td xl-center"));
      parts.push(cell(r, 8, r, 9, esc(row ? (row.flightCompact || row.flight) : ""), "xl-td xl-center"));
      parts.push(cell(r, 10, r, 10, esc(row && row.sibt ? row.sibt : ""), "xl-td xl-center"));
      parts.push(cell(r, 11, r, 12, esc(row ? row.routeTitle : ""), "xl-td xl-center"));
      parts.push(cell(r, 13, r, 13, esc(row ? row.divert : ""), "xl-td xl-center"));
      parts.push(cell(r, 14, r, 14, esc(row && row.actual ? row.actual : ""), "xl-td xl-center"));
    }
    parts.push(cell(sigRow, 12, sigRow, 14, esc(meta.signLine), "xl-sig"));
    parts.push(cell(posRow, 12, posRow, 14, esc(meta.posLine), "xl-sig"));

    const needH = MARGIN.t + scaledH + MARGIN.b;
    const tall = needH > A4_H + 0.5;

    return (
      '<div class="df-a4' + (tall ? " df-tall" : "") + '" id="divertFormPage" data-sheet-h="' + sheetH +
      '" data-scaled-h="' + scaledH + '" data-scaled-w="' + scaledW + '">' +
      '<div class="xl-page" style="width:' + scaledW + "px;height:" + scaledH + 'px">' +
      '<div class="xl-sheet" style="width:' + XL_SHEET_W + "px;height:" + sheetH +
      "px;grid-template-columns:" + colLine + ";grid-template-rows:" + rowLine.join(" ") + '">' +
      parts.join("") +
      "</div></div>" +
      '<div class="df-docfoot">' + esc(DOC_CODE) + "<br>" + esc(DOC_REV) + "</div>" +
      "</div>"
    );
  }

  const FORM_CSS =
    ".df-a4{box-sizing:border-box;position:relative;width:" + A4_W + "px;min-height:" + A4_H + "px;height:" + A4_H + "px;background:#fff;color:#000;" +
    "padding:" + MARGIN.t + "px " + MARGIN.r + "px " + MARGIN.b + "px " + MARGIN.l + "px;overflow:hidden}" +
    ".df-a4.df-tall{height:auto;min-height:" + A4_H + "px;overflow:visible}" +
    ".df-a4 *{box-sizing:border-box}" +
    ".xl-page{overflow:hidden;position:relative}" +
    ".df-a4.df-tall .xl-page{overflow:visible}" +
    ".xl-sheet{display:grid;transform:scale(" + SHEET_SCALE + ");transform-origin:top left;" +
    "font-family:'TH SarabunPSK',Sarabun,'TH Sarabun New','Noto Sans Thai',Tahoma,sans-serif;" +
    "font-size:14pt;font-weight:700;color:#000;background:#fff}" +
    ".xlc{min-width:0;min-height:0;display:flex;align-items:center;padding:0 3px;overflow:hidden;line-height:1.1;white-space:pre}" +
    ".xl-title{justify-content:center;text-align:center;white-space:normal}" +
    ".xl-lab,.xl-val{font-size:14pt}" +
    ".xl-left{justify-content:flex-start;text-align:left}" +
    ".xl-center{justify-content:center;text-align:center}" +
    ".xl-count{justify-content:center;text-align:center;white-space:normal}" +
    ".xl-th,.xl-td{justify-content:center;text-align:center;background:#fff;" +
    "border:none;border-right:0.75px solid #000;border-bottom:0.75px solid #000;" +
    "white-space:normal;line-height:1.05;font-size:14pt;margin:0;padding:0 1px}" +
    ".xl-t{border-top:0.75px solid #000}" +
    ".xl-l{border-left:0.75px solid #000}" +
    ".xl-seq{font-size:16pt}" +
    ".xl-sig{justify-content:center;text-align:center;font-size:14pt;white-space:normal}" +
    ".df-docfoot{position:absolute;right:6px;bottom:" + MARGIN.b + "px;" +
    "text-align:right;font-size:12pt;font-weight:700;line-height:1.15;white-space:nowrap;" +
    "font-family:'TH SarabunPSK',Sarabun,'TH Sarabun New','Noto Sans Thai',Tahoma,sans-serif;color:#000;pointer-events:none;z-index:2}" +
    ".xl-nh{font-size:14pt;line-height:1.15;padding:0;flex-direction:column;justify-content:center;text-align:center}";

  let sheetUrls = [];
  function canShareFile(f) {
    try { return !!(navigator.canShare && navigator.share && navigator.canShare({ files: [f] })); }
    catch (e) { return false; }
  }

  function openSheet(blob, name) {
    const file = new File([blob], name, { type: "image/png" });
    const share = canShareFile(file);
    $("shTitle").textContent = "🖼️ รูปแบบฟอร์มพร้อมแล้ว";
    $("shName").textContent = name;
    sheetUrls.forEach(u => URL.revokeObjectURL(u));
    sheetUrls = [URL.createObjectURL(blob)];
    const box = $("shImgs");
    box.innerHTML = "";
    const im = document.createElement("img");
    im.src = sheetUrls[0]; im.alt = "รูปแบบฟอร์มผู้บริหาร";
    box.appendChild(im);
    $("shHint").textContent = share
      ? "กด “แชร์ / บันทึก” เพื่อส่งเข้าไลน์หรือบันทึกรูป หรือกดดาวน์โหลด"
      : "กดดาวน์โหลดเพื่อบันทึกรูป · แชร์ไฟล์ใช้ได้บนมือถือที่รองรับ";
    $("shShare").classList.remove("hidden");
    $("shShare").onclick = () => {
      if (share) {
        navigator.share({ files: [file], title: name }).catch(e => {
          if (e && e.name !== "AbortError") downloadBlob(blob, name);
        });
      } else {
        downloadBlob(blob, name);
        showToast("บันทึกรูปแล้ว (เครื่องนี้แชร์ไฟล์ไม่ได้)");
      }
    };
    $("shDl").onclick = () => downloadBlob(blob, name);
    $("sheet").classList.remove("hidden");
  }

  async function makeFormPng() {
    const btn = $("btnPng");
    if (btn.disabled) return;
    if (!okToMakeExport(true)) return;
    if (typeof html2canvas !== "function") { showToast("ไม่มี html2canvas"); return; }
    readFormBasics();
    document.querySelectorAll("#flights .flight").forEach(readFlightCard);
    btn.disabled = true;
    const old = btn.textContent;
    btn.textContent = "⏳ กำลังสร้างรูป…";
    let host = null;
    try {
      try { await document.fonts.ready; } catch (e) {}
      const meta = formMeta();
      host = document.createElement("div");
      host.style.cssText = "position:fixed;left:-10000px;top:0;z-index:-1;background:#fff;";
      const style = document.createElement("style");
      style.textContent = FORM_CSS;
      host.appendChild(style);
      const wrap = document.createElement("div");
      wrap.innerHTML = buildFormHTML(meta);
      host.appendChild(wrap);
      document.body.appendChild(host);
      const page = host.querySelector("#divertFormPage");
      // ensure min height for A4
      page.style.minHeight = "1123px";
      // Compensate CSS transform: set page/wrapper height to scaled sheet + margins
      const scaledH = parseFloat(page.getAttribute("data-scaled-h") || "0") || 0;
      const needH = Math.ceil(MARGIN.t + scaledH + MARGIN.b);
      const tall = needH > A4_H + 0.5;
      page.classList.toggle("df-tall", tall);
      page.style.height = tall ? (needH + "px") : (A4_H + "px");
      const xlPage = page.querySelector(".xl-page");
      if (xlPage) xlPage.style.height = Math.ceil(scaledH) + "px";
      const pageH = tall ? needH : A4_H;
      const canvas = await html2canvas(page, {
        scale: 2,
        backgroundColor: "#ffffff",
        useCORS: true,
        logging: false,
        width: A4_W,
        height: pageH,
        windowWidth: A4_W,
        windowHeight: pageH
      });
      let out = canvas;
      if (canvas.width === 1588 && canvas.height < 2246) {
        const c2 = document.createElement("canvas");
        c2.width = 1588; c2.height = 2246;
        const ctx = c2.getContext("2d");
        ctx.fillStyle = "#fff";
        ctx.fillRect(0, 0, 1588, 2246);
        ctx.drawImage(canvas, 0, 0);
        out = c2;
      }
      const blob = await new Promise(r => out.toBlob(r, "image/png"));
      const name = pngFileName();
      openSheet(blob, name);
      showToast("สร้างรูปแล้ว");
      return { blob, name, width: out.width, height: out.height };
    } catch (e) {
      console.error(e);
      showToast("สร้างรูปไม่สำเร็จ");
      throw e;
    } finally {
      if (host) host.remove();
      btn.disabled = false;
      btn.textContent = old;
    }
  }

  window.__divert = {
    getState: () => state,
    setState: (o) => {
      state = Object.assign(defaultState(), o || {});
      delete state.recipient; delete state.fromLine; delete state.subject;
      if (!state.flights || !state.flights.length) state.flights = [blankFlight()];
      state.flights = state.flights.map(f => {
        f = Object.assign(blankFlight(), f || {});
        ["airline", "flight", "originName", "originCode", "destName", "destCode", "reg", "divertName", "divertCode"].forEach(function (k) {
          if (f[k]) f[k] = up(f[k]);
        });
        if (f.flight) f.flight = fmtFlight(f.flight);
        if (f.reg) f.reg = fmtReg(f.reg);
        return f;
      });
      if (o && Object.prototype.hasOwnProperty.call(o, "formDate")) {
        state.formDate = parseFormDateToIso(o.formDate);
      } else {
        state.formDate = parseFormDateToIso(state.formDate);
      }
      if (o && o.d1) $("d1").value = o.d1;
      if (o && o.period) $("period").value = o.period;
      syncFormFromState();
      save();
      AUTO.touch();
    },
    buildText, save, load, blankFlight, defaultState, LS_KEY,
    cleanTimeStr, fmtFlight, fmtReg, datesFromInputs, formMeta, makeFormPng, buildExcelABlob, missingExportFields, paintMiss, okToMakeExport,
    formatFormDateThai, parseFormDateToIso, autoFormDate, todayIso, paintFormDateLabel,
    FIXED_RECIPIENT, FIXED_FROM, FIXED_SUBJECT, FIXED_TO
  };

  function bind() {
    document.querySelectorAll("#kindSeg button").forEach(b => b.addEventListener("click", () => {
      $("kind").value = b.dataset.v; onAnyChange();
    }));
    document.querySelectorAll("#periodSeg button").forEach(b => b.addEventListener("click", () => {
      $("period").value = b.dataset.v; onAnyChange();
    }));
    document.querySelectorAll("#unitSeg button").forEach(b => b.addEventListener("click", () => {
      $("unit").value = b.dataset.v; onAnyChange();
    }));

    ["cause", "formDate", "formTime", "formSigner", "formShift", "name", "shift"].forEach(id => {
      const el = $(id);
      if (!el) return;
      el.addEventListener("input", onAnyChange);
      el.addEventListener("change", onAnyChange);
    });

    $("flights").addEventListener("input", (e) => {
      const t = e.target;
      if (t && t.classList.contains("t")) {
        t.value = t.value.replace(/[:,]/g, ".").replace(/[^0-9.]/g, "").slice(0, 5);
      }
      if (t && t.classList.contains("up")) {
        const start = t.selectionStart, end = t.selectionEnd;
        const v = String(t.value || "").toUpperCase();
        if (t.value !== v) {
          t.value = v;
          try { if (start != null) t.setSelectionRange(start, end); } catch (err) {}
        }
      }
      if (t && t.dataset.k === "originCode") lookupCity(t, t.closest(".flight").querySelector('[data-k="originName"]'));
      if (t && t.dataset.k === "divertCode") lookupCity(t, t.closest(".flight").querySelector('[data-k="divertName"]'));
      if (t && t.dataset.k === "destCode") lookupCity(t, t.closest(".flight").querySelector('[data-k="destName"]'));
      onAnyChange();
    });
    $("flights").addEventListener("change", onAnyChange);
    $("flights").addEventListener("blur", (e) => {
      const t = e.target;
      if (t && t.classList.contains("t")) { t.value = cleanTimeStr(t.value); onAnyChange(); }
      if (t && t.dataset.k === "flight") { t.value = fmtFlight(t.value); onAnyChange(); }
      if (t && t.dataset.k === "reg") { t.value = fmtReg(t.value); onAnyChange(); }
      if (t && t.classList.contains("up")) {
        const v = up(t.value);
        if (t.value !== v) { t.value = v; onAnyChange(); }
      }
    }, true);

    $("flights").addEventListener("click", (e) => {
      const btn = e.target.closest("button[data-act]");
      if (!btn) return;
      const card = btn.closest(".flight");
      const i = Number(card.dataset.i);
      if (btn.dataset.act === "del") {
        if (state.flights.length <= 1) return;
        readFormBasics();
        document.querySelectorAll("#flights .flight").forEach(readFlightCard);
        state.flights.splice(i, 1);
        save(); renderFlights(); refreshOut();
      }
    });

    $("addFlight").addEventListener("click", () => {
      readFormBasics();
      document.querySelectorAll("#flights .flight").forEach(readFlightCard);
      state.flights.push(blankFlight());
      save(); renderFlights(); refreshOut();
    });

    $("btnClear").addEventListener("click", clearForm);
    $("btnCopy").addEventListener("click", copyText);
    $("btnShare").addEventListener("click", shareText);
    $("btnExcelA").addEventListener("click", () => makeExcelA().catch(err => { console.error(err); showToast("สร้าง Excel ไม่ได้"); }));
    $("btnExcelB").addEventListener("click", () => makeExcelB().catch(err => { console.error(err); showToast("สร้าง Excel ไม่ได้"); }));
    $("btnPng").addEventListener("click", () => makeFormPng().catch(() => {}));
    $("shClose").addEventListener("click", () => $("sheet").classList.add("hidden"));
    $("sheet").addEventListener("click", (e) => { if (e.target === $("sheet")) $("sheet").classList.add("hidden"); });

    document.querySelectorAll(".datebox input").forEach(function (i) {
      i.addEventListener("click", function () { try { if (i.showPicker) i.showPicker(); } catch (e) {} });
    });
  }

  // init
  load();
  if (!(state.formDate || "").trim()) state.formDate = autoFormDate();
  if (!(state.formTime || "").trim()) state.formTime = todayThaiParts().time;
  AUTO.set();
  bind();
  AUTO.init();
  syncFormFromState();
  save();
})();
