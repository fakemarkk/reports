/* Service worker — ทำให้เว็บรายงานใช้งานได้แม้ไม่มีเน็ต (ต้องเปิดตอนมีเน็ตอย่างน้อย 1 ครั้ง)
   ⚠️ ทุกครั้งที่ขึ้นเวอร์ชันใหม่ ให้แก้ VERSION ให้ตรงกับเลขเวอร์ชันในหน้าเมนู (index.html)
      และถ้าเพิ่มไฟล์/หน้าใหม่ ให้เพิ่มชื่อไฟล์ใน PRECACHE ด้วย
   กลยุทธ์:
   - หน้าเว็บ (HTML) และไฟล์ .js ของเว็บ: ดึงจากเน็ตก่อน (ได้ของใหม่เสมอเมื่อออนไลน์) ถ้าไม่มีเน็ต/ช้าเกินไปใช้ของที่เก็บไว้
   - ฟอนต์ ไลบรารี และไอคอน (ไฟล์ใหญ่ ไม่ค่อยเปลี่ยน): ใช้ของที่เก็บไว้ก่อน และอัปเดตเบื้องหลัง */
const VERSION = "1.8.1";
const CACHE = "reports-v" + VERSION;
const PRECACHE = [
  "./",
  "changelog/",
  "delay/", "delay/delay.js",
  "delay/word/", "delay/word/template.js", "delay/word/fonts/THSarabunPSK.ttf",
  "return-taxi/",
  "return-flight/",
  "return-flight/word/", "return-flight/word/rfword.js", "return-flight/word/template.js",
  "return-flight/word/fonts/THSarabunPSK-Bold.ttf",
  "return-flight/word/lib/html2canvas.min.js", "return-flight/word/lib/jszip.min.js",
  "aog/",
  "manifest.webmanifest",
  "icon-192.png", "icon-512.png", "apple-touch-icon.png"
];
const NET_TIMEOUT_MS = 4000;   // เน็ตช้าเกินนี้ → แสดงหน้าที่เก็บไว้ (และอัปเดตเก็บไว้เบื้องหลัง)

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE)
    .then(c => c.addAll(PRECACHE.map(u => new Request(u, {cache: "reload"})))));
  // ไม่ skipWaiting เอง: เวอร์ชันใหม่จะรอจนผู้ใช้กดปุ่ม "อัปเดต" บนแถบแจ้งเตือน (ไม่รีโหลดหน้าเองขณะกรอก)
});

self.addEventListener("message", e => {
  const d = e.data || {};
  if (d.type === "SKIP_WAITING") self.skipWaiting();
  if (d.type === "GET_VERSION" && e.ports && e.ports[0]) e.ports[0].postMessage(VERSION);
});

self.addEventListener("activate", e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k.startsWith("reports-") && k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

// ".../delay/index.html" กับ ".../delay/" ถือเป็นหน้าเดียวกัน
function pageKey(url) {
  const u = new URL(url); u.search = ""; u.hash = "";
  if (u.pathname.endsWith("/index.html")) u.pathname = u.pathname.slice(0, -10);
  return u.href;
}

async function networkFirst(req, isPage) {
  const cache = await caches.open(CACHE);
  const key = isPage ? pageKey(req.url) : req;
  const net = fetch(req).then(res => {
    if (res && res.ok && res.type === "basic" && !res.redirected) cache.put(key, res.clone());
    return res;
  });
  net.catch(() => {});                          // กัน error ค้างตอนออฟไลน์
  const timeout = new Promise(r => setTimeout(r, NET_TIMEOUT_MS));
  try {
    const res = await Promise.race([net, timeout.then(() => null)]);
    if (res) return res;
  } catch (err) { /* offline */ }
  const hit = await cache.match(key, {ignoreSearch: true});
  if (hit) return hit;
  return net;                                   // ไม่มีในเครื่อง → รอเน็ตต่อ (หรือ error)
}

async function staleWhileRevalidate(req, event) {
  const cache = await caches.open(CACHE);
  const hit = await cache.match(req, {ignoreSearch: true});
  const net = fetch(req).then(res => {
    if (res && (res.ok || res.type === "opaque")) cache.put(req, res.clone());
    return res;
  }).catch(() => null);
  if (hit) { event.waitUntil(net); return hit; }
  return (await net) || Response.error();
}

self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  const sameOrigin = url.origin === self.location.origin;
  if (sameOrigin && !url.href.startsWith(self.registration.scope)) return;   // นอกเว็บรายงาน ไม่ยุ่ง
  if (req.mode === "navigate" || (sameOrigin && /\.html?$|\/$/.test(url.pathname))) {
    e.respondWith(networkFirst(req, true));
  } else if (sameOrigin && /\.js$|\.webmanifest$/.test(url.pathname) && !/\/lib\//.test(url.pathname)) {
    e.respondWith(networkFirst(req, false));
  } else {
    // ฟอนต์ ไลบรารี ไอคอน และไฟล์จากเว็บภายนอก (CDN) ถ้ามีในอนาคต
    e.respondWith(staleWhileRevalidate(req, e));
  }
});
