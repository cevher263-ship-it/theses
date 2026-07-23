/*
 * YÖK Ulusal Tez Merkezi - Veri Kazıma Aracı (YENİ ARAYÜZ SÜRÜMÜ)  v1.2
 * ---------------------------------------------------------------------------
 * Orijinal araç: https://github.com/mytunca/theses (Muhammet Yunus Tunca, MIT)
 * YÖK Tez Merkezi'nin kart tabanlı yeni arayüzüne uyarlanmıştır.
 *
 * ÖZELLİKLER:
 *  - Metaveri indirme: Excel (.xlsx, filtre okları + istatistik sayfası), CSV, JSON
 *  - Kaynakça dışa aktarma: RIS ve BibTeX (Zotero / Mendeley / EndNote)
 *  - Tez metinleri (PDF) toplu indirme (anlamlı dosya adları, 500 MB'lık ZIP parçaları)
 *  - Biriktirme (IndexedDB): 2000 sınırını aşmak için birden çok aramayı tekrarsız biriktirme
 *  - Filtreleme: biriken listeyi yıl/tür/dil/üniversite/konu/PDF ölçütleriyle süzme
 *  - Yedekle / Geri yükle (JSON) ve 2000 sınırı uyarısı
 *
 * KULLANIM: Arama SONUÇ sayfasında (tezSorguSonucYeni.jsp) F12 > Console açıp
 *           bu dosyayı yapıştırın ya da jsdelivr yükleyici satırını çalıştırın.
 */
(function () {
  "use strict";
  if (window.__yokTezAraci__) { window.__yokTezAraci__.open(); return; }
  var BASE = location.origin + "/UlusalTezMerkezi/";

  /* ---------- CDN bağımlılıkları ---------- */
  function loadScript(src) {
    return new Promise(function (res, rej) {
      var s = document.createElement("script");
      s.src = src; s.onload = res; s.onerror = function () { rej(new Error("Yüklenemedi: " + src)); };
      document.head.appendChild(s);
    });
  }
  function ensureDeps() {
    var t = [];
    if (typeof window.XLSX === "undefined") t.push(loadScript("https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"));
    if (typeof window.JSZip === "undefined") t.push(loadScript("https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"));
    if (typeof window.saveAs === "undefined") t.push(loadScript("https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"));
    return Promise.all(t);
  }

  /* ---------- Yardımcılar ---------- */
  var _tmp = document.createElement("div");
  function stripHtml(html) {
    if (html == null) return "";
    _tmp.innerHTML = String(html);
    var t = _tmp.textContent || _tmp.innerText || "";
    return t.replace(/￾/g, " ").replace(/ /g, " ").replace(/[ \t]+\n/g, "\n").trim();
  }
  function excelSafe(s) { s = s == null ? "" : String(s); return s.length > 32000 ? s.slice(0, 32000) + " …[kesildi]" : s; }
  function cleanUni(yer) { return String(yer || "").split(" / ")[0].replace(/[\s\/]+$/, "").trim(); }
  function stamp() { return new Date().toISOString().slice(0, 19).replace(/[T:]/g, "-"); }

  function parseCountWarning() {
    var m = document.body.innerText.match(/Arama sonucunda\s*([\d.]+)\s*kayıt bulundu\.\s*([\d.]+)\s*tanesi/i);
    if (!m) return null;
    return { total: +m[1].replace(/\./g, ""), shown: +m[2].replace(/\./g, "") };
  }

  /* ---------- Sonuç kartlarından tez listesi ---------- */
  function collectTheses() {
    var RD = (typeof referenceData !== "undefined") ? referenceData : (window.referenceData || {});
    return Array.from(document.querySelectorAll(".result-card")).map(function (card) {
      var idx = card.getAttribute("data-index");
      var meta = (RD[idx] || {}).meta || {};
      var titleEl = card.querySelector(".card-title");
      var engEl = titleEl ? titleEl.nextElementSibling : null;
      var noMatch = card.textContent.match(/Tez No:\s*(\d+)/);
      return {
        index: idx, kayitNo: card.getAttribute("data-kayitno"), tezNo: card.getAttribute("data-tezno"),
        tezNoDuz: noMatch ? noMatch[1] : "",
        baslikTR: stripHtml(meta.title) || stripHtml(titleEl && titleEl.innerHTML),
        baslikEN: stripHtml(engEl && engEl.innerHTML),
        yazar: stripHtml(meta.author), yil: stripHtml(meta.year), tur: stripHtml(meta.type),
        dil: stripHtml(meta.lang), konu: stripHtml(meta.subject), yerKisa: stripHtml(meta.yer)
      };
    });
  }

  /* ---------- Ağ ---------- */
  function fetchDetay(t) {
    return fetch(BASE + "tezBilgiDetay.jsp?kayitNo=" + encodeURIComponent(t.kayitNo) + "&tezNo=" + encodeURIComponent(t.tezNo),
      { headers: { "X-Requested-With": "XMLHttpRequest" }, credentials: "include" })
      .then(function (r) { return r.text(); })
      .then(function (x) { try { return JSON.parse(x.trim()); } catch (e) { return {}; } });
  }
  function fetchPdfLink(t) {
    return fetch(BASE + "getTezPdf.jsp?kayitNo=" + encodeURIComponent(t.kayitNo) + "&tezNo=" + encodeURIComponent(t.tezNo),
      { headers: { "X-Requested-With": "XMLHttpRequest" }, credentials: "include" })
      .then(function (r) { return r.text(); })
      .then(function (h) { var m = h.match(/href=['"]([^'"]*TezGoster[^'"]*)['"]/i); if (!m) return null; return m[1].indexOf("http") === 0 ? m[1] : BASE + m[1].replace(/^\//, ""); })
      .catch(function () { return null; });
  }
  function runPool(items, worker, concurrency, onProgress) {
    return new Promise(function (resolve) {
      var i = 0, done = 0, active = 0, n = items.length;
      if (n === 0) return resolve();
      function next() {
        while (active < concurrency && i < n) {
          (function (it) { active++; Promise.resolve(worker(it)).catch(function () {}).then(function () { active--; done++; if (onProgress) onProgress(done, n); if (done === n) resolve(); else next(); }); })(items[i++]);
        }
      }
      next();
    });
  }
  function buildRow(t, detay, pdfLink) {
    return {
      "Tez No": t.tezNoDuz, "PDF İndirme Linki": pdfLink || "",
      "Tez Adı (Orijinal)": t.baslikTR, "Tez Adı (Çeviri)": t.baslikEN,
      "Yazar": t.yazar, "Danışman": stripHtml(detay.danisman).replace(/^Danışman:\s*/i, ""),
      "Üniversite / Yer Bilgisi": stripHtml(detay.yer) || t.yerKisa, "Konu": t.konu,
      "Dizin (Anahtar Kelimeler)": [stripHtml(detay.anahtarKelimeTr), stripHtml(detay.anahtarKelimeEn)].filter(Boolean).join(" | "),
      "Tür": t.tur, "Dil": t.dil, "Yıl": t.yil,
      "Özet (Türkçe)": excelSafe(stripHtml(detay.trOzet)), "Özet (İngilizce)": excelSafe(stripHtml(detay.enOzet)),
      "kayitNo": t.kayitNo, "tezNo (kodlu)": t.tezNo
    };
  }
  function fetchAllMetadata(theses, onProgress) {
    var rows = new Array(theses.length);
    return runPool(theses, function (t) {
      return Promise.all([fetchDetay(t), fetchPdfLink(t)]).then(function (res) { rows[theses.indexOf(t)] = buildRow(t, res[0] || {}, res[1]); });
    }, 12, onProgress).then(function () { return rows.filter(Boolean); });
  }

  /* ---------- Çıktı biçimleri ---------- */
  var COLUMN_ORDER = ["Tez Adı (Orijinal)", "Yazar", "Tür", "Yıl", "Konu", "Üniversite / Yer Bilgisi", "Danışman", "Dil", "Dizin (Anahtar Kelimeler)", "Tez Adı (Çeviri)", "Tez No", "PDF İndirme Linki", "Özet (Türkçe)", "Özet (İngilizce)", "kayitNo", "tezNo (kodlu)"];
  var COLUMN_WIDTHS = [55, 22, 16, 7, 22, 40, 24, 10, 30, 55, 10, 32, 60, 60, 24, 24];

  function cleanRows(rows) { return rows.map(function (r) { var c = Object.assign({}, r); delete c._key; return c; }); }

  function buildMainSheet(rows) {
    var ws = XLSX.utils.json_to_sheet(rows, { header: COLUMN_ORDER });
    ws["!cols"] = COLUMN_WIDTHS.map(function (w) { return { wch: w }; });
    if (ws["!ref"]) ws["!autofilter"] = { ref: ws["!ref"] };
    return ws;
  }

  // İstatistik sayfası: yıl/tür/dil/konu/üniversiteye göre adet + metin çubuk
  function buildStatsSheet(rows) {
    function tally(getter) { var m = {}; rows.forEach(function (r) { var v = (getter(r) || "—").toString().trim() || "—"; m[v] = (m[v] || 0) + 1; }); return Object.entries(m).sort(function (a, b) { return b[1] - a[1]; }); }
    function bar(n, max) { var w = max ? Math.round(n / max * 30) : 0; return "█".repeat(w); }
    var aoa = [["YÖK TEZ MERKEZİ — İSTATİSTİK ÖZETİ"], ["Toplam tez", rows.length], [""]];
    function section(title, pairs, limit) {
      aoa.push([title]); aoa.push(["Değer", "Adet", ""]);
      var max = pairs.length ? pairs[0][1] : 0;
      pairs.slice(0, limit || pairs.length).forEach(function (p) { aoa.push([p[0], p[1], bar(p[1], max)]); });
      if (limit && pairs.length > limit) aoa.push(["… (+" + (pairs.length - limit) + " diğer)", "", ""]);
      aoa.push([""]);
    }
    section("YILA GÖRE", tally(function (r) { return r["Yıl"]; }));
    section("TÜRE GÖRE", tally(function (r) { return r["Tür"]; }));
    section("DİLE GÖRE", tally(function (r) { return r["Dil"]; }));
    section("KONUYA GÖRE (ilk 25)", tally(function (r) { return r["Konu"]; }), 25);
    section("ÜNİVERSİTEYE GÖRE (ilk 25)", tally(function (r) { return cleanUni(r["Üniversite / Yer Bilgisi"]); }), 25);
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 45 }, { wch: 10 }, { wch: 34 }];
    return ws;
  }

  function exportExcel(rows, prefix) {
    var clean = cleanRows(rows);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, buildMainSheet(clean), "Tezler");
    XLSX.utils.book_append_sheet(wb, buildStatsSheet(clean), "İstatistik");
    XLSX.writeFile(wb, (prefix || "Tez_Metaverileri") + "_" + stamp() + ".xlsx");
  }
  function saveText(text, filename, mime) {
    saveAs(new Blob(["﻿" + text], { type: (mime || "text/plain") + ";charset=utf-8" }), filename);
  }
  function exportCSV(rows, prefix) {
    var csv = XLSX.utils.sheet_to_csv(buildMainSheet(cleanRows(rows)), { FS: ";" });
    saveText(csv, (prefix || "Tez_Metaverileri") + "_" + stamp() + ".csv", "text/csv");
  }
  function exportJSON(rows, prefix) {
    saveText(JSON.stringify(cleanRows(rows), null, 1), (prefix || "Tez_Metaverileri") + "_" + stamp() + ".json", "application/json");
  }

  /* ---------- Kaynakça (RIS / BibTeX) ---------- */
  function risEsc(s) { return String(s || "").replace(/\r?\n/g, " ").trim(); }
  function toRIS(rows) {
    return rows.map(function (r) {
      var L = ["TY  - THES", "TI  - " + risEsc(r["Tez Adı (Orijinal)"])];
      if (r["Tez Adı (Çeviri)"]) L.push("T2  - " + risEsc(r["Tez Adı (Çeviri)"]));
      if (r["Yazar"]) L.push("AU  - " + risEsc(r["Yazar"]));
      if (r["Danışman"]) L.push("A2  - " + risEsc(r["Danışman"]));
      if (r["Yıl"]) L.push("PY  - " + risEsc(r["Yıl"]));
      L.push("PB  - " + risEsc(cleanUni(r["Üniversite / Yer Bilgisi"])));
      if (r["Tür"]) L.push("M3  - " + risEsc(r["Tür"]));
      if (r["Dizin (Anahtar Kelimeler)"]) risEsc(r["Dizin (Anahtar Kelimeler)"]).split(/\s*\|\s*|;\s*/).filter(Boolean).forEach(function (k) { L.push("KW  - " + k); });
      if (r["Özet (Türkçe)"]) L.push("AB  - " + risEsc(r["Özet (Türkçe)"]));
      if (r["Özet (İngilizce)"]) L.push("N2  - " + risEsc(r["Özet (İngilizce)"]));
      if (r["PDF İndirme Linki"]) L.push("UR  - " + risEsc(r["PDF İndirme Linki"]));
      if (r["Tez No"]) L.push("ID  - " + risEsc(r["Tez No"]));
      L.push("ER  - ");
      return L.join("\n");
    }).join("\n\n");
  }
  function texEsc(s) { return String(s || "").replace(/([\\{}%&$#_])/g, "\\$1").replace(/\r?\n/g, " ").trim(); }
  function bibType(t) { return /y[üu]ksek|master/i.test(t) ? "mastersthesis" : "phdthesis"; }
  function bibKey(r, used) {
    var base = ((r["Yazar"] || "yazar").split(/\s+/).pop() + (r["Yıl"] || "")).toLowerCase().replace(/[^a-z0-9]/g, "") || "tez";
    var k = base, i = 1; while (used[k]) k = base + String.fromCharCode(96 + (++i));
    used[k] = 1; return k;
  }
  function toBibTeX(rows) {
    var used = {};
    return rows.map(function (r) {
      var f = [["title", r["Tez Adı (Orijinal)"]], ["author", r["Yazar"]], ["year", r["Yıl"]],
        ["school", cleanUni(r["Üniversite / Yer Bilgisi"])], ["type", r["Tür"]],
        ["keywords", (r["Dizin (Anahtar Kelimeler)"] || "").replace(/\|/g, ",")],
        ["abstract", r["Özet (Türkçe)"]], ["language", r["Dil"]],
        ["note", r["Tez No"] ? "Tez No: " + r["Tez No"] : ""], ["url", r["PDF İndirme Linki"]]];
      var body = f.filter(function (x) { return x[1]; }).map(function (x) { return "  " + x[0] + " = {" + texEsc(x[1]) + "}"; }).join(",\n");
      return "@" + bibType(r["Tür"]) + "{" + bibKey(r, used) + ",\n" + body + "\n}";
    }).join("\n\n");
  }

  /* ---------- Biçim seçimine göre dışa aktar ---------- */
  function exportData(rows, prefix, format) {
    if (!rows.length) { alert("Dışa aktarılacak tez yok."); return; }
    if (format === "csv") exportCSV(rows, prefix);
    else if (format === "json") exportJSON(rows, prefix);
    else if (format === "ris") saveText(toRIS(cleanRows(rows)), (prefix || "Tez") + "_kaynakca_" + stamp() + ".ris", "application/x-research-info-systems");
    else if (format === "bib") saveText(toBibTeX(cleanRows(rows)), (prefix || "Tez") + "_kaynakca_" + stamp() + ".bib", "application/x-bibtex");
    else exportExcel(rows, prefix);
  }

  /* ---------- PDF metinleri (anlamlı dosya adları) ---------- */
  function sanitizeName(s) { return String(s || "").replace(/[\\\/:*?"<>|\n\r\t]+/g, " ").replace(/\s+/g, " ").trim().slice(0, 80); }
  function pdfName(r, used) {
    var base = [r["Yıl"], (r["Yazar"] || "").split(/\s+/).slice(0, 3).join(" "), r["Tez Adı (Orijinal)"]].filter(Boolean).join("_");
    base = sanitizeName(base) || ("tez_" + (r["Tez No"] || "x"));
    var name = base + ".pdf", i = 1;
    while (used[name]) name = base + "_" + (r["Tez No"] || ++i) + ".pdf";
    used[name] = 1; return name;
  }
  function downloadTexts(rows, onProgress, onInfo) {
    var withPdf = rows.filter(function (r) { return r["PDF İndirme Linki"]; });
    if (withPdf.length === 0) { alert("İndirilebilir (erişime açık) PDF bulunamadı."); return Promise.resolve(); }
    var zip = new JSZip(), chunkIndex = 0, chunkSize = 0, total = 0, count = 0, used = {};
    var maxChunk = 500 * 1024 * 1024, seq = Promise.resolve();
    withPdf.forEach(function (r) {
      seq = seq.then(function () {
        return fetch(r["PDF İndirme Linki"], { credentials: "include" })
          .then(function (resp) { if (!resp.ok) throw new Error("HTTP " + resp.status); return resp.blob(); })
          .then(function (blob) {
            zip.file(pdfName(r, used), blob); chunkSize += blob.size; total += blob.size; count++;
            onProgress(Math.round((100 * count) / withPdf.length));
            onInfo(count + " / " + withPdf.length + " tez metni indirildi (" + (total / (1024 * 1024)).toFixed(1) + " MB).\nPDF'ler 500 MB'lık ZIP parçaları hâlinde kaydedilir.");
            if (chunkSize >= maxChunk || count === withPdf.length) {
              return zip.generateAsync({ type: "blob" }).then(function (c) { chunkIndex++; saveAs(c, "Tez_Metinleri_Part_" + chunkIndex + ".zip"); zip = new JSZip(); chunkSize = 0; });
            }
          }).catch(function (e) { console.warn("PDF indirilemedi (" + r["Tez No"] + "):", e.message); });
      });
    });
    return seq;
  }

  /* ---------- Biriktirme deposu (IndexedDB) ---------- */
  var DB_NAME = "yokTezBirikim", STORE = "tezler";
  function dbOpen() { return new Promise(function (res, rej) { var q = indexedDB.open(DB_NAME, 1); q.onupgradeneeded = function () { var db = q.result; if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE, { keyPath: "_key" }); }; q.onsuccess = function () { res(q.result); }; q.onerror = function () { rej(q.error); }; }); }
  function dbPutRows(rows) { return dbOpen().then(function (db) { return new Promise(function (res, rej) { var tx = db.transaction(STORE, "readwrite"), st = tx.objectStore(STORE); rows.forEach(function (r) { st.put(Object.assign({ _key: (r["Tez No"] || r.kayitNo || "").toString() }, r)); }); tx.oncomplete = function () { db.close(); res(rows.length); }; tx.onerror = function () { db.close(); rej(tx.error); }; }); }); }
  function dbGetAll() { return dbOpen().then(function (db) { return new Promise(function (res, rej) { var q = db.transaction(STORE, "readonly").objectStore(STORE).getAll(); q.onsuccess = function () { db.close(); res(q.result || []); }; q.onerror = function () { db.close(); rej(q.error); }; }); }); }
  function dbCount() { return dbOpen().then(function (db) { return new Promise(function (res, rej) { var q = db.transaction(STORE, "readonly").objectStore(STORE).count(); q.onsuccess = function () { db.close(); res(q.result); }; q.onerror = function () { db.close(); rej(q.error); }; }); }); }
  function dbClear() { return dbOpen().then(function (db) { return new Promise(function (res, rej) { var tx = db.transaction(STORE, "readwrite"); tx.objectStore(STORE).clear(); tx.oncomplete = function () { db.close(); res(); }; tx.onerror = function () { db.close(); rej(tx.error); }; }); }); }

  /* ---------- Filtre ---------- */
  function applyFilter(rows, f) {
    return rows.filter(function (r) {
      var y = parseInt(r["Yıl"], 10);
      if (f.yil1 && (!y || y < f.yil1)) return false;
      if (f.yil2 && (!y || y > f.yil2)) return false;
      if (f.tur && (r["Tür"] || "").toLocaleLowerCase("tr").indexOf(f.tur) === -1) return false;
      if (f.dil && (r["Dil"] || "").toLocaleLowerCase("tr").indexOf(f.dil) === -1) return false;
      if (f.uni && (r["Üniversite / Yer Bilgisi"] || "").toLocaleLowerCase("tr").indexOf(f.uni) === -1) return false;
      if (f.konu && (r["Konu"] || "").toLocaleLowerCase("tr").indexOf(f.konu) === -1) return false;
      if (f.onlyPdf && !r["PDF İndirme Linki"]) return false;
      return true;
    });
  }

  /* ===================================================================== */
  /*  Arayüz                                                               */
  /* ===================================================================== */
  var theses = collectTheses();
  var warn = parseCountWarning();

  function buildUI() {
    var css = document.createElement("style");
    css.textContent =
      "#ytz-overlay{position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:2147483646;}" +
      "#ytz-panel{position:fixed;top:16px;right:16px;width:440px;max-width:94vw;max-height:94vh;overflow:auto;background:#fff;border-radius:10px;box-shadow:0 12px 40px rgba(0,0,0,.35);z-index:2147483647;font-family:Segoe UI,Arial,sans-serif;color:#222;}" +
      "#ytz-panel .ytz-head{background:#1f883d;color:#fff;padding:12px 16px;font-size:16px;font-weight:600;display:flex;justify-content:space-between;align-items:center;position:sticky;top:0;z-index:2;}" +
      "#ytz-panel .ytz-x{cursor:pointer;font-size:22px;}" +
      "#ytz-panel .ytz-body{padding:12px 14px;}" +
      "#ytz-panel .ytz-warn{background:#fff3cd;border:1px solid #ffe08a;color:#664d03;padding:8px 10px;border-radius:6px;font-size:12px;margin-bottom:10px;}" +
      "#ytz-panel .ytz-sec{border:1px solid #e3e3e3;border-radius:8px;padding:10px 12px;margin-bottom:11px;}" +
      "#ytz-panel .ytz-sec h4{margin:0 0 8px;font-size:13px;color:#1f883d;}" +
      "#ytz-panel p{margin:0 0 8px;font-size:12.5px;line-height:1.5;white-space:pre-line;}" +
      "#ytz-panel .row{display:flex;gap:6px;flex-wrap:wrap;}" +
      "#ytz-panel .row>*{flex:1;}" +
      "#ytz-panel button.ytz-btn{display:block;width:100%;margin:5px 0;padding:9px;border:0;border-radius:7px;background:#1f883d;color:#fff;font-size:13.5px;font-weight:600;cursor:pointer;}" +
      "#ytz-panel button.ytz-btn.sec{background:#0b5cad;}#ytz-panel button.ytz-btn.gray{background:#555;}" +
      "#ytz-panel button.ytz-btn.warn{background:#b02a37;font-size:12px;padding:7px;}" +
      "#ytz-panel button.ytz-btn:hover{filter:brightness(.92);}#ytz-panel button.ytz-btn:disabled{background:#9e9e9e;cursor:not-allowed;}" +
      "#ytz-panel select,#ytz-panel input{width:100%;padding:6px;border:1px solid #ccc;border-radius:6px;font-size:12.5px;box-sizing:border-box;}" +
      "#ytz-panel label{font-size:11.5px;color:#555;display:block;margin:4px 0 1px;}" +
      "#ytz-panel .ytz-bar{height:13px;background:#e5e5e5;border-radius:7px;overflow:hidden;margin-top:8px;}#ytz-panel .ytz-bar>i{display:block;height:100%;width:0;background:#1f883d;transition:width .2s;}" +
      "#ytz-panel .ytz-label{font-size:12px;color:#444;margin-top:6px;white-space:pre-line;}" +
      "#ytz-panel .ytz-foot{font-size:10px;color:#999;text-align:right;padding:0 16px 10px;}";
    document.head.appendChild(css);

    var overlay = document.createElement("div"); overlay.id = "ytz-overlay";
    var panel = document.createElement("div"); panel.id = "ytz-panel";
    var warnHtml = (warn && warn.total > warn.shown)
      ? '<div class="ytz-warn">⚠️ Bu aramada toplam <b>' + warn.total + '</b> tez var, yalnızca <b>' + warn.shown + '</b> tanesi listeleniyor. Tümünü almak için aramayı (örn. yıl yıl) daraltıp <b>Biriktirme</b> bölümünü kullanın.</div>'
      : '';
    panel.innerHTML =
      '<div class="ytz-head"><span>Tez Merkezi · Veri İndir</span><span class="ytz-x" title="Kapat">×</span></div>' +
      '<div class="ytz-body">' +
        '<p id="ytz-info"></p>' + warnHtml +
        '<label>Çıktı biçimi</label>' +
        '<select id="ytz-format"><option value="xlsx">Excel (.xlsx) — filtreli + istatistik</option><option value="csv">CSV (.csv)</option><option value="json">JSON (.json)</option><option value="ris">RIS — kaynakça (Zotero/Mendeley/EndNote)</option><option value="bib">BibTeX — kaynakça</option></select>' +
        '<div class="ytz-sec"><h4>Bu sayfadaki sonuçlar</h4>' +
          '<button class="ytz-btn" id="ytz-meta">Bu sayfayı indir (seçili biçim)</button>' +
          '<button class="ytz-btn sec" id="ytz-text">Bu sayfanın metinleri (PDF·ZIP)</button>' +
        '</div>' +
        '<div class="ytz-sec"><h4>Biriktirme (2000 sınırını aşmak için)</h4>' +
          '<p style="font-size:11.5px;color:#555;">Aramayı yıl yıl daraltıp her seferinde biriktirin; sonda hepsini tek dosyada indirin. Tekrar eden tezler otomatik ayıklanır.</p>' +
          '<button class="ytz-btn" id="ytz-accum">Bu aramayı biriktir</button>' +
          '<button class="ytz-btn" id="ytz-exportall">Tümünü indir (seçili biçim)</button>' +
          '<button class="ytz-btn sec" id="ytz-textall">Biriktirilenlerin metinleri (PDF·ZIP)</button>' +
          '<div class="row"><button class="ytz-btn gray" id="ytz-backup">Yedekle (JSON)</button><button class="ytz-btn gray" id="ytz-restore">Geri yükle</button></div>' +
          '<button class="ytz-btn warn" id="ytz-clear">Biriktirmeyi temizle</button>' +
          '<input type="file" id="ytz-file" accept=".json" style="display:none;">' +
        '</div>' +
        '<div class="ytz-sec"><h4>Filtrele (biriktirilenler üzerinde)</h4>' +
          '<div class="row"><div><label>Yıl (min)</label><input id="ytz-f-yil1" type="number" placeholder="örn. 2015"></div>' +
          '<div><label>Yıl (max)</label><input id="ytz-f-yil2" type="number" placeholder="örn. 2024"></div></div>' +
          '<div class="row"><div><label>Tür içerir</label><input id="ytz-f-tur" placeholder="doktora…"></div>' +
          '<div><label>Dil içerir</label><input id="ytz-f-dil" placeholder="türkçe…"></div></div>' +
          '<label>Üniversite / Yer içerir</label><input id="ytz-f-uni" placeholder="ankara üniversitesi…">' +
          '<label>Konu içerir</label><input id="ytz-f-konu" placeholder="bilgisayar…">' +
          '<label style="display:flex;align-items:center;gap:6px;margin-top:6px;"><input type="checkbox" id="ytz-f-pdf" style="width:auto;"> Sadece PDF\'i olanlar</label>' +
          '<button class="ytz-btn" id="ytz-filter">Filtreyi uygula</button>' +
          '<button class="ytz-btn" id="ytz-filter-export" disabled>Eşleşenleri indir (seçili biçim)</button>' +
          '<button class="ytz-btn sec" id="ytz-filter-text" disabled>Eşleşenlerin metinleri (PDF·ZIP)</button>' +
        '</div>' +
        '<div id="ytz-prog" style="display:none;"><div class="ytz-bar"><i id="ytz-bar"></i></div><div class="ytz-label" id="ytz-plabel"></div></div>' +
      '</div><div class="ytz-foot">mytunca/theses · yeni arayüz v1.2</div>';
    document.body.appendChild(overlay); document.body.appendChild(panel);

    var $ = function (s) { return panel.querySelector(s); };
    var elInfo = $("#ytz-info"), prog = $("#ytz-prog"), bar = $("#ytz-bar"), plabel = $("#ytz-plabel");
    var fmt = function () { return $("#ytz-format").value; };
    var allBtns = Array.from(panel.querySelectorAll("button.ytz-btn"));
    function setBusy(b) { allBtns.forEach(function (x) { x.disabled = b; }); prog.style.display = "block"; }
    function idle() { allBtns.forEach(function (x) { x.disabled = false; }); refreshInfo(); }
    function setP(p) { bar.style.width = p + "%"; } function setL(t) { plabel.textContent = t; }

    var filtered = null;
    function refreshInfo() {
      dbCount().then(function (c) {
        elInfo.textContent = (theses.length ? theses.length + " tez bu sayfada listeleniyor." : "Bu sayfada tez yok (aracı SONUÇ sayfasında çalıştırın).") + "\nBiriktirilen toplam: " + c + " tez.";
        $("#ytz-exportall").textContent = "Tümünü indir — " + c + " tez (seçili biçim)";
        var noPage = theses.length === 0, empty = c === 0;
        $("#ytz-meta").disabled = $("#ytz-text").disabled = $("#ytz-accum").disabled = noPage;
        $("#ytz-exportall").disabled = $("#ytz-textall").disabled = $("#ytz-clear").disabled = $("#ytz-backup").disabled = empty;
        var hasF = filtered && filtered.length;
        $("#ytz-filter-export").disabled = $("#ytz-filter-text").disabled = !hasF;
      });
    }
    function close() { overlay.remove(); panel.remove(); css.remove(); }
    overlay.onclick = close; $(".ytz-x").onclick = close;

    var pageRows = null;
    function getPageRows() {
      if (pageRows) return Promise.resolve(pageRows);
      setL("Metaveriler indiriliyor…");
      return fetchAllMetadata(theses, function (d, n) { setP(Math.round(100 * d / n)); setL("Metaveri: " + d + " / " + n); }).then(function (r) { pageRows = r; return r; });
    }

    $("#ytz-meta").onclick = function () { setBusy(true); setP(0); getPageRows().then(function (r) { exportData(r, "Tez_Metaverileri", fmt()); setL("Bitti · " + r.length + " tez aktarıldı."); idle(); }); };
    $("#ytz-text").onclick = function () { setBusy(true); setP(0); getPageRows().then(function (r) { setP(0); setL("Metinler indiriliyor…"); return downloadTexts(r, setP, setL); }).then(function () { setL(plabel.textContent + "\nTamamlandı."); idle(); }); };
    $("#ytz-accum").onclick = function () { setBusy(true); setP(0); getPageRows().then(function (r) { setL("Biriktirmeye ekleniyor…"); return dbPutRows(r); }).then(function () { setL("Bu aramadaki tezler biriktirmeye eklendi."); idle(); }); };
    $("#ytz-exportall").onclick = function () { setBusy(true); setL("Hazırlanıyor…"); dbGetAll().then(function (r) { exportData(r, "Tez_BIRIKIM", fmt()); setL("Bitti · " + r.length + " tez aktarıldı."); idle(); }); };
    $("#ytz-textall").onclick = function () { setBusy(true); setP(0); setL("Biriktirilenlerin metinleri indiriliyor…"); dbGetAll().then(function (r) { return downloadTexts(r, setP, setL); }).then(function () { setL(plabel.textContent + "\nTamamlandı."); idle(); }); };
    $("#ytz-backup").onclick = function () { dbGetAll().then(function (r) { saveText(JSON.stringify(r), "Tez_Birikim_Yedek_" + stamp() + ".json", "application/json"); setL(r.length + " tez JSON olarak yedeklendi."); }); };
    $("#ytz-restore").onclick = function () { $("#ytz-file").click(); };
    $("#ytz-file").onchange = function (e) {
      var file = e.target.files[0]; if (!file) return;
      var reader = new FileReader();
      reader.onload = function () { try { var arr = JSON.parse(reader.result); if (!Array.isArray(arr)) throw 0; setBusy(true); setL("Geri yükleniyor…"); dbPutRows(arr).then(function () { setL(arr.length + " tez geri yüklendi (tekrarlar ayıklandı)."); idle(); }); } catch (x) { alert("Geçersiz yedek dosyası."); } };
      reader.readAsText(file); e.target.value = "";
    };
    $("#ytz-clear").onclick = function () { if (!confirm("Biriktirilen tüm veriler silinsin mi? Geri alınamaz.")) return; dbClear().then(function () { filtered = null; setL("Biriktirme temizlendi."); refreshInfo(); }); };

    $("#ytz-filter").onclick = function () {
      var f = {
        yil1: parseInt($("#ytz-f-yil1").value, 10) || 0, yil2: parseInt($("#ytz-f-yil2").value, 10) || 0,
        tur: $("#ytz-f-tur").value.trim().toLocaleLowerCase("tr"), dil: $("#ytz-f-dil").value.trim().toLocaleLowerCase("tr"),
        uni: $("#ytz-f-uni").value.trim().toLocaleLowerCase("tr"), konu: $("#ytz-f-konu").value.trim().toLocaleLowerCase("tr"),
        onlyPdf: $("#ytz-f-pdf").checked
      };
      dbGetAll().then(function (all) {
        if (!all.length) { alert("Önce biriktirin (Biriktirme bölümü)."); return; }
        filtered = applyFilter(all, f);
        setL(filtered.length + " / " + all.length + " tez filtreye uydu.");
        prog.style.display = "block"; refreshInfo();
      });
    };
    $("#ytz-filter-export").onclick = function () { if (!filtered || !filtered.length) return; exportData(filtered, "Tez_FILTRELI", fmt()); setL(filtered.length + " eşleşen tez aktarıldı."); };
    $("#ytz-filter-text").onclick = function () { if (!filtered || !filtered.length) return; setBusy(true); setP(0); setL("Eşleşenlerin metinleri indiriliyor…"); downloadTexts(filtered, setP, setL).then(function () { setL(plabel.textContent + "\nTamamlandı."); idle(); }); };

    refreshInfo();
    return { open: function () { overlay.style.display = "block"; panel.style.display = "block"; }, close: close };
  }

  ensureDeps().then(function () { window.__yokTezAraci__ = buildUI(); })
    .catch(function (e) { alert("Gerekli kütüphaneler yüklenemedi:\n" + e.message); });
})();
