/*
 * YÖK Ulusal Tez Merkezi - Veri Kazıma Aracı (YENİ ARAYÜZ SÜRÜMÜ + BİRİKTİRME)
 * ---------------------------------------------------------------------------
 * Orijinal araç: https://github.com/mytunca/theses (Muhammet Yunus Tunca, MIT)
 * YÖK Tez Merkezi arayüzü 2025-2026'da kart tabanlı yeni tasarıma geçtiği
 * için bu sürüm yeni yapıya (referenceData + tezBilgiDetay.jsp + getTezPdf.jsp)
 * uyarlanmıştır.
 *
 * BİRİKTİRME MODU (2000 sınırını aşmak için):
 *   YÖK tek aramada en fazla 2000 tez listeler. Aramanızı 2000'in altına inecek
 *   parçalara (örn. yıl yıl) bölüp her birinde "Bu aramayı biriktir" derseniz,
 *   veriler tarayıcıda (IndexedDB) Tez No'ya göre tekrarsız birikir. Sonunda
 *   "Tümünü Excel indir" ile hepsini tek dosyada alırsınız. Sayfa yenilense de
 *   birikim korunur.
 *
 * KULLANIM: Arama sonuç sayfasında (tezSorguSonucYeni.jsp) F12 > Console açıp
 *           bu dosyayı yapıştırın ya da jsdelivr yükleyici satırını çalıştırın.
 */
(function () {
  "use strict";

  if (window.__yokTezAraci__) { window.__yokTezAraci__.open(); return; }

  var BASE = location.origin + "/UlusalTezMerkezi/";

  // --- CDN bağımlılıkları --------------------------------------------------
  function loadScript(src) {
    return new Promise(function (resolve, reject) {
      var s = document.createElement("script");
      s.src = src; s.onload = function () { resolve(); };
      s.onerror = function () { reject(new Error("Yüklenemedi: " + src)); };
      document.head.appendChild(s);
    });
  }
  function ensureDeps() {
    var tasks = [];
    if (typeof window.XLSX === "undefined")
      tasks.push(loadScript("https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"));
    if (typeof window.JSZip === "undefined")
      tasks.push(loadScript("https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"));
    if (typeof window.saveAs === "undefined")
      tasks.push(loadScript("https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"));
    return Promise.all(tasks);
  }

  // --- Yardımcılar ---------------------------------------------------------
  var _tmp = document.createElement("div");
  function stripHtml(html) {
    if (html == null) return "";
    _tmp.innerHTML = String(html);
    var t = _tmp.textContent || _tmp.innerText || "";
    return t.replace(/￾/g, " ").replace(/ /g, " ").replace(/[ \t]+\n/g, "\n").trim();
  }
  function excelSafe(str) {
    str = str == null ? "" : String(str);
    return str.length > 32000 ? str.slice(0, 32000) + " …[kesildi]" : str;
  }

  // --- Sonuç kartlarından tez listesi -------------------------------------
  function collectTheses() {
    var RD = (typeof referenceData !== "undefined") ? referenceData : (window.referenceData || {});
    return Array.from(document.querySelectorAll(".result-card")).map(function (card) {
      var idx = card.getAttribute("data-index");
      var meta = (RD[idx] || {}).meta || {};
      var titleEl = card.querySelector(".card-title");
      var engEl = titleEl ? titleEl.nextElementSibling : null;
      var noMatch = card.textContent.match(/Tez No:\s*(\d+)/);
      return {
        index: idx,
        kayitNo: card.getAttribute("data-kayitno"),
        tezNo: card.getAttribute("data-tezno"),
        tezNoDuz: noMatch ? noMatch[1] : "",
        baslikTR: stripHtml(meta.title) || stripHtml(titleEl && titleEl.innerHTML),
        baslikEN: stripHtml(engEl && engEl.innerHTML),
        yazar: stripHtml(meta.author), yil: stripHtml(meta.year),
        tur: stripHtml(meta.type), dil: stripHtml(meta.lang),
        konu: stripHtml(meta.subject), yerKisa: stripHtml(meta.yer)
      };
    });
  }

  // --- Ağ istekleri --------------------------------------------------------
  function fetchDetay(t) {
    return fetch(BASE + "tezBilgiDetay.jsp?kayitNo=" + encodeURIComponent(t.kayitNo) +
      "&tezNo=" + encodeURIComponent(t.tezNo),
      { headers: { "X-Requested-With": "XMLHttpRequest" }, credentials: "include" })
      .then(function (r) { return r.text(); })
      .then(function (txt) { try { return JSON.parse(txt.trim()); } catch (e) { return {}; } });
  }
  function fetchPdfLink(t) {
    return fetch(BASE + "getTezPdf.jsp?kayitNo=" + encodeURIComponent(t.kayitNo) +
      "&tezNo=" + encodeURIComponent(t.tezNo),
      { headers: { "X-Requested-With": "XMLHttpRequest" }, credentials: "include" })
      .then(function (r) { return r.text(); })
      .then(function (html) {
        var m = html.match(/href=['"]([^'"]*TezGoster[^'"]*)['"]/i);
        if (!m) return null;
        return m[1].indexOf("http") === 0 ? m[1] : BASE + m[1].replace(/^\//, "");
      }).catch(function () { return null; });
  }

  function runPool(items, worker, concurrency, onProgress) {
    return new Promise(function (resolve) {
      var i = 0, done = 0, active = 0, n = items.length;
      if (n === 0) return resolve();
      function next() {
        while (active < concurrency && i < n) {
          (function (item) {
            active++;
            Promise.resolve(worker(item)).catch(function () {}).then(function () {
              active--; done++;
              if (onProgress) onProgress(done, n);
              if (done === n) resolve(); else next();
            });
          })(items[i++]);
        }
      }
      next();
    });
  }

  function buildRow(t, detay, pdfLink) {
    var danisman = stripHtml(detay.danisman).replace(/^Danışman:\s*/i, "");
    var yer = stripHtml(detay.yer) || t.yerKisa;
    var dizin = [stripHtml(detay.anahtarKelimeTr), stripHtml(detay.anahtarKelimeEn)].filter(Boolean).join(" | ");
    return {
      "Tez No": t.tezNoDuz,
      "PDF İndirme Linki": pdfLink || "",
      "Tez Adı (Orijinal)": t.baslikTR,
      "Tez Adı (Çeviri)": t.baslikEN,
      "Yazar": t.yazar, "Danışman": danisman,
      "Üniversite / Yer Bilgisi": yer, "Konu": t.konu,
      "Dizin (Anahtar Kelimeler)": dizin,
      "Tür": t.tur, "Dil": t.dil, "Yıl": t.yil,
      "Özet (Türkçe)": excelSafe(stripHtml(detay.trOzet)),
      "Özet (İngilizce)": excelSafe(stripHtml(detay.enOzet)),
      "kayitNo": t.kayitNo, "tezNo (kodlu)": t.tezNo
    };
  }

  function fetchAllMetadata(theses, onProgress) {
    var rows = new Array(theses.length);
    return runPool(theses, function (t) {
      return Promise.all([fetchDetay(t), fetchPdfLink(t)]).then(function (res) {
        rows[theses.indexOf(t)] = buildRow(t, res[0] || {}, res[1]);
      });
    }, 12, onProgress).then(function () { return rows.filter(Boolean); });
  }

  // --- Excel ---------------------------------------------------------------
  // İşe yarayan bilgiler ilk sütunlarda; özetler ve kodlar en sonda.
  var COLUMN_ORDER = [
    "Tez Adı (Orijinal)", "Yazar", "Tür", "Yıl", "Konu",
    "Üniversite / Yer Bilgisi", "Danışman", "Dil",
    "Dizin (Anahtar Kelimeler)", "Tez Adı (Çeviri)", "Tez No",
    "PDF İndirme Linki", "Özet (Türkçe)", "Özet (İngilizce)",
    "kayitNo", "tezNo (kodlu)"
  ];
  var COLUMN_WIDTHS = [55, 22, 16, 7, 22, 40, 24, 10, 30, 55, 10, 32, 60, 60, 24, 24];

  function exportExcel(rows, prefix) {
    var clean = rows.map(function (r) { var c = Object.assign({}, r); delete c._key; return c; });
    var ws = XLSX.utils.json_to_sheet(clean, { header: COLUMN_ORDER });
    ws["!cols"] = COLUMN_WIDTHS.map(function (w) { return { wch: w }; });
    // Başlık satırına filtre okları (her sütunda açılır filtre)
    if (ws["!ref"]) ws["!autofilter"] = { ref: ws["!ref"] };
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Tezler");
    var stamp = new Date().toISOString().slice(0, 19).replace(/[T:]/g, "-");
    XLSX.writeFile(wb, (prefix || "Tez_Metaverileri") + "_" + stamp + ".xlsx");
  }

  // --- PDF metinleri ZIP ---------------------------------------------------
  function downloadTexts(rows, onProgress, onInfo) {
    var withPdf = rows.filter(function (r) { return r["PDF İndirme Linki"]; });
    if (withPdf.length === 0) { alert("İndirilebilir (erişime açık) PDF bulunamadı."); return Promise.resolve(); }
    var zip = new JSZip(), chunkIndex = 0, currentChunkSize = 0, totalSize = 0, count = 0;
    var maxChunk = 500 * 1024 * 1024;
    var seq = Promise.resolve();
    withPdf.forEach(function (r) {
      seq = seq.then(function () {
        return fetch(r["PDF İndirme Linki"], { credentials: "include" })
          .then(function (resp) { if (!resp.ok) throw new Error("HTTP " + resp.status); return resp.blob(); })
          .then(function (blob) {
            zip.file((r["Tez No"] || ("tez_" + count)) + ".pdf", blob);
            currentChunkSize += blob.size; totalSize += blob.size; count++;
            onProgress(Math.round((100 * count) / withPdf.length));
            onInfo(count + " / " + withPdf.length + " tez metni indirildi (" +
              (totalSize / (1024 * 1024)).toFixed(1) + " MB).\nPDF'ler 500 MB'lık ZIP parçaları hâlinde kaydedilir.");
            if (currentChunkSize >= maxChunk || count === withPdf.length) {
              return zip.generateAsync({ type: "blob" }).then(function (content) {
                chunkIndex++; saveAs(content, "Tez_Metinleri_Part_" + chunkIndex + ".zip");
                zip = new JSZip(); currentChunkSize = 0;
              });
            }
          }).catch(function (e) { console.warn("PDF indirilemedi (" + r["Tez No"] + "):", e.message); });
      });
    });
    return seq;
  }

  // =======================================================================
  //  BİRİKTİRME deposu (IndexedDB - büyük veri için)
  // =======================================================================
  var DB_NAME = "yokTezBirikim", STORE = "tezler";
  function dbOpen() {
    return new Promise(function (resolve, reject) {
      var req = indexedDB.open(DB_NAME, 1);
      req.onupgradeneeded = function () {
        var db = req.result;
        if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE, { keyPath: "_key" });
      };
      req.onsuccess = function () { resolve(req.result); };
      req.onerror = function () { reject(req.error); };
    });
  }
  function dbPutRows(rows) {
    return dbOpen().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(STORE, "readwrite"), st = tx.objectStore(STORE);
        rows.forEach(function (r) {
          var key = (r["Tez No"] || r.kayitNo || "").toString();
          st.put(Object.assign({ _key: key }, r));
        });
        tx.oncomplete = function () { db.close(); resolve(rows.length); };
        tx.onerror = function () { db.close(); reject(tx.error); };
      });
    });
  }
  function dbGetAll() {
    return dbOpen().then(function (db) {
      return new Promise(function (resolve, reject) {
        var req = db.transaction(STORE, "readonly").objectStore(STORE).getAll();
        req.onsuccess = function () { db.close(); resolve(req.result || []); };
        req.onerror = function () { db.close(); reject(req.error); };
      });
    });
  }
  function dbCount() {
    return dbOpen().then(function (db) {
      return new Promise(function (resolve, reject) {
        var req = db.transaction(STORE, "readonly").objectStore(STORE).count();
        req.onsuccess = function () { db.close(); resolve(req.result); };
        req.onerror = function () { db.close(); reject(req.error); };
      });
    });
  }
  function dbClear() {
    return dbOpen().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(STORE, "readwrite");
        tx.objectStore(STORE).clear();
        tx.oncomplete = function () { db.close(); resolve(); };
        tx.onerror = function () { db.close(); reject(tx.error); };
      });
    });
  }

  // =======================================================================
  //  Arayüz
  // =======================================================================
  var theses = collectTheses();

  function buildUI() {
    var css = document.createElement("style");
    css.textContent =
      "#ytz-overlay{position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:2147483646;}" +
      "#ytz-panel{position:fixed;top:20px;right:20px;width:430px;max-width:94vw;max-height:92vh;overflow:auto;" +
      "background:#fff;border-radius:10px;box-shadow:0 12px 40px rgba(0,0,0,.35);z-index:2147483647;" +
      "font-family:Segoe UI,Arial,sans-serif;color:#222;}" +
      "#ytz-panel .ytz-head{background:#1f883d;color:#fff;padding:12px 16px;font-size:16px;font-weight:600;" +
      "display:flex;justify-content:space-between;align-items:center;position:sticky;top:0;}" +
      "#ytz-panel .ytz-x{cursor:pointer;font-size:22px;line-height:1;}" +
      "#ytz-panel .ytz-body{padding:14px 16px;}" +
      "#ytz-panel .ytz-sec{border:1px solid #e3e3e3;border-radius:8px;padding:10px 12px;margin-bottom:12px;}" +
      "#ytz-panel .ytz-sec h4{margin:0 0 8px;font-size:13px;color:#1f883d;}" +
      "#ytz-panel p{margin:0 0 10px;font-size:13px;line-height:1.5;white-space:pre-line;}" +
      "#ytz-panel button.ytz-btn{display:block;width:100%;margin:6px 0;padding:10px;border:0;border-radius:7px;" +
      "background:#1f883d;color:#fff;font-size:14px;font-weight:600;cursor:pointer;}" +
      "#ytz-panel button.ytz-btn.sec{background:#0b5cad;}" +
      "#ytz-panel button.ytz-btn.warn{background:#b02a37;font-size:12px;padding:7px;}" +
      "#ytz-panel button.ytz-btn:hover{filter:brightness(.92);}" +
      "#ytz-panel button.ytz-btn:disabled{background:#9e9e9e;cursor:not-allowed;}" +
      "#ytz-panel .ytz-bar{height:14px;background:#e5e5e5;border-radius:7px;overflow:hidden;margin-top:8px;}" +
      "#ytz-panel .ytz-bar>i{display:block;height:100%;width:0;background:#1f883d;transition:width .2s;}" +
      "#ytz-panel .ytz-label{font-size:12px;color:#444;margin-top:6px;white-space:pre-line;}" +
      "#ytz-panel .ytz-foot{font-size:10px;color:#999;text-align:right;padding:0 16px 10px;}";
    document.head.appendChild(css);

    var overlay = document.createElement("div"); overlay.id = "ytz-overlay";
    var panel = document.createElement("div"); panel.id = "ytz-panel";
    panel.innerHTML =
      '<div class="ytz-head"><span>Tez Merkezi · Veri İndir</span><span class="ytz-x" title="Kapat">×</span></div>' +
      '<div class="ytz-body">' +
        '<p id="ytz-info"></p>' +
        '<div class="ytz-sec"><h4>Bu sayfadaki sonuçlar</h4>' +
          '<button class="ytz-btn" id="ytz-meta">Bu sayfayı Excel indir</button>' +
          '<button class="ytz-btn sec" id="ytz-text">Bu sayfanın metinleri (PDF·ZIP)</button>' +
        '</div>' +
        '<div class="ytz-sec"><h4>Biriktirme (2000 sınırını aşmak için)</h4>' +
          '<p style="font-size:12px;color:#555;">Aramanızı yıl yıl daraltıp her seferinde "Bu aramayı biriktir" deyin; sonda hepsini tek dosyada indirin.</p>' +
          '<button class="ytz-btn" id="ytz-accum">Bu aramayı biriktir</button>' +
          '<button class="ytz-btn" id="ytz-exportall">Tümünü Excel indir</button>' +
          '<button class="ytz-btn sec" id="ytz-textall">Biriktirilenlerin metinleri (PDF·ZIP)</button>' +
          '<button class="ytz-btn warn" id="ytz-clear">Biriktirmeyi temizle</button>' +
        '</div>' +
        '<div id="ytz-prog" style="display:none;">' +
          '<div class="ytz-bar"><i id="ytz-bar"></i></div><div class="ytz-label" id="ytz-plabel"></div>' +
        '</div>' +
      '</div>' +
      '<div class="ytz-foot">mytunca/theses · yeni arayüz + biriktirme</div>';
    document.body.appendChild(overlay); document.body.appendChild(panel);

    var $ = function (id) { return panel.querySelector(id); };
    var elInfo = $("#ytz-info"), prog = $("#ytz-prog"), bar = $("#ytz-bar"), plabel = $("#ytz-plabel");
    var btns = ["#ytz-meta", "#ytz-text", "#ytz-accum", "#ytz-exportall", "#ytz-textall", "#ytz-clear"].map($);

    function setBusy(b) { btns.forEach(function (x) { x.disabled = b; }); prog.style.display = "block"; }
    function idle() { btns.forEach(function (x) { x.disabled = false; }); }
    function setProgress(p) { bar.style.width = p + "%"; }
    function setLabel(t) { plabel.textContent = t; }

    function refreshInfo() {
      dbCount().then(function (c) {
        var base = theses.length
          ? theses.length + " tez bu sayfada listeleniyor."
          : "Bu sayfada listelenen tez yok (aracı arama SONUÇ sayfasında çalıştırın).";
        elInfo.textContent = base + "\nBiriktirilen toplam: " + c + " tez.";
        $("#ytz-exportall").textContent = "Tümünü Excel indir (" + c + ")";
        var noData = theses.length === 0;
        $("#ytz-meta").disabled = $("#ytz-text").disabled = $("#ytz-accum").disabled = noData;
        var empty = c === 0;
        $("#ytz-exportall").disabled = $("#ytz-textall").disabled = $("#ytz-clear").disabled = empty;
      });
    }

    function close() { overlay.remove(); panel.remove(); css.remove(); }
    overlay.onclick = close; $(".ytz-x").onclick = close;

    var pageRows = null; // bu sayfanın metaverisi (önbellek)

    function getPageRows(labelPrefix) {
      if (pageRows) return Promise.resolve(pageRows);
      setLabel((labelPrefix || "") + "Metaveriler indiriliyor…");
      return fetchAllMetadata(theses, function (d, n) {
        setProgress(Math.round((100 * d) / n)); setLabel("Metaveri: " + d + " / " + n);
      }).then(function (rows) { pageRows = rows; return rows; });
    }

    $("#ytz-meta").onclick = function () {
      setBusy(true); setProgress(0);
      getPageRows("").then(function (rows) {
        exportExcel(rows, "Tez_Metaverileri");
        setLabel("Bitti · " + rows.length + " tez Excel'e aktarıldı."); idle();
      });
    };
    $("#ytz-text").onclick = function () {
      setBusy(true); setProgress(0);
      getPageRows("").then(function (rows) {
        setProgress(0); setLabel("Tez metinleri (PDF) indiriliyor…");
        return downloadTexts(rows, setProgress, setLabel);
      }).then(function () { setLabel(plabel.textContent + "\nTamamlandı."); idle(); });
    };
    $("#ytz-accum").onclick = function () {
      setBusy(true); setProgress(0);
      getPageRows("Biriktirmeden önce ").then(function (rows) {
        setLabel("Biriktirmeye ekleniyor…");
        return dbPutRows(rows);
      }).then(function () { setLabel("Bu aramadaki tezler biriktirmeye eklendi."); idle(); refreshInfo(); });
    };
    $("#ytz-exportall").onclick = function () {
      setBusy(true); setLabel("Biriktirilenler hazırlanıyor…");
      dbGetAll().then(function (rows) {
        exportExcel(rows, "Tez_Metaverileri_BIRIKIM");
        setLabel("Bitti · " + rows.length + " tez tek Excel'e aktarıldı."); idle();
      });
    };
    $("#ytz-textall").onclick = function () {
      setBusy(true); setProgress(0); setLabel("Biriktirilenlerin metinleri indiriliyor…");
      dbGetAll().then(function (rows) { return downloadTexts(rows, setProgress, setLabel); })
        .then(function () { setLabel(plabel.textContent + "\nTamamlandı."); idle(); });
    };
    $("#ytz-clear").onclick = function () {
      if (!confirm("Biriktirilen tüm veriler silinsin mi? Bu geri alınamaz.")) return;
      dbClear().then(function () { setLabel("Biriktirme temizlendi."); refreshInfo(); });
    };

    refreshInfo();
    return { open: function () { overlay.style.display = "block"; panel.style.display = "block"; }, close: close };
  }

  ensureDeps().then(function () {
    window.__yokTezAraci__ = buildUI();
  }).catch(function (e) {
    alert("Gerekli kütüphaneler yüklenemedi:\n" + e.message + "\n\nİnternet bağlantınızı kontrol edip tekrar deneyin.");
  });
})();
