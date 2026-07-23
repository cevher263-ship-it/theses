/*
 * YÖK Ulusal Tez Merkezi - Veri Kazıma Aracı (YENİ ARAYÜZ SÜRÜMÜ)
 * ------------------------------------------------------------------
 * Orijinal araç: https://github.com/mytunca/theses (Muhammet Yunus Tunca)
 * YÖK Tez Merkezi arayüzü 2025-2026'da kart tabanlı yeni tasarıma geçtiği
 * ve eski "getData() / tezDetay.jsp" yapısı kaldırıldığı için bu sürüm
 * yeni arayüze (referenceData + tezBilgiDetay.jsp + getTezPdf.jsp) uyarlanmıştır.
 *
 * KULLANIM:
 *   1) https://tez.yok.gov.tr/UlusalTezMerkezi/ adresinde aramanızı yapın.
 *   2) Sonuçlar (tezSorguSonucYeni.jsp) listelendiğinde F12 > Console açın.
 *   3) Bu dosyanın tüm içeriğini konsola yapıştırıp Enter'a basın.
 *   4) Sağ üstte açılan "Veri İndir" penceresinden indirmeyi başlatın.
 *
 * Not: Konsol "paste protection" uyarısı verirse, tarayıcının istediği
 *      onay komutunu (örn. Chrome'da "allow pasting") yazıp tekrar yapıştırın.
 */
(function () {
  "use strict";

  // Aynı script iki kez yüklenirse paneli tekrar açmakla yetin.
  if (window.__yokTezAraci__) {
    window.__yokTezAraci__.open();
    return;
  }

  var BASE = location.origin + "/UlusalTezMerkezi/";
  var jq = window.zub || window.jQuery || window.$; // sadece ajax için değil; fetch kullanıyoruz.

  // --- Yardımcı: CDN'den script yükle -------------------------------------
  function loadScript(src) {
    return new Promise(function (resolve, reject) {
      // Zaten yüklüyse tekrar yükleme
      var s = document.createElement("script");
      s.src = src;
      s.onload = function () { resolve(); };
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

  // --- Yardımcı: HTML -> düz metin ---------------------------------------
  var _tmp = document.createElement("div");
  function stripHtml(html) {
    if (html == null) return "";
    _tmp.innerHTML = String(html);
    var t = _tmp.textContent || _tmp.innerText || "";
    // YÖK sayfalarında görülen özel karakterleri ve fazla boşlukları temizle
    return t.replace(/￾/g, " ").replace(/ /g, " ").replace(/[ \t]+\n/g, "\n").trim();
  }

  function excelSafe(str) {
    // Excel hücre sınırı 32767 karakter
    str = str == null ? "" : String(str);
    return str.length > 32000 ? str.slice(0, 32000) + " …[kesildi]" : str;
  }

  // --- Sonuç kartlarından tez listesini oku ------------------------------
  function collectTheses() {
    var RD = (typeof referenceData !== "undefined") ? referenceData
           : (window.referenceData || {});
    var cards = Array.from(document.querySelectorAll(".result-card"));
    return cards.map(function (card) {
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
        yazar: stripHtml(meta.author),
        yil: stripHtml(meta.year),
        tur: stripHtml(meta.type),
        dil: stripHtml(meta.lang),
        konu: stripHtml(meta.subject),
        yerKisa: stripHtml(meta.yer)
      };
    });
  }

  // --- Tek bir tezin detay JSON'unu çek ----------------------------------
  function fetchDetay(t) {
    var url = BASE + "tezBilgiDetay.jsp?kayitNo=" + encodeURIComponent(t.kayitNo) +
              "&tezNo=" + encodeURIComponent(t.tezNo);
    return fetch(url, { headers: { "X-Requested-With": "XMLHttpRequest" }, credentials: "include" })
      .then(function (r) { return r.text(); })
      .then(function (txt) {
        try { return JSON.parse(txt.trim()); } catch (e) { return {}; }
      });
  }

  // --- Tek bir tezin PDF (TezGoster) linkini çek -------------------------
  function fetchPdfLink(t) {
    var url = BASE + "getTezPdf.jsp?kayitNo=" + encodeURIComponent(t.kayitNo) +
              "&tezNo=" + encodeURIComponent(t.tezNo);
    return fetch(url, { headers: { "X-Requested-With": "XMLHttpRequest" }, credentials: "include" })
      .then(function (r) { return r.text(); })
      .then(function (html) {
        var m = html.match(/href=['"]([^'"]*TezGoster[^'"]*)['"]/i);
        if (!m) return null;
        var href = m[1];
        return href.indexOf("http") === 0 ? href : BASE + href.replace(/^\//, "");
      })
      .catch(function () { return null; });
  }

  // --- Eşzamanlılık havuzu (server'ı yormamak için) ----------------------
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

  // --- Bir tez için tam metaveri satırı oluştur --------------------------
  function buildRow(t, detay, pdfLink) {
    var danisman = stripHtml(detay.danisman).replace(/^Danışman:\s*/i, "");
    var yer = stripHtml(detay.yer) || t.yerKisa;
    var dizin = [stripHtml(detay.anahtarKelimeTr), stripHtml(detay.anahtarKelimeEn)]
                  .filter(Boolean).join(" | ");
    return {
      "Tez No": t.tezNoDuz,
      "PDF İndirme Linki": pdfLink || "",
      "Tez Adı (Orijinal)": t.baslikTR,
      "Tez Adı (Çeviri)": t.baslikEN,
      "Yazar": t.yazar,
      "Danışman": danisman,
      "Üniversite / Yer Bilgisi": yer,
      "Konu": t.konu,
      "Dizin (Anahtar Kelimeler)": dizin,
      "Tür": t.tur,
      "Dil": t.dil,
      "Yıl": t.yil,
      "Özet (Türkçe)": excelSafe(stripHtml(detay.trOzet)),
      "Özet (İngilizce)": excelSafe(stripHtml(detay.enOzet)),
      "kayitNo": t.kayitNo,
      "tezNo (kodlu)": t.tezNo
    };
  }

  // --- Tüm metaveriyi topla ----------------------------------------------
  function fetchAllMetadata(theses, onProgress) {
    var rows = new Array(theses.length);
    return runPool(theses, function (t) {
      return Promise.all([fetchDetay(t), fetchPdfLink(t)]).then(function (res) {
        rows[theses.indexOf(t)] = buildRow(t, res[0] || {}, res[1]);
      });
    }, 12, onProgress).then(function () {
      return rows.filter(Boolean);
    });
  }

  // --- Excel dışa aktar ---------------------------------------------------
  function exportExcel(rows) {
    var ws = XLSX.utils.json_to_sheet(rows);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Tezler");
    var stamp = new Date().toISOString().slice(0, 19).replace(/[T:]/g, "-");
    XLSX.writeFile(wb, "Tez_Metaverileri_" + stamp + ".xlsx");
  }

  // --- Tez metinlerini (PDF) ZIP olarak indir ----------------------------
  function downloadTexts(rows, onProgress, onInfo) {
    var withPdf = rows.filter(function (r) { return r["PDF İndirme Linki"]; });
    if (withPdf.length === 0) {
      alert("İndirilebilir (erişime açık) PDF bulunamadı.");
      return Promise.resolve();
    }
    var zip = new JSZip();
    var chunkIndex = 0, currentChunkSize = 0, totalSize = 0, count = 0;
    var maxChunk = 500 * 1024 * 1024; // 500 MB

    // PDF'ler sırayla indirilir (paralel indirme büyük dosyalarda belleği şişirir)
    var seq = Promise.resolve();
    withPdf.forEach(function (r) {
      seq = seq.then(function () {
        return fetch(r["PDF İndirme Linki"], { credentials: "include" })
          .then(function (resp) {
            if (!resp.ok) throw new Error("HTTP " + resp.status);
            return resp.blob();
          })
          .then(function (blob) {
            var name = (r["Tez No"] || ("tez_" + count)) + ".pdf";
            zip.file(name, blob);
            currentChunkSize += blob.size;
            totalSize += blob.size;
            count++;
            onProgress(Math.round((100 * count) / withPdf.length));
            onInfo(count + " / " + withPdf.length + " tez metni indirildi (" +
                   (totalSize / (1024 * 1024)).toFixed(1) + " MB).\n" +
                   "Bellek sorunu yaşamamanız için PDF'ler 500 MB'lık parçalar halinde kaydedilir.");
            if (currentChunkSize >= maxChunk || count === withPdf.length) {
              return zip.generateAsync({ type: "blob" }).then(function (content) {
                chunkIndex++;
                saveAs(content, "Tez_Metinleri_Part_" + chunkIndex + ".zip");
                zip = new JSZip();
                currentChunkSize = 0;
              });
            }
          })
          .catch(function (e) {
            console.warn("PDF indirilemedi (Tez No " + r["Tez No"] + "):", e.message);
          });
      });
    });
    return seq;
  }

  // =======================================================================
  //  Arayüz (kendi modalımız - jQuery UI'a bağımlı değil)
  // =======================================================================
  var theses = collectTheses();

  function buildUI() {
    var css = document.createElement("style");
    css.textContent =
      "#ytz-overlay{position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:2147483646;}" +
      "#ytz-panel{position:fixed;top:24px;right:24px;width:420px;max-width:92vw;background:#fff;" +
      "border-radius:10px;box-shadow:0 12px 40px rgba(0,0,0,.35);z-index:2147483647;font-family:" +
      "Segoe UI,Arial,sans-serif;color:#222;overflow:hidden;}" +
      "#ytz-panel .ytz-head{background:#1f883d;color:#fff;padding:12px 16px;font-size:16px;" +
      "font-weight:600;display:flex;justify-content:space-between;align-items:center;}" +
      "#ytz-panel .ytz-x{cursor:pointer;font-size:20px;line-height:1;opacity:.9;}" +
      "#ytz-panel .ytz-body{padding:16px;}" +
      "#ytz-panel p{margin:0 0 12px;font-size:13px;line-height:1.5;white-space:pre-line;}" +
      "#ytz-panel button.ytz-btn{display:block;width:100%;margin:8px 0;padding:11px;border:0;" +
      "border-radius:7px;background:#1f883d;color:#fff;font-size:14px;font-weight:600;cursor:pointer;}" +
      "#ytz-panel button.ytz-btn:hover{background:#186c31;}" +
      "#ytz-panel button.ytz-btn:disabled{background:#9e9e9e;cursor:not-allowed;}" +
      "#ytz-panel .ytz-bar{height:14px;background:#e5e5e5;border-radius:7px;overflow:hidden;margin-top:10px;}" +
      "#ytz-panel .ytz-bar>i{display:block;height:100%;width:0;background:#1f883d;transition:width .2s;}" +
      "#ytz-panel .ytz-label{font-size:12px;color:#444;margin-top:6px;white-space:pre-line;}" +
      "#ytz-panel .ytz-foot{font-size:10px;color:#999;text-align:right;padding:0 16px 10px;}";
    document.head.appendChild(css);

    var overlay = document.createElement("div");
    overlay.id = "ytz-overlay";

    var panel = document.createElement("div");
    panel.id = "ytz-panel";
    panel.innerHTML =
      '<div class="ytz-head"><span>Tez Merkezi · Veri İndir</span><span class="ytz-x" title="Kapat">×</span></div>' +
      '<div class="ytz-body">' +
        '<p id="ytz-info"></p>' +
        '<button class="ytz-btn" id="ytz-meta">Metaveri İndir (Excel)</button>' +
        '<button class="ytz-btn" id="ytz-text">Tez Metinlerini İndir (PDF · ZIP)</button>' +
        '<div id="ytz-prog" style="display:none;">' +
          '<div class="ytz-bar"><i id="ytz-bar"></i></div>' +
          '<div class="ytz-label" id="ytz-plabel"></div>' +
        '</div>' +
      '</div>' +
      '<div class="ytz-foot">mytunca/theses · yeni arayüz uyarlaması</div>';

    document.body.appendChild(overlay);
    document.body.appendChild(panel);

    var elInfo = panel.querySelector("#ytz-info");
    var btnMeta = panel.querySelector("#ytz-meta");
    var btnText = panel.querySelector("#ytz-text");
    var prog = panel.querySelector("#ytz-prog");
    var bar = panel.querySelector("#ytz-bar");
    var plabel = panel.querySelector("#ytz-plabel");

    if (theses.length) {
      elInfo.textContent = theses.length + " tez listeleniyor.\n\n" +
        "Metaverileri Excel olarak veya erişime açık tez metinlerini (PDF) " +
        "ZIP halinde indirebilirsiniz. Metin indirme işlemi tez sayısına göre " +
        "uzun sürebilir.";
    } else {
      elInfo.textContent = "Sayfada listelenen tez bulunamadı. Bu aracı arama " +
        "sonuç sayfasında (tezSorguSonucYeni.jsp) çalıştırdığınızdan emin olun.";
      btnMeta.disabled = btnText.disabled = true;
    }

    function setBusy(b) { btnMeta.disabled = btnText.disabled = b; prog.style.display = b ? "block" : "block"; }
    function setProgress(pct) { bar.style.width = pct + "%"; }
    function setLabel(txt) { plabel.textContent = txt; }

    function close() {
      overlay.remove(); panel.remove(); css.remove();
    }
    overlay.onclick = close;
    panel.querySelector(".ytz-x").onclick = close;

    var cachedRows = null;

    btnMeta.onclick = function () {
      setBusy(true); prog.style.display = "block";
      setLabel("Metaveriler indiriliyor…");
      fetchAllMetadata(theses, function (d, n) {
        setProgress(Math.round((100 * d) / n));
        setLabel("Metaveri: " + d + " / " + n);
      }).then(function (rows) {
        cachedRows = rows;
        exportExcel(rows);
        setLabel("Bitti · " + rows.length + " tez Excel'e aktarıldı.");
        btnMeta.disabled = btnText.disabled = false;
      });
    };

    btnText.onclick = function () {
      setBusy(true); prog.style.display = "block";
      var pre = cachedRows
        ? Promise.resolve(cachedRows)
        : (setLabel("Önce metaveriler alınıyor…"),
           fetchAllMetadata(theses, function (d, n) {
             setProgress(Math.round((100 * d) / n));
             setLabel("Metaveri: " + d + " / " + n);
           }));
      pre.then(function (rows) {
        cachedRows = rows;
        setProgress(0);
        setLabel("Tez metinleri (PDF) indiriliyor…");
        return downloadTexts(rows, setProgress, setLabel);
      }).then(function () {
        setLabel(function () { return plabel.textContent + "\nTüm indirmeler tamamlandı."; }());
        btnMeta.disabled = btnText.disabled = false;
      });
    };

    return { open: function () { overlay.style.display = "block"; panel.style.display = "block"; }, close: close };
  }

  // Bağımlılıkları yükle, sonra arayüzü aç.
  var ui = null;
  ensureDeps().then(function () {
    ui = buildUI();
    window.__yokTezAraci__ = ui;
  }).catch(function (e) {
    alert("Gerekli kütüphaneler yüklenemedi:\n" + e.message +
          "\n\nİnternet bağlantınızı kontrol edip tekrar deneyin.");
  });
})();
