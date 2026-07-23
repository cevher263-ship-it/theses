/*
 * YÖK Ulusal Tez Merkezi - Veri Kazıma Aracı (YENİ ARAYÜZ SÜRÜMÜ)  v1.9.1
 * ---------------------------------------------------------------------------
 * Orijinal araç: https://github.com/mytunca/theses (Muhammet Yunus Tunca, MIT)
 * YÖK Tez Merkezi'nin kart tabanlı yeni arayüzüne uyarlanmıştır.
 *
 * ÖZELLİKLER:
 *  - Metaveri indirme: Excel (.xlsx, filtre okları + istatistik sayfası), CSV, JSON
 *  - Kaynakça dışa aktarma: RIS ve BibTeX (Zotero / Mendeley / EndNote)
 *  - Tez metinleri (PDF) toplu indirme (anlamlı dosya adları, 500 MB'lık ZIP parçaları)
 *  - Biriktirme (IndexedDB): 2000 sınırını aşmak için birden çok aramayı tekrarsız biriktirme
 *  - Otomatik etiketleme: "Etiket/Bilim Dalı/Anabilim Dalı" sütunları (genel: faktüel dal adı;
 *    özel: İslam Tarihi / Türk İslam Edebiyatı / Türk İslam Sanatları içerikten tespit); +isteğe bağlı kurallar
 *  - Filtreleme: biriken listeyi etiket/yıl/tür/dil/üniversite/konu/PDF ölçütleriyle süzme
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
    // xlsx-js-style: SheetJS API'si + hücre stili (renk/font/kenarlık) desteği (MIT, ücretsiz)
    if (typeof window.XLSX === "undefined") t.push(loadScript("https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js"));
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
    return t.replace(/\uFFFE/g, " ").replace(/ /g, " ").replace(/[ \t]+\n/g, "\n").trim();
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
  function wait(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  // Sunucu yoğunlukta bazen boş/HTML yanıt döndürebilir; yer/özet kritik olduğu için 2 kez yeniden dener.
  function fetchDetay(t, tries) {
    tries = tries || 0;
    return fetch(BASE + "tezBilgiDetay.jsp?kayitNo=" + encodeURIComponent(t.kayitNo) + "&tezNo=" + encodeURIComponent(t.tezNo),
      { headers: { "X-Requested-With": "XMLHttpRequest" }, credentials: "include" })
      .then(function (r) { return r.text(); })
      .then(function (x) {
        var j = null; try { j = JSON.parse(x.trim()); } catch (e) {}
        if (j && (j.yer || j.trOzet || j.danisman || j.enOzet)) return j;
        if (tries < 2) return wait(500 + 400 * tries).then(function () { return fetchDetay(t, tries + 1); });
        return j || {};
      })
      .catch(function () { return tries < 2 ? wait(500 + 400 * tries).then(function () { return fetchDetay(t, tries + 1); }) : {}; });
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
  function detayOk(d) { return !!(d && (d.yer || d.trOzet || d.danisman || d.enOzet)); }
  // collectPdf=false (varsayılan): sadece metaveri (tezBilgiDetay) çekilir → ~2× hızlı.
  //   PDF indirme linki gerektiğinde (PDF·ZIP indirirken) o an alınır.
  // collectPdf=true: PDF linki de metaveriye eklenir (Excel'de "PDF İndirme Linki" dolu gelir, yavaşlar).
  // Dönüş: { rows, failed } — failed: sunucu yoğunluğu vb. nedeniyle verisi eksik kalan tez sayısı.
  function fetchAllMetadata(theses, onProgress, collectPdf) {
    var n = theses.length, rows = new Array(n), failed = new Array(n).fill(false);
    function work(idx) {
      var t = theses[idx];
      return Promise.all([fetchDetay(t), collectPdf ? fetchPdfLink(t) : Promise.resolve(null)])
        .then(function (res) { var d = res[0] || {}; rows[idx] = buildRow(t, d, res[1]); failed[idx] = !detayOk(d); });
    }
    var idxs = theses.map(function (_, i) { return i; });
    return runPool(idxs, work, 10, onProgress).then(function () {
      var retry = idxs.filter(function (i) { return failed[i]; });
      if (!retry.length) return;
      // Nazik ikinci geçiş: eksik kalanları düşük eşzamanlılıkla tekrar dene (throttle'ı aşmak için)
      return runPool(retry, work, 3);
    }).then(function () {
      return { rows: rows.filter(Boolean), failed: failed.filter(Boolean).length };
    });
  }

  /* ---------- Otomatik etiketleme (bilim dalı) ----------
     Etiket, tezin kendi kaydından (Yer Bilgisi) FAKTÜEL olarak doldurulur:
       1) Bilim Dalı belirtilmişse onu yazar (ör. İslam Tarihi, Türk Dili ve Edebiyatı)
       2) Yoksa Anabilim Dalı'nı yazar (ör. Temel İslam Bilimleri, İslam Tarihi ve Sanatları)
       3) Hiç disiplin bilgisi yoksa "Belirsiz"
     İSTEĞE BAĞLI: Panelde anahtar-kelime kuralları tanımlanırsa, birleşik ("X ve Y")
     ya da eksik dallar bu kurallarla içeriğe göre alt-dala ayrılabilir. */
  var DEFAULT_RULES = ""; // varsayılan: kural yok → faktüel bilim/anabilim dalı

  function lc(s) { return String(s || "").toLocaleLowerCase("tr"); }
  function parseRules(text) {
    return String(text || "").split(/\r?\n/).map(function (line) {
      var i = line.indexOf("=");
      if (i === -1) return null;
      var label = line.slice(0, i).trim();
      var kws = line.slice(i + 1).split(",").map(function (s) { return lc(s.trim()); }).filter(Boolean);
      return label && kws.length ? { label: label, keywords: kws } : null;
    }).filter(Boolean);
  }
  function segments(yer) { return String(yer || "").split("/").map(function (s) { return s.trim(); }).filter(Boolean); }
  // NOT: Türkçe "İ" (U+0130), /anabilim dal/i gibi ASCII regexlerle eşleşmez. Bu yüzden
  //      eşleştirme Türkçe küçük harf (lc) + indexOf ile yapılır. lc uzunluğu korur → dilimleme güvenli.
  function bilimDaliSegment(yer) { return segments(yer).find(function (s) { var n = lc(s); return n.indexOf("bilim dal") > -1 && n.indexOf("anabilim dal") === -1; }) || ""; }
  function anabilimDaliSegment(yer) { return segments(yer).find(function (s) { return lc(s).indexOf("anabilim dal") > -1; }) || ""; }
  // Türkçe-duyarlı başharf-büyük (İ/ı doğru; "ve, ile" gibi bağlaçlar küçük). Büyük/küçük harf
  // farkından doğan yinelenen etiketleri (ör. "TÜRK İSLAM EDEBİYATI" = "Türk İslam Edebiyatı") birleştirir.
  var TC_SMALL = { "ve": 1, "ile": 1, "için": 1, "veya": 1, "ya": 1, "ki": 1, "da": 1, "de": 1, " deki": 1 };
  function titleCaseTr(s) {
    return String(s || "").toLocaleLowerCase("tr").split(/\s+/).filter(Boolean).map(function (w, i) {
      if (i > 0 && TC_SMALL[w]) return w;
      return w.split("-").map(function (p) { return p ? p.charAt(0).toLocaleUpperCase("tr") + p.slice(1) : p; }).join("-");
    }).join(" ");
  }
  // "İslam Tarihi Bilim Dalı" -> "İslam Tarihi" ; "İSLAM TARİHİ VE SANATLARI ANABİLİM DALI" -> "İslam Tarihi ve Sanatları"
  function cleanDal(seg) {
    var s = String(seg || "").trim(), n = lc(s);
    var i = n.indexOf("anabilim dal"); if (i === -1) i = n.indexOf("bilim dal");
    var out = i > 0 ? s.slice(0, i).trim() : (i === 0 ? "" : s);
    return titleCaseTr(out);
  }

  // İsteğe bağlı kural eşleşmesi (yalnızca kullanıcı kural girdiyse çağrılır)
  function ruleLabel(row, bd, rules) {
    if (bd && !/\sve\s/i.test(lc(bd))) {
      var bdLc = lc(bd);
      var hits = rules.filter(function (r) { return r.keywords.some(function (kw) { return bdLc.indexOf(kw) > -1; }); });
      if (hits.length === 1) return hits[0].label;
    }
    var hay = lc([row["Konu"], row["Tez Adı (Orijinal)"], row["Dizin (Anahtar Kelimeler)"], row["Özet (Türkçe)"]].join(" "));
    for (var i = 0; i < rules.length; i++) { for (var j = 0; j < rules[i].keywords.length; j++) { if (hay.indexOf(rules[i].keywords[j]) > -1) return rules[i].label; } }
    return null;
  }
  // "İslam Tarihi ve Sanatları" bağlamı mı? (bilim/anabilim dalı hem tarih hem sanat içeriyor)
  function isIslamTASContext(yer) {
    var d = lc(bilimDaliSegment(yer) + " | " + anabilimDaliSegment(yer));
    return d.indexOf("tarih") > -1 && d.indexOf("sanat") > -1 && d.indexOf("islam") > -1;
  }
  function isCombinedIslamName(s) { s = lc(s); return s.indexOf("tarih") > -1 && s.indexOf("sanat") > -1; }
  // ÖZEL: İslam Tarihi ve Sanatları bağlamındaki tezi içerikten 3 alt-dala ayır (illaki bir etiket verir)
  function detectIslam3(row) {
    // 1) KONU en güçlü sinyal (YÖK'ün kendi sınıflaması): Edebiyat / Sanat
    var konu = lc(row["Konu"]);
    if (/sanat/.test(konu)) return "Türk İslam Sanatları";                 // Sanat Tarihi, Güzel Sanatlar, El Sanatları…
    if (/edebiyat|dil ve edeb/.test(konu)) return "Türk İslam Edebiyatı";   // Türk Dili ve Edebiyatı…
    // 2) İçerik anahtar kelimeleri (başlık + özet + anahtar kelime)
    var hay = lc([row["Konu"], row["Tez Adı (Orijinal)"], row["Tez Adı (Çeviri)"], row["Dizin (Anahtar Kelimeler)"], row["Özet (Türkçe)"]].join(" "));
    if (/hüsn[-\s]?i hat|hattat|hat sanat|hat san'at|tezhip|tezyin|minyatür|\bebru\b|\bçini\b|kaligrafi|mushaf|cilt san|kitâbe|kitabe|süsleme|mimari|mimarî|türbe|hânkâh|külliye|kaligraf|sanat tarih|hüsn-i/.test(hay)) return "Türk İslam Sanatları";
    if (/edebiyat|edebî|edebi |divan|dîvân|\bmesnevi\b|\bgazel\b|kaside|kasîde|na't|na’t|naat|mevlid|mevlit|münşeat|tekke şiir|\bşiir|\bşair|şerh|şârih|nüsha|müellif hatt|çeviri yazı|çeviriyazı|belagat|belâgat|metin neşri|dîvânçe|tercüme met/.test(hay)) return "Türk İslam Edebiyatı";
    return "İslam Tarihi"; // bu bağlamda kalan her şey İslam Tarihi kabul edilir
  }
  // Nihai etiket:
  //  1) net spesifik bilim dalı (birleşik İslam T&S hariç) → faktüel yaz (GENEL: her alan için)
  //  2) İslam Tarihi ve Sanatları bağlamı → içerikten 3 alt-dala ayır (ÖZEL)
  //  3) (varsa) kullanıcı kuralları
  //  4) faktüel: birleşik bilim dalı → anabilim dalı → Belirsiz
  // Kanonik etiket: İslam sanat/edebiyat/tarih dallarının farklı yazımlarını tek ortak etikette birleştirir.
  // Farklı disiplinleri (Din Musikisi, İslam Mezhepleri, İslam Medeniyeti, Mevlânâ vb.) DIŞTA bırakır.
  function canonicalEtiket(label) {
    var n = lc(label);
    if (n.indexOf("islam") === -1 && n.indexOf("islâm") === -1) return label; // İslam ile ilgili değilse dokunma
    if (/musik|mezhep|medeniyet|mevlan|mevlân|tasavvuf|hukuk|iktisat|ekonom|felsefe|kelam|tefsir|hadis|akaid/.test(n)) return label; // ayrı disiplinler
    if (/sanat/.test(n)) return "Türk İslam Sanatları";                        // İslam Sanatları, Türk-İslam Sanatı, Türk ve İslam Sanatı…
    if (/edebiyat|edebiyât/.test(n)) return "Türk İslam Edebiyatı";
    if (/tarih/.test(n)) return "İslam Tarihi";                                // Siyer-i Nebi ve İslam Tarihi…
    return label;
  }
  function classifyRow(row, rules) {
    var yer = row["Üniversite / Yer Bilgisi"];
    var bd = bilimDaliSegment(yer), abd = anabilimDaliSegment(yer);
    var bdC = cleanDal(bd), abdC = cleanDal(abd);
    var res;
    if (bdC && !isCombinedIslamName(bdC)) res = bdC;           // ör. "İslam Tarihi", "Türk Dili ve Edebiyatı", "Temel İslam Bilimleri"
    else if (isIslamTASContext(yer)) res = detectIslam3(row);  // ÖZEL: 3 İslam alt-dalı
    else if (rules && rules.length && ruleLabel(row, bd, rules)) res = ruleLabel(row, bd, rules);
    else if (bdC) res = bdC;
    else if (abdC) res = abdC;
    else return "Belirsiz";
    return canonicalEtiket(res);
  }
  var getRules = function () { return parseRules(DEFAULT_RULES); }; // UI hazır olunca panele bağlanır
  function tagRows(rows) {
    var rules = getRules();
    return rows.map(function (r) {
      var c = Object.assign({}, r); delete c._key;
      c["Anabilim Dalı"] = cleanDal(anabilimDaliSegment(c["Üniversite / Yer Bilgisi"]));
      c["Bilim Dalı"] = cleanDal(bilimDaliSegment(c["Üniversite / Yer Bilgisi"]));
      c["Etiket"] = classifyRow(c, rules);
      return c;
    });
  }

  /* ---------- Çıktı biçimleri ---------- */
  var COLUMN_ORDER = ["Tez Adı (Orijinal)", "Yazar", "Etiket", "Bilim Dalı", "Anabilim Dalı", "Tür", "Yıl", "Konu", "Üniversite / Yer Bilgisi", "Danışman", "Dil", "Dizin (Anahtar Kelimeler)", "Tez Adı (Çeviri)", "Tez No", "PDF İndirme Linki", "Özet (Türkçe)", "Özet (İngilizce)", "kayitNo", "tezNo (kodlu)"];
  var COLUMN_WIDTHS = [55, 22, 24, 26, 30, 16, 7, 22, 40, 24, 10, 30, 55, 10, 32, 60, 60, 24, 24];

  function cleanRows(rows) { return rows.map(function (r) { var c = Object.assign({}, r); delete c._key; return c; }); }

  /* ---------- Stil yardımcıları (xlsx-js-style) ---------- */
  var GREEN = "1F883D", GREEN_DK = "186C31", ZEBRA = "EAF3EC", HEAD_TXT = "FFFFFF", BORDER = "D6D6D6";
  function thinBorder(color) { var b = { style: "thin", color: { rgb: color || BORDER } }; return { top: b, bottom: b, left: b, right: b }; }
  var HEADER_STYLE = { font: { bold: true, color: { rgb: HEAD_TXT }, sz: 11 }, fill: { fgColor: { rgb: GREEN } }, alignment: { horizontal: "center", vertical: "center", wrapText: true }, border: thinBorder(GREEN_DK) };
  function styleDataSheet(ws) {
    if (!ws["!ref"]) return ws;
    var range = XLSX.utils.decode_range(ws["!ref"]);
    for (var c = range.s.c; c <= range.e.c; c++) {
      var h = XLSX.utils.encode_cell({ r: 0, c: c });
      if (ws[h]) ws[h].s = HEADER_STYLE;
    }
    for (var r = 1; r <= range.e.r; r++) {
      var zebra = (r % 2 === 0);
      var rowStyle = { alignment: { vertical: "top", wrapText: false }, font: { sz: 10 }, border: thinBorder() };
      if (zebra) rowStyle.fill = { fgColor: { rgb: ZEBRA } };
      for (var c2 = range.s.c; c2 <= range.e.c; c2++) {
        var a = XLSX.utils.encode_cell({ r: r, c: c2 });
        if (ws[a]) ws[a].s = rowStyle;
      }
    }
    ws["!rows"] = [{ hpt: 26 }]; // başlık satırı biraz yüksek
    return ws;
  }
  function buildMainSheet(rows, styled) {
    var ws = XLSX.utils.json_to_sheet(rows, { header: COLUMN_ORDER });
    ws["!cols"] = COLUMN_WIDTHS.map(function (w) { return { wch: w }; });
    if (ws["!ref"]) ws["!autofilter"] = { ref: ws["!ref"] };
    if (!styled) return ws;
    ws["!freeze"] = { xSplit: 0, ySplit: 1, topLeftCell: "A2", activePane: "bottomLeft", state: "frozen" };
    return styleDataSheet(ws);
  }

  // İstatistik sayfası: yıl/tür/dil/konu/üniversiteye göre adet + renkli çubuk
  function buildStatsSheet(rows) {
    function tally(getter) { var m = {}; rows.forEach(function (r) { var v = (getter(r) || "—").toString().trim() || "—"; m[v] = (m[v] || 0) + 1; }); return Object.entries(m).sort(function (a, b) { return b[1] - a[1]; }); }
    function bar(n, max) { var w = max ? Math.round(n / max * 30) : 0; return "█".repeat(w); }
    var aoa = [], kind = []; // kind[r]: 'title' | 'sub' | 'data' | 'blank'
    function push(row, k) { aoa.push(row); kind.push(k); }
    push(["YÖK TEZ MERKEZİ — İSTATİSTİK ÖZETİ", ""], "title");
    push(["Toplam tez", rows.length], "sub");
    push([""], "blank");
    function section(title, pairs, limit) {
      push([title, ""], "title");
      push(["Değer", "Adet", ""], "sub");
      var max = pairs.length ? pairs[0][1] : 0;
      pairs.slice(0, limit || pairs.length).forEach(function (p) { push([p[0], p[1], bar(p[1], max)], "data"); });
      if (limit && pairs.length > limit) push(["… (+" + (pairs.length - limit) + " diğer)", "", ""], "data");
      push([""], "blank");
    }
    if (rows.some(function (r) { return r["Etiket"]; })) section("ETİKETE GÖRE (ilk 30)", tally(function (r) { return r["Etiket"]; }), 30);
    section("ANABİLİM DALINA GÖRE (ilk 25)", tally(function (r) { return r["Anabilim Dalı"] || "(belirtilmemiş)"; }), 25);
    section("YILA GÖRE", tally(function (r) { return r["Yıl"]; }));
    section("TÜRE GÖRE", tally(function (r) { return r["Tür"]; }));
    section("DİLE GÖRE", tally(function (r) { return r["Dil"]; }));
    section("KONUYA GÖRE (ilk 25)", tally(function (r) { return r["Konu"]; }), 25);
    section("ÜNİVERSİTEYE GÖRE (ilk 25)", tally(function (r) { return cleanUni(r["Üniversite / Yer Bilgisi"]); }), 25);
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 45 }, { wch: 10 }, { wch: 34 }];
    ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }];
    // stiller
    var titleStyle = { font: { bold: true, color: { rgb: HEAD_TXT }, sz: 12 }, fill: { fgColor: { rgb: GREEN } }, alignment: { vertical: "center" } };
    var subStyle = { font: { bold: true, color: { rgb: GREEN_DK }, sz: 10 }, fill: { fgColor: { rgb: ZEBRA } } };
    var barStyle = { font: { color: { rgb: GREEN }, sz: 10 } };
    kind.forEach(function (k, r) {
      if (k === "title") { for (var c = 0; c <= 2; c++) { var a = XLSX.utils.encode_cell({ r: r, c: c }); if (!ws[a]) ws[a] = { t: "s", v: "" }; ws[a].s = titleStyle; } ws["!rows"] = ws["!rows"] || []; ws["!rows"][r] = { hpt: 20 }; }
      else if (k === "sub") { for (var c2 = 0; c2 <= 2; c2++) { var a2 = XLSX.utils.encode_cell({ r: r, c: c2 }); if (ws[a2]) ws[a2].s = subStyle; } }
      else if (k === "data") { var ab = XLSX.utils.encode_cell({ r: r, c: 2 }); if (ws[ab]) ws[ab].s = barStyle; }
    });
    return ws;
  }

  function exportExcel(rows, prefix) {
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, buildMainSheet(rows, true), "Tezler");
    XLSX.utils.book_append_sheet(wb, buildStatsSheet(rows), "İstatistik");
    XLSX.writeFile(wb, (prefix || "Tez_Metaverileri") + "_" + stamp() + ".xlsx");
  }
  function saveText(text, filename, mime) {
    saveAs(new Blob(["\uFEFF" + text], { type: (mime || "text/plain") + ";charset=utf-8" }), filename);
  }
  function exportCSV(rows, prefix) {
    var csv = XLSX.utils.sheet_to_csv(buildMainSheet(rows), { FS: ";" });
    saveText(csv, (prefix || "Tez_Metaverileri") + "_" + stamp() + ".csv", "text/csv");
  }
  function exportJSON(rows, prefix) {
    saveText(JSON.stringify(rows, null, 1), (prefix || "Tez_Metaverileri") + "_" + stamp() + ".json", "application/json");
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
      if (r["Etiket"] && r["Etiket"] !== "Belirsiz") L.push("KW  - " + risEsc(r["Etiket"]));
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
        ["keywords", [r["Etiket"] && r["Etiket"] !== "Belirsiz" ? r["Etiket"] : "", (r["Dizin (Anahtar Kelimeler)"] || "").replace(/\|/g, ",")].filter(Boolean).join(", ")],
        ["abstract", r["Özet (Türkçe)"]], ["language", r["Dil"]],
        ["note", r["Tez No"] ? "Tez No: " + r["Tez No"] : ""], ["url", r["PDF İndirme Linki"]]];
      var body = f.filter(function (x) { return x[1]; }).map(function (x) { return "  " + x[0] + " = {" + texEsc(x[1]) + "}"; }).join(",\n");
      return "@" + bibType(r["Tür"]) + "{" + bibKey(r, used) + ",\n" + body + "\n}";
    }).join("\n\n");
  }

  /* ---------- bibliometrix / biblioshiny (WoS düz-metin) ---------- */
  function oneLine(s) { return String(s || "").replace(/\s+/g, " ").trim(); }
  function authorWos(name) { // "BEGÜM SÖNMEZ" -> "Sönmez, Begüm"
    var p = oneLine(name).split(" ").filter(Boolean);
    if (p.length < 2) return titleCaseTr(name);
    var soyad = p.pop();
    return titleCaseTr(soyad) + ", " + titleCaseTr(p.join(" "));
  }
  function docTypeWos(tur) { var t = lc(tur); return t.indexOf("doktora") > -1 ? "Doctoral Thesis" : (/y[üu]ksek|master/.test(t) ? "Master Thesis" : (t.indexOf("uzmanl") > -1 ? "Medical Speciality Thesis" : "Thesis")); }
  var WOS_TITLES = { prof: 1, dr: 1, "doç": 1, doc: 1, "doçent": 1, docent: 1, yrd: 1, yard: 1, "öğr": 1, ogr: 1, "gör": 1, gor: 1, "üyesi": 1, uyesi: 1, uzm: 1, op: 1 };
  function stripTitle(name) { // "DOÇ. DR. BÜNYAMİN AYÇİÇEĞİ" -> "BÜNYAMİN AYÇİÇEĞİ"
    return oneLine(String(name || "").split(/\s+/).filter(function (w) { return !WOS_TITLES[lc(w).replace(/[.̇]/g, "")]; }).join(" "));
  }
  var wosAdvisorAsAuthor = false; // panelden ayarlanır
  function toWos(rows) {
    var out = ["FN Clarivate Analytics Web of Science", "VR 1.0"];
    rows.forEach(function (r) {
      var aus = [authorWos(r["Yazar"])];
      if (wosAdvisorAsAuthor && r["Danışman"]) { var adv = authorWos(stripTitle(r["Danışman"])); if (adv && adv !== aus[0]) aus.push(adv); }
      var de = oneLine(r["Dizin (Anahtar Kelimeler)"]).split(/\s*[|;]\s*/).filter(Boolean).join("; ");
      var id = [r["Etiket"], r["Konu"], r["Bilim Dalı"]].filter(function (x) { return x && x !== "Belirsiz" && x !== "—"; })
        .join("; ").split(/\s*;\s*/).filter(function (v, i, a) { return v && a.indexOf(v) === i; }).join("; ");
      var ab = oneLine(r["Özet (İngilizce)"]) || oneLine(r["Özet (Türkçe)"]);
      out.push("PT J");
      out.push("AU " + aus[0]); for (var i1 = 1; i1 < aus.length; i1++) out.push("   " + aus[i1]);
      out.push("AF " + aus[0]); for (var i2 = 1; i2 < aus.length; i2++) out.push("   " + aus[i2]);
      if (r["Danışman"]) out.push("C3 Danışman: " + oneLine(r["Danışman"])); // ek alan (bilgi amaçlı)
      out.push("TI " + oneLine(r["Tez Adı (Orijinal)"]));
      out.push("SO " + (cleanUni(r["Üniversite / Yer Bilgisi"]) || "YOK Ulusal Tez Merkezi"));
      if (r["Dil"]) out.push("LA " + r["Dil"]);
      out.push("DT " + docTypeWos(r["Tür"]));
      if (de) out.push("DE " + de);
      if (id) out.push("ID " + id);
      if (ab) out.push("AB " + ab);
      out.push("C1 " + oneLine(r["Üniversite / Yer Bilgisi"]));
      if (r["Yıl"]) out.push("PY " + r["Yıl"]);
      out.push("TC 0");
      out.push("UT YOK:" + (r["Tez No"] || (r["kayitNo"] || "")));
      out.push("ER");
      out.push("");
    });
    out.push("EF");
    return out.join("\n");
  }
  function saveNoBom(text, filename) { saveAs(new Blob([text], { type: "text/plain;charset=utf-8" }), filename); }

  /* ---------- Biçim seçimine göre dışa aktar ---------- */
  function exportData(rows, prefix, format) {
    if (!rows.length) { alert("Dışa aktarılacak tez yok."); return; }
    var data = tagRows(rows); // temizle + Etiket/Bilim Dalı sütunlarını ekle
    if (format === "csv") exportCSV(data, prefix);
    else if (format === "json") exportJSON(data, prefix);
    else if (format === "ris") saveText(toRIS(data), (prefix || "Tez") + "_kaynakca_" + stamp() + ".ris", "application/x-research-info-systems");
    else if (format === "bib") saveText(toBibTeX(data), (prefix || "Tez") + "_kaynakca_" + stamp() + ".bib", "application/x-bibtex");
    else if (format === "wos") saveNoBom(toWos(data), (prefix || "Tez") + "_bibliometrix_" + stamp() + ".txt"); // WoS plaintext (BOM'suz!)
    else exportExcel(data, prefix);
  }

  /* ---------- Özet rapor (grafikli, kendi içinde bağımsız HTML) ---------- */
  function repEsc(s) { return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function repTally(rows, getter) { var m = {}; rows.forEach(function (r) { var v = (getter(r) || "—").toString().trim() || "—"; m[v] = (m[v] || 0) + 1; }); return Object.entries(m).sort(function (a, b) { return b[1] - a[1]; }); }
  function repBars(pairs, limit) {
    pairs = pairs.slice(0, limit || pairs.length);
    if (!pairs.length) return '<p class="empty">Veri yok.</p>';
    var max = Math.max.apply(null, pairs.map(function (p) { return p[1]; })) || 1;
    return '<div class="bars">' + pairs.map(function (p) {
      var pct = Math.max(1, Math.round(p[1] / max * 100));
      return '<div class="bar-row"><div class="bar-label" title="' + repEsc(p[0]) + '">' + repEsc(p[0]) + '</div>' +
        '<div class="bar-track"><div class="bar-fill" style="width:' + pct + '%"></div></div><div class="bar-val">' + p[1] + '</div></div>';
    }).join("") + '</div>';
  }
  var REP_STOP = "ve ile bir bu için olarak göre ait dair olan ki da de mi mı ya veya çok daha en gibi kadar sonra önce adlı isimli üzerine bağlamında örneği örneğinde açısından incelenmesi incelemesi analizi araştırma araştırması değerlendirilmesi karşılaştırmalı karşılaştırılması ilişkin yönelik ması mesi arasında ının inin unun ünün nın nin nun nün ile ilgili".split(" ").reduce(function (o, w) { o[w] = 1; return o; }, {});
  function repWords(rows) {
    var freq = {};
    rows.forEach(function (r) {
      lc(r["Tez Adı (Orijinal)"]).split(/[^a-zçğıöşü0-9âîû]+/).forEach(function (w) {
        if (w.length >= 4 && !REP_STOP[w] && !/^\d+$/.test(w)) freq[w] = (freq[w] || 0) + 1;
      });
    });
    return Object.entries(freq).sort(function (a, b) { return b[1] - a[1]; }).slice(0, 35);
  }
  function buildReportHTML(rows, kapsam) {
    var years = repTally(rows, function (r) { return r["Yıl"]; }).filter(function (p) { return /^\d{4}$/.test(p[0]); }).sort(function (a, b) { return a[0] - b[0]; });
    var ynums = years.map(function (p) { return +p[0]; });
    var uni = {}, dan = {}, etk = {};
    rows.forEach(function (r) { uni[cleanUni(r["Üniversite / Yer Bilgisi"])] = 1; if (r["Danışman"]) dan[r["Danışman"]] = 1; etk[r["Etiket"]] = 1; });
    var words = repWords(rows), maxW = words.length ? words[0][1] : 1;
    var cloud = words.map(function (p) { return '<span style="font-size:' + (13 + Math.round(p[1] / maxW * 24)) + 'px;opacity:' + (0.55 + p[1] / maxW * 0.45).toFixed(2) + '">' + repEsc(p[0]) + '</span>'; }).join(" ");
    function box(l, v) { return '<div class="stat"><div class="stat-v">' + v + '</div><div class="stat-l">' + l + '</div></div>'; }
    function sec(t, body) { return '<section><h2>' + t + '</h2>' + body + '</section>'; }
    var css = "*{box-sizing:border-box}body{font-family:Segoe UI,Arial,sans-serif;color:#222;max-width:900px;margin:24px auto;padding:0 18px;line-height:1.5}" +
      "h1{color:#1f883d;margin:0 0 4px}.sub{color:#666;margin:0 0 18px;font-size:13px}h2{color:#186c31;font-size:16px;border-bottom:2px solid #eaf3ec;padding-bottom:4px;margin:26px 0 12px}" +
      ".stats{display:flex;flex-wrap:wrap;gap:10px;margin:10px 0 6px}.stat{flex:1;min-width:130px;background:#f4f9f5;border:1px solid #dcece0;border-radius:10px;padding:12px 14px;text-align:center}" +
      ".stat-v{font-size:24px;font-weight:700;color:#1f883d}.stat-l{font-size:12px;color:#555;margin-top:2px}" +
      ".bars{display:flex;flex-direction:column;gap:5px}.bar-row{display:flex;align-items:center;gap:8px;font-size:13px}" +
      ".bar-label{width:230px;flex:0 0 230px;text-align:right;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#333}" +
      ".bar-track{flex:1;background:#eef1ef;border-radius:5px;height:18px;overflow:hidden}.bar-fill{height:100%;background:linear-gradient(90deg,#1f883d,#2ea24a);border-radius:5px}" +
      ".bar-val{width:52px;flex:0 0 52px;font-variant-numeric:tabular-nums;color:#444;font-weight:600}" +
      ".cloud{line-height:2;background:#f9fbfa;border:1px solid #e6efe9;border-radius:10px;padding:14px}.cloud span{color:#1f883d;margin:0 6px;display:inline-block}" +
      ".foot{margin:30px 0 10px;color:#999;font-size:11px;text-align:right}.empty{color:#999;font-size:13px}" +
      "@media print{body{margin:0}section{page-break-inside:avoid}}@media(max-width:560px){.bar-label{width:130px;flex-basis:130px}}";
    return '<!doctype html><html lang="tr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>YÖK Tez — Özet Rapor</title><style>' + css + '</style></head><body>' +
      '<h1>YÖK Tez — Özet Rapor</h1><p class="sub">' + repEsc(kapsam) + ' · ' + rows.length + ' tez · oluşturulma: ' + stamp().replace(/-/g, ".") + '</p>' +
      '<div class="stats">' + box("Toplam tez", rows.length) + box("Yıl aralığı", ynums.length ? Math.min.apply(null, ynums) + "–" + Math.max.apply(null, ynums) : "—") + box("Üniversite", Object.keys(uni).length) + box("Danışman", Object.keys(dan).length) + box("Disiplin/Etiket", Object.keys(etk).length) + '</div>' +
      sec("Yıllara göre dağılım", repBars(years)) +
      sec("Etikete göre (bilim dalı)", repBars(repTally(rows, function (r) { return r["Etiket"]; }), 20)) +
      sec("Türe göre", repBars(repTally(rows, function (r) { return r["Tür"]; }))) +
      sec("Dile göre", repBars(repTally(rows, function (r) { return r["Dil"]; }))) +
      sec("En çok tez veren üniversiteler (ilk 15)", repBars(repTally(rows, function (r) { return cleanUni(r["Üniversite / Yer Bilgisi"]); }), 15)) +
      sec("En üretken danışmanlar (ilk 15)", repBars(repTally(rows, function (r) { return r["Danışman"]; }).filter(function (p) { return p[0] !== "—"; }), 15)) +
      sec("En sık konular (ilk 15)", repBars(repTally(rows, function (r) { return r["Konu"]; }), 15)) +
      sec("Başlıklarda sık geçen kelimeler", '<div class="cloud">' + (cloud || '<span class="empty">Veri yok.</span>') + '</div>') +
      '<p class="foot">mytunca/theses · yeni arayüz — Özet Rapor (yazdır → PDF olarak kaydedebilirsiniz)</p></body></html>';
  }
  function exportReport(rows, kapsam) {
    if (!rows.length) { alert("Rapor için tez yok."); return; }
    var html = buildReportHTML(tagRows(rows), kapsam);
    var blob = new Blob([html], { type: "text/html;charset=utf-8" });
    var url = URL.createObjectURL(blob);
    var w = window.open(url, "_blank");
    if (!w) saveAs(blob, "Tez_Ozet_Rapor_" + stamp() + ".html"); // popup engellenirse indir
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
  // Satırdaki PDF linki yoksa (metaveri aşamasında toplanmadıysa) o an getTezPdf ile alır.
  function resolvePdfLink(r) {
    if (r["PDF İndirme Linki"]) return Promise.resolve(r["PDF İndirme Linki"]);
    var kayitNo = r["kayitNo"], tezNo = r["tezNo (kodlu)"];
    if (!kayitNo || !tezNo) return Promise.resolve(null);
    return fetchPdfLink({ kayitNo: kayitNo, tezNo: tezNo });
  }
  function downloadTexts(rows, onProgress, onInfo) {
    var zip = new JSZip(), chunkIndex = 0, chunkSize = 0, total = 0, count = 0, downloaded = 0, used = {};
    var n = rows.length, maxChunk = 500 * 1024 * 1024, seq = Promise.resolve();
    rows.forEach(function (r) {
      seq = seq.then(function () {
        return resolvePdfLink(r).then(function (link) {
          count++;
          onProgress(Math.round((100 * count) / n));
          if (!link) { onInfo(count + " / " + n + " tez tarandı, " + downloaded + " PDF indirildi (" + (total / (1024 * 1024)).toFixed(1) + " MB). (Erişime kapalı olanlar atlanır.)"); return; }
          return fetch(link, { credentials: "include" })
            .then(function (resp) { if (!resp.ok) throw new Error("HTTP " + resp.status); return resp.blob(); })
            .then(function (blob) {
              zip.file(pdfName(r, used), blob); chunkSize += blob.size; total += blob.size; downloaded++;
              onInfo(count + " / " + n + " tez tarandı, " + downloaded + " PDF indirildi (" + (total / (1024 * 1024)).toFixed(1) + " MB).\nPDF'ler 500 MB'lık ZIP parçaları hâlinde kaydedilir.");
              if (chunkSize >= maxChunk) {
                return zip.generateAsync({ type: "blob" }).then(function (c) { chunkIndex++; saveAs(c, "Tez_Metinleri_Part_" + chunkIndex + ".zip"); zip = new JSZip(); chunkSize = 0; });
              }
            }).catch(function (e) { console.warn("PDF indirilemedi (" + r["Tez No"] + "):", e.message); });
        });
      });
    });
    return seq.then(function () {
      if (downloaded === 0) { alert("İndirilebilir (erişime açık) PDF bulunamadı."); return; }
      if (chunkSize > 0) return zip.generateAsync({ type: "blob" }).then(function (c) { chunkIndex++; saveAs(c, "Tez_Metinleri_Part_" + chunkIndex + ".zip"); });
    });
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
    var rules = f.etiket ? getRules() : null;
    return rows.filter(function (r) {
      var y = parseInt(r["Yıl"], 10);
      if (f.yil1 && (!y || y < f.yil1)) return false;
      if (f.yil2 && (!y || y > f.yil2)) return false;
      if (f.etiket && lc(classifyRow(r, rules)).indexOf(f.etiket) === -1) return false;
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
      "#ytz-panel select,#ytz-panel input,#ytz-panel textarea{width:100%;padding:6px;border:1px solid #ccc;border-radius:6px;font-size:12.5px;box-sizing:border-box;font-family:inherit;}" +
      "#ytz-panel textarea{min-height:88px;resize:vertical;line-height:1.4;}" +
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
        '<select id="ytz-format"><option value="xlsx">Excel (.xlsx) — filtreli + istatistik</option><option value="csv">CSV (.csv)</option><option value="json">JSON (.json)</option><option value="ris">RIS — kaynakça (Zotero/Mendeley/EndNote)</option><option value="bib">BibTeX — kaynakça</option><option value="wos">bibliometrix / biblioshiny (WoS düz-metin)</option></select>' +
        '<label style="display:flex;align-items:center;gap:6px;margin-top:8px;"><input type="checkbox" id="ytz-pdflink" style="width:auto;"> Excel\'e PDF indirme linki sütununu da ekle <span style="color:#b02a37;">(≈2× yavaşlar)</span></label>' +
        '<label id="ytz-wos-adv-row" style="display:none;align-items:center;gap:6px;margin-top:6px;"><input type="checkbox" id="ytz-wos-advisor" style="width:auto;"> Danışmanı 2. yazar yap <span style="color:#666;">(bibliometrix işbirliği/danışman ağı için; yazar üretkenliği metriklerini karıştırır)</span></label>' +
        '<div class="ytz-sec"><h4>Bu sayfadaki sonuçlar</h4>' +
          '<button class="ytz-btn" id="ytz-meta">Bu sayfayı indir (seçili biçim)</button>' +
          '<button class="ytz-btn sec" id="ytz-text">Bu sayfanın metinleri (PDF·ZIP)</button>' +
          '<button class="ytz-btn" id="ytz-report" style="background:#6b3fa0">📊 Özet rapor (grafikli)</button>' +
        '</div>' +
        '<div class="ytz-sec"><h4>Biriktirme (2000 sınırını aşmak için)</h4>' +
          '<p style="font-size:11.5px;color:#555;">Aramayı yıl yıl daraltıp her seferinde biriktirin; sonda hepsini tek dosyada indirin. Tekrar eden tezler otomatik ayıklanır.</p>' +
          '<button class="ytz-btn" id="ytz-accum">Bu aramayı biriktir</button>' +
          '<button class="ytz-btn" id="ytz-exportall">Tümünü indir (seçili biçim)</button>' +
          '<button class="ytz-btn sec" id="ytz-textall">Biriktirilenlerin metinleri (PDF·ZIP)</button>' +
          '<button class="ytz-btn" id="ytz-reportall" style="background:#6b3fa0">📊 Biriktirilenlerin özet raporu</button>' +
          '<div class="row"><button class="ytz-btn gray" id="ytz-backup">Yedekle (JSON)</button><button class="ytz-btn gray" id="ytz-restore">Geri yükle</button></div>' +
          '<button class="ytz-btn warn" id="ytz-clear">Biriktirmeyi temizle</button>' +
          '<input type="file" id="ytz-file" accept=".json" style="display:none;">' +
        '</div>' +
        '<div class="ytz-sec"><h4>Etiketleme (bilim dalı)</h4>' +
          '<p style="font-size:11.5px;color:#555;">Çıktıya otomatik <b>Etiket</b>, <b>Bilim Dalı</b>, <b>Anabilim Dalı</b> sütunları eklenir.<br>• <b>Genel:</b> Etiket = bilim dalı (varsa) → anabilim dalı → yoksa <b>Belirsiz</b>.<br>• <b>Özel:</b> <b>İslam Tarihi</b>, <b>Türk İslam Edebiyatı</b>, <b>Türk İslam Sanatları</b> için bilim dalı eksik/birleşik olsa da içerikten mutlaka birine atanır.<br><b>İsteğe bağlı</b>: başka alanları da içerikten ayırmak için kural yazın — <b>Etiket = kelime1, kelime2, …</b> (boş bırakılabilir).</p>' +
          '<textarea id="ytz-rules" spellcheck="false" placeholder="Örnek (isteğe bağlı):&#10;Türk İslam Edebiyatı = divan edebiyat, mesnevi, na\'t, mevlid&#10;Türk İslam Sanatları = hüsn-i hat, tezhip, minyatür, çini&#10;İslam Tarihi = siyer, sahabe, emevi, abbasi, endülüs"></textarea>' +
        '</div>' +
        '<div class="ytz-sec"><h4>Filtrele (biriktirilenler üzerinde)</h4>' +
          '<div class="row"><div><label>Yıl (min)</label><input id="ytz-f-yil1" type="number" placeholder="örn. 2015"></div>' +
          '<div><label>Yıl (max)</label><input id="ytz-f-yil2" type="number" placeholder="örn. 2024"></div></div>' +
          '<label>Etiket içerir</label><input id="ytz-f-etiket" placeholder="islam tarihi… (yukarıdaki kurallara göre)">' +
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
      '</div><div class="ytz-foot">mytunca/theses · yeni arayüz v1.9.1</div>';
    document.body.appendChild(overlay); document.body.appendChild(panel);

    var $ = function (s) { return panel.querySelector(s); };
    $("#ytz-rules").value = DEFAULT_RULES;
    getRules = function () { return parseRules(($("#ytz-rules").value || "").trim() || DEFAULT_RULES); };
    var elInfo = $("#ytz-info"), prog = $("#ytz-prog"), bar = $("#ytz-bar"), plabel = $("#ytz-plabel");
    var fmt = function () { return $("#ytz-format").value; };
    // Danışman-2.-yazar kutusu yalnızca bibliometrix (wos) biçiminde görünür
    $("#ytz-format").onchange = function () { $("#ytz-wos-adv-row").style.display = (fmt() === "wos") ? "flex" : "none"; };
    $("#ytz-wos-advisor").onchange = function () { wosAdvisorAsAuthor = this.checked; };
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
        $("#ytz-meta").disabled = $("#ytz-text").disabled = $("#ytz-accum").disabled = $("#ytz-report").disabled = noPage;
        $("#ytz-exportall").disabled = $("#ytz-textall").disabled = $("#ytz-clear").disabled = $("#ytz-backup").disabled = $("#ytz-reportall").disabled = empty;
        var hasF = filtered && filtered.length;
        $("#ytz-filter-export").disabled = $("#ytz-filter-text").disabled = !hasF;
      });
    }
    function close() { overlay.remove(); panel.remove(); css.remove(); }
    overlay.onclick = close; $(".ytz-x").onclick = close;

    var pageRows = null, pageRowsPdf = false, pageFailed = 0;
    function failNote() { return pageFailed > 0 ? " (" + pageFailed + " tezde sunucu yoğunluğu nedeniyle veri eksik kalmış olabilir.)" : ""; }
    function getPageRows() {
      var wantPdf = $("#ytz-pdflink").checked;
      if (pageRows && pageRowsPdf === wantPdf) return Promise.resolve(pageRows);
      setL("Metaveriler indiriliyor…");
      return fetchAllMetadata(theses, function (d, n) { setP(Math.round(100 * d / n)); setL("Metaveri: " + d + " / " + n); }, wantPdf).then(function (res) { pageRows = res.rows; pageRowsPdf = wantPdf; pageFailed = res.failed; return res.rows; });
    }

    $("#ytz-report").onclick = function () { setBusy(true); setP(0); getPageRows().then(function (r) { exportReport(r, "Bu sayfadaki arama"); setL("Özet rapor oluşturuldu (yeni sekme)." + failNote()); idle(); }); };
    $("#ytz-reportall").onclick = function () { setBusy(true); setL("Rapor hazırlanıyor…"); dbGetAll().then(function (r) { exportReport(r, "Biriktirilen tüm aramalar"); setL("Özet rapor oluşturuldu (yeni sekme)."); idle(); }); };
    $("#ytz-meta").onclick = function () { setBusy(true); setP(0); getPageRows().then(function (r) { exportData(r, "Tez_Metaverileri", fmt()); setL("Bitti · " + r.length + " tez aktarıldı." + failNote()); idle(); }); };
    $("#ytz-text").onclick = function () { setBusy(true); setP(0); getPageRows().then(function (r) { setP(0); setL("Metinler indiriliyor…"); return downloadTexts(r, setP, setL); }).then(function () { setL(plabel.textContent + "\nTamamlandı."); idle(); }); };
    $("#ytz-accum").onclick = function () { setBusy(true); setP(0); getPageRows().then(function (r) { setL("Biriktirmeye ekleniyor…"); return dbPutRows(r); }).then(function () { setL("Bu aramadaki tezler biriktirmeye eklendi." + failNote()); idle(); }); };
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
        etiket: $("#ytz-f-etiket").value.trim().toLocaleLowerCase("tr"),
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
