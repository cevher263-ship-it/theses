// Uzantı simgesine tıklanınca aracı, açık olan YÖK Tez sayfasına enjekte eder.
// Kütüphaneler (XLSX/JSZip/FileSaver) uzantı içine gömülüdür; CDN'e ihtiyaç yoktur.
// Sırayla: gömülü kütüphaneler -> araç. Araç, tanımlı olan kütüphaneleri tekrar yüklemez.
chrome.action.onClicked.addListener(function (tab) {
  if (!tab || !tab.id) return;
  var url = tab.url || "";

  if (!/tez\.yok\.gov\.tr\/UlusalTezMerkezi/i.test(url)) {
    chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: function () {
        alert("Bu aracı YÖK Ulusal Tez Merkezi'nde ARAMA SONUÇ sayfasında açın.\n\n" +
          "1) https://tez.yok.gov.tr/UlusalTezMerkezi/ adresinde aramanızı yapın.\n" +
          "2) Sonuçlar listelendiğinde uzantı simgesine tekrar tıklayın.");
      }
    }).catch(function (e) { console.error("YÖKTez uyarı hatası:", e); });
    return;
  }

  chrome.scripting.executeScript({
    target: { tabId: tab.id },
    func: function () {
      try {
        // Zaten açıksa yeniden yükleme; paneli aç.
        if (window.__yokTezAraci__) { window.__yokTezAraci__.open(); return; }
        var files = ["lib/xlsx.bundle.js", "lib/jszip.min.js", "lib/FileSaver.min.js", "yoktez-veri-indir.js"];
        var i = 0;
        (function next() {
          if (i >= files.length) return;
          var s = document.createElement("script");
          s.src = chrome.runtime.getURL(files[i]);
          s.onload = function () { s.remove(); i++; next(); };
          s.onerror = function () { console.error("YÖKTez yükleme hatası (araç CDN'e düşecek):", files[i]); s.remove(); i++; next(); };
          (document.head || document.documentElement).appendChild(s);
        })();
      } catch (e) { console.error("YÖKTez enjeksiyon hatası:", e); }
    }
  }).catch(function (e) { console.error("YÖKTez executeScript hatası:", e); });
});
