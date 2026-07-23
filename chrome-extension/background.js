// Uzantı simgesine tıklanınca aracı, açık olan YÖK Tez sayfasına enjekte eder.
// Yöntem: sayfaya, uzantı içindeki yoktez-veri-indir.js'yi işaret eden bir <script>
// etiketi eklenir. Böylece araç sayfanın kendi bağlamında (MAIN world) çalışır ve
// sayfadaki referenceData'ya erişir. (Kanıtlanmış, tüm Chrome sürümlerinde çalışan yöntem.)
// Araç ikinci kez enjekte edilirse kendi __yokTezAraci__ koruması paneli yeniden açar.
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
        var s = document.createElement("script");
        s.src = chrome.runtime.getURL("yoktez-veri-indir.js");
        s.onload = function () { s.remove(); };
        (document.head || document.documentElement).appendChild(s);
      } catch (e) { console.error("YÖKTez enjeksiyon hatası:", e); }
    }
  }).catch(function (e) { console.error("YÖKTez executeScript hatası:", e); });
});
