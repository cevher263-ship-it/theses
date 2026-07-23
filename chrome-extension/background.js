// Uzantı simgesine tıklanınca aracı, açık olan YÖK Tez sonuç sayfasına enjekte eder.
// Araç sayfanın kendi bağlamında (MAIN world) çalışır; böylece sayfadaki referenceData'ya erişir.
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
    }).catch(function () {});
    return;
  }
  chrome.scripting.executeScript({
    target: { tabId: tab.id },
    world: "MAIN",
    files: ["yoktez-veri-indir.js"]
  }).catch(function (e) { console.error("Enjeksiyon hatası:", e); });
});
