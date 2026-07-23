## YÖK Tez Merkezi Veri Kazıma Aracı

Bu araç, Yükseköğretim Kurulu (YÖK) Tez Merkezi'nden veri indirmek için geliştirilmiştir.

> ### ⚠️ Yeni arayüz sürümü (2026)
>
> YÖK Tez Merkezi arayüzü kart tabanlı yeni tasarıma geçtiği için orijinal `dist/index.js` artık çalışmıyor. Yeni arayüze uyarlanmış sürüm bu depodaki [`yoktez-veri-indir.js`](yoktez-veri-indir.js) dosyasıdır.
>
> **Kullanım:** Arama sonuç sayfasında (F12 → Console) aşağıdaki satırı yapıştırıp Enter'a basın:
>
> ```js
> document.head.appendChild(Object.assign(document.createElement("script"),{src:"https://cdn.jsdelivr.net/gh/cevher263-ship-it/theses@v1.4/yoktez-veri-indir.js"}))
> ```
>
> Sağ üstte açılan pencereden metaverileri **Excel** olarak veya erişime açık tez metinlerini **PDF/ZIP** olarak indirebilirsiniz.
>
> **2000 sınırını aşma:** YÖK tek aramada en çok 2000 tez listeler. "Biriktirme" bölümündeki **Bu aramayı biriktir** düğmesiyle aramanızı yıl yıl daraltıp her parçayı biriktirebilir, sonda **Tümünü Excel indir** ile hepsini tek dosyada alabilirsiniz (veriler tarayıcıda IndexedDB'de tekrarsız birikir).
>
> *Orijinal araç ve tüm emek [@mytunca](https://github.com/mytunca/theses)'ya aittir; bu sürüm yalnızca yeni arayüz uyumu için uyarlanmıştır.*

### 🧩 Chrome Eklentisi olarak kullanım (en pratik — tek tık)

Konsola kod yapıştırmadan, tek tıkla çalışır:

1. [`YOKTez-Eklenti.zip`](YOKTez-Eklenti.zip) dosyasını indirip bir klasöre çıkarın (ya da `chrome-extension` klasörünü indirin).
2. Chrome'da `chrome://extensions/` adresine gidin, sağ üstten **Geliştirici modu**nu açın.
3. **Paketlenmemiş öğe yükle** → çıkardığınız klasörü seçin.
4. YÖK Tez arama **sonuç sayfasında** araç çubuğundaki uzantı simgesine tıklayın → araç açılır.

> Not: Eklenti yalnızca `tez.yok.gov.tr` üzerinde çalışır, başka veri toplamaz. Kod, konsol sürümüyle birebir aynıdır (`yoktez-veri-indir.js`). Excel/ZIP kütüphaneleri çalışma anında CDN'den yüklenir (internet gerekir).

---

_Aşağıdaki orijinal açıklama eski arayüz içindir:_

YÖK Tez Merkezi'nde sorgu yapıldığında, tezlerin metaverileri (tez no, yazar, yıl, tez adı, üniversite, dil, tez türü ve konu) bir tablo halinde listelenir. Ancak, daha ayrıntılı bilgilere erişmek için her bir tezin detay sayfasına tek tek tıklamak gereklidir. Bu araç sizin yerinize bu işlemi hızlı bir şekilde tamamlayarak veriyi indirilebilecek şekilde size sunar.

Bu araç, araştırmacılar ve veri bilimciler için YÖK  Tez Merkezi'nden veri toplamayı ve analiz etmeyi kolaylaştırmayı amaçlamaktadır.

### Özellikler

✅ Ücretsizdir.  
✅ Üyelik gerektirmez.  
✅ Yazılım bilgisi gerektirmez.  
✅ Tezlerin tüm metaverilerini ve metin dosyalarını indirmenize olanak tanır.  
✅ JavaScript dilini kullanarak tarayıcıda çalışır, ayrı bir yazılım yüklemenizi gerektirmez.

### Kullanım

1.  [YÖK Tez Merkezi](https://tez.yok.gov.tr/UlusalTezMerkezi/)'nde istediğiniz aramayı yapın.
2.  Sonuçların listelendiği https://tez.yok.gov.tr/UlusalTezMerkezi/tezSorguSonucYeni.jsp sayfası açıldığında tarayıcının adres çubuğundaki mevcut adresi silip yerine elle `javascript:` yazın ve arkasına aşağıdaki kodu yapıştırın.
    ```js
    document.head.appendChild(
        Object.assign(
    	document.createElement("script"),
    	{src:"https://cdn.jsdelivr.net/gh/mytunca/theses@latest/dist/index.js"}
        )
    )
    ```
    **Enter** tuşuna bastığınızda sayfaya **Veri İndir** menüsü eklenecektir.
    ![Adres çubuğu ekran kaydı](src/assets/images/screencast1.gif)

	 <details>
	  <summary>Alternatif yöntem</summary>
	
	  #### Yukarıdaki yöntem çalışmazsa
	
	  - https://tez.yok.gov.tr/UlusalTezMerkezi/tezSorguSonucYeni.jsp sayfası açıkken F12 tuşuna basarak tarayıcınızın geliştirici araçlarını açın ve **Console** sekmesine geçin.
	  - Yukarıdaki kodu kopyalayıp konsola yapıştırın, ardından **Enter** tuşuna basın.
	    > **Not**
	    >
	    >Konsola daha önce kod yapıştırmadıysanız yapıştırma engeliyle **(paste protection)** karşılaşmanız olasıdır. Bu engeli kaldırmak için tarayıcınız bir öneride bulunacaktır. Örneğin Google Chrome için **allow pasting** komutunu kullanmanız gerekecektir. Engeli kaldırdıktan sonra kodu tekrar yapıştırın.
	</details>

3.  Menüye tıklandığında açılan diyalog penceresinden o anda listelenmekte olan tezlerin metaverilerini ve metin dosyalarını indirebilirsiniz. (Metin dosyalarını indirme işlemi bir saat civarı sürebilir.)
![UI kullanımı ekran kaydı](src/assets/images/screencast2.gif)

### Chrome Uzantısı Olarak Kullanım
1. [chrome-extension.rar](/chrome-extension.rar) dosyasını bilgisayarınıza [indirip](https://cdn.jsdelivr.net/gh/mytunca/theses@latest/chrome-extension.rar) sıkıştırılmış dosyadan çıkarın.
2. Google Chrome Uzantıları'nı açmak için Google Chrome tarayıcınızda [chrome://extensions/](chrome://extensions/) adresine gidin.
3. Sayfanın sağ üst kısmından **Geliştirici modu**nu aktif hale getirin.
4. **Paketlenmemiş öğe yükle** seçeneğini tıklayarak çıkardığınız klasörü seçin.
5. Uzantı yüklendikten sonra Tez Merkezi Sorgu Sonuç Sayfası'nda uzantıyı tıklayın.
6. Uzantının tez.yok.gov.tr üzerinde her zaman değişiklik yapmasına izin verin. 


### Notlar

1. Bu kodun kullanımı, YÖK Tez Merkezi'nin kullanım koşullarına uygun olmalıdır. Veri toplama ve kullanma konusunda YÖK'ün politikalarını göz önünde bulundurun.
2. YÖK Tez Merkezi'nde yapılan aramalar, tek seferde en fazla 2000 tezin verisini listelemektedir.

   > Tarama sonucunda 13472 kayıt bulundu. 2000 tanesi görüntülenmektedir.

   şeklinde bir uyarıyla karşılaşmanız halinde arama kriterlerinizi genişleterek listelenen tez sayısını 2000'in altına düşürmeye çalışın.

   Örneğin arama yaptığınız kriterleri sağlayan toplam 5000 civarı tez varsa onar yıllık periyotlar halinde sorgulama yaparsanız tek seferde listelenen tez sayısı muhtemelen 2000'in altına düşecektir.

3. Bu kod yalnızca Google Chrome (125.0.6422.142) üzerinde test edilmiştir. 19.06.2024 tarihi itibarıyla çalışmaktadır. Ancak sayfanın kodlarında değişiklik yapılması durumunda kodun çalışmayı durdurma ihtimali her zaman mevcuttur.

### Sorumluluk Reddi
Bu yazılım, "olduğu gibi" ve "mevcut haliyle" sağlanmaktadır. Yazar, bu yazılımın kullanılmasından kaynaklanabilecek herhangi bir zarar veya kayıptan sorumlu değildir. Kullanıcılar, bu yazılımı kullanmanın tüm risklerini kabul eder ve tüm sorumluluğu üstlenir.

**1. Garanti Yokluğu:** Bu yazılım, belirli bir amaca uygunluk, ticari elverişlilik ve ihlal etmeme dahil ancak bunlarla sınırlı olmamak üzere, açık veya zımni herhangi bir garanti olmaksızın sağlanmaktadır. Yazar, yazılımın hatasız veya kesintisiz çalışacağına dair hiçbir garanti vermez.

**2. Kullanım Riski:** Kullanıcı, bu yazılımı kullanmanın tüm risklerini üstlenir. Bu yazılımın kullanımı sırasında ortaya çıkabilecek herhangi bir veri kaybı, iş kaybı veya diğer zararlar için yazar(lar) sorumlu değildir.

**3. Güncellemeler ve Değişiklikler:** Bu yazılımın sürekli güncellenmesi veya desteklenmesi garanti edilmez. Yazar, bu yazılımı herhangi bir zamanda değiştirme veya durdurma hakkını saklı tutar.

**4. Üçüncü Taraf Bileşenler:** Bu yazılım, üçüncü taraf bileşenler veya kitaplıklar içerebilir. Bu üçüncü taraf bileşenlerin kullanımından kaynaklanan herhangi bir sorun veya zarar için yazar sorumlu değildir.

**5. Kullanım Kısıtlamaları:** Bu yazılımı kullanarak, tüm ilgili yasalara ve düzenlemelere uyacağınızı kabul edersiniz. Yazılımın yasadışı veya zararlı amaçlarla kullanılması yasaktır.

**6. Destek:** Bu yazılımın kullanımıyla ilgili herhangi bir destek veya yardım sağlama yükümlülüğü yoktur. Ancak, topluluk veya yazar gönüllü olarak yardım sağlayabilir.
