# SJUOZ Word → LaTeX Dönüştürücü

`sjuoz_word2latex.py`, makale (manuscript) Word dosyasını ve kapak mektubu (cover letter)
Word dosyasını okuyup **Science Journal of University of Zakho** şablonunda tek bir
XeLaTeX dosyası üretir.

## Kurulum

```bash
pip install lxml          # zorunlu
pip install Pillow        # önerilir (TIFF/GIF/BMP resimleri, kenar kırpma)
pip install matplotlib    # önerilir (Word grafiklerini resme çevirmek için)
```

EMF/WMF/SVG resimleri için bilgisayarda LibreOffice, Inkscape veya ImageMagick'ten biri
kurulu olmalıdır. PDF'in otomatik derlenmesi için XeLaTeX (MiKTeX / TeX Live) gerekir.

## Kullanım

```bash
python sjuoz_word2latex.py makale.docx kapak_mektubu.docx \
    --volume 14 --issue 4 --year 2026 --month October --first-page 571 \
    --received "12 May 2026" --accepted "1 Aug 2026" --published "15 Oct 2026"
```

Tek komut her şeyi yapar:

```
1993/
├── 1993.tex          ← resimleri figures/… yoluyla çağırır
├── 1993.pdf          ← xelatex kuruluysa otomatik derlenir (iki kez)
├── SJUOZ.png, orcid.png
└── figures/          ← Word'deki TÜM resimler ve grafikler
    ├── Figure1.png
    ├── Figure2.png   (Word grafiği → PNG)
    └── …
1993_latex.zip        ← yukarıdakilerin hepsi tek zip'te (Overleaf'e doğrudan yüklenebilir)
```

* Word'e eklenmiş resimler (PNG, JPG, TIFF, GIF, BMP, EMF, WMF, SVG…) çıkarılır ve gerekirse PNG'ye çevrilir.
* Word grafikleri (sütun, çubuk, çizgi, alan, pasta, dağılım) verisinden yeniden çizilir.
* Gruplanmış resimler, metin kutusundaki resimler ve tablo hücresindeki resimler de alınır.
* Aynı resim iki kez eklenmişse bir kez kaydedilir.
* 34 satırdan uzun tablolar sayfalara bölünür; başlık satırı her sayfada tekrarlanır.

İki dosyanın sırası önemli değil; kapak mektubu içeriğine göre otomatik tanınır.
Kapak mektubu verilmezse ORCID bilgileri olmadan çalışır.

| Seçenek | Açıklama |
|---|---|
| `--article-id` | Makale no (DOI ve QR kod). Verilmezse dosya adındaki sayı alınır. |
| `--volume`, `--issue`, `--year`, `--month` | Cilt, sayı, yıl, ay. |
| `--first-page`, `--last-page` | Sayfa aralığı. `--last-page` verilmezse LaTeX otomatik hesaplar. |
| `--received/--accepted/--published` | Tarihler (varsayılan `XX`). |
| `--article-type` | `Original Article` (varsayılan), `Review Article` … |
| `--keep-heading-case` | Ana bölüm başlıklarını BÜYÜK harfe çevirme. |
| `--keep-title-case` | Makale başlığını BÜYÜK harfe çevirme. |
| `--main-font`, `--header-font` | Varsayılan: Times New Roman / Garamond. Bilgisayarda yoksa benzer yazı tipi otomatik seçilir. |
| `--assets` | `SJUOZ.png` ve `orcid.png` klasörü (verilmezse makalenin, programın ve çalışma klasöründe aranır). `orcid.png` yoksa ORCID simgesi çizilir. |
| `--no-compile` | PDF derleme (varsayılan: xelatex varsa derlenir). |
| `--no-zip` | Zip üretme (varsayılan: zip üretilir). |
| `--wide-figs` | İki sütunu kaplayacak şekiller: `3,5` veya `all`. |
| `--no-links` | Atıf/şekil/tablo/denklem bağlantılarını ekleme. |

Dosya adı verilmeden çalıştırılırsa (`python sjuoz_word2latex.py`) dosya seçme penceresi açılır.

## Logolar

`SJUOZ.png` (sağ üst köşedeki dergi logosu) ve `orcid.png` bu klasörde hazırdır; program
bunları kendiliğinden çıktı klasörüne kopyalar. Bu dosyalar örnek PDF'ten çıkarıldığı için
düşük çözünürlüklüdür (143×72 piksel); elinizde daha yüksek çözünürlüklü orijinali varsa
aynı adla bu klasöre koyun.

## Nereden ne alınıyor?

* **Makaleden:** başlık, yazarlar ve üst simge numaraları, affiliasyonlar (aynı olanlar
  birleştirilir ve numaralar yeniden düzenlenir), sorumlu yazar e-posta/telefon, özet,
  anahtar kelimeler, numaralı bölümler/alt bölümler, paragraflar (kalın, italik, alt/üst simge,
  bağlantılar), denklemler, tablolar, şekiller ve başlıkları, son bölümler
  (Acknowledgements, Ethical Statement, Author Contributions, Conflict of Interest, Funding)
  ve kaynakça.
* **Kapak mektubundan:** her yazarın ORCID, e-posta, telefon ve affiliasyon bilgisi
  (yazarlar ada göre eşleştirilir; makale ile çelişen bilgi varsa uyarı verilir).

## Bütün sayıyı tek seferde işlemek

Her makale bir alt klasöre konur (klasör adı makale numarasıdır):

```
sayi_14_4/
├── sayi.json            ← isteğe bağlı
├── 1993/  makale.docx  kapak_mektubu.docx
├── 2003/  ...
└── 2002/  ...
```

```bash
python sjuoz_word2latex.py --sayi sayi_14_4
```

`sayi.json` (verilmezse klasörler ada göre sıralanır, değerler komut satırından alınır):

```json
{
  "cilt": 14, "sayi": 4, "yil": 2026, "ay": "October", "ilk_sayfa": 571,
  "makaleler": [
    {"klasor": "1993", "gelis": "12 May 2026", "kabul": "1 Aug 2026", "yayin": "15 Oct 2026"},
    {"klasor": "2003", "tur": "Review Article"},
    {"klasor": "2002", "secenekler": ["--wide-figs", "3"]}
  ]
}
```

Makaleler bu sırayla derlenir ve **sayfa numaraları zincirlenir** (bir makale 576'da bitiyorsa
sonraki 577'den başlar). `sayi_14_4/_cikti/` içinde şunlar oluşur:

* her makalenin klasörü, zip'i ve kontrol raporu
* `sayi_raporu.html` — bütün makalelerin sayfa aralıkları, durumları ve bakılacak noktaları tek sayfada
* `icindekiler.csv` — makale no, başlık, yazarlar, sayfalar, DOI (Excel'de açılır; OJS/Crossref girişi için)
* `sayi_tam.pdf` — bütün sayı tek PDF'te (`pip install pypdf` veya `pdfunite` gerekir)

Sayfa zinciri için XeLaTeX gerekir; yoksa her makale `ilk_sayfa`'dan başlar.

## Kontrol raporu

Her dönüşümden sonra `<makale no>_kontrol_raporu.html` üretilir ve özeti ekrana yazılır.
Genel durum **TAMAM**, **KONTROL** (bakılması gereken şeyler var) veya **HATA** olur.
Rapor, Word dosyasının kendisinden yapılan sayımları LaTeX/PDF çıktısıyla karşılaştırır:

| Kontrol | Ne yapar |
|---|---|
| Resimler / tablolar / denklemler | Word'deki sayı ile LaTeX'teki sayı aynı mı; metinde "Figure 9"a atıf var ama 8 şekil mi var? |
| Başlıksız şekil | Başlığı (caption) bulunamayan şekiller |
| Metin bütünlüğü | Word'deki **her paragraf ve her tablo hücresi** PDF'te (PDF yoksa .tex'te) var mı? Kaybolan metin HATA verir. |
| Atıflar | Kaynakçada bulunamayan atıflar, metinde atıf yapılmamış kaynaklar |
| Ön bilgiler | Özet, anahtar kelime, yazar, ORCID, sorumlu yazar e-postası |
| Derleme | LaTeX hataları, bulunamayan dosyalar, yazı tipinde olmayan (PDF'te görünmeyen) karakterler, sütundan taşan satırlar (.tex satır numarasıyla), sayfadan taşan içerik |
| Uyarılar | Makale ile kapak mektubu arasındaki farklar vb. |

`--strict` verilirse raporda HATA olduğunda program 2 çıkış koduyla biter (toplu işlemler için).

## Testler

```bash
pip install lxml Pillow matplotlib pytest
python -m pytest tests -v                  # bütün testler (xelatex yoksa derleme testleri atlanır)
python -m pytest tests --update-golden     # bilerek yapılan değişiklikten sonra beklenen çıktıları güncelle
```

Test seti (`tests/fixtures/`):

| Makale | Neyi sınar |
|---|---|
| `1993` | Gerçek SJUOZ makalesi + kapak mektubu: yazar-yıl atıfları, SEQ alan kodlu denklemler, birleştirilecek affiliasyonlar, ORCID |
| `sentetik_temel` | Heading stilleri, madde/numaralı listeler, satır içi ve ayrı denklem, tablo, şekil, özel karakterler (`_ & %`), kapak mektubu yok |
| `sentetik_grafik` | Word sütun grafiği (chart) ve EMF resim |
| `sentetik_uzun_tablo` | 80 satırlık tablo: sayfalara bölünmeli, başlık her sayfada tekrarlanmalı, satır kaybolmamalı |
| `sentetik_numerik` | IEEE tarzı `[n]` atıflar, "Review Article", üst simgesiz yazar satırı, tek affiliasyon, satır içi "Abstract:", "Index Terms—", dipnot, birleştirilmiş başlık hücreli tablo |

Her makale için üç şey denetlenir:

1. **Golden karşılaştırma:** üretilen `.tex`, daha önce kontrol edilip `tests/golden/` içine
   saklanmış çıktıyla birebir karşılaştırılır. Kodda yapılan bir değişiklik başka bir makaleyi
   bozarsa test farkı satır satır gösterir.
2. **Sayımlar:** `case.json`'daki beklenen yazar, ORCID, affiliasyon, şekil, tablo, denklem,
   kaynak ve bağlantı sayıları.
3. **Derleme:** PDF hatasız derlenmeli, sayfa sayısı beklendiği gibi olmalı, raporda HATA
   olmamalı ve hiçbir şey sayfadan taşmamalı.

Ayrıca sayı testleri (sıra ve sayfa zinciri), birim testleri (denklem çevirici, yazar ayrıştırma, atıf eşleştirme, log analizi) ve
raporun gerçekten kaybı yakaladığını gösteren bir test (çıktıdan bir paragraf ve bir resim
bilerek silinir) vardır.

GitHub'da her push'ta `.github/workflows/sjuoz-word2latex.yml` bütün testleri çalıştırır ve
test makalelerinin PDF'lerini ve kontrol raporlarını indirilebilir dosya olarak ekler.

### Test setine yeni makale eklemek

1. `tests/fixtures/<ad>/` klasörü açın; makale ve (varsa) kapak mektubu `.docx` dosyalarını koyun.
2. Aynı klasöre `case.json` yazın (diğerlerini örnek alın): komut satırı seçenekleri ve beklenen sayılar.
3. `python sjuoz_word2latex.py tests/fixtures/<ad>/*.docx ...` ile PDF'i üretin ve **gözle kontrol edin**.
4. Çıktı doğruysa `python -m pytest tests --update-golden` ile beklenen `.tex`'i kaydedin.

Dergiye gelen sorunlu her makale test setine eklenirse, aynı sorun bir daha sessizce geri gelemez.

## Bağlantılar

Metindeki atıflar — (Yazar *et al.*, 2020), (A & B, 2016), Yazar (2019), [3] gibi — kaynakçadaki
kayda, "Figure 2 / Table 1 / Equation 5" ifadeleri ilgili şekil/tablo/denkleme tıklanabilir
bağlantı olur. Kaynakçada bulunamayan atıflar ve metinde hiç atıf yapılmamış kaynaklar
uyarı olarak yazılır.

## Eski `docx2sjuoz.py` koduna göre düzeltilen hatalar

1993 makalesi ve kapak mektubuyla iki kod da çalıştırıldı. Eski kodda görülen hatalar ve
yeni koddaki durumu:

| Eski kodda | Yeni kodda |
|---|---|
| "ABASTRACT" yazım hatası yüzünden özet bulunamıyor, uyarı veriyor | Özet bulunuyor (`Abstract`, `ABASTRACT`, `Abstract:` …) |
| Denklemlerde `SEQ Equation \* ARABIC` yazısı çıkıyor, denklem numarası yok | Alan kodları atlanıyor, denklemler (1)–(6) numaralı |
| Denklem (1) metin içinde `)1(` olarak bozuk kalıyor | Numaralı ayrı denklem |
| `ī` (üst çizgi) yerine `î` çıkıyor | `\bar{i}` |
| "white region" → "whiteregion" | Boşluk korunuyor |
| Aynı iki affiliasyon ayrı ayrı yazılıyor (1, 2, 3) | Birleştiriliyor (1, 1, 2) |
| ORCID eklenmiyor (kapak mektubu kullanılmıyor) | ORCID'ler kapak mektubundan ekleniyor |
| Kapak mektubu "Click or tap here to enter text." yazılarıyla ayrı sayfalar olarak basılıyor (9 sayfa) | Kapak mektubu sadece bilgi kaynağı; makale 6 sayfa |
| Şekil 6 resmi iki kez basılıyor | Bir kez |
| Şekil 7 ve 8 başlıkları bulunamıyor, Şekil 7 başlığı düz metin oluyor | Bütün başlıklar doğru şekilde |
| Sadece "Introduction" ve "Experimental Results" numaralı bölüm, diğerleri numarasız alt bölüm | Word'deki numaralandırma: 1–5 |
| Telefon "(Tel: …))" çift parantez | Düzgün |
| 21 kaynaktan 6'sına atıf bağlantısı kurulamıyor ("Mohammed E. Safi & Abbas", "Wisam H Ali et al.") | 21 / 21 bağlantı |
| `\enlargethispage` kısa özetlerde ilk sayfa düzenini bozabiliyor | İlk sayfa metin alanı güvenli şekilde kısaltılıyor |

Eski koddaki iyi özellikler (atıf bağlantıları, PDF derleme, zip, dosya seçme penceresi,
geniş şekiller) yeni koda da eklendi.

## Önceki `.tex` dosyasına göre düzeltilen hatalar

* Denklemlerdeki `SEQ Equation \* ARABIC` alan kodları artık LaTeX'e sızmıyor; Word'deki
  denklem numaraları `equation` ortamıyla doğru numaralanıyor (`)1(` gibi ters numaralar dahil).
* Denklem (1) satır içi bozuk metin yerine düzgün bir denklem olarak çıkıyor.
* Bütün bölümler `1.` diye numaralanmıyor; Word'deki numaralandırmaya göre 1, 2, 3 … ve
  alt bölümler 1.1 … şeklinde.
* `REFERENCES` başlığı numarasız.
* `https://doi.org/https://doi.org/…` gibi çift DOI'ler temizleniyor, DOI'ler `DOI:` bağlantısı oluyor.
* E-posta bağlantısı `mailto:` ile çalışıyor.
* Aynı affiliasyonlar birleştiriliyor (1 ve 2 aynıysa tek affiliasyon).
* Yazar katkıları tek paragrafta, adlar kısaltılmış (M. E. Safi …) ve fazladan tırnak/numara yok.
* Tablolar sayfa genişliğine sığıyor (`tabularx`), uzun başlıklar satıra bölünüyor.
* Tekrarlanan paketler kaldırıldı; ilk sayfadaki dergi bilgi kutusu, metin ne kadar uzun
  olursa olsun ilk sayfanın altında yer alıyor ve elle `\pagebreak`/`\newgeometry` gerekmiyor.
* Word'deki grup/metin kutusu içindeki resimler iki kez çıkarılmıyor.

## Dikkat

Dönüştürücü uyarıları ekrana yazar (ör. makale ve kapak mektubundaki e-postalar farklı).
Derlemeden önce bunlara bakın. Word dosyasındaki içerik hataları (ör. metinde yanlış tablo
numarasına atıf) olduğu gibi aktarılır.

`ornek_cikti_1993/` klasöründe 1993 numaralı makale için üretilmiş örnek `.tex`, şekiller ve
PDF bulunur (PDF, Times New Roman yerine Liberation Serif yazı tipiyle derlendi).
