# SJUOZ Word → LaTeX Dönüştürücü

Makale (manuscript) ve kapak mektubu (cover letter) Word dosyalarını
**Science Journal of University of Zakho (SJUOZ)** şablonunda XeLaTeX'e çevirir. Tek seferde:

* SJUOZ biçiminde `.tex` ve **PDF** (ilk sayfa logosu, QR kodlu bilgi kutusu, iki sütun, başlık/altlık),
* Word'deki **bütün resimler ve grafikler** (`figures/`), Overleaf'e yüklenebilir **zip**,
* editör için **kontrol raporu** (eksik resim/tablo, kaybolan metin, atıf sorunları, LaTeX hataları),
* **Crossref XML** (DOI kaydı, resmi şemayla doğrulanmış), **JATS XML** (indeksleme), **BibTeX**.

Bütün bir sayı da tek komutla işlenir; sayfa numaraları makaleden makaleye zincirlenir.

---

## 1. Kurulum ve çalıştırma

### a) Windows: hazır program (Python gerekmez)
GitHub → **Actions** → *SJUOZ Word2LaTeX testleri* → son çalışma → **SJUOZ-Word2LaTeX-Windows**
dosyasını indirin. İçinde iki program vardır:

* `SJUOZ-Word2LaTeX.exe` — pencere arayüzü (çift tıklayın),
* `sjuoz-word2latex.exe` — komut satırı.

PDF'in otomatik derlenmesi için bilgisayarda **MiKTeX** (veya TeX Live) kurulu olmalıdır;
yoksa üretilen zip'i Overleaf'e yükleyin.

### b) Python paketi olarak
```bash
pip install ".[tam]"          # lxml + Pillow + matplotlib + pypdf
sjuoz-gui                     # pencere arayüzü
sjuoz-word2latex makale.docx kapak.docx --volume 14 --issue 4 --month October --first-page 571
```

### c) Kurmadan, eskisi gibi
```bash
pip install lxml Pillow matplotlib
python sjuoz_word2latex.py makale.docx kapak.docx --first-page 571
```

İsteğe bağlı araçlar: EMF/WMF/SVG resimleri için LibreOffice, Inkscape veya ImageMagick;
Arapça/Kürtçe metin için Amiri veya Noto Naskh Arabic yazı tipi (Windows'ta Times New Roman yeterli).

## 2. Pencere arayüzü

![arayüz](docs/arayuz.png)

* **Tek makale** sekmesi: makale ve (isteğe bağlı) kapak mektubunu seçin. Makale no dosya
  adından, çıktı klasörü makalenin yanından otomatik doldurulur.
* **Bütün sayı** sekmesi: sayı klasörünü seçin (bkz. bölüm 5).
* Cilt, sayı, ay, ilk sayfa, tarihler ve Crossref e-postası girilir; bir sonraki açılışta hatırlanır.
* **Dönüştür** → işlem günlüğü renkli akar (uyarılar turuncu, hatalar kırmızı). Bitince
  **Raporu aç / PDF'i aç / Klasörü aç** düğmeleri etkinleşir.

## 3. Komut satırı

```bash
sjuoz-word2latex makale.docx kapak_mektubu.docx \
    --volume 14 --issue 4 --year 2026 --month October --first-page 571 \
    --received "12 May 2026" --accepted "1 Aug 2026" --published "15 Oct 2026" \
    --depositor-email editor@uoz.edu.krd
```

İki dosyanın sırası önemli değil; kapak mektubu içeriğinden tanınır.

| Seçenek | Açıklama |
|---|---|
| `--article-id` | Makale no (DOI, QR kod, dosya adları). Verilmezse dosya adındaki sayı. |
| `--volume`, `--issue`, `--year`, `--month` | Cilt, sayı, yıl, ay. |
| `--first-page`, `--last-page` | Sayfa aralığı. `--last-page` verilmezse LaTeX hesaplar. |
| `--received`, `--accepted`, `--published` | Tarihler (ör. `12 May 2026`; varsayılan `XX`). |
| `--article-type` | `Original Article` (varsayılan), `Review Article` … |
| `--depositor-name`, `--depositor-email` | Crossref gönderen bilgisi. |
| `--wide-figs 3,5` / `all` | İki sütunu kaplayacak şekiller. |
| `--landscape-tables 2,3` | Yatay sayfaya konacak tablolar (Word'de yatay bölümdekiler ve 10+ sütunlular zaten yatay). |
| `--keep-heading-case`, `--keep-title-case` | Başlıkları BÜYÜK harfe çevirme. |
| `--main-font`, `--header-font` | Varsayılan Times New Roman / Garamond; yoksa benzeri seçilir. |
| `--assets KLASÖR` | Başka bir `SJUOZ.png` / `orcid.png` kullan. |
| `--no-compile`, `--no-zip`, `--no-export`, `--no-report`, `--no-links` | İlgili adımı atla. |
| `--strict` | Raporda HATA varsa çıkış kodu 2 (toplu işlemler/CI için). |
| `--sayi KLASÖR` | Bütün sayıyı işle (bölüm 5). |

## 4. Ne üretir?

```
1993/
├── 1993.tex                  ← SJUOZ şablonu; resimleri figures/… yoluyla çağırır
├── 1993.pdf                  ← xelatex iki kez çalıştırılır
├── figures/                  ← Word'deki bütün resimler ve grafikler (Figure1.png, Figure6a.png …)
├── SJUOZ.png, orcid.png
├── 1993_kontrol_raporu.html  ← editör için kontrol raporu
├── 1993_crossref.xml         ← DOI kaydı (Crossref 5.3.1, şemayla doğrulanmış)
├── 1993_jats.xml             ← JATS 1.3: ön bilgiler, gövde, şekiller, tablolar, denklemler, kaynaklar
├── 1993_kaynaklar.bib        ← kaynakça BibTeX alanlarına ayrılmış
└── 1993_atif.bib             ← bu makaleye atıf için BibTeX
1993_latex.zip                ← .tex + figures/ + logolar + PDF (Overleaf'e doğrudan yüklenir)
```

### Word'den neler aktarılır?

* **Ön bilgiler:** başlık, yazarlar ve üst simge numaraları, affiliasyonlar (aynı olanlar
  birleştirilir), sorumlu yazar, özet, anahtar kelimeler; **kapak mektubundan** ORCID, e-posta, telefon.
* **Gövde:** numaralı bölümler/alt bölümler, kalın/italik/alt-üst simge, listeler, dipnotlar,
  bağlantılar, son bölümler (Acknowledgements, Ethical Statement, Author Contributions …), kaynakça.
* **Denklemler:** Word denklemleri (OMML) LaTeX'e çevrilir; `SEQ` alan kodlu, `)1(` ya da `(1)`
  biçimindeki numaralar tanınır.
* **Şekiller:** PNG/JPG/TIFF/GIF/BMP/EMF/WMF/SVG (gerekirse PNG'ye çevrilir), **Word grafikleri**
  (veriden yeniden çizilir), gruplanmış/metin kutusundaki resimler, **alt şekiller (a)(b)(c)**
  (kenarlıksız tabloyla dizilmiş etiketli resimler, aynı paragraftaki ya da ardışık paragraflardaki resimler).
* **Tablolar:** birleştirilmiş hücreler, iki satırlı başlıklar, hücre içi resimler; **34 satırdan uzun
  tablolar sayfalara bölünür** (başlık her sayfada tekrarlanır); Word'de **yatay bölümdeki** tablo ve
  şekiller ile 10+ sütunlu tablolar **yatay sayfaya** konur.
* **Sağdan sola metin:** Kürtçe (Sorani) ve Arapça kelime ve paragraflar doğru yön ve yazı tipiyle.
* **Bağlantılar:** (Yazar *et al.*, 2020), Yazar (2019), [3] gibi atıflar kaynakçaya; "Figure 2",
  "Table 1", "Equation 5" ilgili yere tıklanabilir bağlantı olur.

## 5. Bütün sayıyı tek seferde işlemek

```
sayi_14_4/
├── sayi.json            ← isteğe bağlı
├── 1993/  makale.docx  kapak_mektubu.docx
├── 2003/  ...
└── 2002/  ...
```

```bash
sjuoz-word2latex --sayi sayi_14_4 --depositor-email editor@uoz.edu.krd
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

Makaleler bu sırayla derlenir ve sayfa numaraları zincirlenir. `sayi_14_4/_cikti/` içinde:
her makalenin klasörü, zip'i ve raporu; `sayi_raporu.html` (bütün makalelerin durumu tek sayfada),
`icindekiler.csv` (Excel'de açılır), `sayi_tam.pdf` (bütün sayı) ve **`sayi_crossref.xml`**
(bütün sayının DOI kaydı tek dosyada).

## 6. Kontrol raporu

Her dönüşümde `<makale no>_kontrol_raporu.html` üretilir; genel durum **TAMAM**, **KONTROL**
veya **HATA** olur. Rapor, Word dosyasının kendisinden yapılan sayımları çıktıyla karşılaştırır:

| Kontrol | Ne yapar |
|---|---|
| Resim / tablo / denklem sayıları | Word'deki ile LaTeX'teki aynı mı; "Figure 9"a atıf var ama 8 şekil mi var? |
| Metin bütünlüğü | Word'deki **her paragraf ve her tablo hücresi** PDF'te var mı? Kaybolan metin HATA. |
| Atıflar | Kaynakçada bulunamayan atıflar, atıf yapılmamış kaynaklar |
| Ön bilgiler | Özet, anahtar kelime, yazar, ORCID, sorumlu yazar e-postası |
| Derleme | LaTeX hataları, bulunamayan dosyalar, PDF'te görünmeyen karakterler, taşan satır/içerik |
| Crossref | XML resmi şemaya uyuyor mu; DOI kaydı için eksik bilgi (e-posta, yayın tarihi, son sayfa) |
| Uyarılar | Makale ile kapak mektubu arasındaki farklar vb. |

## 7. Proje yapısı

```
sjuoz_word2latex/
├── sjuoz/                 ← paket
│   ├── cli.py             komut satırı giriş noktası
│   ├── gui.py             pencere arayüzü
│   ├── pipeline.py        dönüştürme akışı ve seçenekler
│   ├── issue.py           bütün sayıyı işleme
│   ├── docx.py            Word okuyucu (paragraf, tablo, bölüm, dipnot)
│   ├── omml.py            Word denklemleri → LaTeX
│   ├── parse.py           kapak mektubu ve makale yapısı
│   ├── render.py          biçimli metin → LaTeX, atıf bağlantıları
│   ├── figures.py         resimler, alt şekiller, Word grafikleri
│   ├── writer.py          gövdeyi LaTeX'e yazma
│   ├── template.py        SJUOZ şablonu
│   ├── export.py          Crossref, JATS, BibTeX
│   ├── report.py          kontrol raporu
│   ├── common.py          ortak yardımcılar
│   └── data/              SJUOZ.png, orcid.png, Crossref şemaları
├── sjuoz_word2latex.py    eski kullanım için giriş noktası
├── packaging/             Windows .exe derleme betikleri (PyInstaller)
├── tests/                 test seti (bölüm 8)
├── ornek_cikti_1993/      1993 makalesi için örnek çıktı
└── pyproject.toml
```

Dergi bilgileri (ISSN, DOI öneki, yayıncı, lisans) `sjuoz/export.py` içindeki `JOURNAL`
sözlüğündedir. `SJUOZ.png` örnek PDF'ten çıkarıldığı için düşük çözünürlüklüdür (143×72);
orijinali varsa `sjuoz/data/SJUOZ.png` yerine koyun.

## 8. Testler

```bash
pip install ".[test]"
python -m pytest tests -v                  # xelatex yoksa derleme testleri atlanır
python -m pytest tests --update-golden     # bilerek yapılan değişiklikten sonra beklenen çıktıları güncelle
```

| Test makalesi | Neyi sınar |
|---|---|
| `1993` | Gerçek SJUOZ makalesi + kapak mektubu: yazar-yıl atıfları, SEQ alanlı denklemler, birleştirilecek affiliasyonlar, ORCID |
| `sentetik_temel` | Heading stilleri, listeler, satır içi/ayrı denklem, tablo, şekil, özel karakterler, kapak mektubu yok |
| `sentetik_grafik` | Word sütun grafiği ve EMF resim |
| `sentetik_numerik` | IEEE tarzı `[n]` atıflar, üst simgesiz yazar satırı, satır içi "Abstract:", "Index Terms—", dipnot, birleştirilmiş başlıklı tablo |
| `sentetik_uzun_tablo` | 80 satırlık tablo: sayfalara bölünmeli, satır kaybolmamalı |
| `sentetik_alt_sekil` | Alt şekiller: düzen tablosu, aynı paragrafta ve ardışık paragraflarda resimler |
| `sentetik_yatay` | Word'deki yatay bölüm ve 12 sütunlu tablo |
| `sentetik_rtl` | Kürtçe (Sorani) ve Arapça paragraf ve kelimeler |

Her makale için: üretilen `.tex` saklanmış doğru çıktıyla (`tests/golden/`) birebir
karşılaştırılır; yazar/ORCID/şekil/tablo/denklem/kaynak sayıları; PDF hatasız derlenmeli ve hiçbir
şey taşmamalı; Crossref XML şemaya uymalı; JATS sayıları makaleyle tutarlı olmalı. Ayrıca raporun
gerçek bir kaybı yakaladığı (çıktıdan bilerek paragraf, resim ve tablo satırı silinir), sayı modunun
sayfa zinciri ve birim testleri vardır. Yapay test makaleleri `tests/fixtures/make_synthetic.py`
ile yeniden üretilebilir.

GitHub'da her push'ta: kurulum, kod denetimi (pyflakes), bütün testler, örnek PDF ve raporlar,
ayrıca Windows `.exe` derlenip bir makaleyle denenir.

**Test setine yeni makale eklemek:** `tests/fixtures/<ad>/` klasörüne `.docx` dosyalarını ve
`case.json`'u koyun (diğerlerini örnek alın), PDF'i üretip gözle kontrol edin, sonra
`python -m pytest tests --update-golden`. Dergiye gelen sorunlu her makale test setine
eklenirse, aynı sorun bir daha sessizce geri gelemez.

## 9. Windows programını elle derlemek

```powershell
pip install pyinstaller lxml Pillow matplotlib pypdf
powershell -ExecutionPolicy Bypass -File packaging\build_windows.ps1   # dist\ altında iki .exe
```

## 10. Önceki sürümlere göre düzeltilenler

Eski `docx2sjuoz.py` ve elle hazırlanmış `.tex` ile karşılaştırıldığında: özet "ABASTRACT"
yazım hatasına rağmen bulunuyor; denklemlere `SEQ Equation \* ARABIC` sızmıyor ve (1)–(6)
numaralı; `ī` doğru (`\bar{i}`), "white region" boşluklu; aynı affiliasyonlar birleşiyor; ORCID'ler
kapak mektubundan geliyor; kapak mektubu ayrı sayfa olarak basılmıyor; Şekil 6 iki kez basılmıyor;
bütün şekil başlıkları bulunuyor; bölümler Word'deki gibi 1–5 numaralı, REFERENCES numarasız;
çift DOI'ler temizleniyor; e-posta bağlantısı çalışıyor; 21 kaynağın hepsine atıf bağlantısı var;
ilk sayfadaki bilgi kutusu her durumda doğru yerde.

`ornek_cikti_1993/` klasöründe 1993 makalesi için üretilmiş örnek çıktı bulunur (PDF bu ortamda
Times New Roman yerine Liberation Serif ile derlendi).
