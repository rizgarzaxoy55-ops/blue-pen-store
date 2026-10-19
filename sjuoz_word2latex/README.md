# SJUOZ Word → LaTeX Dönüştürücü

`sjuoz_word2latex.py`, makale (manuscript) Word dosyasını ve kapak mektubu (cover letter)
Word dosyasını okuyup **Science Journal of University of Zakho** şablonunda tek bir
XeLaTeX dosyası üretir.

## Kurulum

```bash
pip install lxml          # zorunlu
pip install Pillow        # isteğe bağlı (TIFF/GIF/BMP resimleri için)
```

## Kullanım

```bash
python sjuoz_word2latex.py makale.docx kapak_mektubu.docx -o 1993 \
    --article-id 1993 --volume 14 --issue 4 --year 2026 --month October \
    --first-page 571 --received "12 May 2026" --accepted "1 Aug 2026" --published "15 Oct 2026" \
    --assets klasor_logolar     # SJUOZ.png ve orcid.png buradan kopyalanır
cd 1993
xelatex 1993.tex
xelatex 1993.tex          # ikinci çalıştırma son sayfa numarası için gerekli
```

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
| `--assets` | `SJUOZ.png` ve `orcid.png` klasörü. `orcid.png` yoksa ORCID simgesi çizilir. |

## Nereden ne alınıyor?

* **Makaleden:** başlık, yazarlar ve üst simge numaraları, affiliasyonlar (aynı olanlar
  birleştirilir ve numaralar yeniden düzenlenir), sorumlu yazar e-posta/telefon, özet,
  anahtar kelimeler, numaralı bölümler/alt bölümler, paragraflar (kalın, italik, alt/üst simge,
  bağlantılar), denklemler, tablolar, şekiller ve başlıkları, son bölümler
  (Acknowledgements, Ethical Statement, Author Contributions, Conflict of Interest, Funding)
  ve kaynakça.
* **Kapak mektubundan:** her yazarın ORCID, e-posta, telefon ve affiliasyon bilgisi
  (yazarlar ada göre eşleştirilir; makale ile çelişen bilgi varsa uyarı verilir).

## Önceki çıktıya göre düzeltilen hatalar

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
