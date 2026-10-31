# İş Takip – Parsel Sorgu düzeltme alanı

Bu dal, mevcut `fatihay16/parsel-sorgu` projesine İş Takip entegrasyonu için ayrılmıştır.

## Bu klasörde
- `parsel_aktarim_duzeltmesi.py`: İş Takip'teki açık kayıtları filtreleme ve ada/parsel bilgisini ayıklama yardımcıları.
- Ayrıştırma kuralları:
  - `727/2-3` → ada 727, parsel 2 (yalnızca ilk parsel)
  - `383-404` → ada 0, parsel 383 (yalnızca ilk parsel)
  - `727/2` → ada 727, parsel 2
- Tescillendi, askıda ve iptal durumundaki kayıtlar açık işler listesine alınmaz.

## Önemli
Bu dal, önceki entegrasyon paketindeki kaybolan il/ilçe/mahalle açılır listelerini geri getirme çalışması için ayrılmıştır. Açılır listelerin tam arayüz bağlantısı henüz bu yardımcı dosyada yapılmış değildir; eski çalışan TKGM listesi arayüzü korunarak ana entegrasyon modülüne bağlanmalıdır. Bu dosya tek başına tam masaüstü uygulaması değildir.

## Kontrol
`python parsel_aktarim_duzeltmesi.py` komutu ayrıştırma örneklerini test eder.
