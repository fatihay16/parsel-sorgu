# -*- coding: utf-8 -*-
"""İş Takip kayıtlarından ada/parsel bilgisini güvenli biçimde ayıklar.

Örnekler:
  727/2-3   -> ada=727, parsel=2
  383-404   -> ada=0,   parsel=383
  727/2     -> ada=727, parsel=2

Bir kayıttaki ilk parsel dışında kalan numaralar aktarılmaz.
"""
import re

def ada_parsel_ayikla(deger):
    metin = str(deger or "").strip()
    metin = re.sub(r"\s+", "", metin)
    if not metin:
        return None

    # Ada/parsel biçimi: 727/2-3, 727/2,3 veya 727/2
    m = re.match(r"^(\d+)\s*/\s*(\d+)(?:[-,;].*)?$", metin)
    if m:
        return {"ada": m.group(1), "parsel": m.group(2)}

    # Ada sıfır biçimi: 383-404, 383,404 veya tek parsel 383
    m = re.match(r"^(\d+)(?:[-,;]\d+)*$", metin)
    if m:
        return {"ada": "0", "parsel": m.group(1)}

    return None


def acik_isleri_getir(connection):
    """İş Takip'teki açık kayıtları döndürür; kapalı/iptal kayıtları dışarıda bırakır."""
    return connection.execute("""
        SELECT id, mahalle, ada_parsel
        FROM kayit
        WHERE lower(trim(coalesce(onay,''))) NOT LIKE 'tescillendi%'
          AND lower(trim(coalesce(onay,''))) NOT LIKE 'askıda%'
          AND lower(trim(coalesce(onay,''))) NOT LIKE 'iptal%'
        ORDER BY id
    """).fetchall()


def parsel_satirlarini_olustur(kayitlar, il="Bursa", ilce="Mustafakemalpaşa"):
    """Açık işlerden yalnızca her kaydın ilk parselini aktarır."""
    sonuc = []
    for kayit in kayitlar:
        mahalle = str(kayit.get("mahalle") or "").strip()
        ayiklanan = ada_parsel_ayikla(kayit.get("ada_parsel"))
        if not mahalle or not ayiklanan:
            continue
        sonuc.append({
            "work_id": kayit.get("id"),
            "il": il,
            "ilce": ilce,
            "mahalle": mahalle,
            "ada": ayiklanan["ada"],
            "parsel": ayiklanan["parsel"],
        })
    return sonuc


if __name__ == "__main__":
    ornekler = {
        "727/2-3": ("727", "2"),
        "383-404": ("0", "383"),
        "727/2": ("727", "2"),
        "383": ("0", "383"),
    }
    for giris, beklenen in ornekler.items():
        sonuc = ada_parsel_ayikla(giris)
        assert (sonuc["ada"], sonuc["parsel"]) == beklenen, (giris, sonuc)
    print("Ada/parsel ayrıştırma testleri başarılı.")
