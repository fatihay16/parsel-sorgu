# -*- coding: utf-8 -*-
"""İş Takip kayıtlarından ilk ada/parsel bilgisini ayıklar."""
import re

def ada_parsel_ayikla(deger):
    metin = str(deger or "").strip()
    if not metin:
        return None
    # Ada/parsel: "Atatürk 727/2-3" -> 727 / 2
    m = re.search(r"(?<!\d)(\d+)\s*/\s*(\d+)", metin)
    if m:
        return {"ada": m.group(1), "parsel": m.group(2)}
    # Ada belirtilmemiş parsel dizisi: "383-404" -> 0 / 383
    m = re.search(r"(?<!\d)(\d+)\s*(?:[-,;]|$)", metin)
    if m:
        return {"ada": "0", "parsel": m.group(1)}
    return None

def acik_isleri_getir(connection):
    """Tescillendi, askıda ve iptal kayıtlarını dışarıda bırakarak açık işleri getirir."""
    return connection.execute("""
        SELECT id, mahalle, ada_parsel
        FROM kayit
        WHERE lower(trim(coalesce(onay,''))) NOT LIKE 'tescillendi%'
          AND lower(trim(coalesce(onay,''))) NOT LIKE 'askıda%'
          AND lower(trim(coalesce(onay,''))) NOT LIKE 'iptal%'
        ORDER BY id
    """).fetchall()

def parsel_satirlarini_olustur(kayitlar, il="Bursa", ilce="Mustafakemalpaşa"):
    sonuc = []
    for kayit in kayitlar:
        work_id, mahalle, ada_parsel = kayit
        parsel = ada_parsel_ayikla(ada_parsel)
        if not str(mahalle or "").strip() or not parsel:
            continue
        sonuc.append({
            "work_id": work_id, "il": il, "ilce": ilce,
            "mahalle": str(mahalle).strip(),
            "ada": parsel["ada"], "parsel": parsel["parsel"]
        })
    return sonuc

if __name__ == "__main__":
    tests = {
        "Atatürk 727/2-3": ("727", "2"),
        "383-404": ("0", "383"),
        "727/2": ("727", "2"),
        "383": ("0", "383"),
    }
    for value, expected in tests.items():
        got = ada_parsel_ayikla(value)
        assert (got["ada"], got["parsel"]) == expected, (value, got)
    print("4 ayrıştırma testi başarılı.")
