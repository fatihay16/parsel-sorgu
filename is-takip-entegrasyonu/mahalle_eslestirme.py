# -*- coding: utf-8 -*-
"""İş Takip mahalle adlarını değiştirmeden TKGM mahalle koduna eşleştirir.

ÖNEMLİ:
- source_mahalle: İş Takip tablosundaki özgün değer; asla değiştirilmez.
- tkgm_mahalle: TKGM'nin kendi idari listesindeki resmi ad.
- mahalle_id: TKGM parsel API'sine gönderilen kod.
- Eşleşme il + ilçe kapsamında yapılır. Belirsiz eşleşme reddedilir.
"""
import re
import unicodedata

API = "https://cbsapi.tkgm.gov.tr/megsiswebapi.v3.1/api"
_IDARI_CACHE = {}

def norm_mahalle(value):
    value = unicodedata.normalize("NFKD", str(value or "").strip().casefold())
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = value.replace("ı", "i")
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return " ".join(value.split())

def feature_items(data):
    if isinstance(data, dict) and data.get("type") == "FeatureCollection":
        return [f.get("properties", {}) for f in data.get("features", [])]
    if isinstance(data, list):
        return data
    return []

def get_cached_json(url, getter):
    if url not in _IDARI_CACHE:
        _IDARI_CACHE[url] = getter(url)
    return _IDARI_CACHE[url]

def item_name(item, keys):
    for key in keys:
        value = item.get(key)
        if value not in (None, ""):
            return str(value).strip()
    return ""

def item_id(item):
    for key in ("id", "mahalleKodu", "ilceKodu", "ilKodu"):
        value = item.get(key)
        if value not in (None, ""):
            return str(value)
    return ""

def strict_find(items, name_keys, wanted, label):
    wanted_norm = norm_mahalle(wanted)
    exact = [(item_name(x, name_keys), item_id(x)) for x in items
             if norm_mahalle(item_name(x, name_keys)) == wanted_norm]
    exact = [(n, i) for n, i in exact if i]
    if len(exact) == 1:
        return exact[0]
    if len(exact) > 1:
        raise ValueError(f"{label} adı TKGM listesinde birden fazla kayıtla eşleşti: {wanted}")
    # Yalnızca tek ve çok yakın (normalize edilmiş tam eşleşme sonrası) sonuç varsa
    # harf/boşluk farklarını düzelt; parça eşleşmesiyle yanlış mahalle seçme.
    raise ValueError(
        f"'{wanted}' adı seçilen il/ilçe altındaki TKGM {label} listesinde bulunamadı. "
        "TKGM'deki resmi mahalle adını seçerek eşleştirme yapın."
    )

def idari_liste(tur, ust_id, getter):
    if tur == "il":
        url = f"{API}/idariYapi/ilListe"
    elif tur == "ilce":
        url = f"{API}/idariYapi/ilceListe/{ust_id}"
    elif tur == "mahalle":
        url = f"{API}/idariYapi/mahalleListe/{ust_id}"
    else:
        raise ValueError("Bilinmeyen idari liste türü: " + str(tur))
    return get_cached_json(url, getter)

def mahalleleri_getir(il, ilce, getter):
    iller = feature_items(idari_liste("il", None, getter))
    _il_adi, il_id = strict_find(iller, ("text", "ad", "name", "ilAdi"), il, "il")
    ilceler = feature_items(idari_liste("ilce", il_id, getter))
    _ilce_adi, ilce_id = strict_find(ilceler, ("text", "ilceAdi", "ad", "name"), ilce, "ilçe")
    mahalleler = feature_items(idari_liste("mahalle", ilce_id, getter))
    return mahalleler

def mahalle_eslestir(source_mahalle, il, ilce, getter, resmi_mahalle=None):
    """İş Takip mahalle adını TKGM listesine eşler, özgün değeri korur.

    resmi_mahalle verilirse kullanıcı tarafından seçilmiş TKGM adıdır.
    """
    items = mahalleleri_getir(il, ilce, getter)
    if resmi_mahalle:
        secim = resmi_mahalle
    else:
        secim = source_mahalle
    wanted = norm_mahalle(secim)
    matches = []
    for item in items:
        name = item_name(item, ("text", "mahalleAdi", "ad", "name"))
        code = item_id(item)
        if name and code and norm_mahalle(name) == wanted:
            matches.append((name, code))
    if len(matches) == 1:
        name, code = matches[0]
    elif len(matches) > 1:
        raise ValueError(f"'{secim}' TKGM listesinde birden fazla mahalleyle eşleşti.")
    else:
        names = sorted({item_name(x, ("text", "mahalleAdi", "ad", "name"))
                        for x in items if item_name(x, ("text", "mahalleAdi", "ad", "name"))})
        raise ValueError(
            f"İş Takip mahalle adı '{source_mahalle}' TKGM listesinde otomatik eşleşmedi. "
            "İş Takip'teki mahalle adı değiştirilmedi. Sorgu için TKGM listesinden doğru "
            "mahalleyi seçin. Bu ilçedeki TKGM mahalle sayısı: " + str(len(names))
        )
    return {
        "source_mahalle": str(source_mahalle or "").strip(),
        "mahalle": str(source_mahalle or "").strip(),
        "tkgm_mahalle": name,
        "mahalle_id": code,
        "il": il,
        "ilce": ilce,
    }

def query_parsel(row, getter):
    """Sorgu satırını TKGM endpoint'ine gönderir; mahalle alanını değiştirmez."""
    matched = mahalle_eslestir(
        row.get("source_mahalle", row.get("mahalle", "")),
        row.get("il", ""),
        row.get("ilce", ""),
        getter,
        resmi_mahalle=row.get("tkgm_mahalle")
    )
    url = f"{API}/parsel/{matched['mahalle_id']}/{row['ada']}/{row['parsel']}"
    data = getter(url)
    if isinstance(data, dict) and data.get("Message"):
        raise RuntimeError(str(data["Message"]))
    if not isinstance(data, dict) or data.get("type") != "Feature":
        raise RuntimeError("TKGM parsel servisi beklenen sonucu döndürmedi.")
    result = dict(row)
    result.update(matched)
    result["geometry"] = data.get("geometry") or {}
    props = data.get("properties") or {}
    result["alan"] = str(props.get("alan") or "")
    result["nitelik"] = str(props.get("nitelik") or "")
    return result
