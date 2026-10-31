# -*- coding: utf-8 -*-
"""Toplu Parsel Sorgulama v2 | TKGM

Her parsel için:
  * parsel AKTİF ise  -> alan ve nitelik gösterilir,
  * parsel PASİF ise  -> "pasife alınmıştır" ibaresi, işlem türü/tarihi ve oluşan YENİ parsellerin
                         tamamı (Parsel Sorgu'daki hareket listesi gibi) gösterilir.
Günlük istek sayacı vardır; üst sınıra gelince toplu sorgu durur. İl/ilçe/mahalle kimlikleri bir kez
çözülüp saklanır (eskiden her parsel için 3 ek istek atılıyordu).
"""
import os, re, json, time, sqlite3, threading
from pathlib import Path
from datetime import datetime, date, timedelta
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from PIL import Image, ImageDraw, ImageFont

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "Sonuclar"
DB_PATH = APP_DIR / "parsel_kayitlari.db"
API = "https://cbsapi.tkgm.gov.tr/megsiswebapi.v3.1/api"
WEB = "https://parselsorgu.tkgm.gov.tr/"
HEADERS = {
    "Accept": "application/json",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36",
    "Origin": "https://parselsorgu.tkgm.gov.tr",
    "Referer": WEB,
}
TIMEOUT = 25
VARSAYILAN_LIMIT = 90       # günlük istek üst sınırı (TKGM sınırı 100 olduğu için pay bırakıldı)
VARSAYILAN_YENILE_GUN = 3   # Aktif görünen parseller kaç günde bir yeniden kontrol edilsin
BEKLEME_SN = 1.2            # istekler arası bekleme

YENI_SUTUNLAR = {"pasif_mesaj": "TEXT", "islem_turu": "TEXT", "islem_tarihi": "TEXT", "yeni_parseller": "TEXT",
                 "hareketler": "TEXT", "ham_parsel": "TEXT", "ham_hareket": "TEXT", "devam_parseller": "TEXT"}
SUTUNLAR = ["il", "ilce", "mahalle", "ada", "parsel", "mahalle_id", "alan", "nitelik", "koordinatlar", "map_file",
            "status", "detail", "updated"] + list(YENI_SUTUNLAR)


class LimitHatasi(RuntimeError):
    pass


class ErisimHatasi(RuntimeError):
    pass


# ------------------------------------------------------------------ veritabanı
def db():
    con = sqlite3.connect(DB_PATH)
    con.execute("""CREATE TABLE IF NOT EXISTS parcels (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        il TEXT, ilce TEXT, mahalle TEXT, ada TEXT, parsel TEXT,
        mahalle_id TEXT, alan TEXT, nitelik TEXT, koordinatlar TEXT,
        map_file TEXT, status TEXT, detail TEXT, updated TEXT,
        UNIQUE(il,ilce,mahalle,ada,parsel)
    )""")
    var = {r[1] for r in con.execute("PRAGMA table_info(parcels)")}
    for k, t in YENI_SUTUNLAR.items():  # eski veritabanını bozmadan yeni sütunları ekle
        if k not in var:
            con.execute(f"ALTER TABLE parcels ADD COLUMN {k} {t}")
    con.execute("CREATE TABLE IF NOT EXISTS ayar (k TEXT PRIMARY KEY, v TEXT)")
    con.execute("CREATE TABLE IF NOT EXISTS sayac (gun TEXT PRIMARY KEY, adet INTEGER)")
    con.execute("CREATE TABLE IF NOT EXISTS liste_onbellek (url TEXT PRIMARY KEY, veri TEXT)")
    con.commit()
    return con


def ayar_oku(k, varsayilan=""):
    con = db()
    r = con.execute("SELECT v FROM ayar WHERE k=?", (k,)).fetchone()
    con.close()
    return r[0] if r and r[0] is not None else varsayilan


def ayar_yaz(k, v):
    con = db()
    con.execute("INSERT OR REPLACE INTO ayar VALUES (?,?)", (k, str(v)))
    con.commit()
    con.close()


def limit_degeri():
    try:
        return max(int(ayar_oku("gunluk_limit", VARSAYILAN_LIMIT)), 1)
    except ValueError:
        return VARSAYILAN_LIMIT


def sayac_oku():
    con = db()
    r = con.execute("SELECT adet FROM sayac WHERE gun=?", (date.today().isoformat(),)).fetchone()
    con.close()
    return r[0] if r else 0


def sayac_artir():
    con = db()
    con.execute("INSERT INTO sayac VALUES (?,1) ON CONFLICT(gun) DO UPDATE SET adet=adet+1", (date.today().isoformat(),))
    con.commit()
    con.close()


# ------------------------------------------------------------------ servis istekleri
def get_json(url, say=True):
    """say=True: istek günlük sayaca işlenir ve üst sınırda LimitHatasi atılır."""
    if say:
        if sayac_oku() >= limit_degeri():
            raise LimitHatasi(f"Günlük üst sınıra ({limit_degeri()} istek) ulaşıldı. Yarın devam edin ya da sınırı artırın.")
        sayac_artir()
    r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
    if r.status_code == 403:
        raise ErisimHatasi("TKGM erişimi 403 ile reddetti (erişim/limit kısıtı). Bir süre sonra yeniden deneyin.")
    if r.status_code == 429:
        raise ErisimHatasi("TKGM geçici istek sınırına ulaşıldı (429). Daha sonra tekrar deneyin.")
    r.raise_for_status()
    return r.json()


def get_json_onbellekli(url):
    """İl/ilçe/mahalle listeleri nadiren değişir: bir kez çekilip saklanır, sayaca işlenmez."""
    con = db()
    r = con.execute("SELECT veri FROM liste_onbellek WHERE url=?", (url,)).fetchone()
    con.close()
    if r:
        return json.loads(r[0])
    veri = get_json(url, say=False)
    con = db()
    con.execute("INSERT OR REPLACE INTO liste_onbellek VALUES (?,?)", (url, json.dumps(veri, ensure_ascii=False)))
    con.commit()
    con.close()
    return veri


def feature_items(data):
    if isinstance(data, dict) and data.get("type") == "FeatureCollection":
        return [f.get("properties", {}) for f in data.get("features", [])]
    if isinstance(data, list):
        return data
    return []


def norm_ad(x):
    return " ".join(str(x or "").strip().casefold().replace("i̇", "i").split())


def find_id(items, name_keys, wanted):
    w = norm_ad(wanted)
    for item in items:
        name = next((item.get(k) for k in name_keys if item.get(k) not in (None, "")), "")
        if norm_ad(name) == w:
            return item.get("id") or item.get("mahalleKodu") or item.get("ilceKodu")
    matches = []
    for item in items:
        name = next((item.get(k) for k in name_keys if item.get(k) not in (None, "")), "")
        if w and (w in norm_ad(name) or norm_ad(name) in w):
            matches.append(item.get("id") or item.get("mahalleKodu") or item.get("ilceKodu"))
    if len(matches) == 1:
        return matches[0]
    raise RuntimeError(f"'{wanted}' idari birimi TKGM listesinde tekil olarak bulunamadı.")


def mahalle_id_bul(row):
    if row.get("mahalle_id"):
        return row["mahalle_id"]
    con = db()  # aynı il/ilçe/mahalle için daha önce çözülmüş kimlik varsa tekrar istek atma
    kayitlar = con.execute("SELECT il,ilce,mahalle,mahalle_id FROM parcels WHERE mahalle_id IS NOT NULL AND mahalle_id!=''").fetchall()
    con.close()
    hedef = (norm_ad(row["il"]), norm_ad(row["ilce"]), norm_ad(row["mahalle"]))
    for il, ilce, mah, mid in kayitlar:
        if (norm_ad(il), norm_ad(ilce), norm_ad(mah)) == hedef:
            return mid
    provinces = feature_items(get_json_onbellekli(f"{API}/idariYapi/ilListe"))
    il_id = find_id(provinces, ["text", "ad", "name", "ilAdi"], row["il"])
    districts = feature_items(get_json_onbellekli(f"{API}/idariYapi/ilceListe/{il_id}"))
    ilce_id = find_id(districts, ["text", "ilceAdi", "ad", "name"], row["ilce"])
    neighborhoods = feature_items(get_json_onbellekli(f"{API}/idariYapi/mahalleListe/{ilce_id}"))
    return find_id(neighborhoods, ["text", "mahalleAdi", "ad", "name"], row["mahalle"])


# ------------------------------------------------------------------ aktif / pasif / hareket çözümleme
_KATLA = str.maketrans("İıĞğÜüŞşÖöÇç", "IiGgUuSsOoCc")  # 1'e 1 harf eşleme: indeksler korunur


def katla(s):
    return str(s if s is not None else "").translate(_KATLA).lower()


NO_RE = re.compile(r"^\s*\d+\s*/\s*\d+\s*$")
ISLEM_RE = re.compile(r"(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\s+tarihli\s+(.+?)\s+islemi")


def metinleri_topla(obj, cikti=None):
    cikti = [] if cikti is None else cikti
    if isinstance(obj, dict):
        for v in obj.values():
            metinleri_topla(v, cikti)
    elif isinstance(obj, list):
        for v in obj:
            metinleri_topla(v, cikti)
    elif isinstance(obj, str):
        cikti.append(obj)
    return cikti


def pasif_ibaresi(obj):
    """Yanıtın herhangi bir yerinde 'pasife alınmıştır' içeren metni bulur."""
    for s in metinleri_topla(obj):
        if "pasife alinmistir" in katla(s):
            return s.strip()
    return ""


def islem_ayikla(metin):
    """'20/10/2026 Tarihli Tevhit İşlemi (Tsm)' -> ('20.10.2026', 'Tevhit')"""
    k = katla(metin)
    m = ISLEM_RE.search(k)
    if not m:
        return "", ""
    tarih = re.sub(r"[/-]", ".", metin[m.start(1):m.end(1)])
    tur = metin[m.start(2):m.end(2)].strip()
    return tarih, tur[:1].upper() + tur[1:]


def _durum_degeri(v):
    k = katla(v).strip()
    return "Aktif" if k == "aktif" else "Pasif" if k == "pasif" else None


def hareket_satirlari(obj):
    """Hareket listesindeki satırları (açıklama, durum, taşınmaz no) şekil gözetmeden bulur:
    'Aktif/Pasif' değeri ile '727/13' gibi bir numarası olan her kayıt bir satırdır (iç içe olsa da)."""
    satirlar, goruldu = [], set()

    def gez(x):
        if isinstance(x, dict):
            durum = no = None
            ack = ""
            ada = parsel = None
            for k, v in x.items():
                kk = katla(k)
                if isinstance(v, bool):
                    if "pasif" in kk and v:
                        durum = durum or "Pasif"
                    elif "aktif" in kk:
                        durum = durum or ("Aktif" if v else "Pasif")
                elif isinstance(v, str):
                    d = _durum_degeri(v)
                    if d and durum is None:
                        durum = d
                    elif NO_RE.match(v) and no is None:
                        no = re.sub(r"\s+", "", v)
                    elif "tarihli" in katla(v) and not ack:
                        ack = v.strip()
                    elif v.strip() in ("-", "–") and not ack:
                        ack = "-"
                if kk in ("ada", "adano") and isinstance(v, (int, str)) and str(v).strip().isdigit():
                    ada = str(v).strip()
                if kk in ("parsel", "parselno") and isinstance(v, (int, str)) and str(v).strip().isdigit():
                    parsel = str(v).strip()
            if no is None and ada and parsel:
                no = f"{ada}/{parsel}"
            if durum and no:
                anahtar = (no, durum, ack)
                if anahtar not in goruldu:
                    goruldu.add(anahtar)
                    satirlar.append({"aciklama": ack, "durum": durum, "no": no})
            for v in x.values():
                gez(v)
        elif isinstance(x, list):
            for v in x:
                gez(v)

    gez(obj)
    return satirlar


def props_durum(props):
    """Parsel yanıtında doğrudan 'Aktif/Pasif' bilgisi varsa onu verir (alan adı bilinmediği için geniş taranır)."""
    if not isinstance(props, dict):
        return None
    for k, v in props.items():
        kk = katla(k)
        if isinstance(v, bool):
            if "pasif" in kk and v:
                return "Pasif"
            if "aktif" in kk and not v:
                return "Pasif"
        elif isinstance(v, str) and ("durum" in kk or "aktif" in kk or "pasif" in kk or "statu" in kk):
            d = _durum_degeri(v)
            if d:
                return d
    return None


def zemin_id_bul(props):
    for k, v in (props or {}).items():
        kk = katla(k)
        if "zemin" in kk and "id" in kk and isinstance(v, (int, str)) and str(v).strip():
            return str(v).strip()
    return ""


def analiz_et(ada, parsel, props, hareket):
    """Parsel yanıtı (+ varsa hareket yanıtı) -> durum, ibare, işlem, yeni parseller, hareket satırları."""
    kendi = f"{str(ada).strip()}/{str(parsel).strip()}"
    kaynaklar = [props] + ([hareket] if hareket is not None else [])
    satirlar, ibare = [], ""
    for v in kaynaklar:
        for s in hareket_satirlari(v):
            if s not in satirlar:
                satirlar.append(s)
        ibare = ibare or pasif_ibaresi(v)
    kendi_satir = next((s for s in satirlar if s["no"] == kendi), None)
    durum = kendi_satir["durum"] if kendi_satir else None
    if durum is None and ibare:
        durum = "Pasif"
    if durum is None:
        durum = props_durum(props)
    tarih = tur = ""
    for metin in [ibare] + [s["aciklama"] for s in ([kendi_satir] if kendi_satir else []) + satirlar]:
        if metin and not tarih:
            tarih, tur = islem_ayikla(metin)
    if durum == "Pasif" and not ibare and tarih:
        ibare = f"Bu taşınmaz {tarih} Tarihli {tur} İşlemi ile pasife alınmıştır."
    yeni = []
    for s in satirlar:
        if s["durum"] == "Aktif" and s["no"] != kendi and s["no"] not in yeni:
            yeni.append(s["no"])
    return {"durum": durum, "ibare": ibare, "islem_tarihi": tarih, "islem_turu": tur, "yeni": yeni, "satirlar": satirlar}


def hareket_gecerli(h):
    if h is None:
        return False
    if isinstance(h, dict) and (h.get("Message") or h.get("message") or h.get("error")):
        return False
    return isinstance(h, (dict, list))


def hareket_adresi(sablon, row, props, mahalle_id):
    zemin = zemin_id_bul(props)
    if "{zemin_id}" in sablon and not zemin:
        raise RuntimeError("Hareket adresi {zemin_id} istiyor ama parsel yanıtında zemin kimliği bulunamadı.")
    degerler = {"mahalle_id": mahalle_id, "ada": row["ada"], "parsel": row["parsel"], "zemin_id": zemin,
                "il_id": (props or {}).get("ilId", ""), "ilce_id": (props or {}).get("ilceId", "")}
    for k, v in degerler.items():
        sablon = sablon.replace("{" + k + "}", str(v))
    return sablon


def geometri_ozeti(geom):
    coords = []
    raw = (geom or {}).get("coordinates") or []
    if (geom or {}).get("type") == "Polygon" and raw:
        coords = raw[0]
    elif (geom or {}).get("type") == "MultiPolygon" and raw and raw[0]:
        coords = raw[0][0]
    try:
        return "; ".join(f"{p[1]:.7f}, {p[0]:.7f}" for p in coords)
    except (TypeError, IndexError, ValueError):
        return ""


def area_text(v):
    return "" if v is None else str(v).strip()


def query_one(row):
    """Bir parseli sorgular; row sözlüğünü sonuçla doldurur. (Parsel + gerekirse hareket: en çok 2 istek)"""
    mahalle_id = mahalle_id_bul(row)
    data = get_json(f"{API}/parsel/{mahalle_id}/{row['ada']}/{row['parsel']}")
    if isinstance(data, dict) and data.get("Message"):
        raise RuntimeError(str(data["Message"]))
    if not isinstance(data, dict) or data.get("type") != "Feature":
        raise RuntimeError("TKGM beklenmeyen yanıt döndürdü; parsel bulunamadı veya servis erişimi kısıtlı.")
    props = data.get("properties") or {}
    geom = data.get("geometry") or {}
    sonuc = analiz_et(row["ada"], row["parsel"], props, None)
    hareket, ham_h, hareket_hata = None, "", ""
    sablon = ayar_oku("hareket_url")
    if sablon and (sonuc["durum"] is None or (sonuc["durum"] == "Pasif" and not sonuc["yeni"])):
        try:
            hareket = get_json(hareket_adresi(sablon, row, props, mahalle_id))
            ham_h = json.dumps(hareket, ensure_ascii=False)
            if not hareket_gecerli(hareket):
                hareket_hata = "Hareket servisi geçerli bir liste döndürmedi."
                hareket = None
            else:
                sonuc = analiz_et(row["ada"], row["parsel"], props, hareket)
        except (LimitHatasi, ErisimHatasi):
            raise
        except Exception as e:
            hareket_hata = f"Hareket servisi: {e}"
    ayrinti = []
    if sonuc["durum"] is None:
        sonuc["durum"] = "Belirsiz"
            ayrinti.append(hareket_hata or ("Aktif/pasif bilgisi servis yanıtında bulunamadı. "
                                             "'Hareket servisini ayarla' bölümünü kullanın."))
    elif hareket_hata:
        ayrinti.append(hareket_hata)
    if sonuc["durum"] == "Pasif" and not sonuc["yeni"]:
        ayrinti.append("Yeni parsel listesi alınamadı (hareket servisi ayarlı değil veya liste boş).")
    ozet = {k: props.get(k) for k in ("ozet", "mevkii", "pafta", "zeminId", "zeminKmdurum") if props.get(k) not in (None, "")}
    ham = json.dumps({"properties": props, "geometri": (geom or {}).get("type") or None}, ensure_ascii=False)
    devam = []
    devam_hatalari = []
    if sonuc["durum"] == "Pasif":
        for no in sonuc["yeni"]:
            try:
                parca = no.split("/", 1)
                if len(parca) != 2 or not all(x.strip().isdigit() for x in parca):
                    continue
                yeni_ada, yeni_parsel = parca[0].strip(), parca[1].strip()
                yeni_data = get_json(f"{API}/parsel/{mahalle_id}/{yeni_ada}/{yeni_parsel}")
                if not isinstance(yeni_data, dict) or yeni_data.get("type") != "Feature":
                    devam_hatalari.append(f"{no}: TKGM kaydı alınamadı")
                    continue
                yeni_props = yeni_data.get("properties") or {}
                yeni_geom = yeni_data.get("geometry") or {}
                yeni_durum = props_durum(yeni_props)
                if yeni_durum == "Pasif":
                    continue
                devam.append({
                    "ada": yeni_ada, "parsel": yeni_parsel,
                    "durum": yeni_durum or "Belirsiz",
                    "alan": area_text(yeni_props.get("alan")) or "TKGM verisinde bulunamadı",
                    "nitelik": str(yeni_props.get("nitelik") or "TKGM verisinde bulunamadı"),
                    "koordinatlar": geometri_ozeti(yeni_geom),
                    "geometry": yeni_geom
                })
            except (LimitHatasi, ErisimHatasi):
                raise
            except Exception as e:
                devam_hatalari.append(f"{no}: {e}")
    if devam_hatalari:
        ayrinti.append("Devam parselleri: " + "; ".join(devam_hatalari))
    ham = json.dumps({"properties": props, "geometri": (geom or {}).get("type") or None}, ensure_ascii=False)
    row.update({
        "mahalle_id": str(mahalle_id), "alan": area_text(props.get("alan")) or "TKGM verisinde bulunamadı", "nitelik": str(props.get("nitelik") or "TKGM verisinde bulunamadı"),
        "koordinatlar": geometri_ozeti(geom), "geometry": geom, "status": sonuc["durum"],
        "pasif_mesaj": sonuc["ibare"], "islem_turu": sonuc["islem_turu"], "islem_tarihi": sonuc["islem_tarihi"],
        "yeni_parseller": "; ".join(sonuc["yeni"]), "hareketler": json.dumps(sonuc["satirlar"], ensure_ascii=False),
        "ham_parsel": ham[:60000], "ham_hareket": ham_h[:60000],
        "devam_parseller": json.dumps(devam, ensure_ascii=False),
        "detail": " ".join(ayrinti + ([" • ".join(f"{k}: {v}" for k, v in ozet.items())] if ozet else [])),
    })
    return row


def make_map(row, geometry, out_path):
    coords = (geometry or {}).get("coordinates") or []
    if (geometry or {}).get("type") == "MultiPolygon":
        coords = coords[0] if coords else []
    ring = coords[0] if coords and isinstance(coords[0], list) else coords
    if not ring or not isinstance(ring[0], (list, tuple)):
        raise RuntimeError("Harita için parsel geometrisi bulunamadı.")
    xs = [float(p[0]) for p in ring]
    ys = [float(p[1]) for p in ring]
    minx, maxx, miny, maxy = min(xs), max(xs), min(ys), max(ys)
    w, h, margin = 1200, 900, 110
    dx, dy = max(maxx - minx, 1e-9), max(maxy - miny, 1e-9)
    scale = min((w - 2 * margin) / dx, (h - 2 * margin) / dy)
    ox, oy = (w - dx * scale) / 2, (h - dy * scale) / 2
    pts = [(ox + (x - minx) * scale, h - (oy + (y - miny) * scale)) for x, y in zip(xs, ys)]
    im = Image.new("RGB", (w, h), "#f4f6f8")
    d = ImageDraw.Draw(im)
    for x in range(60, w, 120):
        d.line((x, 70, x, h - 70), fill="#e3e7eb", width=1)
    for y in range(70, h, 120):
        d.line((60, y, w - 60, y), fill="#e3e7eb", width=1)
    d.polygon(pts, fill="#a8d8b9", outline="#176b3a")
    d.line(pts + [pts[0]], fill="#176b3a", width=7)
    try:
        font = ImageFont.truetype("arial.ttf", 30)
        small = ImageFont.truetype("arial.ttf", 22)
    except Exception:
        font = small = ImageFont.load_default()
    d.text((60, 25), f"{row['mahalle']} | Ada {row['ada']} / Parsel {row['parsel']}", fill="#17212b", font=font)
    d.text((60, h - 55), f"Alan: {row.get('alan', '-')} m²   •   {row.get('nitelik', '')}", fill="#17212b", font=small)
    d.text((w - 280, h - 55), "TKGM parsel geometrisi", fill="#59636e", font=small)
    im.save(out_path, quality=94)


def hareket_sablonu_yap(url, row):
    """F12'den kopyalanan adresteki mahalle kimliği/ada/parsel değerlerini {mahalle_id},{ada},{parsel} ile değiştirir."""
    from urllib.parse import urlsplit, urlunsplit
    p = urlsplit(url.strip())
    esle = {str(row.get("mahalle_id") or ""): "{mahalle_id}", str(row.get("ada") or ""): "{ada}", str(row.get("parsel") or ""): "{parsel}"}
    esle.pop("", None)
    parcalar = [esle.get(s, s) for s in p.path.split("/")]
    sorgu = "&".join(f"{kv.split('=')[0]}={esle.get(kv.split('=')[1], kv.split('=')[1])}" if "=" in kv else kv
                     for kv in p.query.split("&")) if p.query else ""
    return urlunsplit((p.scheme, p.netloc, "/".join(parcalar), sorgu, "")).replace("%7B", "{").replace("%7D", "}")


def son_kontrol(r):
    try:
        return datetime.fromisoformat(r.get("updated") or "")
    except ValueError:
        return datetime.min


# ------------------------------------------------------------------ arayüz
RENKLER = {"Aktif": "#e3f4e8", "Pasif": "#fde9c9", "Belirsiz": "#fff6cc", "Hata": "#fde2e2"}


class Detay(tk.Toplevel):
    """Parsel Sorgu'daki 'hareket listesi' penceresi gibi: Özet / Tablo / Grafik / Ham veri."""

    def __init__(self, master, row):
        super().__init__(master)
        self.row = row
        self.title(f"{row['ada']} ada {row['parsel']} parsel hareket listesi")
        self.geometry("820x560")
        tk.Label(self, text=f"{row['ada']} ada {row['parsel']} parsel hareket listesi", bg="#337ab7", fg="white",
                 font=("Segoe UI", 15), anchor="w", padx=16, pady=12).pack(fill="x")
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)
        self.ozet_tab, self.tablo_tab, self.grafik_tab, self.ham_tab = (ttk.Frame(nb) for _ in range(4))
        for f, ad in ((self.ozet_tab, "Özet"), (self.tablo_tab, "Tablo"), (self.grafik_tab, "Grafik"), (self.ham_tab, "Ham veri")):
            nb.add(f, text=ad)
        self.satirlar = []
        try:
            self.satirlar = json.loads(row.get("hareketler") or "[]")
        except ValueError:
            pass
        self._ozet()
        self._tablo()
        self._grafik()
        self._ham()

    def _ozet(self):
        r = self.row
        durum = r.get("status") or "Bekliyor"
        renk = {"Aktif": "#dff0d8", "Pasif": "#fcf8e3", "Belirsiz": "#fff6cc", "Hata": "#f2dede"}.get(durum, "#eeeeee")
        if durum == "Pasif":
            metin = r.get("pasif_mesaj") or "Bu taşınmaz pasife alınmıştır."
        elif durum == "Aktif":
            metin = "Bu parsel AKTİF görünüyor."
        elif durum == "Belirsiz":
            metin = "Aktif/pasif durumu belirlenemedi.\n" + (r.get("detail") or "")
        else:
            metin = r.get("detail") or durum
        tk.Label(self.ozet_tab, text=metin, bg=renk, fg="#6b5a2a", wraplength=720, justify="center", padx=16, pady=18,
                 relief="solid", bd=1).pack(fill="x", padx=18, pady=(18, 10))
        satirlar = [("Parsel", f"{r['mahalle']} / {r['ada']} ada {r['parsel']} parsel"),
                    ("Alan (m²)", r.get("alan") or "-"), ("Nitelik", r.get("nitelik") or "-")]
        if durum == "Pasif":
            satirlar.insert(1, ("İşlem", f"{r.get('islem_turu') or '-'}  ({r.get('islem_tarihi') or '-'})"))
            yeni = [x for x in (r.get("yeni_parseller") or "").split("; ") if x]
            satirlar.append((f"Oluşan yeni parseller ({len(yeni)} adet)" if yeni else "Oluşan yeni parseller",
                             ",  ".join(yeni) if yeni else "Alınamadı"))
        for k, v in satirlar:
            f = ttk.Frame(self.ozet_tab)
            f.pack(fill="x", padx=24, pady=3)
            ttk.Label(f, text=k + ":", width=30, font=("Segoe UI", 10, "bold")).pack(side="left")
            ttk.Label(f, text=v, font=("Segoe UI", 10)).pack(side="left")
        if durum == "Pasif":
            ttk.Label(self.ozet_tab, text="Not: Pasif parselin alan/niteliği son kayıtlı bilgidir.", foreground="#777").pack(anchor="w", padx=24, pady=(10, 0))

    def _tablo(self):
        t = ttk.Treeview(self.tablo_tab, columns=("aciklama", "durum", "no"), show="headings", height=10)
        for c, ad, w in (("aciklama", "Açıklama", 380), ("durum", "Durum", 90), ("no", "Taşınmaz No", 120)):
            t.heading(c, text=ad)
            t.column(c, width=w, anchor="w")
        t.tag_configure("Pasif", foreground="#aaaaaa")
        for s in self.satirlar:
            t.insert("", "end", values=(s.get("aciklama") or "-", s.get("durum"), s.get("no")), tags=(s.get("durum"),))
        t.pack(fill="both", expand=True, padx=12, pady=12)
        if not self.satirlar:
            ttk.Label(self.tablo_tab, text="Hareket kaydı alınamadı.").pack()

    def _grafik(self):
        self.cv = tk.Canvas(self.grafik_tab, bg="white", highlightthickness=0)
        self.cv.pack(fill="both", expand=True)
        self.cv.bind("<Configure>", lambda e: self.grafik_ciz())
        self.grafik_ciz()

    def grafik_ciz(self):
        cv, w = self.cv, max(self.cv.winfo_width(), 700)
        cv.delete("all")
        pasif = [s["no"] for s in self.satirlar if s["durum"] == "Pasif"]
        aktif = [s["no"] for s in self.satirlar if s["durum"] == "Aktif"]
        if not self.satirlar:
            cv.create_text(w / 2, 100, text="Grafik için hareket kaydı yok", fill="#888")
            return
        ust_y, alt_y, orta = 60, 190, 125
        kutu = lambda x, y, metin, dolgu, yazi: (
            cv.create_rectangle(x - 48, y - 15, x + 48, y + 15, fill=dolgu, outline="#666"),
            cv.create_text(x, y, text=metin, fill=yazi, font=("Segoe UI", 10)))
        px = {n: w * (i + 1) / (len(pasif) + 1) for i, n in enumerate(pasif)}
        ax = {n: w * (i + 1) / (len(aktif) + 1) for i, n in enumerate(aktif)}
        for p in pasif:
            for a in aktif:
                cv.create_line(px[p], ust_y + 15, px[p], orta, ax[a], orta, ax[a], alt_y - 15, fill="#777", width=2)
        for n, x in px.items():
            kutu(x, ust_y, n, "#d3d3d3", "#000")
        for n, x in ax.items():
            kutu(x, alt_y if pasif else ust_y, n, "#7a9a3a", "#fff")

    def _ham(self):
        t = tk.Text(self.ham_tab, wrap="none", font=("Consolas", 9))
        t.pack(fill="both", expand=True, padx=8, pady=8)
        parca = []
        for ad, v in (("PARSEL YANITI", self.row.get("ham_parsel")), ("HAREKET YANITI", self.row.get("ham_hareket"))):
            try:
                parca.append(f"=== {ad} ===\n" + (json.dumps(json.loads(v), ensure_ascii=False, indent=2) if v else "(yok)"))
            except ValueError:
                parca.append(f"=== {ad} ===\n{v}")
        self.ham_metin = "\n\n".join(parca)
        t.insert("1.0", self.ham_metin)
        ttk.Button(self.ham_tab, text="Panoya kopyala", command=self.kopyala).pack(pady=(0, 8))

    def kopyala(self):
        self.clipboard_clear()
        self.clipboard_append(self.ham_metin)


class HareketAyar(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Hareket servisini ayarla")
        self.geometry("760x470")
        yardim = ("Pasif parselin hareket listesini (yeni parselleri) getiren servis adresi bir kez tanımlanır.\n\n"
                  "1) Tarayıcıda parselsorgu.tkgm.gov.tr açın, F12 tuşuna basın, 'Ağ / Network' sekmesinde 'Fetch/XHR' seçin.\n"
                  "2) Pasif bir parseli sorgulayıp 'Parsel hareketlerini göster' bağlantısına tıklayın.\n"
                  "3) Listede yeni beliren isteğe tıklayın, 'Request URL' adresini kopyalayıp aşağıya yapıştırın.\n"
                  "4) Bu programda o parselin satırını seçip 'Örnek adresi şablona çevir'e basın, sonra Kaydet.\n"
                  "   (Adresteki mahalle kimliği, ada ve parsel değerleri {mahalle_id} {ada} {parsel} ile değiştirilir.)")
        ttk.Label(self, text=yardim, justify="left", wraplength=720).pack(anchor="w", padx=16, pady=14)
        self.var = tk.StringVar(value=ayar_oku("hareket_url"))
        ttk.Entry(self, textvariable=self.var, width=110).pack(padx=16, fill="x")
        f = ttk.Frame(self)
        f.pack(pady=12)
        ttk.Button(f, text="Örnek adresi şablona çevir", command=self.cevir).pack(side="left", padx=6)
        ttk.Button(f, text="Kaydet", command=self.kaydet).pack(side="left", padx=6)
        ttk.Button(f, text="Temizle (servisi kapat)", command=lambda: self.var.set("")).pack(side="left", padx=6)
        ttk.Label(self, text="Şablon örneği:  https://.../api/.../{mahalle_id}/{ada}/{parsel}", foreground="#666").pack(anchor="w", padx=16)

    def cevir(self):
        sec = self.app.tree.selection()
        if not sec:
            messagebox.showinfo("Bilgi", "Önce ana listede örnek parselin satırını seçin.", parent=self)
            return
        self.var.set(hareket_sablonu_yap(self.var.get(), self.app.rows[int(sec[0])]))

    def kaydet(self):
        ayar_yaz("hareket_url", self.var.get().strip())
        messagebox.showinfo("Kaydedildi", "Hareket servisi ayarı kaydedildi.", parent=self)
        self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Toplu Parsel Sorgulama | TKGM")
        self.geometry("1280x780")
        self.minsize(1050, 650)
        self.rows = []
        self.busy = False
        self._style()
        self._ui()
        self.load_saved()
        self.sayac_goster()
        self.after(150, self.illeri_yukle)

    def _style(self):
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except Exception:
            pass
        st.configure("Treeview", rowheight=28, font=("Segoe UI", 9))
        st.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        st.configure("TButton", padding=7, font=("Segoe UI", 9))
        st.configure("Header.TLabel", font=("Segoe UI", 18, "bold"), foreground="#18334d")
        st.configure("Sub.TLabel", font=("Segoe UI", 9), foreground="#586575")

    def _ui(self):
        top = ttk.Frame(self, padding=(18, 14))
        top.pack(fill="x")
        ttk.Label(top, text="Toplu Parsel Sorgulama", style="Header.TLabel").pack(anchor="w")
        ttk.Label(top, text="Aktif parselde alan ve nitelik, pasif parselde pasife alınma ibaresi ve oluşan yeni parseller gösterilir. "
                            "Satıra çift tıklayınca hareket listesi açılır.", style="Sub.TLabel").pack(anchor="w", pady=(4, 0))
        bar = ttk.Frame(self, padding=(18, 5))
        bar.pack(fill="x")
        self.btn(bar, "Excel şablonu oluştur", self.template).pack(side="left", padx=(0, 7))
        self.btn(bar, "Excel içe aktar", self.import_excel).pack(side="left", padx=7)
        self.btn(bar, "Seçili satırı sil", self.delete_selected).pack(side="left", padx=7)
        self.btn(bar, "Listeyi temizle", self.clear_rows).pack(side="left", padx=7)
        self.btn(bar, "Hareket servisini ayarla", lambda: HareketAyar(self)).pack(side="left", padx=7)
        self.btn(bar, "Excel'e aktar", self.export_excel).pack(side="right", padx=(7, 0))
        form = ttk.LabelFrame(self, text="Elle parsel ekle", padding=10)
        form.pack(fill="x", padx=18, pady=8)
        self.vars = {}
        fields = [("İl", "il", 15), ("İlçe", "ilce", 18), ("Mahalle", "mahalle", 20), ("Ada", "ada", 10), ("Parsel", "parsel", 10)]
        for i, (label, key, width) in enumerate(fields):
            ttk.Label(form, text=label).grid(row=0, column=i, sticky="w", padx=5)
            v = tk.StringVar()
            self.vars[key] = v
            if key in ("il", "ilce", "mahalle"):
                widget = ttk.Combobox(form, textvariable=v, width=width)
                widget.grid(row=1, column=i, sticky="ew", padx=5, pady=(3, 0))
                if key == "il":
                    widget.bind("<<ComboboxSelected>>", self.il_secildi)
                elif key == "ilce":
                    widget.bind("<<ComboboxSelected>>", self.ilce_secildi)
                setattr(self, key + "_combo", widget)
            else:
                ttk.Entry(form, textvariable=v, width=width).grid(row=1, column=i, sticky="ew", padx=5, pady=(3, 0))
            form.columnconfigure(i, weight=1)
        self.btn(form, "Mahalle listesini yenile", self.mahalleleri_yenile).grid(row=1, column=6, padx=6)
        self.btn(form, "Listeye ekle", self.add_manual).grid(row=1, column=5, padx=10)
        frame = ttk.Frame(self, padding=(18, 5))
        frame.pack(fill="both", expand=True)
        cols = ("il", "ilce", "mahalle", "ada", "parsel", "durum", "alan", "nitelik", "yeni", "islem", "kontrol")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="extended")
        labels = {"il": "İl", "ilce": "İlçe", "mahalle": "Mahalle", "ada": "Ada", "parsel": "Parsel", "durum": "Parsel durumu",
                  "alan": "Alan (m²)", "nitelik": "Nitelik", "yeni": "Yeni parseller", "islem": "Pasif işlemi", "kontrol": "Son kontrol"}
        widths = {"il": 80, "ilce": 120, "mahalle": 110, "ada": 60, "parsel": 60, "durum": 100, "alan": 90, "nitelik": 120,
                  "yeni": 190, "islem": 170, "kontrol": 120}
        for c in cols:
            self.tree.heading(c, text=labels[c])
            self.tree.column(c, width=widths[c], anchor="w")
        for ad, renk in RENKLER.items():
            self.tree.tag_configure(ad, background=renk)
        vs = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        self.tree.bind("<Double-1>", lambda e: self.detay_ac())
        bottom = ttk.Frame(self, padding=(18, 8))
        bottom.pack(fill="x")
        self.progress = ttk.Progressbar(bottom, mode="determinate")
        self.progress.pack(fill="x", pady=(0, 7))
        self.status = tk.StringVar(value="Hazır. Excel şablonundaki sütunlar: İl, İlçe, Mahalle, Ada, Parsel.")
        ttk.Label(bottom, textvariable=self.status).pack(side="left")
        self.btn(bottom, "Bekleyenleri sorgula", self.start_queries).pack(side="right")
        self.btn(bottom, "Seçili satırları sorgula", lambda: self.start_queries(secili=True)).pack(side="right", padx=7)
        self.sayac_var = tk.StringVar()
        ttk.Label(bottom, textvariable=self.sayac_var, foreground="#18334d", font=("Segoe UI", 9, "bold")).pack(side="right", padx=14)
        self.limit_var = tk.StringVar(value=str(limit_degeri()))
        sp = ttk.Spinbox(bottom, from_=10, to=100, width=5, textvariable=self.limit_var, command=self.limit_degisti)
        sp.pack(side="right")
        sp.bind("<Return>", lambda e: self.limit_degisti())
        sp.bind("<FocusOut>", lambda e: self.limit_degisti())
        ttk.Label(bottom, text="Günlük üst sınır:").pack(side="right", padx=(14, 4))

    def btn(self, parent, text, command):
        return ttk.Button(parent, text=text, command=command)

    @staticmethod
    def liste_adlari(veri):
        alanlar = ("text", "ad", "name", "ilAdi", "ilceAdi", "mahalleAdi")
        sonuc = []
        for item in feature_items(veri):
            if isinstance(item, dict):
                ad = next((str(item[k]).strip() for k in alanlar if item.get(k) not in (None, "")), "")
                if ad and ad not in sonuc:
                    sonuc.append(ad)
        return sorted(sonuc, key=norm_ad)

    def idari_liste(self, tur, ust_id=None):
        if tur == "il":
            url = API + "/idariYapi/ilListe"
        elif tur == "ilce":
            url = API + "/idariYapi/ilceListe/" + str(ust_id)
        else:
            url = API + "/idariYapi/mahalleListe/" + str(ust_id)
        return get_json_onbellekli(url)

    def illeri_yukle(self):
        try:
            iller = self.idari_liste("il")
            self.il_combo["values"] = self.liste_adlari(iller)
        except Exception as e:
            self.status.set("İl listesi alınamadı: " + str(e))

    def il_secildi(self, _event=None):
        try:
            iller = self.idari_liste("il")
            il_id = find_id(feature_items(iller), ["text", "ad", "name", "ilAdi"], self.vars["il"].get())
            ilceler = self.idari_liste("ilce", il_id)
            self.ilce_combo["values"] = self.liste_adlari(ilceler)
            self.vars["ilce"].set("")
            self.mahalle_combo["values"] = ()
            self.vars["mahalle"].set("")
        except Exception as e:
            messagebox.showerror("İlçe listesi alınamadı", str(e))

    def ilce_secildi(self, _event=None):
        self.mahalleleri_yenile()

    def mahalleleri_yenile(self):
        try:
            iller = self.idari_liste("il")
            il_id = find_id(feature_items(iller), ["text", "ad", "name", "ilAdi"], self.vars["il"].get())
            ilceler = self.idari_liste("ilce", il_id)
            ilce_id = find_id(feature_items(ilceler), ["text", "ilceAdi", "ad", "name"], self.vars["ilce"].get())
            mahalleler = self.idari_liste("mahalle", ilce_id)
            self.mahalle_combo["values"] = self.liste_adlari(mahalleler)
            self.status.set(str(len(self.mahalle_combo["values"])) + " mahalle listelendi.")
        except Exception as e:
            messagebox.showerror("Mahalle listesi alınamadı", str(e))

    # ---- sayaç / limit
    def limit_degisti(self):
        try:
            v = max(1, min(int(self.limit_var.get()), 100))
        except ValueError:
            v = VARSAYILAN_LIMIT
        self.limit_var.set(str(v))
        ayar_yaz("gunluk_limit", v)
        self.sayac_goster()

    def sayac_goster(self):
        self.sayac_var.set(f"Bugün: {sayac_oku()} / {limit_degeri()} istek")

    # ---- liste işlemleri
    def add_manual(self):
        row = {k: v.get().strip() for k, v in self.vars.items()}
        if not all(row.values()):
            messagebox.showwarning("Eksik bilgi", "İl, ilçe, mahalle, ada ve parsel alanlarının tamamını doldurun.")
            return
        self.add_row(row)
        for v in self.vars.values():
            v.set("")

    def add_row(self, row):
        key = tuple(str(row.get(k, "")).strip() for k in ("il", "ilce", "mahalle", "ada", "parsel"))
        if any(tuple(str(x.get(k, "")).strip() for k in ("il", "ilce", "mahalle", "ada", "parsel")) == key for x in self.rows):
            return
        row.setdefault("status", "Bekliyor")
        self.rows.append(row)
        self.refresh()
        self.save_row(row)

    def refresh(self):
        for x in self.tree.get_children():
            self.tree.delete(x)
        for i, r in enumerate(self.rows):
            st = r.get("status") or "Bekliyor"
            islem = f"{r.get('islem_turu') or ''} {r.get('islem_tarihi') or ''}".strip()
            kontrol = (r.get("updated") or "")[:16].replace("T", " ")
            self.tree.insert("", "end", iid=str(i), tags=(st,), values=(
                r.get("il", ""), r.get("ilce", ""), r.get("mahalle", ""), r.get("ada", ""), r.get("parsel", ""), st,
                r.get("alan", "") or "", r.get("nitelik", "") or "", (r.get("yeni_parseller") or "").replace("; ", ", "), islem, kontrol))

    def detay_ac(self):
        sec = self.tree.selection()
        if sec:
            Detay(self, self.rows[int(sec[0])])

    def template(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="Parsel_Sorgu_Sablonu.xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "Parsel Listesi"
        ws.append(["İl", "İlçe", "Mahalle", "Ada", "Parsel"])
        ws.append(["Bursa", "Mustafakemalpaşa", "Atatürk", "727", "2"])
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="244766")
        for col, w in zip("ABCDE", [22, 25, 28, 14, 14]):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = "A2"
        wb.save(path)
        messagebox.showinfo("Tamamlandı", "Excel şablonu oluşturuldu.")

    def import_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if not path:
            return
        try:
            ws = load_workbook(path, data_only=True).active
            headers = [str(c.value or "").strip().casefold() for c in ws[1]]
            aliases = {"il": ["il", "şehir", "sehir"], "ilce": ["ilçe", "ilce"], "mahalle": ["mahalle", "köy", "koy"],
                       "ada": ["ada", "ada no", "ada no."], "parsel": ["parsel", "parsel no", "parsel no."]}
            idx = {}
            for k, names in aliases.items():
                for j, h in enumerate(headers):
                    if h in names:
                        idx[k] = j
                        break
            if len(idx) != 5:
                raise ValueError("İlk satırda İl, İlçe, Mahalle, Ada ve Parsel sütunları bulunmalı.")
            count = 0
            for vals in ws.iter_rows(min_row=2, values_only=True):
                row = {k: str(vals[j]).strip() if j < len(vals) and vals[j] is not None else "" for k, j in idx.items()}
                if all(row.values()):
                    before = len(self.rows)
                    self.add_row(row)
                    count += len(self.rows) - before
            self.status.set(f"{count} yeni parsel eklendi.")
        except Exception as e:
            messagebox.showerror("Excel okunamadı", str(e))

    def delete_selected(self):
        selected = self.tree.selection()
        if not selected:
            return
        for iid in sorted(selected, key=int, reverse=True):
            r = self.rows.pop(int(iid))
            con = db()
            con.execute("DELETE FROM parcels WHERE il=? AND ilce=? AND mahalle=? AND ada=? AND parsel=?",
                        (r["il"], r["ilce"], r["mahalle"], r["ada"], r["parsel"]))
            con.commit()
            con.close()
        self.refresh()

    def clear_rows(self):
        if self.busy:
            return
        if messagebox.askyesno("Onay", "Ekrandaki liste temizlensin mi? Kayıtlı sonuç veritabanı korunur."):
            self.rows = []
            self.refresh()

    # ---- veritabanı
    def save_row(self, r):
        con = db()
        r["updated"] = datetime.now().isoformat(timespec="seconds")
        deger = [r.get(k) for k in SUTUNLAR]
        yer = ",".join("?" * len(SUTUNLAR))
        guncelle = ",".join(f"{k}=excluded.{k}" for k in SUTUNLAR[5:])
        con.execute(f"INSERT INTO parcels({','.join(SUTUNLAR)}) VALUES({yer}) ON CONFLICT(il,ilce,mahalle,ada,parsel) DO UPDATE SET {guncelle}", deger)
        con.commit()
        con.close()

    def load_saved(self):
        con = db()
        veri = con.execute(f"SELECT {','.join(SUTUNLAR)} FROM parcels ORDER BY id").fetchall()
        con.close()
        for x in veri:
            r = dict(zip(SUTUNLAR, x))
            if r.get("status") in (None, "", "Başarılı"):  # eski sürümde pasif/aktif ayrımı yoktu: yeniden sorgulanmalı
                r["status"] = "Bekliyor"
            self.rows.append(r)
        self.refresh()

    # ---- sorgulama
    def bekleyenler(self):
        sinir = datetime.now() - timedelta(days=VARSAYILAN_YENILE_GUN)
        out = []
        for r in self.rows:
            st = r.get("status") or "Bekliyor"
            if st in ("Bekliyor", "Hata", "Belirsiz") or (st == "Aktif" and son_kontrol(r) < sinir):
                out.append(r)  # Pasif olanlar tekrar sorgulanmaz (işlem tamamlanmış)
        return out

    def start_queries(self, secili=False):
        if self.busy:
            return
        if secili:
            pending = [self.rows[int(i)] for i in self.tree.selection()]
        else:
            pending = self.bekleyenler()
        if not pending:
            messagebox.showinfo("Bilgi", "Sorgulanacak parsel yok. (Pasif olanlar ve yakın zamanda kontrol edilen aktif parseller atlanır; "
                                         "belirli satırları sorgulamak için seçip 'Seçili satırları sorgula'ya basın.)")
            return
        kalan = max(limit_degeri() - sayac_oku(), 0)
        if kalan <= 0:
            messagebox.showwarning("Günlük sınır", "Bugünkü istek üst sınırına ulaşıldı. Yarın devam edin ya da sınırı artırın.")
            return
        if not messagebox.askyesno("Toplu sorgu", f"{len(pending)} parsel sorgulanacak (bugün kalan istek hakkı: {kalan}). "
                                                  "İstekler arası bekleme uygulanır. Başlatılsın mı?"):
            return
        self.busy = True
        self.progress["maximum"] = len(pending)
        self.progress["value"] = 0
        threading.Thread(target=self.worker, args=(pending,), daemon=True).start()

    def worker(self, pending):
        sayim = {"Aktif": 0, "Pasif": 0, "Belirsiz": 0, "Hata": 0}
        durdu = ""
        for n, row in enumerate(pending, 1):
            try:
                self.after(0, self.status.set, f"Sorgulanıyor: {row['mahalle']} {row['ada']}/{row['parsel']}")
                query_one(row)
                geom = row.pop("geometry", None)
                try:
                    devamlar = json.loads(row.get("devam_parseller") or "[]")
                except (TypeError, ValueError):
                    devamlar = []
                for alt in devamlar:
                    alt_geom = alt.pop("geometry", None)
                    if alt_geom:
                        try:
                            DATA_DIR.mkdir(exist_ok=True)
                            safe_alt = "_".join(str(row[k]).replace("/", "-").replace(" ", "_") for k in ("il", "ilce", "mahalle")) + f"_Ada_{alt['ada']}_Parsel_{alt['parsel']}"
                            alt_path = DATA_DIR / (safe_alt + ".png")
                            alt_row = dict(row, ada=alt.get("ada"), parsel=alt.get("parsel"), alan=alt.get("alan"), nitelik=alt.get("nitelik"))
                            make_map(alt_row, alt_geom, alt_path)
                            alt["map_file"] = str(alt_path)
                        except Exception:
                            alt["map_file"] = ""
                row["devam_parseller"] = json.dumps(devamlar, ensure_ascii=False)
                if row["status"] == "Aktif" and geom:
                    try:
                        DATA_DIR.mkdir(exist_ok=True)
                        safe = "_".join(str(row[k]).replace("/", "-").replace(" ", "_") for k in ("il", "ilce", "mahalle", "ada", "parsel"))
                        make_map(row, geom, DATA_DIR / (safe + ".png"))
                        row["map_file"] = str(DATA_DIR / (safe + ".png"))
                    except Exception:
                        pass  # harita üretilememesi sorgu sonucunu bozmaz
                sayim[row["status"]] = sayim.get(row["status"], 0) + 1
            except (LimitHatasi, ErisimHatasi) as e:
                durdu = str(e)
                row["detail"] = durdu
                self.save_row(row)
                break
            except Exception as e:
                row["status"] = "Hata"
                row["detail"] = str(e)
                sayim["Hata"] += 1
            self.save_row(row)
            self.after(0, self.refresh)
            self.after(0, self.sayac_goster)
            self.after(0, self.progress.configure, {"value": n})
            time.sleep(BEKLEME_SN)
        self.busy = False
        ozet = f"Aktif: {sayim['Aktif']} | Pasif: {sayim['Pasif']} | Belirsiz: {sayim['Belirsiz']} | Hatalı: {sayim['Hata']}"
        self.after(0, self.refresh)
        self.after(0, self.sayac_goster)
        self.after(0, self.status.set, ("Durduruldu: " + durdu + "  " if durdu else "Tamamlandı. ") + ozet)
        self.after(0, lambda: messagebox.showinfo("Sorgu bitti", ("DURDURULDU: " + durdu + "\n\n" if durdu else "") + ozet))

    def export_excel(self):
        if not self.rows:
            messagebox.showinfo("Bilgi", "Aktarılacak kayıt yok.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="Parsel_Sorgu_Sonuclari.xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "Sorgu Sonuçları"
        ws.append(["İl", "İlçe", "Mahalle", "Ada", "Parsel", "Parsel durumu", "Alan (m²)", "Nitelik", "Pasif ibaresi", "İşlem türü",
                   "İşlem tarihi", "Yeni parseller", "Son kontrol", "Açıklama", "Koordinatlar (Enlem, Boylam)", "Harita Görseli"])
        for r in self.rows:
            ws.append([r.get("il"), r.get("ilce"), r.get("mahalle"), r.get("ada"), r.get("parsel"), r.get("status"), r.get("alan"),
                       r.get("nitelik"), r.get("pasif_mesaj"), r.get("islem_turu"), r.get("islem_tarihi"),
                       (r.get("yeni_parseller") or "").replace("; ", ", "), (r.get("updated") or "").replace("T", " "),
                       r.get("detail"), r.get("koordinatlar"), r.get("map_file")])
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="244766")
            c.alignment = Alignment(wrap_text=True)
        for i, w in enumerate([14, 20, 20, 10, 10, 14, 12, 16, 50, 12, 13, 24, 17, 50, 40, 40], 1):
            ws.column_dimensions[chr(64 + i)].width = w
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        devam_ws = wb.create_sheet("Aktif Devam Parselleri")
        devam_ws.append(["Kaynak İl", "Kaynak İlçe", "Kaynak Mahalle", "Eski Ada", "Eski Parsel", "İşlem Türü", "İşlem Tarihi", "Yeni Ada", "Yeni Parsel", "Durum", "Alan (m²)", "Nitelik", "Koordinatlar (Enlem, Boylam)"])
        for kaynak in self.rows:
            try:
                altlar = json.loads(kaynak.get("devam_parseller") or "[]")
            except (TypeError, ValueError):
                altlar = []
            for alt in altlar:
                devam_ws.append([kaynak.get("il"), kaynak.get("ilce"), kaynak.get("mahalle"), kaynak.get("ada"), kaynak.get("parsel"),
                                 kaynak.get("islem_turu"), kaynak.get("islem_tarihi"), alt.get("ada"), alt.get("parsel"), alt.get("durum"),
                                 alt.get("alan"), alt.get("nitelik"), alt.get("koordinatlar")])
        for c in devam_ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="244766")
            c.alignment = Alignment(wrap_text=True)
        for i, w in enumerate([14, 20, 20, 10, 10, 14, 13, 10, 10, 12, 12, 22, 42], 1):
            devam_ws.column_dimensions[chr(64 + i)].width = w
        devam_ws.freeze_panes = "A2"
        devam_ws.auto_filter.ref = devam_ws.dimensions
        wb.save(path)
        messagebox.showinfo("Tamamlandı", "Excel sonuç dosyası oluşturuldu. Aktif devam parselleri ayrı sayfaya eklendi.")


if __name__ == "__main__":
    App().mainloop()
