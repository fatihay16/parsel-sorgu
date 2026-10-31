# -*- coding: utf-8 -*-
"""İş Takip içindeki Parsel Sorgu penceresi için TKGM idari liste yardımcıları.

Ana pencereye eklemek için:
  self.idari_secicileri_kur(parent)
  self.illeri_yukle()
Metotlar bu sınıfın içine veya bir mixin üzerinden eklenebilir.
"""
import json
import tkinter as tk
from tkinter import ttk

API = "https://cbsapi.tkgm.gov.tr/megsiswebapi.v3.1/api"


def norm_ad(x):
    return " ".join(str(x or "").strip().casefold().replace("i\u0307", "i").split())


def feature_items(data):
    if isinstance(data, dict) and data.get("type") == "FeatureCollection":
        return [f.get("properties", {}) for f in data.get("features", [])]
    return data if isinstance(data, list) else []


def liste_adlari(data):
    keys = ("text", "ad", "name", "ilAdi", "ilceAdi", "mahalleAdi")
    names = []
    for item in feature_items(data):
        if not isinstance(item, dict):
            continue
        name = next((str(item[k]).strip() for k in keys if item.get(k) not in (None, "")), "")
        if name and name not in names:
            names.append(name)
    return sorted(names, key=norm_ad)


def find_id(items, name_keys, wanted):
    target = norm_ad(wanted)
    for item in items:
        name = next((item.get(k) for k in name_keys if item.get(k) not in (None, "")), "")
        if norm_ad(name) == target:
            return item.get("id") or item.get("mahalleKodu") or item.get("ilceKodu")
    matches = []
    for item in items:
        name = next((item.get(k) for k in name_keys if item.get(k) not in (None, "")), "")
        if target and (target in norm_ad(name) or norm_ad(name) in target):
            matches.append(item.get("id") or item.get("mahalleKodu") or item.get("ilceKodu"))
    if len(matches) == 1:
        return matches[0]
    raise RuntimeError(f"'{wanted}' TKGM idari listesinde tekil bulunamadı.")


class TkgmIdariListeMixin:
    """Tkinter App sınıfına eklenecek il/ilçe/mahalle açılır liste davranışı."""

    def idari_secicileri_kur(self, parent):
        self.idari_vars = {k: tk.StringVar() for k in ("il", "ilce", "mahalle")}
        self.idari_combos = {}
        self.idari_tam_listeler = {k: [] for k in ("il", "ilce", "mahalle")}
        self._idari_oneri = {}
        bar = ttk.LabelFrame(parent, text="TKGM Konum Seçimi", padding=8)
        bar.pack(fill="x", padx=18, pady=(4, 8))
        for col, key, label in ((0, "il", "İl"), (1, "ilce", "İlçe"), (2, "mahalle", "Mahalle")):
            ttk.Label(bar, text=label).grid(row=0, column=col, sticky="w", padx=5)
            cb = ttk.Combobox(bar, textvariable=self.idari_vars[key], state="normal")
            cb.grid(row=1, column=col, sticky="ew", padx=5, pady=(2, 0))
            cb.bind("<<ComboboxSelected>>", lambda e, k=key: self._idari_secildi(k))
            cb.bind("<KeyRelease>", lambda e, k=key: self._idari_filtrele(k, e))
            self.idari_combos[key] = cb
            bar.columnconfigure(col, weight=1)
        self.after(100, self.illeri_yukle)
        return bar

    def idari_liste(self, tur, ust_id=None):
        if tur == "il":
            url = f"{API}/idariYapi/ilListe"
        elif tur == "ilce":
            url = f"{API}/idariYapi/ilceListe/{ust_id}"
        else:
            url = f"{API}/idariYapi/mahalleListe/{ust_id}"
        # Ana modülde get_json_onbellekli varsa onu kullan; yoksa doğrudan iste.
        getter = globals().get("get_json_onbellekli") or globals().get("get_json")
        return getter(url)

    def _idari_filtrele(self, key, event=None):
        if event and event.keysym in ("Up", "Down", "Left", "Right", "Return", "Escape", "Tab"):
            if event.keysym == "Escape":
                self._idari_oneri_kapat(key)
            return
        all_names = self.idari_tam_listeler.get(key, [])
        typed = self.idari_vars[key].get()
        q = norm_ad(typed)
        matches = [x for x in all_names if q in norm_ad(x)] if q else all_names
        self.idari_combos[key]["values"] = matches
        # Combobox kendi açılır listesini kullanır; ek pencere açılmaz.
        if matches and self.idari_combos[key].focus_get() == self.idari_combos[key]:
            try:
                self.idari_combos[key].event_generate("<Down>")
            except tk.TclError:
                pass

    def _idari_oneri_kapat(self, key):
        win = self._idari_oneri.get(key)
        if win:
            try:
                win.destroy()
            except tk.TclError:
                pass
        self._idari_oneri[key] = None

    def illeri_yukle(self):
        try:
            data = self.idari_liste("il")
            self.idari_tam_listeler["il"] = liste_adlari(data)
            self.idari_combos["il"]["values"] = self.idari_tam_listeler["il"]
            saved = globals().get("ayar_oku", lambda k: "") 
            il = saved("secim_il")
            ilce = saved("secim_ilce")
            mahalle = saved("secim_mahalle")
            if il in self.idari_tam_listeler["il"]:
                self.idari_vars["il"].set(il)
                self._idari_il_secildi(restore=True, ilce=ilce, mahalle=mahalle)
        except Exception as exc:
            if hasattr(self, "status"):
                self.status.set("İl listesi alınamadı: " + str(exc))

    def _idari_secildi(self, key):
        if key == "il":
            self._idari_il_secildi()
        elif key == "ilce":
            self._idari_mahalleleri_yukle()
        self._idari_ayar_kaydet()

    def _idari_il_secildi(self, restore=False, ilce="", mahalle=""):
        try:
            provinces = feature_items(self.idari_liste("il"))
            il_id = find_id(provinces, ("text", "ad", "name", "ilAdi"), self.idari_vars["il"].get())
            districts = self.idari_liste("ilce", il_id)
            self.idari_tam_listeler["ilce"] = liste_adlari(districts)
            self.idari_combos["ilce"]["values"] = self.idari_tam_listeler["ilce"]
            self.idari_vars["ilce"].set(ilce if restore and ilce in self.idari_tam_listeler["ilce"] else "")
            self.idari_tam_listeler["mahalle"] = []
            self.idari_combos["mahalle"]["values"] = ()
            self.idari_vars["mahalle"].set("")
            if restore and self.idari_vars["ilce"].get():
                self._idari_mahalleleri_yukle(restore_mahalle=mahalle)
        except Exception as exc:
            if hasattr(self, "status"):
                self.status.set("İlçe listesi alınamadı: " + str(exc))

    def _idari_mahalleleri_yukle(self, restore_mahalle=""):
        try:
            provinces = feature_items(self.idari_liste("il"))
            il_id = find_id(provinces, ("text", "ad", "name", "ilAdi"), self.idari_vars["il"].get())
            districts = feature_items(self.idari_liste("ilce", il_id))
            ilce_id = find_id(districts, ("text", "ilceAdi", "ad", "name"), self.idari_vars["ilce"].get())
            data = self.idari_liste("mahalle", ilce_id)
            self.idari_tam_listeler["mahalle"] = liste_adlari(data)
            self.idari_combos["mahalle"]["values"] = self.idari_tam_listeler["mahalle"]
            if restore_mahalle in self.idari_tam_listeler["mahalle"]:
                self.idari_vars["mahalle"].set(restore_mahalle)
            if hasattr(self, "status"):
                self.status.set(f"{len(self.idari_tam_listeler['mahalle'])} mahalle listelendi.")
        except Exception as exc:
            if hasattr(self, "status"):
                self.status.set("Mahalle listesi alınamadı: " + str(exc))

    def _idari_ayar_kaydet(self):
        save = globals().get("ayar_yaz")
        if save:
            for key in ("il", "ilce", "mahalle"):
                save("secim_" + key, self.idari_vars[key].get())
