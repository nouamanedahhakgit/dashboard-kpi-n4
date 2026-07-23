# -*- coding: utf-8 -*-
"""
KPI Parc Conteneurs - Interface graphique (Navis N4)
====================================================
Charge les donnees UNE fois (API ou cache local), puis filtre instantanement
par periode et par critere, sans rappeler l'API. Export Excel du perimetre filtre.

Lancer :  python kpi_gui.py

Les identifiants sont lus dans .env (N4_USER / N4_PASSWORD / N4_URL).
Le premier chargement API est mis en cache dans  cache.xml  : les fois suivantes,
l'appli demarre instantanement sur le cache. Bouton "Rafraichir depuis l'API"
pour recharger des donnees fraiches.
"""

import os
import sys
import threading
import queue
import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import kpi_stats as ks   # reutilise fetch_api, parse_xml, Data, build_excel, load_dotenv

CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache.xml")

# mois francais abreges -> numero
MONTHS = {"janv": 1, "févr": 2, "fevr": 2, "mars": 3, "avr": 4, "mai": 5, "juin": 6,
          "juil": 7, "juill": 7, "août": 8, "aout": 8, "sept": 9, "oct": 10,
          "nov": 11, "déc": 12, "dec": 12}


def parse_dt(s):
    """Convertit '26-juil.-12 1125'  ->  datetime(2026,7,12,11,25).  Format YY-mois-JJ HHMM."""
    if not s:
        return None
    try:
        parts = s.strip().split()
        datepart = parts[0]
        timepart = parts[1] if len(parts) > 1 else ""
        yy, mon, dd = datepart.split("-")
        mon = mon.strip(".").lower()
        m = MONTHS.get(mon) or MONTHS.get(mon[:4]) or MONTHS.get(mon[:3])
        if not m:
            return None
        hh = int(timepart[:2]) if len(timepart) >= 2 else 0
        mi = int(timepart[2:4]) if len(timepart) >= 4 else 0
        return datetime.datetime(2000 + int(yy), m, int(dd), hh, mi)
    except Exception:
        return None


class App:
    def __init__(self, root):
        self.root = root
        root.title("KPI Parc Conteneurs - Navis N4")
        root.geometry("1080x720")
        self.data = None          # ks.Data (toutes lignes)
        self.filtered = []        # lignes filtrees
        self.q = queue.Queue()
        ks.load_dotenv()

        self._build_ui()

        # demarrage : charge le cache s'il existe
        if os.path.exists(CACHE):
            self._load_file(CACHE, "cache local")
        else:
            self.set_status("Aucun cache. Clique « Rafraichir depuis l'API » pour charger les donnees.")

    # ---------------------------------------------------------------- UI
    def _build_ui(self):
        top = ttk.Frame(self.root, padding=8)
        top.pack(fill="x")
        ttk.Button(top, text="Rafraichir depuis l'API", command=self.refresh_api).pack(side="left")
        ttk.Button(top, text="Charger un fichier XML…", command=self.load_dialog).pack(side="left", padx=6)
        ttk.Button(top, text="Exporter Excel (perimetre filtre)", command=self.export).pack(side="left")
        self.status = ttk.Label(top, text="", foreground="#555")
        self.status.pack(side="left", padx=12)

        # ---- filtres ----
        fl = ttk.LabelFrame(self.root, text="Filtres", padding=8)
        fl.pack(fill="x", padx=8, pady=6)

        ttk.Label(fl, text="Colonne date :").grid(row=0, column=0, sticky="w")
        self.datecol = ttk.Combobox(fl, width=16, state="readonly",
                                    values=["Last Move", "EC-In Time", "Complex InTime"])
        self.datecol.set("Last Move")
        self.datecol.grid(row=0, column=1, padx=4)

        ttk.Label(fl, text="Du (AAAA-MM-JJ) :").grid(row=0, column=2, sticky="w", padx=(12, 0))
        self.dfrom = ttk.Entry(fl, width=12)
        self.dfrom.grid(row=0, column=3, padx=4)
        ttk.Label(fl, text="Au :").grid(row=0, column=4, sticky="w")
        self.dto = ttk.Entry(fl, width=12)
        self.dto.grid(row=0, column=5, padx=4)

        # raccourcis periode
        for i, (lab, days) in enumerate([("Aujourd'hui", 0), ("7 j", 7), ("30 j", 30), ("90 j", 90)]):
            ttk.Button(fl, text=lab, width=10,
                       command=lambda dd=days: self.quick_period(dd)).grid(row=0, column=6 + i, padx=2)

        ttk.Label(fl, text="Categorie :").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.f_cat = ttk.Combobox(fl, width=16, state="readonly")
        self.f_cat.grid(row=1, column=1, pady=(8, 0))
        ttk.Label(fl, text="Armateur :").grid(row=1, column=2, sticky="w", pady=(8, 0), padx=(12, 0))
        self.f_line = ttk.Combobox(fl, width=16, state="readonly")
        self.f_line.grid(row=1, column=3, pady=(8, 0))
        ttk.Label(fl, text="Plein/Vide :").grid(row=1, column=4, sticky="w", pady=(8, 0))
        self.f_fk = ttk.Combobox(fl, width=12, state="readonly")
        self.f_fk.grid(row=1, column=5, pady=(8, 0))

        ttk.Button(fl, text="Appliquer", command=self.apply).grid(row=1, column=8, pady=(8, 0), padx=2)
        ttk.Button(fl, text="Reinitialiser", command=self.reset).grid(row=1, column=9, pady=(8, 0))

        # ---- KPIs ----
        self.kpiframe = ttk.LabelFrame(self.root, text="Indicateurs", padding=8)
        self.kpiframe.pack(fill="x", padx=8, pady=6)
        self.kpi_labels = {}
        kpis = ["Conteneurs", "Import", "Export", "Storage", "Plein", "Vide",
                "Reefers", "Dangereux", "Hors gabarit", "Bloques navire", "Bloques route",
                "Dwell median (j)", "Aging >120j"]
        for i, name in enumerate(kpis):
            cell = ttk.Frame(self.kpiframe, relief="solid", borderwidth=1, padding=6)
            cell.grid(row=i // 7, column=i % 7, sticky="nsew", padx=3, pady=3)
            ttk.Label(cell, text=name, foreground="#666", font=("Segoe UI", 8)).pack()
            v = ttk.Label(cell, text="-", font=("Segoe UI", 15, "bold"))
            v.pack()
            self.kpi_labels[name] = v
        for c in range(7):
            self.kpiframe.columnconfigure(c, weight=1)

        # ---- tableau (apercu 200 lignes) ----
        tb = ttk.LabelFrame(self.root, text="Apercu (200 premieres lignes du perimetre filtre)", padding=6)
        tb.pack(fill="both", expand=True, padx=8, pady=6)
        self.tree = ttk.Treeview(tb, show="headings", height=10)
        vs = ttk.Scrollbar(tb, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(tb, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        tb.rowconfigure(0, weight=1)
        tb.columnconfigure(0, weight=1)

    # ---------------------------------------------------------------- data
    def set_status(self, txt):
        self.status.config(text=txt)

    def refresh_api(self):
        user = os.environ.get("N4_USER")
        pwd = os.environ.get("N4_PASSWORD")
        url = os.environ.get("N4_URL", ks.DEFAULT_URL)
        if not user or not pwd:
            messagebox.showwarning("Identifiants manquants",
                                   "Renseigne N4_USER et N4_PASSWORD dans le fichier .env.")
            return
        self.set_status("Appel de l'API en cours… (peut prendre 1-2 min)")
        threading.Thread(target=self._fetch_thread, args=(url, user, pwd), daemon=True).start()
        self.root.after(200, self._poll)

    def _fetch_thread(self, url, user, pwd):
        try:
            text = ks.fetch_api(url, user, pwd)
            with open(CACHE, "w", encoding="utf-8") as f:
                f.write(text)
            self.q.put(("ok", text))
        except Exception as e:
            self.q.put(("err", str(e)))

    def _poll(self):
        try:
            kind, payload = self.q.get_nowait()
        except queue.Empty:
            self.root.after(200, self._poll)
            return
        if kind == "err":
            self.set_status("❌ API : %s" % payload)
            messagebox.showerror("Erreur API", payload)
        else:
            self._ingest(payload, "API (mis en cache)")

    def load_dialog(self):
        p = filedialog.askopenfilename(filetypes=[("XML", "*.xml"), ("Tous", "*.*")])
        if p:
            self._load_file(p, os.path.basename(p))

    def _load_file(self, path, label):
        try:
            with open(path, encoding="utf-8", errors="replace") as f:
                self._ingest(f.read(), label)
        except Exception as e:
            messagebox.showerror("Erreur", str(e))

    def _ingest(self, text, label):
        cols, rows = ks.parse_xml(text)
        if not rows:
            self.set_status("Reponse vide ou inattendue.")
            return
        self.data = ks.Data(cols, rows)
        # remplir les combos
        def opts(name):
            return [""] + [k for k, _ in self.data.count_by(name, 30)]
        self.f_cat["values"] = opts("Category")
        self.f_line["values"] = opts("Line Op")
        self.f_fk["values"] = opts("Frght Kind")
        for c in (self.f_cat, self.f_line, self.f_fk):
            c.set("")
        # colonnes du tableau
        self.tree["columns"] = cols
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=90, anchor="w", stretch=False)
        self.set_status("Charge : %s — %d conteneurs." % (label, len(rows)))
        self.apply()

    # ---------------------------------------------------------------- filtres
    def quick_period(self, days):
        today = datetime.date.today()
        self.dto.delete(0, "end"); self.dto.insert(0, today.isoformat())
        start = today - datetime.timedelta(days=days)
        self.dfrom.delete(0, "end"); self.dfrom.insert(0, start.isoformat())
        self.apply()

    def reset(self):
        for e in (self.dfrom, self.dto):
            e.delete(0, "end")
        for c in (self.f_cat, self.f_line, self.f_fk):
            c.set("")
        self.apply()

    def _date_bound(self, entry, end=False):
        t = entry.get().strip()
        if not t:
            return None
        try:
            y, m, d = [int(x) for x in t.split("-")]
            return datetime.datetime(y, m, d, 23, 59 if end else 0, 59 if end else 0)
        except Exception:
            return None

    def apply(self):
        if not self.data:
            return
        d = self.data
        dcol = self.datecol.get()
        jdate = d.idx.get(dcol, -1)
        dfrom = self._date_bound(self.dfrom)
        dto = self._date_bound(self.dto, end=True)
        cat = self.f_cat.get(); line = self.f_line.get(); fk = self.f_fk.get()
        jcat = d.idx.get("Category", -1); jline = d.idx.get("Line Op", -1); jfk = d.idx.get("Frght Kind", -1)

        out = []
        for r in d.rows:
            if cat and jcat >= 0 and (r[jcat] if jcat < len(r) else "") != cat:
                continue
            if line and jline >= 0 and (r[jline] if jline < len(r) else "") != line:
                continue
            if fk and jfk >= 0 and (r[jfk] if jfk < len(r) else "") != fk:
                continue
            if (dfrom or dto) and jdate >= 0:
                dt = parse_dt(r[jdate] if jdate < len(r) else "")
                if dt is None:
                    continue
                if dfrom and dt < dfrom:
                    continue
                if dto and dt > dto:
                    continue
            out.append(r)
        self.filtered = out
        self._update_kpis()
        self._update_table()

    # ---------------------------------------------------------------- affichage
    def _cnt(self, rows, name, value=None, truthy=False):
        j = self.data.idx.get(name, -1)
        if j < 0:
            return 0
        if truthy:
            return sum(1 for r in rows if j < len(r) and r[j].lower() == "true")
        return sum(1 for r in rows if j < len(r) and r[j] == value)

    def _update_kpis(self):
        rows = self.filtered
        n = len(rows)
        L = self.kpi_labels
        L["Conteneurs"].config(text="{:,}".format(n))
        L["Import"].config(text="{:,}".format(self._cnt(rows, "Category", "Import")))
        L["Export"].config(text="{:,}".format(self._cnt(rows, "Category", "Export")))
        L["Storage"].config(text="{:,}".format(self._cnt(rows, "Category", "Storage")))
        L["Plein"].config(text="{:,}".format(self._cnt(rows, "Frght Kind", "FCL")))
        L["Vide"].config(text="{:,}".format(self._cnt(rows, "Frght Kind", "Empty")))
        L["Reefers"].config(text="{:,}".format(self._cnt(rows, "Reqs Power", truthy=True)))
        L["Dangereux"].config(text="{:,}".format(self._cnt(rows, "Hazardous?", truthy=True)))
        L["Hors gabarit"].config(text="{:,}".format(self._cnt(rows, "Is OOG", truthy=True)))
        L["Bloques navire"].config(text="{:,}".format(self._cnt(rows, "Stop-Vsl", truthy=True)))
        L["Bloques route"].config(text="{:,}".format(self._cnt(rows, "Stop-Road", truthy=True)))
        jd = self.data.idx.get("Dwell", -1)
        med, aging = "-", 0
        if jd >= 0:
            vals = sorted(float(r[jd]) for r in rows if jd < len(r) and _isnum(r[jd]))
            if vals:
                med = int(vals[len(vals) // 2])
                aging = sum(1 for x in vals if x > 120)
        L["Dwell median (j)"].config(text=str(med))
        L["Aging >120j"].config(text="{:,}".format(aging))

    def _update_table(self):
        self.tree.delete(*self.tree.get_children())
        for r in self.filtered[:200]:
            self.tree.insert("", "end", values=r)

    # ---------------------------------------------------------------- export
    def export(self):
        if not self.filtered:
            messagebox.showinfo("Rien a exporter", "Aucune ligne dans le perimetre filtre.")
            return
        p = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                         initialfile="kpi_%s.xlsx" % datetime.datetime.now().strftime("%Y%m%d_%H%M"),
                                         filetypes=[("Excel", "*.xlsx")])
        if not p:
            return
        try:
            ks.build_excel(ks.Data(self.data.cols, self.filtered), p)
            messagebox.showinfo("Export termine", "%d lignes exportees vers :\n%s" % (len(self.filtered), p))
        except Exception as e:
            messagebox.showerror("Erreur export", str(e))


def _isnum(s):
    try:
        float(s); return True
    except Exception:
        return False


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
