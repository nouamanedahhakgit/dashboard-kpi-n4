# -*- coding: utf-8 -*-
"""
KPI Parc Conteneurs - Serveur web local (Navis N4)
==================================================
Serveur sur http://localhost:8000 :
  - auth Basic + appel API cote SERVEUR (pas de CORS), cache dans cache.xml ;
  - le FILTRE pilote la requete : le navigateur envoie les criteres a /api/stats,
    le serveur filtre en memoire et ne renvoie que les KPI + agregats des graphiques
    (quelques Ko), jamais les 30 000 lignes.

Lancer :  python kpi_web.py    (le navigateur s'ouvre sur http://localhost:8000)
Identifiants dans .env (N4_USER / N4_PASSWORD / N4_URL).
"""

import os
import json
import datetime
import threading
import webbrowser
from urllib.parse import urlparse, parse_qs
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import kpi_stats as ks

HERE = os.path.dirname(os.path.abspath(__file__))
CACHE = os.path.join(HERE, "cache.xml")
INDEX = os.path.join(HERE, "index.html")
PORT = 8000
AGING = 120  # jours

ks.load_dotenv()

STATE = {"columns": [], "rows": [], "idx": {}, "source": "aucune"}
LOCK = threading.Lock()

CAT_KEYS = ["Category", "Frght Kind", "Line Op", "V-State", "T-State", "POD",
            "Reqs Power", "Hazardous?", "Is OOG", "Stop-Vsl"]
DATE_COLS = ["Last Move", "EC-In Time", "Complex InTime"]

MONTHS = {"janv": 1, "févr": 2, "fevr": 2, "mars": 3, "avr": 4, "mai": 5, "juin": 6,
          "juil": 7, "juill": 7, "août": 8, "aout": 8, "sept": 9, "oct": 10,
          "nov": 11, "déc": 12, "dec": 12}


def parse_dt(s):
    if not s:
        return None
    try:
        parts = s.strip().split()
        yy, mon, dd = parts[0].split("-")
        mon = mon.strip(".").lower()
        m = MONTHS.get(mon) or MONTHS.get(mon[:4]) or MONTHS.get(mon[:3])
        if not m:
            return None
        t = parts[1] if len(parts) > 1 else ""
        hh = int(t[:2]) if len(t) >= 2 else 0
        mi = int(t[2:4]) if len(t) >= 4 else 0
        return datetime.datetime(2000 + int(yy), m, int(dd), hh, mi)
    except Exception:
        return None


# --------------------------------------------------------------- data lifecycle
def ingest_xml(text, source):
    cols, rows = ks.parse_xml(text)
    with LOCK:
        STATE["columns"] = cols
        STATE["rows"] = rows
        STATE["idx"] = {c: i for i, c in enumerate(cols)}
        STATE["source"] = source
    return len(rows)


def load_cache():
    if os.path.exists(CACHE):
        try:
            with open(CACHE, encoding="utf-8", errors="replace") as f:
                print("Cache charge : %d conteneurs." % ingest_xml(f.read(), "cache local"))
        except Exception as e:
            print("Cache illisible :", e)
    else:
        print("Pas de cache. Clique « Rafraichir depuis l'API » dans la page.")


def refresh_from_api():
    user, pwd = os.environ.get("N4_USER"), os.environ.get("N4_PASSWORD")
    url = os.environ.get("N4_URL", ks.DEFAULT_URL)
    if not user or not pwd:
        raise RuntimeError("N4_USER / N4_PASSWORD manquants dans .env")
    text = ks.fetch_api(url, user, pwd)
    with open(CACHE, "w", encoding="utf-8") as f:
        f.write(text)
    return ingest_xml(text, "API (mis en cache)")


# ------------------------------------------------------------------ agregation
def gidx(name):
    return STATE["idx"].get(name, -1)


def filter_rows(q):
    """q : dict de listes (parse_qs). Filtre les lignes en memoire."""
    rows = STATE["rows"]
    # filtres categoriels
    conds = []
    for k in CAT_KEYS:
        v = q.get(k, [""])[0]
        if v:
            j = gidx(k)
            if j >= 0:
                conds.append((j, v))
    # periode
    dcol = q.get("datecol", [DATE_COLS[0]])[0]
    jd = gidx(dcol)
    dfrom = _bound(q.get("from", [""])[0], False)
    dto = _bound(q.get("to", [""])[0], True)
    # dwell / recherche
    dmin = q.get("dwellMin", [""])[0]
    dmin = float(dmin) if dmin not in ("", None) else None
    jDwell = gidx("Dwell")
    search = (q.get("search", [""])[0] or "").upper()
    jUnit = gidx("Unit Nbr")

    out = []
    for r in rows:
        ok = True
        for j, v in conds:
            if (r[j] if j < len(r) else "") != v:
                ok = False
                break
        if not ok:
            continue
        if (dfrom or dto) and jd >= 0:
            dt = parse_dt(r[jd] if jd < len(r) else "")
            if dt is None or (dfrom and dt < dfrom) or (dto and dt > dto):
                continue
        if dmin is not None and jDwell >= 0:
            try:
                if float(r[jDwell]) < dmin:
                    continue
            except (ValueError, IndexError):
                continue
        if search and jUnit >= 0 and search not in (r[jUnit] if jUnit < len(r) else "").upper():
            continue
        out.append(r)
    return out


def _bound(txt, end):
    if not txt:
        return None
    try:
        y, m, d = [int(x) for x in txt.split("-")]
        return datetime.datetime(y, m, d, 23 if end else 0, 59 if end else 0, 59 if end else 0)
    except Exception:
        return None


def count_by(rows, name, top=None):
    j = gidx(name)
    if j < 0:
        return []
    m = {}
    for r in rows:
        v = (r[j] if j < len(r) else "") or "(vide)"
        m[v] = m.get(v, 0) + 1
    a = sorted(m.items(), key=lambda kv: -kv[1])
    return a[:top] if top else a


def count_true(rows, name):
    j = gidx(name)
    if j < 0:
        return 0
    return sum(1 for r in rows if j < len(r) and r[j].lower() == "true")


def count_val(rows, name, val):
    j = gidx(name)
    if j < 0:
        return 0
    return sum(1 for r in rows if j < len(r) and r[j] == val)


def compute_stats(rows):
    n = len(rows)
    # dwell
    jd = gidx("Dwell")
    dv = []
    if jd >= 0:
        for r in rows:
            try:
                dv.append(float(r[jd]))
            except (ValueError, IndexError):
                pass
    dv.sort()
    edges = [0, 1, 3, 7, 14, 30, 60, 90, 120, 180, 10 ** 9]
    hist = [0] * (len(edges) - 1)
    for x in dv:
        for i in range(len(edges) - 1):
            if edges[i] <= x < edges[i + 1]:
                hist[i] += 1
                break
    dwell = None
    if dv:
        dwell = {"med": dv[len(dv) // 2], "p90": dv[int(len(dv) * 0.9)], "max": dv[-1],
                 "avg": sum(dv) / len(dv), "aging": sum(1 for x in dv if x > AGING), "hist": hist}
    return {
        "total": n,
        "counts": {
            "import": count_val(rows, "Category", "Import"),
            "export": count_val(rows, "Category", "Export"),
            "storage": count_val(rows, "Category", "Storage"),
            "full": count_val(rows, "Frght Kind", "FCL") + count_val(rows, "Frght Kind", "Full"),
            "empty": count_val(rows, "Frght Kind", "Empty"),
            "reefer": count_true(rows, "Reqs Power"),
            "haz": count_true(rows, "Hazardous?"),
            "oog": count_true(rows, "Is OOG"),
            "stopvsl": count_true(rows, "Stop-Vsl"),
            "stoproad": count_true(rows, "Stop-Road"),
        },
        "byCategory": count_by(rows, "Category"),
        "byFreight": count_by(rows, "Frght Kind"),
        "byLine": count_by(rows, "Line Op", 10),
        "byPOD": count_by(rows, "POD", 10),
        "byTState": count_by(rows, "T-State"),
        "byISO": count_by(rows, "ISO Grp", 8),
        "dwell": dwell,
    }


def build_meta():
    opts = {}
    for k in CAT_KEYS:
        opts[k] = count_by(STATE["rows"], k, 60)
    return {
        "source": STATE["source"],
        "total": len(STATE["rows"]),
        "columns": STATE["columns"],
        "dateCols": [c for c in DATE_COLS if gidx(c) >= 0],
        "hasDwell": gidx("Dwell") >= 0,
        "options": opts,
        "aging": AGING,
    }


# ------------------------------------------------------------------ HTTP
class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        u = urlparse(self.path)
        path = u.path
        if path == "/" or path.startswith("/index"):
            try:
                with open(INDEX, encoding="utf-8") as f:
                    self._send(200, f.read(), "text/html; charset=utf-8")
            except FileNotFoundError:
                self._send(500, "index.html introuvable", "text/plain; charset=utf-8")
            return
        if path == "/api/meta":
            with LOCK:
                self._send(200, json.dumps(build_meta()))
            return
        if path == "/api/stats":
            q = parse_qs(u.query)
            with LOCK:
                rows = filter_rows(q)
                self._send(200, json.dumps(compute_stats(rows)))
            return
        if path == "/api/refresh":
            try:
                n = refresh_from_api()
                self._send(200, json.dumps({"ok": True, "count": n}))
            except Exception as e:
                self._send(200, json.dumps({"ok": False, "error": str(e)}))
            return
        self._send(404, json.dumps({"error": "not found"}))


def main():
    load_cache()
    srv = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    url = "http://localhost:%d" % PORT
    print("\nServeur KPI demarre  ->  %s\n(Ctrl+C pour arreter)\n" % url)
    try:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\nArret.")
        srv.shutdown()


if __name__ == "__main__":
    main()
