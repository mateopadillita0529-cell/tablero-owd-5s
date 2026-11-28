"""Sincroniza el tablero OWD · 5S · Acciones con DPO Evolution (dpo.arvolution.com).

Lo ejecuta GitHub Actions cada 15 minutos (.github/workflows/arvolution.yml).
Credenciales: secretos del repositorio ARVOLUTION_EMAIL y ARVOLUTION_PASSWORD
(nunca se escriben en el código ni en los datos publicados).

Escribe en data/:
  owd_ejecucion.csv   reporte OWD del mes (mismas columnas que la pestaña "Ejecucion")
  5s_ejecucion.csv    reporte 5S del mes (mismas columnas que la pestaña "5S Ejecucion")
  acciones.csv        acciones del año, Villavicencio + Yopal (pestaña "Acciones")
  estado.json         hora y resultado de cada fuente

Uso local para probar el procesamiento sin conectarse:
  python3 sync/arvolution_sync.py --owd-xlsx a.xlsx --s5-xlsx b.xlsx --acc-xlsx c.xlsx d.xlsx
"""
import os, sys, io, re, csv, json, time, argparse, unicodedata
import datetime as dt
import openpyxl
import pandas as pd

BASE = "https://dpo.arvolution.com"
SUBZONE = "colombia-sur"
CDS = ["cd-villavicencio", "cd-yopal"]
TZ = dt.timezone(dt.timedelta(hours=-5))          # Bogotá
OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")


def log(*a):
    print(time.strftime("%H:%M:%S"), *a, flush=True)


# ---------------------------------------------------------------- conexión
def sesion():
    import requests
    email = os.environ.get("ARVOLUTION_EMAIL", "").strip()
    pw = os.environ.get("ARVOLUTION_PASSWORD", "")
    if not email or not pw:
        raise RuntimeError("faltan los secretos ARVOLUTION_EMAIL / ARVOLUTION_PASSWORD")
    s = requests.Session()
    s.headers.update({"User-Agent": "tablero-owd-5s (GitHub Actions)", "Accept": "application/json, text/plain, */*",
                      "Origin": BASE, "Referer": BASE + "/"})
    ultimo = ""
    for campo in ("email", "username", "user"):
        r = s.post(BASE + "/api/v1/users/login", json={campo: email, "password": pw}, timeout=60)
        claves = []
        try:
            j = r.json()
            claves = list(j.keys()) if isinstance(j, dict) else []
        except Exception:
            j = None
        ultimo = f"login campo={campo} -> HTTP {r.status_code} · claves respuesta {claves[:8]}"
        log(ultimo)
        if r.status_code in (200, 201):
            # si la respuesta trae un token, úsalo como Bearer (además de la cookie)
            tok = None
            if isinstance(j, dict):
                for k, v in list(j.items()) + list((j.get("data") or {}).items() if isinstance(j.get("data"), dict) else []):
                    if isinstance(v, str) and "token" in k.lower() and len(v) > 20:
                        tok = v
                        break
            if tok:
                s.headers["Authorization"] = "Bearer " + tok
            chk = s.get(BASE + "/api/v1/users", timeout=60)
            log(f"verificación de sesión -> HTTP {chk.status_code}")
            if chk.status_code == 200:
                return s
        if r.status_code in (401, 403) and campo == "email":
            # credenciales rechazadas con el campo correcto: no seguir probando
            body = (r.text or "")[:200].lower()
            if "mfa" in body or "code" in body or "otp" in body:
                raise RuntimeError("Arvolution pide un código de verificación (MFA): no se puede automatizar")
    raise RuntimeError("no se pudo iniciar sesión en Arvolution (" + ultimo + ")")


def rango_mes(hoy):
    ini = dt.datetime(hoy.year, hoy.month, 1, tzinfo=TZ)
    fin = (dt.datetime(hoy.year + (hoy.month == 12), hoy.month % 12 + 1, 1, tzinfo=TZ) - dt.timedelta(milliseconds=1))
    return ini, fin


def iso(d):
    return d.astimezone(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.") + f"{d.microsecond // 1000:03d}Z"


def pedir_excel(s, ruta, params, intentos=6, espera=30):
    """Llama al endpoint de reporte. Puede devolver un enlace .xlsx (texto) o el Excel directamente."""
    ultimo = ""
    for k in range(1, intentos + 1):
        try:
            r = s.get(BASE + ruta, params=params, timeout=240)
            ct = r.headers.get("content-type", "")
            if r.status_code == 200:
                if r.content[:2] == b"PK":
                    log(f"{ruta}: Excel directo ({len(r.content)//1024} KB, intento {k})")
                    return r.content
                t = r.text.strip().strip('"')
                if re.match(r"^https://\S+\.xlsx(\?\S*)?$", t, re.I):
                    f = s.get(t, timeout=240)
                    if f.status_code == 200 and f.content[:2] == b"PK":
                        log(f"{ruta}: Excel por enlace ({len(f.content)//1024} KB, intento {k})")
                        return f.content
                    ultimo = f"enlace HTTP {f.status_code}"
                else:
                    ultimo = f"respuesta inesperada ({ct}, {len(r.content)} bytes)"
            else:
                ultimo = f"HTTP {r.status_code}"
        except Exception as e:
            ultimo = type(e).__name__
        log(f"{ruta}: intento {k} sin éxito ({ultimo}); reintento en {espera}s")
        time.sleep(espera)
    raise RuntimeError(f"{ruta}: {ultimo}")


# ---------------------------------------------------------------- utilidades
def filas_xlsx(data, hoja=0):
    wb = openpyxl.load_workbook(io.BytesIO(data) if isinstance(data, bytes) else data, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[hoja]]
    ws.reset_dimensions()
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    while rows and not any(v not in (None, "") for v in rows[-1]):
        rows.pop()
    return rows


def escribir_csv(nombre, header, rows):
    os.makedirs(OUT, exist_ok=True)
    p = os.path.join(OUT, nombre)
    with open(p, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(header)
        for r in rows:
            w.writerow(["" if v is None or (isinstance(v, float) and v != v) else v for v in r])
    return p


def num(v):
    try:
        return float(str(v).replace(",", "."))
    except Exception:
        return None


# ---------------------------------------------------------------- OWD
COLS_OWD = ["Agencia", "Evaluador", "Posición Evaluador", "Evaluado", "Fecha evaluación", "Type", "Pillar",
            "Proceso", "Tarea", "Descripción", "Puntuación"]


def procesar_owd(data):
    rows = filas_xlsx(data)
    h = [str(x).strip() if x else "" for x in rows[0]]
    falt = [c for c in COLS_OWD if c not in h]
    if falt:
        raise RuntimeError("reporte OWD sin columnas: " + ", ".join(falt))
    I = {c: h.index(c) for c in COLS_OWD}
    out = []
    for r in rows[1:]:
        if not r or not r[I["Evaluador"]]:
            continue
        v = [r[I[c]] for c in COLS_OWD]
        nok = str(r[I["Puntuación"]] or "").strip().lower() == "no ok"
        v[COLS_OWD.index("Descripción")] = r[I["Descripción"]] if nok else ""
        v[COLS_OWD.index("Tarea")] = r[I["Tarea"]] if nok else ""
        out.append(v)
    escribir_csv("owd_ejecucion.csv", COLS_OWD, out)
    owd = len({(r[4], r[7], r[1], r[3]) for r in out})
    return {"filas": len(out), "owd": owd}


# ---------------------------------------------------------------- 5S
COLS_5S = ["Agencia", "Área", "Subárea", "Evaluador", "Fecha evaluación", "S", "Ponderación (%)", "Placa de la unidad",
           "Título", "Puntuación", "Ponderación (#)", "Máx. resultado", "Descripción"]


def procesar_5s(data):
    rows = filas_xlsx(data)
    h = [str(x).strip() if x else "" for x in rows[0]]
    falt = [c for c in COLS_5S if c not in h]
    if falt:
        raise RuntimeError("reporte 5S sin columnas: " + ", ".join(falt))
    I = [h.index(c) for c in COLS_5S]
    e5 = [[r[i] for i in I] for r in rows[1:] if r and r[h.index("Evaluador")]]
    # puntaje máximo por pregunta = Máx. resultado / n.º de preguntas de esa S en esa auditoría;
    # la Descripción (texto de la pregunta) solo se conserva cuando la pregunta perdió puntos
    from collections import Counter
    ka = lambda r: (r[3], r[4], r[2], r[7], r[5])
    nq = Counter(ka(r) for r in e5)
    for r in e5:
        sc, mx = num(r[10]), num(r[11])
        qmax = mx / nq[ka(r)] if mx and nq[ka(r)] else None
        perdio = sc is not None and qmax is not None and sc < qmax - 1e-9 and "no aplica" not in str(r[9]).lower()
        r[12] = str(r[12] or "").strip().strip('"').strip() if perdio else ""
    escribir_csv("5s_ejecucion.csv", COLS_5S, e5)
    aud = len({(r[3], r[4], r[2], r[7]) for r in e5})
    return {"filas": len(e5), "auditorias": aud}


# ---------------------------------------------------------------- Acciones (misma lógica de acciones_dpo.py)
COLS_ACC = ["Folio", "CD", "Dueno", "Posicion", "Pilar", "Rutina", "KPI", "Herramienta", "Descripcion", "Estatus",
            "Creacion", "Compromiso", "Cierre", "Origen"]
PILAR = {"SAFETY": "Seguridad", "DISTRIBUTION": "Distribución", "FLEET": "Flota", "WAREHOUSE": "Almacén",
         "PEOPLE": "Gente", "MANAGEMENT": "Gestión", "PLANNING": "Planeación", "FINANZAS": "Finanzas"}
ORIGEN = {"ACTION_LOG": "Plan de acción", "OWD": "OWD", "FIVE_S": "5S"}


def _limpio(x, n=None):
    s = "" if x is None or (isinstance(x, float) and x != x) else re.sub(r"\s+", " ", str(x)).strip()
    s = s.lstrip("=+@").strip()
    if n and len(s) > n:
        s = s[:n - 1].rstrip() + "…"
    return s


def _llave(s):
    s = unicodedata.normalize("NFD", s.lower())
    return re.sub(r"[^a-z0-9]", "", "".join(c for c in s if unicodedata.category(c) != "Mn"))


def _bonito(s):
    s = re.sub(r"\s+", " ", s.replace("_", " ")).strip()
    if not s:
        return ""
    if s.upper() == s and " " not in s and len(s) <= 5:
        return s
    if s.upper() == s or s.lower() == s:
        s = s.capitalize()
    return s[0].upper() + s[1:]


def _unificar(serie, vacio):
    b = serie.map(lambda v: _bonito(_limpio(v)))
    ll = b.map(_llave)
    forma = b[b != ""].groupby(ll[b != ""]).agg(lambda x: x.value_counts().index[0])
    return [forma.get(k, vacio) if k else vacio for k in ll]


def _fecha(x):
    if x is None or str(x).strip() == "":
        return ""
    d = pd.to_datetime(x, errors="coerce")
    return "" if pd.isna(d) else d.strftime("%Y-%m-%d")


def procesar_acciones(datas):
    frames = []
    for data in datas:
        x = pd.ExcelFile(io.BytesIO(data) if isinstance(data, bytes) else data)
        for sh in x.sheet_names:
            frames.append(x.parse(sh))
    d = pd.concat(frames, ignore_index=True)
    req = ["Folio Number", "Agencia", "Dueño", "Pilar", "Routina", "Estatus", "Fecha creación", "Fecha compromiso"]
    falt = [c for c in req if c not in d.columns]
    if falt:
        raise RuntimeError("reporte de acciones sin columnas: " + ", ".join(falt))
    d = d.drop_duplicates("Folio Number", keep="last")
    vac = lambda: pd.Series([""] * len(d), index=d.index)
    st = d["Estatus"].astype(str).str.upper().str.strip()
    out = pd.DataFrame({
        "Folio": d["Folio Number"].map(_limpio),
        "CD": d["Agencia"].map(_limpio),
        "Dueno": d["Dueño"].map(_limpio).str.title(),
        "Posicion": d.get("Posición de dueño", vac()).map(_limpio).str.capitalize(),
        "Pilar": d["Pilar"].astype(str).str.upper().str.strip().map(lambda p: PILAR.get(p, p.capitalize())),
        "Rutina": _unificar(d["Routina"], "Sin rutina"),
        "KPI": _unificar(d.get("KPI", vac()), "Sin indicador"),
        "Herramienta": [_bonito(_limpio(v)) for v in d.get("herramienta de gestión", vac())],
        "Descripcion": d["Descripción"].map(lambda v: _limpio(v, 1000)),
        "Estatus": st,
        "Creacion": d["Fecha creación"].map(_fecha),
        "Compromiso": d["Fecha compromiso"].map(_fecha),
        "Cierre": [(_fecha(u) or _fecha(w)) if s_ == "COMPLETED" else "" for u, w, s_ in
                   zip(d.get("Últ. actualización de estatus", vac()), d.get("Última actualización", vac()), st)],
        "Origen": d.get("Origen", vac()).map(lambda o: ORIGEN.get(str(o), _limpio(o))),
    })[COLS_ACC].sort_values(["CD", "Creacion"])
    escribir_csv("acciones.csv", COLS_ACC, out.itertuples(index=False, name=None))
    return {"filas": len(out), "abiertas": int((out.Estatus != "COMPLETED").sum())}


# ---------------------------------------------------------------- principal
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--owd-xlsx"); ap.add_argument("--s5-xlsx"); ap.add_argument("--acc-xlsx", nargs="*")
    ap.add_argument("--solo", choices=["owd", "5s", "acciones"], nargs="*",
                    help="qué fuentes actualizar (por defecto: OWD y 5S siempre; acciones una vez por hora)")
    a = ap.parse_args()
    ahora = dt.datetime.now(TZ)
    os.makedirs(OUT, exist_ok=True)
    pe = os.path.join(OUT, "estado.json")
    try:
        estado = json.load(open(pe, encoding="utf-8"))
    except Exception:
        estado = {"fuentes": {}}
    estado.setdefault("fuentes", {})
    local = bool(a.owd_xlsx or a.s5_xlsx or a.acc_xlsx)
    fuentes = a.solo or (["owd", "5s"] + (["acciones"] if ahora.minute < 15 or
                                         "acciones" not in estado["fuentes"] or
                                         not estado["fuentes"]["acciones"].get("ok") else []))
    s = None
    if not local:
        try:
            s = sesion()
        except Exception as e:
            log("ERROR sesión:", e)
            estado["ultimo_intento"] = ahora.isoformat(timespec="seconds")
            estado["error_sesion"] = str(e)
            json.dump(estado, open(pe, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
            sys.exit(1)
        estado.pop("error_sesion", None)
    ini, fin = rango_mes(ahora)
    pm = {"subzone": SUBZONE, "year": str(ahora.year), "month": str(ahora.month), "startTime": iso(ini),
          "endTime": iso(fin), "timezoneOffsetStart": "300", "timezoneOffsetEnd": "300", "timezoneOffset": "300"}
    fallos = 0
    for f in fuentes:
        t0 = time.time()
        try:
            if f == "owd":
                data = a.owd_xlsx if local else pedir_excel(s, "/api/v1/owd/excel", pm)
                res = procesar_owd(data)
            elif f == "5s":
                data = a.s5_xlsx if local else pedir_excel(s, "/api/v1/5s/excel", pm)
                res = procesar_5s(data)
            else:
                if local:
                    datas = a.acc_xlsx
                else:
                    y0 = dt.datetime(ahora.year, 1, 1, tzinfo=TZ)
                    y1 = dt.datetime(ahora.year + 1, 1, 1, tzinfo=TZ) - dt.timedelta(milliseconds=1)
                    datas = [pedir_excel(s, "/api/v1/actionlogs-report",
                                         {"distributionCenter": cd, "year": str(ahora.year), "startTime": iso(y0),
                                          "endTime": iso(y1), "timezoneOffsetStart": "300",
                                          "timezoneOffsetEnd": "300", "allMonths": "true"}, intentos=4)
                             for cd in CDS]
                res = procesar_acciones(datas)
            estado["fuentes"][f] = {"ok": True, "actualizado": dt.datetime.now(TZ).isoformat(timespec="seconds"),
                                    "segundos": round(time.time() - t0), **res}
            log(f"OK {f}: {res}")
        except Exception as e:
            fallos += 1
            prev = estado["fuentes"].get(f, {})
            prev.update({"ok": False, "error": str(e)[:300], "intento": dt.datetime.now(TZ).isoformat(timespec="seconds")})
            estado["fuentes"][f] = prev
            log(f"ERROR {f}: {e}")
    estado["ultimo_intento"] = dt.datetime.now(TZ).isoformat(timespec="seconds")
    json.dump(estado, open(pe, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    if fallos == len(fuentes):
        sys.exit(1)


if __name__ == "__main__":
    main()
