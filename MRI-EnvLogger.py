#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Registro Parametri Ambientali MRI — dashboard + upload a surge.sh
- Salvataggio parametri su SQLite
- Dashboard HTML responsive (ultima lettura + storico 30gg + grafico)
- Supporto **offline** per Chart.js (chart.umd.min.js)
- **NUOVO**: alla pressione di "Salva registrazione":
    1) Genera/aggiorna la dashboard in APP_DIR/dashboardmri/index.html
    2) Carica automaticamente la cartella su https://dashboardmri.surge.sh
       cercando 'surge' o 'npx' anche se non sono nel PATH.
"""
import os
import sys
import json
import shutil
import sqlite3
import subprocess
import time
import threading
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# ---- SUBPROCESS WINDOWS NO-CONSOLE HELPER ----
def _popen_no_window(cmd, **kwargs):
    """
    Avvia un processo figlio senza aprire finestre console su Windows.
    Su altri sistemi operativi si comporta come subprocess.Popen.
    - Non usare shell=True (qui impostiamo sempre shell=False).
    - Catturiamo stdout/stderr a PIPE per poter aggiornare la GUI.
    """
    import subprocess, sys
    kwargs = dict(kwargs)  # copy
    kwargs.setdefault("shell", False)
    # Assicura cattura I/O (così niente console a schermo)
    if "stdout" not in kwargs:
        kwargs["stdout"] = subprocess.PIPE
    if "stderr" not in kwargs:
        kwargs["stderr"] = subprocess.STDOUT
    if "stdin" not in kwargs and kwargs.get("text") and kwargs.get("input") is not None:
        kwargs["stdin"] = subprocess.PIPE

    if sys.platform.startswith("win"):
        # HIDE window
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        si.wShowWindow = 0  # SW_HIDE
        CREATE_NO_WINDOW = 0x08000000
        kwargs["startupinfo"] = si
        kwargs["creationflags"] = (kwargs.get("creationflags", 0) | CREATE_NO_WINDOW)
    return subprocess.Popen(cmd, **kwargs)


# ===================== Percorsi / Costanti =====================

APP_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
DB_NAME = "mr_envlog.db"
DB_PATH = os.path.join(APP_DIR, DB_NAME)
CONFIG_PATH = os.path.join(APP_DIR, "config.json")

# Surge config richiesto dall'utente
SURGE_DOMAIN = ""
SURGE_EMAIL = ""
SURGE_PASSWORD = ""

# Cartella dashboard fissa richiesta
DASH_FOLDER_NAME = "dashboardmri"
DASH_DIR = os.path.join(APP_DIR, DASH_FOLDER_NAME)

# Colonne per tabelle/esporti
COLUMNS = [
    ("timestamp", "Data/Ora"),
    ("o2", "O2 (%)"),
    ("rh1", "RH Umidità 1 (%)"),
    ("temp1", "Temperatura 1 (°C)"),
    ("rh2", "RH Umidità 2 (%)"),
    ("temp2", "Temperatura 2 (°C)"),
    ("elio_ok", "Livello di elio (SI/NO)"),
    ("aspirazione_ok", "Aspirazione forzata (SI/NO)"),
    ("operatore", "Operatore"),
]

PLOT_NUM_KEYS = [
    ("o2", "O2 (%)"),
    ("rh1", "RH Umidità 1 (%)"),
    ("temp1", "Temperatura 1 (°C)"),
    ("rh2", "RH Umidità 2 (%)"),
    ("temp2", "Temperatura 2 (°C)"),
]

# ===================== Config =====================

DEFAULT_CONFIG = {"dashboard_dir": "", "chart_offline": False}

def load_config():
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg = json.load(f)
            for k, v in DEFAULT_CONFIG.items():
                cfg.setdefault(k, v)
            return cfg
    except Exception:
        return dict(DEFAULT_CONFIG)

def save_config(cfg: dict):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception as e:
        messagebox.showwarning("Configurazione", f"Impossibile salvare la configurazione:\n{e}")

# ===================== DB =====================

def init_db(db_path=DB_PATH):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT NOT NULL,
            o2 REAL NOT NULL,
            rh1 REAL NOT NULL,
            temp1 REAL NOT NULL,
            rh2 REAL NOT NULL,
            temp2 REAL NOT NULL,
            elio_ok TEXT NOT NULL CHECK(elio_ok IN ('SI','NO')),
            aspirazione_ok TEXT NOT NULL CHECK(aspirazione_ok IN ('SI','NO')),
            operatore TEXT NOT NULL
        )
        """
    )
    conn.commit()
    conn.close()

def insert_record(o2, rh1, temp1, rh2, temp2, elio_ok, aspirazione_ok, operatore, db_path=DB_PATH):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    now_iso = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur.execute(
        """
        INSERT INTO logs (timestamp, o2, rh1, temp1, rh2, temp2, elio_ok, aspirazione_ok, operatore)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (now_iso, o2, rh1, temp1, rh2, temp2, elio_ok, aspirazione_ok, operatore),
    )
    conn.commit()
    conn.close()
    return now_iso

def fetch_records(start=None, end=None, db_path=DB_PATH):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    query = "SELECT timestamp, o2, rh1, temp1, rh2, temp2, elio_ok, aspirazione_ok, operatore FROM logs"
    params = []

    if start and end:
        query += " WHERE timestamp BETWEEN ? AND ?"
        params.extend([normalize_to_iso(start, True), normalize_to_iso(end, False)])
    elif start:
        query += " WHERE timestamp >= ?"
        params.append(normalize_to_iso(start, True))
    elif end:
        query += " WHERE timestamp <= ?"
        params.append(normalize_to_iso(end, False))

    query += " ORDER BY timestamp ASC"
    cur.execute(query, params)
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows

def fetch_last_record():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute(
        "SELECT timestamp, o2, rh1, temp1, rh2, temp2, elio_ok, aspirazione_ok, operatore FROM logs ORDER BY timestamp DESC LIMIT 1"
    )
    row = cur.fetchone()
    conn.close()
    return dict(row) if row else None

# ===================== Utils =====================

def parse_float(value_str: str, field_name: str) -> float:
    if value_str is None:
        raise ValueError(f"Il campo '{field_name}' è vuoto.")
    s = value_str.strip().replace(",", ".")
    if s == "":
        raise ValueError(f"Il campo '{field_name}' è vuoto.")
    try:
        return float(s)
    except Exception:
        raise ValueError(f"Il campo '{field_name}' deve essere un numero. Valore dato: '{value_str}'")

def parse_it_date(s: str):
    s = s.strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    for fmt in ("%d/%m/%y %H:%M:%S", "%d/%m/%y %H:%M", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    raise ValueError("Formato data/ora non valido. Usa es. 'gg/mm/aa' o 'gg/mm/aaaa' (opzionale 'HH:MM').")

def normalize_to_iso(s: str, start_of_day: bool) -> str:
    dt = parse_it_date(s)
    only_date = (len(s.strip()) in (8, 10)) and ("/" in s or "-" in s)
    if only_date:
        if start_of_day:
            dt = dt.replace(hour=0, minute=0, second=0)
        else:
            dt = dt.replace(hour=23, minute=59, second=59)
    return dt.strftime("%Y-%m-%d %H:%M:%S")

def it_ts_display(ts_iso: str) -> str:
    try:
        dt = datetime.strptime(ts_iso, "%Y-%m-%d %H:%M:%S")
        return dt.strftime("%d/%m/%y %H:%M")
    except Exception:
        return ts_iso

def format_num(x) -> str:
    try:
        return f"{float(x):.2f}"
    except Exception:
        return str(x)

# ===================== Export elenco (HTML/PDF) =====================

def export_html(records, filepath, start=None, end=None):
    title = "Registro Parametri Ambientali MRI"
    period = ""
    if start or end:
        period = f"Intervallo: {start or '-'} → {end or '-'}"

    head = f"""<!doctype html>
<html lang="it">
<head>
<meta charset="utf-8">
<title>{title}</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
body {{ font-family: 'Segoe UI', Arial, Helvetica, sans-serif; margin: 40px; background:#f8f9fa; color:#2c3e50; }}
h1 {{ margin-bottom: .2rem; text-align:center; }}
p.meta {{ color: #555; margin-top: 0; text-align:center; }}
table {{ border-collapse: collapse; width: 100%; margin-top: 16px; background:#fff; box-shadow:0 2px 10px rgba(0,0,0,.05);}}
th, td {{ border-bottom: 1px solid #ecf0f1; padding: 10px; text-align: center; font-size: 14px; }}
th {{ background: #3498db; color:#fff; position:sticky; top:0; }}
tr:nth-child(even) {{ background: #f8f9fa; }}
.footer {{ margin-top: 24px; font-size: 12px; color: #777; text-align:center; }}
@media print {{
  body {{ margin: 0; }}
  h1 {{ font-size: 20px; }}
}}
</style>
</head>
<body>
<h1>🏥 {title}</h1>
<p class="meta">{period}</p>
<table>
<thead>
<tr>"""
    headers = "".join([f"<th>{label}</th>" for _, label in COLUMNS])
    rows_html = []
    for r in records:
        rows_html.append(f"""
<tr>
  <td>{it_ts_display(r['timestamp'])}</td>
  <td>{format_num(r['o2'])}</td>
  <td>{format_num(r['rh1'])}</td>
  <td>{format_num(r['temp1'])}</td>
  <td>{format_num(r['rh2'])}</td>
  <td>{format_num(r['temp2'])}</td>
  <td>{r['elio_ok']}</td>
  <td>{r['aspirazione_ok']}</td>
  <td>{r['operatore']}</td>
</tr>""")

    tail = f"""</tr>
</thead>
<tbody>
{''.join(rows_html)}
</tbody>
</table>
<div class="footer">Generato il {datetime.now().strftime('%d/%m/%y %H:%M')}</div>
</body>
</html>"""
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(head + headers + tail)

def export_pdf(records, filepath, start=None, end=None):
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib.styles import getSampleStyleSheet
    except Exception:
        raise RuntimeError(
            "Per l'esportazione in PDF è necessario installare 'reportlab'.\n"
            "Apri un terminale e digita: python -m pip install reportlab"
        )

    title = "Registro Parametri Ambientali MRI"
    period = ""
    if start or end:
        period = f"Intervallo: {start or '-'} → {end or '-'}"

    doc = SimpleDocTemplate(filepath, pagesize=landscape(A4), leftMargin=24, rightMargin=24, topMargin=24, bottomMargin=24)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph(f"🏥 {title}", styles["Title"]))
    if period:
        story.append(Paragraph(period, styles["Normal"]))
    story.append(Spacer(1, 12))

    data = [[label for _, label in COLUMNS]]
    for r in records:
        data.append([
            it_ts_display(r["timestamp"]),
            format_num(r["o2"]),
            format_num(r["rh1"]),
            format_num(r["temp1"]),
            format_num(r["rh2"]),
            format_num(r["temp2"]),
            r["elio_ok"],
            r["aspirazione_ok"],
            r["operatore"],
        ])

    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#3498db")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8f9fa")]),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
        ("RIGHTPADDING", (0,0), (-1,-1), 6),
        ("TOPPADDING", (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
    ]))

    story.append(table)
    story.append(Spacer(1, 12))
    story.append(Paragraph(f"Generato il {datetime.now().strftime('%d/%m/%y %H:%M')}", styles["Normal"]))
    doc.build(story)

# ===================== Dashboard HTML (con offline Chart.js) =====================

def generate_dashboard_html(out_dir: str, use_offline: bool = False):
    """
    Scrive:
      - dashboard_latest.html
      - dashboard_YYYYmmdd_HHMMSS.html
    Contenuto:
      - Ultima lettura completa
      - Storico 30 giorni precedenti
      - Grafico Chart.js; se use_offline=True prova a usare chart.umd.min.js locale
    """
    if not out_dir:
        raise RuntimeError("Cartella dashboard non impostata.")
    if not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    last = fetch_last_record()
    if not last:
        raise RuntimeError("Nessun record presente nel database.")

    last_dt = datetime.strptime(last["timestamp"], "%Y-%m-%d %H:%M:%S")
    start_dt = last_dt - timedelta(days=30)
    records = fetch_records(start_dt.strftime("%Y-%m-%d %H:%M:%S"), last_dt.strftime("%Y-%m-%d %H:%M:%S"))

    labels = []
    o2_vals, rh1_vals, t1_vals, rh2_vals, t2_vals = [], [], [], [], []
    table_rows = []
    for r in records:
        labels.append(it_ts_display(r["timestamp"]))
        o2_vals.append(float(r["o2"]))
        rh1_vals.append(float(r["rh1"]))
        t1_vals.append(float(r["temp1"]))
        rh2_vals.append(float(r["rh2"]))
        t2_vals.append(float(r["temp2"]))

        table_rows.append(f"""
<tr>
  <td>{it_ts_display(r['timestamp'])}</td>
  <td>{format_num(r['o2'])}</td>
  <td>{format_num(r['rh1'])}</td>
  <td>{format_num(r['temp1'])}</td>
  <td>{format_num(r['rh2'])}</td>
  <td>{format_num(r['temp2'])}</td>
  <td>{r['elio_ok']}</td>
  <td>{r['aspirazione_ok']}</td>
  <td>{r['operatore']}</td>
</tr>""")

    import json as _json

    # Decide script tag per Chart.js
    local_chart_in_outdir = os.path.join(out_dir, "chart.umd.min.js")
    local_chart_in_app = os.path.join(APP_DIR, "chart.umd.min.js")
    if use_offline and os.path.exists(local_chart_in_outdir):
        chart_tag = '<script src="chart.umd.min.js"></script>'
    elif use_offline and os.path.exists(local_chart_in_app):
        try:
            shutil.copy2(local_chart_in_app, local_chart_in_outdir)
            chart_tag = '<script src="chart.umd.min.js"></script>'
        except Exception:
            chart_tag = '<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>'
    else:
        chart_tag = '<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>'

    html = f"""<!doctype html>
<html lang="it">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Dashboard Parametri Ambientali MRI - Istituto di Cura Città di Pavia - Siemens Sola 1.5T</title>
{chart_tag}
<style>
  :root {{
    --bg:#0f172a; --card:#111827; --muted:#94a3b8; --accent:#38bdf8; --text:#e5e7eb;
  }}
  * {{ box-sizing: border-box; }}
  body {{ margin:0; font-family: system-ui,-apple-system,Segoe UI,Roboto,Ubuntu,Arial; background:var(--bg); color:var(--text); }}
  .wrap {{ max-width: 1100px; margin:0 auto; padding:16px; }}
  .title {{ font-size: clamp(20px, 3vw, 28px); font-weight:700; margin:8px 0 2px; }}
  .subtitle {{ color: var(--muted); margin-bottom: 14px; }}
  .grid {{ display:grid; grid-template-columns: repeat(12, 1fr); gap: 12px; }}
  .card {{ background:var(--card); border-radius:16px; padding:14px; box-shadow: 0 6px 24px rgba(0,0,0,.25); }}
  .span-12 {{ grid-column: span 12; }}
  .span-6 {{ grid-column: span 6; }}
  @media (max-width: 860px) {{ .span-6 {{ grid-column: span 12; }} }}
  .kv {{ display:grid; grid-template-columns: 1fr auto; gap:6px; padding:8px 0; border-bottom: 1px dashed rgba(255,255,255,.08); }}
  .kv:last-child {{ border-bottom:none; }}
  .k {{ color:var(--muted); }}
  .v {{ font-weight:700; }}
  .badge {{ display:inline-block; padding:4px 8px; border-radius:9999px; font-size:12px; }}
  .ok {{ background: rgba(34,197,94,.15); color:#86efac; }}
  .no {{ background: rgba(239,68,68,.15); color:#fecaca; }}
  table {{ width:100%; border-collapse: collapse; }}
  th, td {{ text-align:center; padding:8px 6px; border-bottom:1px solid rgba(255,255,255,.08); font-size:13px; }}
  th {{ color:var(--muted); position:sticky; top:0; background:var(--card); }}
  tr:nth-child(even) td {{ background: rgba(255,255,255,.02); }}
  .foot {{ color:var(--muted); font-size:12px; text-align:center; margin-top:14px; }}
  .chartbox {{ height: 260px; }}
</style>
</head>
<body>
  <div class="wrap">
    <div class="title">🏥 Dashboard Parametri Ambientali MRI - Istituto di Cura Città di Pavia - Siemens Sola 1.5T</div>
    <div class="subtitle">Aggiornato al {it_ts_display(last['timestamp'])} — Operatore: {last['operatore']}</div>

    <div class="grid">
      <div class="card span-6">
        <div class="title" style="font-size:18px;">Ultima lettura</div>
        <div class="kv"><div class="k">Data/Ora</div><div class="v">{it_ts_display(last['timestamp'])}</div></div>
        <div class="kv"><div class="k">O2 (%)</div><div class="v">{format_num(last['o2'])}</div></div>
        <div class="kv"><div class="k">RH 1 (%)</div><div class="v">{format_num(last['rh1'])}</div></div>
        <div class="kv"><div class="k">Temp 1 (°C)</div><div class="v">{format_num(last['temp1'])}</div></div>
        <div class="kv"><div class="k">RH 2 (%)</div><div class="v">{format_num(last['rh2'])}</div></div>
        <div class="kv"><div class="k">Temp 2 (°C)</div><div class="v">{format_num(last['temp2'])}</div></div>
        <div class="kv"><div class="k">Elio</div><div class="v"><span class="badge {"ok" if last['elio_ok']=="SI" else "no"}">{last['elio_ok']}</span></div></div>
        <div class="kv"><div class="k">Aspirazione</div><div class="v"><span class="badge {"ok" if last['aspirazione_ok']=="SI" else "no"}">{last['aspirazione_ok']}</span></div></div>
      </div>

      <div class="card span-6">
        <div class="title" style="font-size:18px;">Andamento ultimi 30 giorni</div>
        <div class="chartbox"><canvas id="chart"></canvas></div>
      </div>

      <div class="card span-12">
        <div class="title" style="font-size:18px;">Storico (30 giorni)</div>
        <div style="overflow:auto; max-height: 46vh;">
          <table>
            <thead>
              <tr>
                <th>Data/Ora</th>
                <th>O2 (%)</th>
                <th>RH 1 (%)</th>
                <th>Temp 1 (°C)</th>
                <th>RH 2 (%)</th>
                <th>Temp 2 (°C)</th>
                <th>Elio</th>
                <th>Aspirazione</th>
                <th>Operatore</th>
              </tr>
            </thead>
            <tbody>
              {''.join(table_rows)}
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <div class="foot">Generato automaticamente — {datetime.now().strftime('%d/%m/%y %H:%M')}</div>
  </div>

<script>
  const labels = {_json.dumps(labels, ensure_ascii=False)};
  const dataO2 = {_json.dumps(o2_vals)};
  const dataRH1 = {_json.dumps(rh1_vals)};
  const dataT1 = {_json.dumps(t1_vals)};
  const dataRH2 = {_json.dumps(rh2_vals)};
  const dataT2 = {_json.dumps(t2_vals)};

  const ctx = document.getElementById('chart').getContext('2d');
  new Chart(ctx, {{
    type: 'line',
    data: {{
      labels: labels,
      datasets: [
        {{ label: 'O2 (%)', data: dataO2, tension: .25 }},
        {{ label: 'RH 1 (%)', data: dataRH1, tension: .25 }},
        {{ label: 'Temp 1 (°C)', data: dataT1, tension: .25 }},
        {{ label: 'RH 2 (%)', data: dataRH2, tension: .25 }},
        {{ label: 'Temp 2 (°C)', data: dataT2, tension: .25 }}
      ]
    }},
    options: {{
      responsive: true,
      maintainAspectRatio: false,
      scales: {{
        x: {{ ticks: {{ maxRotation: 0, autoSkip: true }} }},
        y: {{ beginAtZero: false }}
      }}
    }}
  }});
</script>
</body>
</html>"""

    ts_name = datetime.now().strftime("%Y%m%d_%H%M%S")
    path_latest = os.path.join(out_dir, "dashboard_latest.html")
    path_ts = os.path.join(out_dir, f"dashboard_{ts_name}.html")
    with open(path_latest, "w", encoding="utf-8") as f:
        f.write(html)
    with open(path_ts, "w", encoding="utf-8") as f:
        f.write(html)
    return path_latest, path_ts

# ===================== Surge helpers =====================

def _possible_node_dirs():
    # Alcuni percorsi tipici su Windows per node/npm/npx
    dirs = []
    env = os.environ
    userprofile = env.get("USERPROFILE") or env.get("HOMEPATH") or ""
    appdata = env.get("APPDATA") or ""
    localapp = env.get("LOCALAPPDATA") or ""
    program_files = env.get("ProgramFiles") or ""
    program_files_x86 = env.get("ProgramFiles(x86)") or ""

    # standard npm bin
    if appdata:
        dirs.append(os.path.join(appdata, "npm"))
    if localapp:
        dirs.append(os.path.join(localapp, "Programs", "npm"))
        dirs.append(os.path.join(localapp, "Programs", "nodejs"))
    if program_files:
        dirs.append(os.path.join(program_files, "nodejs"))
    if program_files_x86:
        dirs.append(os.path.join(program_files_x86, "nodejs"))
    if userprofile:
        dirs.append(os.path.join(userprofile, "AppData", "Roaming", "npm"))
        # cache di npx per pacchetti scaricati
        dirs.append(os.path.join(userprofile, ".npm", "_npx"))
    return [d for d in dict.fromkeys(dirs) if d and os.path.isdir(d)]

def _which_executable(names):
    """Cerca un eseguibile tra PATH e cartelle note. Ritorna percorso completo o None."""
    from shutil import which
    # prova PATH
    for n in names:
        p = which(n)
        if p:
            return p
    # prova estensioni tipiche su Windows
    suffixes = ["", ".cmd", ".exe", ".bat"]
    for base in names:
        for d in _possible_node_dirs():
            for suf in suffixes:
                cand = os.path.join(d, base + suf)
                if os.path.isfile(cand):
                    return cand
    return None

def _run_subprocess(cmd, input_text=None, timeout=180, env=None):
    try:
        proc = _popen_no_window(
            cmd,
            stdin=subprocess.PIPE if input_text is not None else None,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            env=env or os.environ.copy(),
            cwd=APP_DIR,
        )
        out, _ = proc.communicate(input=input_text, timeout=timeout)
        return proc.returncode, out
    except subprocess.TimeoutExpired:
        try:
            proc.kill()
        except Exception:
            pass
        return 1, "Timeout esecuzione comando."
    except FileNotFoundError:
        return 1, "Comando non trovato."
    except Exception as e:
        return 1, f"Errore esecuzione comando: {e}"

def _surge_login_if_needed(surge_cmd):
    """Esegue 'surge whoami' e, se non autenticato, prova login non interattivo inviando email e password."""
    rc, out = _run_subprocess([surge_cmd, "whoami"], timeout=30)
    if rc == 0 and out and ("email" in out.lower() or "you are" in out.lower() or "logged" in out.lower()):
        return True, out
    # prova login: surge login -> invia email e password (ognuna seguita da newline)
    login_input = f"{SURGE_EMAIL}\n{SURGE_PASSWORD}\n"
    rc2, out2 = _run_subprocess([surge_cmd, "login"], input_text=login_input, timeout=60)
    # ritenta whoami
    rc3, out3 = _run_subprocess([surge_cmd, "whoami"], timeout=30)
    ok = (rc2 == 0 or rc3 == 0)
    return ok, (out or "") + "\n" + (out2 or "") + "\n" + (out3 or "")

def deploy_to_surge(folder: str) -> (bool, str):
    """
    Carica 'folder' su SURGE_DOMAIN.
    Tenta in ordine:
      1) surge (se installato globalmente)
      2) npx surge (senza necessità di installazione globale)
    Effettua login automatico se necessario.
    """
    # 1) prova surge
    surge_cmd = _which_executable(["surge"])
    if surge_cmd:
        ok_login, log1 = _surge_login_if_needed(surge_cmd)
        if not ok_login:
            # proseguiamo comunque, surge farà prompt/errore
            pass
        rc, out = _run_subprocess([surge_cmd, folder, SURGE_DOMAIN, "--yes"], timeout=240)
        if rc == 0:
            return True, out
        # continua provando npx se fallito
        last = f"[surge] rc={rc}\n{out}"
    else:
        last = "[surge] non trovato; provo npx"

    # 2) prova npx surge
    npx_cmd = _which_executable(["npx"])
    if not npx_cmd:
        return False, last + "\n[npx] non trovato. Installa Node.js (npx) o aggiungi surge alla PATH."

    # tentativo login con npx (potrebbe non essere necessario se auth presente in %USERPROFILE%\.surge\)
    # Non tutti gli ambienti accettano login non interattivo, ma ci proviamo.
    rc_login, out_login = _run_subprocess([npx_cmd, "surge", "whoami"], timeout=60)
    if rc_login != 0 or ("not logged" in (out_login or "").lower() or "no token" in (out_login or "").lower()):
        _run_subprocess([npx_cmd, "surge", "login"], input_text=f"{SURGE_EMAIL}\n{SURGE_PASSWORD}\n", timeout=120)

    rc2, out2 = _run_subprocess([npx_cmd, "surge", folder, SURGE_DOMAIN, "--yes"], timeout=300)
    if rc2 == 0:
        return True, out2
    return False, last + f"\n[npx surge] rc={rc2}\n{out2}"

# ===================== Finestra grafici (opzionale) =====================

class ChartWindow(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("📊 Grafico parametri - Intervallo")
        self.geometry("980x640")
        self.resizable(True, True)

        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")

        ttk.Label(top, text="Dal (gg/mm/aa o gg/mm/aaaa [HH:MM]):").grid(row=0, column=0, sticky="w", padx=(0,8))
        self.start_entry = ttk.Entry(top, width=20)
        self.start_entry.grid(row=0, column=1, sticky="w", padx=(0,16))

        ttk.Label(top, text="Al (gg/mm/aa o gg/mm/aaaa [HH:MM]):").grid(row=0, column=2, sticky="w", padx=(0,8))
        self.end_entry = ttk.Entry(top, width=20)
        self.end_entry.grid(row=0, column=3, sticky="w", padx=(0,16))

        self.chk_vars = {}
        chk_frame = ttk.LabelFrame(self, text="Parametri da plottare", padding=8)
        chk_frame.pack(fill="x", padx=8, pady=(0,8))
        for i, (key, label) in enumerate(PLOT_NUM_KEYS):
            var = tk.BooleanVar(value=(i == 0))
            self.chk_vars[key] = var
            ttk.Checkbutton(chk_frame, text=label, variable=var).grid(row=0, column=i, sticky="w", padx=8)

        btns = ttk.Frame(self, padding=8)
        btns.pack(fill="x")
        ttk.Button(btns, text="Disegna grafico", command=self.draw_chart).pack(side="left")
        self.btn_save_png = ttk.Button(btns, text="Salva PNG", command=self.save_png, state="disabled")
        self.btn_save_png.pack(side="left", padx=8)

        self.canvas_frame = ttk.Frame(self, padding=8)
        self.canvas_frame.pack(fill="both", expand=True)

        self._mpl_ready = False
        self._embed_ok = False
        self.figure = None
        self.canvas = None
        self.ax = None

    def _ensure_mpl(self):
        if self._mpl_ready:
            return True
        try:
            import matplotlib
            from matplotlib.figure import Figure
            try:
                matplotlib.use("TkAgg", force=True)
                from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
                self._FigureCanvasTkAgg = FigureCanvasTkAgg
                self._embed_ok = True
            except Exception:
                self._embed_ok = False
            self._Figure = Figure
            self._mpl_ready = True
            return True
        except ModuleNotFoundError:
            messagebox.showwarning(
                "matplotlib non disponibile",
                "Installa matplotlib con:\n\npython -m pip install matplotlib",
                parent=self
            )
            return False

    def draw_chart(self):
        if not self._ensure_mpl():
            return
        start = self.start_entry.get().strip() or None
        end = self.end_entry.get().strip() or None
        try:
            records = fetch_records(start, end)
        except Exception as e:
            messagebox.showerror("Errore filtri", str(e), parent=self)
            return

        selected_keys = [k for k,v in self.chk_vars.items() if v.get()]
        if not selected_keys:
            messagebox.showinfo("Selezione vuota", "Seleziona almeno un parametro da plottare.", parent=self)
            return

        xs, series = [], {k: [] for k in selected_keys}
        for r in records:
            try:
                dt = datetime.strptime(r["timestamp"], "%Y-%m-%d %H:%M:%S")
                xs.append(dt)
                for k in selected_keys:
                    series[k].append(float(r[k]))
            except Exception:
                pass

        if not xs:
            messagebox.showinfo("Nessun dato", "Nessun record nell'intervallo.", parent=self)
            return

        if self.canvas:
            try:
                self.canvas.get_tk_widget().destroy()
            except Exception:
                pass
            self.canvas = None

        self.figure = self._Figure(figsize=(11, 5.8), dpi=100)
        self.ax = self.figure.add_subplot(111)

        label_map = dict(PLOT_NUM_KEYS)
        for k in selected_keys:
            self.ax.plot(xs, series[k], label=label_map.get(k, k))

        self.ax.set_xlabel("Data/Ora")
        self.ax.set_ylabel("Valore")
        self.ax.legend()
        self.ax.grid(True)
        try:
            self.figure.autofmt_xdate()
        except Exception:
            pass

        if self._embed_ok:
            try:
                self.canvas = self._FigureCanvasTkAgg(self.figure, master=self.canvas_frame)
                self.canvas.draw()
                self.canvas.get_tk_widget().pack(fill="both", expand=True)
            except Exception:
                self._embed_ok = False

        self.btn_save_png.config(state="normal")

        if not self._embed_ok:
            messagebox.showinfo(
                "Nota",
                "Il backend TkAgg non è disponibile: il grafico non può essere integrato nella finestra.\n"
                "Puoi comunque salvarlo come PNG con il pulsante 'Salva PNG'.",
                parent=self
            )

    def save_png(self):
        if not self.figure:
            messagebox.showinfo("Nessun grafico", "Disegna prima un grafico.", parent=self)
            return
        filepath = filedialog.asksaveasfilename(
            parent=self,
            title="Salva grafico come PNG",
            defaultextension=".png",
            filetypes=[("PNG", "*.png")],
            initialfile="grafico_parametri.png",
        )
        if not filepath:
            return
        try:
            self.figure.savefig(filepath, dpi=150, bbox_inches="tight")
            messagebox.showinfo("Fatto", f"Grafico salvato in:\n{filepath}", parent=self)
        except Exception as e:
            messagebox.showerror("Errore salvataggio", f"Impossibile salvare il grafico:\n{e}")

# ===================== Stili UI =====================

class ModernCard(ttk.Frame):
    def __init__(self, parent, title="", **kwargs):
        super().__init__(parent, **kwargs)
        header = ttk.Frame(self, style="CardHeader.TFrame")
        header.pack(fill="x", pady=(0, 10))
        ttk.Label(header, text=title, style="CardTitle.TLabel").pack(pady=10)
        self.content = ttk.Frame(self, style="Card.TFrame")
        self.content.pack(fill="both", expand=True, padx=10, pady=(0, 10))

class AnimatedButton(ttk.Button):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
    def _on_enter(self, _e):
        self.configure(style="Hover.TButton")
    def _on_leave(self, _e):
        self.configure(style="Modern.TButton")

def configure_styles(root):
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass

    color_primary = '#2c3e50'
    color_secondary = '#3498db'
    color_success = '#2ecc71'
    color_warning = '#f39c12'
    color_danger = '#e74c3c'
    color_white = '#ffffff'

    style.configure("Main.TFrame", background="#f8f9fa")
    style.configure("Card.TFrame", background=color_white, relief="flat")
    style.configure("CardHeader.TFrame", background=color_primary)
    style.configure("CardTitle.TLabel", background=color_primary, foreground=color_white, font=("Segoe UI", 14, "bold"))
    style.configure("Modern.TLabel", background=color_white, foreground=color_primary, font=("Segoe UI", 10))

    style.configure("Modern.TButton",
                    font=("Segoe UI", 10, "bold"),
                    foreground=color_white,
                    background=color_secondary,
                    borderwidth=0,
                    padding=(15, 8))
    style.map("Modern.TButton", background=[("pressed", color_primary), ("active", color_primary)])
    style.configure("Hover.TButton", background=color_primary, padding=(15, 8))
    style.configure("Success.TButton", background=color_success, foreground=color_white, padding=(15, 8))
    style.configure("Warning.TButton", background=color_warning, foreground=color_white, padding=(15, 8))
    style.configure("Danger.TButton", background=color_danger, foreground=color_white, padding=(15, 8))

    style.configure("Modern.TEntry", fieldbackground=color_white, borderwidth=2, relief="solid", padding=6)
    style.configure("Modern.TRadiobutton", background=color_white, foreground=color_primary, font=("Segoe UI", 10))

# ===================== Viewer (elenco + export) =====================

class Viewer(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("📋 Visualizza / Esporta registro")
        self.geometry("1100x640")
        self.configure(bg="#f8f9fa")
        self.resizable(True, True)

        frm_filters = ttk.Frame(self, style="Main.TFrame", padding=8)
        frm_filters.pack(fill="x")

        ttk.Label(frm_filters, text="Dal (gg/mm/aa o gg/mm/aaaa [HH:MM]):").grid(row=0, column=0, sticky="w", padx=(0,8))
        self.start_entry = ttk.Entry(frm_filters, width=22, style="Modern.TEntry")
        self.start_entry.grid(row=0, column=1, sticky="w", padx=(0,16))

        ttk.Label(frm_filters, text="Al (gg/mm/aa o gg/mm/aaaa [HH:MM]):").grid(row=0, column=2, sticky="w", padx=(0,8))
        self.end_entry = ttk.Entry(frm_filters, width=22, style="Modern.TEntry")
        self.end_entry.grid(row=0, column=3, sticky="w", padx=(0,16))

        self.btn_refresh = AnimatedButton(frm_filters, text="Aggiorna elenco", command=self.refresh, style="Modern.TButton")
        self.btn_refresh.grid(row=0, column=4, padx=4)

        self.btn_export_html = AnimatedButton(frm_filters, text="Esporta HTML", command=self.do_export_html, style="Warning.TButton")
        self.btn_export_html.grid(row=0, column=5, padx=4)

        self.btn_export_pdf = AnimatedButton(frm_filters, text="Esporta PDF", command=self.do_export_pdf, style="Modern.TButton")
        self.btn_export_pdf.grid(row=0, column=6, padx=4)

        self.btn_chart = AnimatedButton(frm_filters, text="Grafico…", command=self.open_chart, style="Modern.TButton")
        self.btn_chart.grid(row=0, column=7, padx=4)

        frm_table = ttk.Frame(self, style="Main.TFrame", padding=(8,0,8,8))
        frm_table.pack(fill="both", expand=True)

        cols = [key for key, _ in COLUMNS]
        self.tree = ttk.Treeview(frm_table, columns=cols, show="headings", selectmode="extended")
        for key, label in COLUMNS:
            self.tree.heading(key, text=label, anchor="center")
            width = 150 if key == "timestamp" else 130
            self.tree.column(key, width=width, anchor="center")

        vsb = ttk.Scrollbar(frm_table, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(frm_table, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscroll=vsb.set, xscroll=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        frm_table.rowconfigure(0, weight=1)
        frm_table.columnconfigure(0, weight=1)

        self.refresh()

    def get_filters(self):
        start = self.start_entry.get().strip() or None
        end = self.end_entry.get().strip() or None
        return start, end

    def refresh(self):
        try:
            start, end = self.get_filters()
            records = fetch_records(start, end)
        except Exception as e:
            messagebox.showerror("Errore filtri", str(e), parent=self)
            return

        for iid in self.tree.get_children():
            self.tree.delete(iid)

        for r in records:
            values = [
                it_ts_display(r["timestamp"]),
                format_num(r["o2"]),
                format_num(r["rh1"]),
                format_num(r["temp1"]),
                format_num(r["rh2"]),
                format_num(r["temp2"]),
                r["elio_ok"],
                r["aspirazione_ok"],
                r["operatore"],
            ]
            self.tree.insert("", "end", values=values)

    def do_export_html(self):
        start, end = self.get_filters()
        try:
            records = fetch_records(start, end)
        except Exception as e:
            messagebox.showerror("Errore", str(e), parent=self)
            return
        if not records:
            messagebox.showinfo("Nessun dato", "Non ci sono record per l'intervallo selezionato.", parent=self)
            return
        default_name = (
            f"registro_MRI_{(start or 'inizio')}_to_{(end or 'fine')}.html"
            .replace(" ", "_").replace(":", "-").replace("/", "-")
        )
        filepath = filedialog.asksaveasfilename(
            parent=self,
            title="Salva come HTML",
            defaultextension=".html",
            filetypes=[("HTML", "*.html")],
            initialfile=default_name,
        )
        if not filepath:
            return
        try:
            export_html(records, filepath, start, end)
            messagebox.showinfo("Fatto", f"Esportazione HTML completata:\n{filepath}", parent=self)
        except Exception as e:
            messagebox.showerror("Errore esportazione", str(e), parent=self)

    def do_export_pdf(self):
        start, end = self.get_filters()
        try:
            records = fetch_records(start, end)
        except Exception as e:
            messagebox.showerror("Errore", str(e), parent=self)
            return
        if not records:
            messagebox.showinfo("Nessun dato", "Non ci sono record per l'intervallo selezionato.", parent=self)
            return
        default_name = (
            f"registro_MRI_{(start or 'inizio')}_to_{(end or 'fine')}.pdf"
            .replace(" ", "_").replace(":", "-").replace("/", "-")
        )
        filepath = filedialog.asksaveasfilename(
            parent=self,
            title="Salva come PDF",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile=default_name,
        )
        if not filepath:
            return
        try:
            export_pdf(records, filepath, start, end)
            messagebox.showinfo("Fatto", f"Esportazione PDF completata:\n{filepath}", parent=self)
        except RuntimeError as re:
            messagebox.showwarning("ReportLab mancante", str(re), parent=self)
        except Exception as e:
            messagebox.showerror("Errore esportazione", str(e), parent=self)

    def open_chart(self):
        ChartWindow(self)

# ===================== App principale =====================


class ProgressDialog(tk.Toplevel):
    def __init__(self, master, message="Operazione in corso…"):
        super().__init__(master)
        self.title("Caricamento…")
        self.resizable(False, False)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", lambda: None)  # disabilita chiusura
        self.geometry("+{}+{}".format(self.winfo_screenwidth()//2 - 180, self.winfo_screenheight()//2 - 60))
        frm = ttk.Frame(self, padding=16)
        frm.pack(fill="both", expand=True)
        self.label = ttk.Label(frm, text=message, anchor="center", justify="center")
        self.label.pack(fill="x", pady=(0,10))
        self.pb = ttk.Progressbar(frm, mode="indeterminate", length=320)
        self.pb.pack(fill="x")
        try:
            self.pb.start(10)
        except Exception:
            pass
        self.update_idletasks()

    def set_message(self, msg:str):
        self.label.config(text=msg)
        self.update_idletasks()

    def close(self):
        try:
            self.pb.stop()
        except Exception:
            pass
        self.grab_release()
        self.destroy()

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("🏥 Registro Parametri Ambientali - MRI (dashboard + surge)")
        try:
            self.state("zoomed")
        except tk.TclError:
            try:
                self.attributes("-zoomed", True)
            except tk.TclError:
                self.geometry("1080x720")
        self.configure(bg="#f8f9fa")
        self.resizable(True, True)

        configure_styles(self)

        # Config
        self.config_data = load_config()

        main_container = ttk.Frame(self, style="Main.TFrame", padding=16)
        main_container.pack(fill="both", expand=True)

        # Header
        header_frame = ttk.Frame(main_container, style="Main.TFrame")
        header_frame.pack(fill="x", pady=(0, 12))
        title_bar = ttk.Frame(header_frame, style="CardHeader.TFrame")
        title_bar.pack(fill="x")
        ttk.Label(title_bar, text="🏥 SISTEMA MONITORAGGIO AMBIENTALE MRI", style="CardTitle.TLabel", font=("Segoe UI", 18, "bold")).pack(pady=12)
        ttk.Label(header_frame, text="Registrazione parametri ambientali per Risonanza Magnetica - Istituto di Cura Città di Pavia - Danilo Savioni 2025", style="Modern.TLabel", font=("Segoe UI", 11, "italic"), background="#f8f9fa").pack(pady=(8, 0))

        # Layout 2 colonne
        content_frame = ttk.Frame(main_container, style="Main.TFrame")
        content_frame.pack(fill="both", expand=True, pady=(16, 0))

        # Sidebar (form)
        sidebar = ModernCard(content_frame, title="📝 INSERIMENTO PARAMETRI")
        sidebar.pack(side="left", fill="y", padx=(0, 20), ipadx=12, ipady=8)

        # Variabili
        self.o2 = tk.StringVar()
        self.rh1 = tk.StringVar()
        self.temp1 = tk.StringVar()
        self.rh2 = tk.StringVar()
        self.temp2 = tk.StringVar()
        self.elio_var = tk.StringVar(value="NO")
        self.aspirazione_var = tk.StringVar(value="NO")
        self.operatore_var = tk.StringVar()

        # Form fields
        form = ttk.Frame(sidebar.content, style="Card.TFrame")
        form.pack(fill="x")

        def add_field(label_text, var):
            fr = ttk.Frame(form, style="Card.TFrame")
            fr.pack(fill="x", pady=6)
            ttk.Label(fr, text=label_text, style="Modern.TLabel").pack(anchor="w", pady=(0, 3))
            ttk.Entry(fr, textvariable=var, style="Modern.TEntry").pack(fill="x")

        add_field("🌡️ Percentuale O2 (%)", self.o2)
        add_field("💧 RH Umidità 1 (%)", self.rh1)
        add_field("🌡️ Temperatura 1 (°C)", self.temp1)
        add_field("💧 RH Umidità 2 (%)", self.rh2)
        add_field("🌡️ Temperatura 2 (°C)", self.temp2)

        def add_radio(label_text, var):
            fr = ttk.Frame(form, style="Card.TFrame")
            fr.pack(fill="x", pady=6)
            ttk.Label(fr, text=label_text, style="Modern.TLabel").pack(anchor="w", pady=(0, 3))
            rfr = ttk.Frame(fr, style="Card.TFrame"); rfr.pack(fill="x")
            ttk.Radiobutton(rfr, text="SI", value="SI", variable=var, style="Modern.TRadiobutton").pack(side="left", padx=6)
            ttk.Radiobutton(rfr, text="NO", value="NO", variable=var, style="Modern.TRadiobutton").pack(side="left", padx=6)

        add_radio("⚡ Controllo Livello di elio", self.elio_var)
        add_radio("🌪️ Controllo Aspirazione", self.aspirazione_var)

        add_field("👨‍⚕️ Operatore (max 10 caratteri)", self.operatore_var)

        # Colonna destra
        right = ttk.Frame(content_frame, style="Main.TFrame")
        right.pack(side="right", fill="both", expand=True)

        right.grid_rowconfigure(0, weight=0)
        right.grid_rowconfigure(1, weight=1)
        right.grid_columnconfigure(0, weight=1)

        info_card = ModernCard(right, title="ℹ️ INFORMAZIONI SISTEMA")
        info_card.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        info_card.grid_propagate(False)
        info_card.configure(height=190)

        info_text = (
            "📋 Sistema di monitoraggio ambientale per Risonanza Magnetica\n"
            "🔍 Acquisizione manuale dei parametri\n"
            "📊 Filtri, grafici e report HTML/PDF\n"
            "🌐 Dashboard pubblicata su surge.sh ad ogni salvataggio\n"
            "🔌 Modalità OFFLINE Chart.js (se abilitata e presente chart.umd.min.js)\n"
            "💾 Backup e ripristino database"
        )
        ttk.Label(info_card.content, text=info_text, style="Modern.TLabel", justify="left", anchor="w").pack(anchor="w", padx=10, pady=6)

        # Toggle offline Chart.js
        self.chart_offline_var = tk.BooleanVar(value=bool(self.config_data.get("chart_offline")))
        toggle = ttk.Checkbutton(info_card.content, text="Usa Chart.js offline (se presente in cartella)", variable=self.chart_offline_var, command=self._toggle_offline)
        toggle.pack(anchor="w", padx=10, pady=(0,6))

        reg_card = ModernCard(right, title="📋 REGISTRO — Tutti i record")
        reg_card.grid(row=1, column=0, sticky="nsew")
        reg_card.grid_propagate(False)
        reg_card.configure(height=420)

        toolbar = ttk.Frame(reg_card.content, style="Card.TFrame")
        toolbar.pack(fill="x", pady=(6, 0))
        AnimatedButton(toolbar, text="Aggiorna", command=lambda: self.refresh_main_registry(), style="Modern.TButton").pack(side="left")
        self.lbl_dash = ttk.Label(toolbar, text=self._dashboard_label_text(), style="Modern.TLabel")
        self.lbl_dash.pack(side="right")

        table_frame = ttk.Frame(reg_card.content, style="Card.TFrame")
        table_frame.pack(fill="both", expand=True, pady=8)

        cols = [key for key, _ in COLUMNS]
        self.main_tree = ttk.Treeview(table_frame, columns=cols, show="headings", selectmode="browse")
        for key, label in COLUMNS:
            self.main_tree.heading(key, text=label, anchor="center")
            width = 150 if key == "timestamp" else 130
            self.main_tree.column(key, width=width, anchor="center")

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.main_tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.main_tree.xview)
        self.main_tree.configure(yscroll=vsb.set, xscroll=hsb.set)

        self.main_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        # Barra inferiore
        bottom_bar = ttk.Frame(main_container, style="Main.TFrame")
        bottom_bar.pack(fill="x", pady=(10, 0))

        for i in range(4):
            bottom_bar.columnconfigure(i, weight=1, uniform="btns")

        b1 = AnimatedButton(bottom_bar, text="💾 Salva registrazione", command=self.save_record, style="Success.TButton")
        b2 = AnimatedButton(bottom_bar, text="📋 Visualizza / Esporta registro", command=self.open_viewer, style="Modern.TButton")
        b3 = AnimatedButton(bottom_bar, text="📊 Grafico", command=self.open_chart, style="Modern.TButton")
        b4 = AnimatedButton(bottom_bar, text="📈 Salva PNG grafico", command=self.open_chart_and_prompt_save, style="Modern.TButton")
        b5 = AnimatedButton(bottom_bar, text="💾 Backup database", command=self.backup_database, style="Modern.TButton")
        b6 = AnimatedButton(bottom_bar, text="🔄 Ripristino database", command=self.restore_database, style="Danger.TButton")
        b7 = AnimatedButton(bottom_bar, text="🌐 Imposta cartella dashboard…", command=self.set_dashboard_dir, style="Warning.TButton")
        b8 = AnimatedButton(bottom_bar, text="🧪 Rigenera dashboard ora", command=self.manual_generate_dashboard, style="Modern.TButton")
        b9 = AnimatedButton(bottom_bar, text="⬇️ Copia Chart.js offline…", command=self.copy_chart_js_to_dashboard, style="Modern.TButton")

        b1.grid(row=0, column=0, sticky="ew", padx=6, pady=4)
        b2.grid(row=0, column=1, sticky="ew", padx=6, pady=4)
        b3.grid(row=0, column=2, sticky="ew", padx=6, pady=4)
        b4.grid(row=0, column=3, sticky="ew", padx=6, pady=4)
        b5.grid(row=1, column=0, sticky="ew", padx=6, pady=4)
        b6.grid(row=1, column=1, sticky="ew", padx=6, pady=4)
        b7.grid(row=1, column=2, sticky="ew", padx=6, pady=4)
        b8.grid(row=1, column=3, sticky="ew", padx=6, pady=4)
        b9.grid(row=2, column=0, columnspan=4, sticky="ew", padx=6, pady=4)

        self.status = tk.StringVar(value="Pronto")
        ttk.Label(main_container, textvariable=self.status, relief="sunken", anchor="w").pack(fill="x", pady=(10,0))

        self.operatore_var.trace_add("write", self._limit_operatore)
        self.refresh_main_registry()

        # Aggiorna label posizione dashboard predefinita (fissa in APP_DIR)
        self.lbl_dash.config(text=self._dashboard_label_text())

    # ------- Helpers UI -------

    def _limit_operatore(self, *args):
        s = self.operatore_var.get()
        if len(s) > 10:
            self.operatore_var.set(s[:10])

    def _toggle_offline(self):
        self.config_data["chart_offline"] = bool(self.chart_offline_var.get())
        save_config(self.config_data)
        self.status.set("Preferenza Chart.js offline aggiornata")

    def _dashboard_label_text(self):
        # usa cartella fissa APP_DIR/dashboardmri come richiesto
        return f"Cartella dashboard (fissa): {DASH_DIR}"

    # ------- Azioni -------

    def save_record(self):
        try:
            o2 = parse_float(self.o2.get(), "Percentuale O2")
            rh1 = parse_float(self.rh1.get(), "RH Umidità 1")
            temp1 = parse_float(self.temp1.get(), "Temperatura 1")
            rh2 = parse_float(self.rh2.get(), "RH Umidità 2")
            temp2 = parse_float(self.temp2.get(), "Temperatura 2")
            elio_ok = self.elio_var.get()
            aspirazione_ok = self.aspirazione_var.get()
            operatore = (self.operatore_var.get() or "").strip()
            if not operatore:
                raise ValueError("Il campo 'Operatore' non può essere vuoto.")
            if len(operatore) > 10:
                operatore = operatore[:10]

            ts_iso = insert_record(o2, rh1, temp1, rh2, temp2, elio_ok, aspirazione_ok, operatore)
            self.status.set(f"Registrazione salvata alle {it_ts_display(ts_iso)}")
            messagebox.showinfo("Salvato", f"Registrazione salvata alle {it_ts_display(ts_iso)}")
            self._clear_fields()
            self.refresh_main_registry()

            # 1) genera/aggiorna dashboard in APP_DIR/dashboardmri + index.html
            latest, tsfile = self._generate_dashboard_to_fixed_dir()

            
            # 2) esegue upload su surge.sh con finestra di attesa
            self.status.set("Caricamento su surge.sh in corso…")
            progress = ProgressDialog(self, "Attendere il caricamento online della dashboard…\nNon chiudere l'applicazione.")
            def _worker():
                ok, log = deploy_to_surge(DASH_DIR)
                def _done():
                    try:
                        progress.close()
                    except Exception:
                        pass
                    if ok:
                        self.status.set("Dashboard pubblicata su surge.sh")
                        messagebox.showinfo("Surge", "✅ Pubblicazione completata su:\n" + SURGE_DOMAIN)
                    else:
                        self.status.set("Errore pubblicazione surge.sh")
                        snippet = (log or "").strip()
                        if len(snippet) > 1500:
                            snippet = snippet[:1500] + "..."
                        messagebox.showwarning("Surge", "⚠️ Pubblicazione non riuscita.\n\nDettagli:\n" + snippet)
                self.after(0, _done)
            threading.Thread(target=_worker, daemon=True).start()

        except ValueError as ve:
            messagebox.showerror("Errore di validazione", str(ve))
        except Exception as e:
            messagebox.showerror("Errore", f"Si è verificato un errore: {e}")

    def _generate_dashboard_to_fixed_dir(self):
        # Crea cartella se manca
        os.makedirs(DASH_DIR, exist_ok=True)
        # Genera dashboard (usa preferenza offline)
        latest, tsfile = generate_dashboard_html(DASH_DIR, use_offline=bool(self.config_data.get("chart_offline")))
        # Copia/aggiorna index.html dalla versione latest
        try:
            index_path = os.path.join(DASH_DIR, "index.html")
            shutil.copy2(latest, index_path)
        except Exception:
            # Se per qualche motivo non riesce a copiare, tenta a scrivere direttamente
            with open(os.path.join(DASH_DIR, "index.html"), "w", encoding="utf-8") as f:
                with open(latest, "r", encoding="utf-8") as src:
                    f.write(src.read())
        return latest, tsfile

    def _clear_fields(self):
        self.o2.set("")
        self.rh1.set("")
        self.temp1.set("")
        self.rh2.set("")
        self.temp2.set("")
        self.elio_var.set("NO")
        self.aspirazione_var.set("NO")

    def open_viewer(self):
        Viewer(self)

    def open_chart(self):
        ChartWindow(self)

    def open_chart_and_prompt_save(self):
        ChartWindow(self)

    def refresh_main_registry(self):
        try:
            records = fetch_records()
        except Exception as e:
            messagebox.showerror("Errore lettura registro", str(e), parent=self)
            return

        for iid in self.main_tree.get_children():
            self.main_tree.delete(iid)

        for r in records:
            values = [
                it_ts_display(r["timestamp"]),
                format_num(r["o2"]),
                format_num(r["rh1"]),
                format_num(r["temp1"]),
                format_num(r["rh2"]),
                format_num(r["temp2"]),
                r["elio_ok"],
                r["aspirazione_ok"],
                r["operatore"],
            ]
            self.main_tree.insert("", "end", values=values)

    # ------- Backup / Ripristino --------

    def backup_database(self):
        try:
            if not os.path.exists(DB_PATH):
                init_db()
            default_name = f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
            filepath = filedialog.asksaveasfilename(
                parent=self,
                title="Salva backup database",
                defaultextension=".db",
                filetypes=[("Database SQLite", "*.db"), ("Tutti i file", "*.*")],
                initialfile=default_name,
            )
            if not filepath:
                return
            shutil.copy2(DB_PATH, filepath)
            messagebox.showinfo("Backup completato", f"Backup salvato in:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Errore backup", f"Impossibile eseguire il backup:\n{e}")

    def restore_database(self):
        try:
            filepath = filedialog.askopenfilename(
                parent=self,
                title="Seleziona file di backup",
                filetypes=[("Database SQLite", "*.db *.sqlite *.bak *.dbbak"), ("Tutti i file", "*.*")],
            )
            if not filepath:
                return
            if not os.path.exists(filepath):
                messagebox.showerror("File non trovato", "Il file selezionato non esiste.")
                return
            if not messagebox.askyesno(
                "Conferma ripristino",
                "Questa operazione sovrascriverà il database corrente.\nProcedere?",
                icon="warning",
            ):
                return
            shutil.copy2(filepath, DB_PATH)
            messagebox.showinfo(
                "Ripristino completato",
                "Il database è stato ripristinato con successo.\nRiavvia l'applicazione per essere certo di leggere i dati aggiornati."
            )
            self.refresh_main_registry()
        except Exception as e:
            messagebox.showerror("Errore ripristino", f"Impossibile ripristinare il database:\n{e}")

    # ------- Dashboard config --------

    def set_dashboard_dir(self):
        # Manteniamo la funzione per compatibilità, ma informiamo che ora è fissa.
        messagebox.showinfo(
            "Dashboard",
            "La dashboard ora viene sempre generata nella cartella fissa:\n\n"
            f"{DASH_DIR}\n\nIl file principale è 'index.html'.\n"
            "Aggiungi qui dentro 'chart.umd.min.js' se vuoi la modalità OFFLINE."
        )

    def manual_generate_dashboard(self):
        try:
            latest, tsfile = self._generate_dashboard_to_fixed_dir()
            messagebox.showinfo("Dashboard", f"Dashboard rigenerata:\n{os.path.join(DASH_DIR,'index.html')}\n\nCopia storica:\n{tsfile}")
            self.status.set(f"Dashboard aggiornata: {os.path.join(DASH_DIR,'index.html')}")
        except Exception as e:
            messagebox.showwarning("Dashboard", f"Impossibile generare la dashboard:\n{e}")

    def copy_chart_js_to_dashboard(self):
        src = filedialog.askopenfilename(
            parent=self,
            title="Seleziona chart.umd.min.js",
            filetypes=[("Chart.js UMD", "chart.umd.min.js;*.js"), ("Tutti i file", "*.*")]
        )
        if not src:
            return
        try:
            os.makedirs(DASH_DIR, exist_ok=True)
            dst = os.path.join(DASH_DIR, "chart.umd.min.js")
            shutil.copy2(src, dst)
            messagebox.showinfo("Chart.js", f"Copiato in:\n{dst}\n\nOra puoi abilitare 'Usa Chart.js offline'.")
        except Exception as e:
            messagebox.showerror("Chart.js", f"Errore durante la copia:\n{e}")

# ===================== Main =====================

def main():
    init_db()
    try:
        os.makedirs(DASH_DIR, exist_ok=True)
    except Exception:
        pass
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()