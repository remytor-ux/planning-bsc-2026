#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Génère une page HTML (index.html) reproduisant fidèlement la mise en page du
fichier BSC_Planning_2026V2.xlsx : grille journalière + tableau horaire récap,
un onglet par mois, en lecture seule.

Usage :
    python3 generate_planning_v2.py chemin/vers/BSC_Planning_2026V2.xlsx dossier_sortie/
"""

import re
import sys
import json
import unicodedata
import openpyxl
from pathlib import Path
from datetime import datetime

MOIS_ORDRE = [
    "JANVIER", "FEVRIER", "MARS", "AVRIL", "MAI", "JUIN",
    "JUILLET", "AOUT", "SEPTEMBRE", "OCTOBRE", "NOVEMBRE", "DECEMBRE",
]

# Couleur par mot-clé détecté dans le libellé de colonne (palette navy/amber/red/blue, pas de vert)
COLOR_RULES = [
    ("19H00-9H00", "#8B1E2B"),   # nuit -> rouge
    ("19H-9H", "#8B1E2B"),
    ("9H00-21H00", "#1F3A5F"),   # 12h -> navy
    ("9H21", "#1F3A5F"),
    ("9H00-19H00", "#C99A3B"),   # jour -> amber
    ("9H-19H", "#C99A3B"),
    ("FFI", "#3B6EA5"),          # bleu
    ("UHCD", "#5A4B81"),         # violet
    ("RUF", "#2E6E6E"),          # sarcelle sobre
    ("DPI", "#6B5B3E"),          # brun
    ("USC", "#6B5B3E"),          # brun (même famille que DPI, poste équivalent)
    ("CA", "#7A7A7A"),           # gris
    ("INDISPO", "#B0413E"),
    ("FORMATION", "#4A5D8A"),
    ("AT", "#A0522D"),
]
VACANT_COLOR = "#8B1E2B"  # rouge : poste vacant (cellule rouge sans nom dans le fichier source)


def color_for_label(label):
    up = label.upper().replace("\n", " ")
    for key, color in COLOR_RULES:
        if key in up:
            return color
    return "#555555"


def is_red_fill(cell):
    """Détecte une cellule remplie en rouge (poste vacant) dans le fichier source."""
    fg = cell.fill.fgColor
    try:
        if fg.type == "rgb" and fg.rgb and str(fg.rgb).endswith("FF0000"):
            return True
    except Exception:
        pass
    return False



# ---------------------------------------------------------------------------
# Reconnaissance des noms dans les cases du planning (pour le sélecteur de médecin)
# Les cases sont en texte libre : "adrien/IMAD", "TODY J AFZAL N", "hach", etc.
# Les variantes d'orthographe se règlent ici, en une ligne par variante.
# ---------------------------------------------------------------------------
# variante (minuscules, sans accent) -> nom retenu (minuscules, sans accent)
NAME_ALIASES = {
    "hach": "hachimi", "hachim": "hachimi", "hchij": "hachimi",
    "kanche": "kantche",
    "rodrigues": "rodrigue",
    "lolaj": "lola",
    "nesrinej": "nesrine",
    "sid": "sid ahmed", "ahmed": "sid ahmed",
    "remy": "torregrossa",   # Dr Torregrossa écrit parfois "remy" dans les cases
}
# mots qui ne sont pas des noms de personnes
NAME_STOPWORDS = {
    "jour", "nuit", "bloc", "usc", "interne", "pas", "journee", "evc", "la",
    "uh", "am", "pm",
}
# libellés avec accents pour l'affichage (nom normalisé -> libellé)
NAME_DISPLAY = {"stephane": "Stéphane", "chloe": "Chloé"}


def _norm(text):
    text = unicodedata.normalize("NFD", str(text))
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    return text.lower()


def names_in(text):
    """Liste (sans doublon) des personnes citées dans un texte de case."""
    found = []
    for tok in re.findall(r"[a-z]+", _norm(text)):
        tok = NAME_ALIASES.get(tok, tok)
        if len(tok) < 3 or tok in NAME_STOPWORDS:
            continue
        if tok not in found:
            found.append(tok)
    return found


def parse_month_sheet(ws):
    # --- Colonnes journalières : lues dynamiquement depuis la ligne d'en-têtes (row 3)
    # et les regroupements de la ligne 2 (LIGNE 1 / 12H / LIGNE 2 / AUTRES / UHCD...),
    # car la structure varie légèrement d'un mois à l'autre dans le fichier source.
    group_row = [ws.cell(row=2, column=c).value for c in range(1, 26)]
    label_row = [ws.cell(row=3, column=c).value for c in range(1, 26)]

    # Reconstruit le groupe courant pour chaque colonne (les groupes sont fusionnés,
    # donc seule la première colonne du groupe porte la valeur -> on la propage).
    current_group = None
    groups = []
    for v in group_row:
        if v:
            current_group = str(v).strip()
        groups.append(current_group)

    day_columns = []
    for idx in range(2, len(label_row)):  # colonnes C (index 2, 0-based) et suivantes
        label = label_row[idx]
        if not label:
            continue
        col = idx + 1  # 1-based pour openpyxl
        label_clean = str(label).replace("\n", " ").strip()
        day_columns.append({
            "col": col,
            "label": label_clean,
            "group": groups[idx] or "AUTRES",
            "color": color_for_label(label_clean),
        })

    days = []
    r = 4
    while True:
        jour = ws.cell(row=r, column=1).value
        date_val = ws.cell(row=r, column=2).value
        if jour is None and date_val is None:
            break
        entry = {"jour": jour, "date": None, "cells": {}}
        if isinstance(date_val, datetime):
            entry["date"] = date_val.strftime("%d/%m")
        for dc in day_columns:
            cell = ws.cell(row=r, column=dc["col"])
            if cell.value:
                val = str(cell.value).strip()
                entry["cells"][dc["col"]] = {"val": val, "vacant": False, "names": names_in(val)}
            elif is_red_fill(cell):
                entry["cells"][dc["col"]] = {"val": None, "vacant": True}
            else:
                entry["cells"][dc["col"]] = None
        days.append(entry)
        r += 1
        if r > 400:
            break

    # Tableau horaire récap : commence après la ligne "Nb semaines" (col C).
    # Les colonnes varient légèrement d'un onglet à l'autre (ex: DPI/USC selon
    # le mois), donc on lit les en-têtes réels ligne par ligne plutôt que de
    # supposer des indices de colonnes fixes.
    hours_start = None
    local_headers = []
    for rr in range(r, r + 40):
        if ws.cell(row=rr, column=3).value == "Nb semaines":
            header_row = rr + 1
            hours_start = rr + 2
            for col in range(4, 16):
                h = ws.cell(row=header_row, column=col).value
                if h:
                    local_headers.append((col, str(h).strip()))
            break
    hours = []
    if hours_start:
        rr = hours_start
        blanks_in_a_row = 0
        while blanks_in_a_row < 3:
            name = ws.cell(row=rr, column=3).value
            percent = ws.cell(row=rr, column=4).value
            if not name or not isinstance(name, str) or not isinstance(percent, (int, float)):
                blanks_in_a_row += 1
                rr += 1
                if rr - hours_start > 60:
                    break
                continue
            blanks_in_a_row = 0
            row = {"nom": str(name).strip()}
            _k = names_in(name)
            row["key"] = _k[0] if _k else None
            for col, h in local_headers:
                v = ws.cell(row=rr, column=col).value
                if isinstance(v, float):
                    v = round(v, 1)
                row[h] = v
            hours.append(row)
            rr += 1
            if rr - hours_start > 60:
                break
    headers_out = ["Nom"] + [h for _, h in local_headers]
    return days, day_columns, hours, headers_out


# ---------------------------------------------------------------------------
# Récapitulatifs par quadrimestre (heures + CA), comme dans le fichier Excel :
# ils figurent en bas du DERNIER mois de chaque quadrimestre (AVRIL, AOUT,
# DECEMBRE) ; le bilan annuel des CA est celui de DECEMBRE.
# Les colonnes ne sont pas au même endroit d'un onglet à l'autre : on repère les
# tableaux par leurs en-têtes ("A faire / Fait / Delta" et "A prendre / Pris / Rest").
# ---------------------------------------------------------------------------
QUADS = [
    ("AVRIL", "Q1", "Janvier – Avril"),
    ("AOUT", "Q2", "Mai – Août"),
    ("DECEMBRE", "Q3", "Septembre – Décembre"),
]


def _is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _txt(v):
    return str(v).strip() if isinstance(v, str) and v.strip() else None


def parse_recap_hours(ws):
    """Récap horaire du quadrimestre : colonnes 'A faire / Fait / Delta' (colonnes A à E)."""
    hdr = None
    for r in range(1, ws.max_row + 1):
        for c in range(1, 6):
            trio = (_txt(ws.cell(r, c).value), _txt(ws.cell(r, c + 1).value), _txt(ws.cell(r, c + 2).value))
            if trio == ("A faire", "Fait", "Delta"):
                hdr = (r, c)  # on garde le dernier trouvé (le récap est sous le tableau du mois)
    if not hdr:
        return []
    hr, c = hdr
    rows, blanks, rr = [], 0, hr + 1
    while blanks < 3 and rr <= ws.max_row and rr < hr + 80:
        texts = [_txt(ws.cell(rr, col).value) for col in range(1, c)]
        texts = [t for t in texts if t]
        a_faire = ws.cell(rr, c).value
        if texts and _is_num(a_faire):
            nom = texts[-1]
            note = " ".join(texts[:-1])
            fait = ws.cell(rr, c + 1).value
            delta = ws.cell(rr, c + 2).value
            delta24 = ws.cell(rr, c + 3).value
            k = names_in(nom)
            rows.append({
                "nom": nom,
                "key": k[0] if k else None,
                "note": note,
                "a_faire": round(a_faire, 1),
                "fait": round(fait, 1) if _is_num(fait) else None,
                "delta": round(delta, 1) if _is_num(delta) else None,
                "delta24": round(delta24, 2) if _is_num(delta24) else None,
            })
            blanks = 0
        else:
            blanks += 1
        rr += 1
    return rows


def parse_recap_ca(ws):
    """Tableau CA cumulé depuis janvier : colonnes '% / A prendre / Pris / Rest'."""
    hdr = None
    for r in range(1, ws.max_row + 1):
        for c in range(3, 20):
            trio = (_txt(ws.cell(r, c).value), _txt(ws.cell(r, c + 1).value), _txt(ws.cell(r, c + 2).value))
            if trio == ("A prendre", "Pris", "Rest"):
                hdr = (r, c)
    if not hdr:
        return []
    hr, c = hdr
    rows, blanks, rr = [], 0, hr + 1
    while blanks < 3 and rr <= ws.max_row and rr < hr + 80:
        nom = _txt(ws.cell(rr, c - 2).value)
        pris = ws.cell(rr, c + 1).value
        if nom and _is_num(pris):
            pct = ws.cell(rr, c - 1).value
            a_prendre = ws.cell(rr, c).value
            rest = ws.cell(rr, c + 2).value
            rows.append({
                "nom": nom,
                "pct": round(pct, 1) if _is_num(pct) else None,
                "a_prendre": round(a_prendre, 1) if _is_num(a_prendre) else None,
                "pris_cumul": round(pris, 1),
                "rest": round(rest, 1) if _is_num(rest) else None,
            })
            blanks = 0
        else:
            blanks += 1
        rr += 1
    return rows


def build_recaps(wb):
    """{ 'AVRIL': {...}, 'AOUT': {...}, 'DECEMBRE': {...} } : récap heures + CA par quadrimestre.
    Le 'Pris' de l'Excel est cumulé depuis janvier : CA du quad = cumul - cumul du quad précédent."""
    raw = {}
    for sheet, quad, periode in QUADS:
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        raw[sheet] = {
            "hours": parse_recap_hours(ws),
            "ca": {r["nom"].upper(): r for r in parse_recap_ca(ws)},
            "ca_order": [r["nom"] for r in parse_recap_ca(ws)],
        }

    out = {}
    for idx, (sheet, quad, periode) in enumerate(QUADS):
        if sheet not in raw:
            continue
        hours = raw[sheet]["hours"]
        ca_rows = []
        quads_done = [q for s, q, _ in QUADS[: idx + 1] if s in raw]
        for nom in raw[sheet]["ca_order"]:
            cur = raw[sheet]["ca"][nom.upper()]
            per_quad, prev = [], 0.0
            for s, q, _ in QUADS[: idx + 1]:
                if s not in raw:
                    continue
                cum = raw[s]["ca"].get(nom.upper(), {}).get("pris_cumul", 0.0)
                per_quad.append(round(cum - prev, 1))
                prev = cum
            k = names_in(nom)
            ca_rows.append({
                "nom": nom,
                "key": k[0] if k else None,
                "pct": cur["pct"],
                "a_prendre": cur["a_prendre"],
                "quads": per_quad,
                "total": cur["pris_cumul"],
                "rest": cur["rest"],
            })
        if hours or ca_rows:
            out[sheet] = {
                "quad": quad,
                "periode": periode,
                "annual": sheet == "DECEMBRE",
                "quads": quads_done,
                "hours": hours,
                "ca": ca_rows,
            }
    return out


def build_data(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    data = {}
    recaps = build_recaps(wb)
    for name in MOIS_ORDRE:
        if name not in wb.sheetnames:
            continue
        days, day_columns, hours, hours_headers = parse_month_sheet(wb[name])
        data[name] = {
            "days": days,
            "columns": day_columns,
            "hours": hours,
            "hours_headers": hours_headers,
            "recap": recaps.get(name),
        }
    return data


# Personnes ajoutées au sélecteur alors qu'elles ne figurent pas dans le tableau horaire
# (le menu ne propose que l'équipe médicale : tableau horaire + cette liste).
EXTRA_TEAM = {"adel": "ADEL"}


def build_people(data):
    """Liste des personnes proposées dans le sélecteur : clé -> {label, main, count}.
    Équipe médicale uniquement : noms du tableau horaire + EXTRA_TEAM."""
    people = {}
    for month in data.values():
        for h in month["hours"]:
            if h.get("key"):
                people.setdefault(h["key"], {"label": h["nom"].upper(), "main": True, "count": 0})
    for key, label in EXTRA_TEAM.items():
        people.setdefault(key, {"label": label, "main": True, "count": 0})
    for month in data.values():
        for day in month["days"]:
            for cell in day["cells"].values():
                if not cell:
                    continue
                for key in cell.get("names", []):
                    if key in people:
                        people[key]["count"] += 1
    return people


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Planning Urgences · SMUR · UHCD — CH Bagnols-sur-Cèze</title>
<style>
  :root {
    --navy: #1F3A5F;
    --amber: #C99A3B;
    --red: #8B1E2B;
    --blue: #3B6EA5;
    --bg: #F0F1F4;
    --card: #FFFFFF;
    --text: #1D2430;
    --muted: #6B7280;
    --border: #DADFE5;
  }
  * { box-sizing: border-box; }
  body { margin: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; background: var(--bg); color: var(--text); }
  header { background: var(--navy); color: #fff; padding: 14px 18px; box-shadow: 0 2px 6px rgba(0,0,0,0.15); }
  header h1 { margin: 0; font-size: 17px; font-weight: 600; }
  header .sub { font-size: 12px; opacity: 0.85; margin-top: 2px; }
  header .updated { font-size: 11px; opacity: 0.7; margin-top: 4px; }
  #stickyTop { position: sticky; top: 0; z-index: 20; background: #fff; box-shadow: 0 2px 6px rgba(0,0,0,0.12); }
  nav { display: flex; overflow-x: auto; background: #fff; border-bottom: 1px solid var(--border); -webkit-overflow-scrolling: touch; }
  nav button { flex: 0 0 auto; padding: 9px 14px; background: none; border: none; font-size: 12.5px; font-weight: 600; color: var(--muted); cursor: pointer; border-bottom: 3px solid transparent; white-space: nowrap; }
  nav button.active { color: var(--navy); border-bottom-color: var(--amber); }
  main { padding: 10px; max-width: 100%; overflow-x: auto; }
  .month { display: none; }
  .month.active { display: block; }
  .table-wrap { overflow-x: auto; background: var(--card); border-radius: 8px; border: 1px solid var(--border); margin-bottom: 18px; }
  table { border-collapse: collapse; width: 100%; min-width: 900px; font-size: 12px; }
  thead th { background: var(--navy); color: #fff; padding: 6px 8px; text-align: center; font-weight: 600; border: 1px solid rgba(255,255,255,0.15); position: sticky; top: 0; }
  thead tr.group-row th { background: #16293f; font-size: 10.5px; letter-spacing: 0.03em; text-transform: uppercase; opacity: 0.9; }
  tbody td { padding: 5px 7px; border: 1px solid var(--border); text-align: center; white-space: nowrap; }
  tbody td.jour-cell { text-align: left; font-weight: 600; background: #FAFAFB; white-space: nowrap; }
  tbody td.date-cell { color: var(--muted); background: #FAFAFB; }
  tbody tr:nth-child(even) td:not(.jour-cell):not(.date-cell) { background: #FBFBFC; }
  .we td.jour-cell, .we td.date-cell { background: #EFEFF3; }
  .poste { display: inline-block; padding: 2px 6px; border-radius: 5px; color: #fff; font-weight: 600; font-size: 11.5px; }
  .poste.vacant { background: #1D2430; color: #fff; font-weight: 700; border: 2px solid #8B1E2B; }
  tr.row-vacant td { background: #FBEAEA !important; }
  tr.row-vacant td.jour-cell, tr.row-vacant td.date-cell { background: #F6D9D9 !important; box-shadow: inset 4px 0 0 #8B1E2B; }
  h2.month-title { font-size: 14px; color: var(--navy); margin: 4px 4px 8px; }
  h3.section-title { font-size: 13px; color: var(--navy); margin: 4px 4px 8px; }
  .hours-table td.nom { text-align: left; font-weight: 600; }
  .hours-table td.delta-pos { color: var(--blue); font-weight: 600; }
  .hours-table td.delta-neg { color: var(--red); font-weight: 600; }
  .hours-table td.num { font-variant-numeric: tabular-nums; }
  .hours-table td.note { text-align: left; color: var(--muted); font-size: 11px; white-space: normal; }
  .hours-table td.total { font-weight: 700; background: #F3F5F9; }
  .recap-block { margin: 6px 0 22px; border-top: 3px solid var(--amber); padding-top: 10px; }
  .recap-block h2.recap-title { font-size: 14px; color: var(--navy); margin: 0 4px 10px; }
  .recap-block .hours-table { min-width: 0; }
  .recap-block .hours-table thead th { position: static; }

  /* --- Sélecteur de médecin --- */
  #filterBar { display: flex; align-items: center; flex-wrap: wrap; gap: 8px 10px; padding: 8px 12px; background: #F7F8FA; border-bottom: 1px solid var(--border); }
  #filterBar label { font-size: 12px; font-weight: 600; color: var(--navy); }
  #docSelect { font-size: 14px; padding: 6px 8px; border: 1px solid var(--border); border-radius: 6px; background: #fff; color: var(--text); min-width: 190px; max-width: 100%; }
  #clearSel { font-size: 12px; padding: 6px 10px; border: 1px solid var(--border); border-radius: 6px; background: #fff; color: var(--navy); cursor: pointer; }
  #selCount { font-size: 12.5px; color: var(--text); font-weight: 600; }
  body.has-sel tbody tr[data-names]:not(.sel-hit) { opacity: 0.38; }
  tr.sel-hit td { background: #FFF1BF !important; }
  tr.sel-hit td.jour-cell, tr.sel-hit td.date-cell { background: #F2B705 !important; color: #1D2430; font-weight: 700; box-shadow: inset 5px 0 0 var(--navy) !important; }
  .poste.sel-chip { box-shadow: 0 0 0 2px #fff, 0 0 0 4px var(--navy); font-weight: 700; }
  .hours-table tr.sel-hours-row td { background: #FFF1BF; font-weight: 700; }
  footer { text-align: center; font-size: 11px; color: var(--muted); padding: 18px; }
</style>
</head>
<body>
<header>
  <h1>Planning Urgences · SMUR · UHCD</h1>
  <div class="sub">CH Bagnols-sur-Cèze — lecture seule</div>
  <div class="updated">Dernière mise à jour : __GENERATED_AT__</div>
</header>
<div id="stickyTop">
  <nav id="monthNav"></nav>
  <div id="filterBar">
    <label for="docSelect">Médecin</label>
    <select id="docSelect"><option value="">— Tous —</option></select>
    <button id="clearSel" type="button" hidden>✕ Effacer</button>
    <span id="selCount"></span>
  </div>
</div>
<main id="monthContainer"></main>
<footer>Généré automatiquement depuis BSC_Planning_2026V2.xlsx — contactez Dr Torregrossa pour toute correction.</footer>

<script>
const DATA = __DATA_JSON__;
const HOURS_HEADERS = __HOURS_HEADERS_JSON__;
const PEOPLE = __PEOPLE_JSON__;
let selected = "";
const MONTH_FR = {FEVRIER: "Février", AOUT: "Août", DECEMBRE: "Décembre"};

const nav = document.getElementById("monthNav");
const container = document.getElementById("monthContainer");
const months = Object.keys(DATA);

function groupSpans(dayColumns) {
  const spans = [];
  let last = null;
  dayColumns.forEach(c => {
    if (last && last.group === c.group) {
      last.span += 1;
    } else {
      last = {group: c.group, span: 1};
      spans.push(last);
    }
  });
  return spans;
}

months.forEach((month, idx) => {
  const btn = document.createElement("button");
  btn.textContent = monthLabel(month);
  btn.dataset.month = month;
  if (idx === 0) btn.classList.add("active");
  btn.addEventListener("click", () => showMonth(month));
  nav.appendChild(btn);

  const section = document.createElement("div");
  section.className = "month" + (idx === 0 ? " active" : "");
  section.id = "month-" + month;

  // --- Grille journalière ---
  const dayColumns = DATA[month].columns;
  const wrap1 = document.createElement("div");
  wrap1.className = "table-wrap";
  const table = document.createElement("table");

  const thead = document.createElement("thead");
  const groupRow = document.createElement("tr");
  groupRow.className = "group-row";
  groupRow.innerHTML = "<th colspan='2'></th>";
  groupSpans(dayColumns).forEach(g => {
    groupRow.innerHTML += `<th colspan="${g.span}">${g.group}</th>`;
  });
  thead.appendChild(groupRow);

  const headRow = document.createElement("tr");
  headRow.innerHTML = "<th>Jour</th><th>Date</th>" + dayColumns.map(c => `<th>${c.label}</th>`).join("");
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  DATA[month].days.forEach(day => {
    const tr = document.createElement("tr");
    const hasVacant = Object.values(day.cells).some(c => c && c.vacant);
    const dayNames = [];
    Object.values(day.cells).forEach(c => { if (c && c.names) c.names.forEach(n => { if (!dayNames.includes(n)) dayNames.push(n); }); });
    tr.dataset.names = dayNames.join("|");
    if (hasVacant) {
      tr.className = "row-vacant";
    } else if (day.jour === "Samedi" || day.jour === "Dimanche") {
      tr.className = "we";
    }
    let rowHtml = `<td class="jour-cell">${day.jour || ""}</td><td class="date-cell">${day.date || ""}</td>`;
    dayColumns.forEach(c => {
      const cell = day.cells[c.col];
      if (cell && cell.val) {
        rowHtml += `<td><span class="poste" data-names="${(cell.names || []).join("|")}" style="background:${c.color}">${cell.val}</span></td>`;
      } else if (cell && cell.vacant) {
        rowHtml += `<td><span class="poste vacant">⚠ Vacant</span></td>`;
      } else {
        rowHtml += "<td></td>";
      }
    });
    tr.innerHTML = rowHtml;
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  wrap1.appendChild(table);
  section.appendChild(wrap1);

  // --- Tableau horaire récap ---
  if (DATA[month].hours && DATA[month].hours.length) {
    const h3 = document.createElement("h3");
    h3.className = "section-title";
    h3.textContent = "Tableau horaire";
    section.appendChild(h3);

    const monthHeaders = DATA[month].hours_headers;
    const wrap2 = document.createElement("div");
    wrap2.className = "table-wrap";
    const htable = document.createElement("table");
    htable.className = "hours-table";
    const hthead = document.createElement("thead");
    hthead.innerHTML = "<tr>" + monthHeaders.map(h => `<th>${h}</th>`).join("") + "</tr>";
    htable.appendChild(hthead);
    const htbody = document.createElement("tbody");
    DATA[month].hours.forEach(row => {
      const tr = document.createElement("tr");
      if (row.key) tr.dataset.key = row.key;
      let rowHtml = `<td class="nom">${row.nom}</td>`;
      monthHeaders.slice(1).forEach(h => {
        let v = row[h];
        v = (v === null || v === undefined) ? "" : v;
        let cls = "";
        if (h === "Delta" && typeof v === "number") cls = v < 0 ? "delta-neg" : "delta-pos";
        rowHtml += `<td class="${cls}">${v}</td>`;
      });
      tr.innerHTML = rowHtml;
      htbody.appendChild(tr);
    });
    htable.appendChild(htbody);
    wrap2.appendChild(htable);
    section.appendChild(wrap2);
  }

  if (DATA[month].recap) renderRecap(section, DATA[month].recap);

  container.appendChild(section);
});

// --- Récapitulatif de quadrimestre (heures + CA), affiché sous le dernier mois du quad ---
function fmtNum(v) { return (v === null || v === undefined) ? "" : v; }

function buildTable(headers, rowsHtml) {
  const wrap = document.createElement("div");
  wrap.className = "table-wrap";
  const t = document.createElement("table");
  t.className = "hours-table";
  t.innerHTML = "<thead><tr>" + headers.map(h => `<th>${h}</th>`).join("") + "</tr></thead>" +
                "<tbody>" + rowsHtml.join("") + "</tbody>";
  wrap.appendChild(t);
  return wrap;
}

function signClass(v) { return (typeof v === "number") ? (v < 0 ? "delta-neg" : "delta-pos") : ""; }

function renderRecap(section, rc) {
  const box = document.createElement("div");
  box.className = "recap-block";

  const h2 = document.createElement("h2");
  h2.className = "recap-title";
  h2.textContent = "Récapitulatif " + rc.quad + " (" + rc.periode + ")";
  box.appendChild(h2);

  // Heures du quadrimestre
  if (rc.hours && rc.hours.length) {
    const h3 = document.createElement("h3");
    h3.className = "section-title";
    h3.textContent = "Heures " + rc.quad;
    box.appendChild(h3);
    const hasNote = rc.hours.some(r => r.note);
    const heads = ["Nom", "A faire", "Fait", "Delta", "Delta ÷ 24"].concat(hasNote ? ["Note"] : []);
    const rows = rc.hours.map(r => {
      const k = r.key ? ` data-key="${r.key}"` : "";
      return `<tr${k}><td class="nom">${r.nom}</td>` +
        `<td class="num">${fmtNum(r.a_faire)}</td><td class="num">${fmtNum(r.fait)}</td>` +
        `<td class="num ${signClass(r.delta)}">${fmtNum(r.delta)}</td>` +
        `<td class="num ${signClass(r.delta24)}">${fmtNum(r.delta24)}</td>` +
        (hasNote ? `<td class="note">${r.note || ""}</td>` : "") + `</tr>`;
    });
    box.appendChild(buildTable(heads, rows));
  }

  // CA : par quadrimestre, puis bilan annuel sur décembre
  if (rc.ca && rc.ca.length) {
    const h3 = document.createElement("h3");
    h3.className = "section-title";
    h3.textContent = rc.annual ? "Congés annuels (CA) — par quadrimestre et bilan annuel"
                               : "Congés annuels (CA) — jours pris";
    box.appendChild(h3);
    const heads = ["Nom", "%", "A prendre (année)"].concat(rc.quads.map(q => "Pris " + q))
      .concat([rc.annual ? "Total année" : "Cumul", "Reste"]);
    const rows = rc.ca.map(r => {
      const k = r.key ? ` data-key="${r.key}"` : "";
      return `<tr${k}><td class="nom">${r.nom}</td>` +
        `<td class="num">${r.pct === null ? "" : r.pct + " %"}</td>` +
        `<td class="num">${fmtNum(r.a_prendre)}</td>` +
        r.quads.map(v => `<td class="num">${fmtNum(v)}</td>`).join("") +
        `<td class="num total">${fmtNum(r.total)}</td>` +
        `<td class="num ${r.rest !== null && r.rest < 0 ? "delta-neg" : ""}">${fmtNum(r.rest)}</td></tr>`;
    });
    box.appendChild(buildTable(heads, rows));
  }

  section.appendChild(box);
}

let currentMonth = months[0];

function showMonth(month) {
  currentMonth = month;
  document.querySelectorAll("#monthNav button").forEach(b => b.classList.toggle("active", b.dataset.month === month));
  document.querySelectorAll(".month").forEach(s => s.classList.toggle("active", s.id === "month-" + month));
  updateCount();
}

function monthLabel(m) { return MONTH_FR[m] || (m.charAt(0) + m.slice(1).toLowerCase()); }

function updateCount() {
  const el = document.getElementById("selCount");
  if (!selected) { el.textContent = ""; return; }
  const n = DATA[currentMonth].days.filter(d =>
    Object.values(d.cells).some(c => c && c.names && c.names.includes(selected))).length;
  el.textContent = n === 0
    ? "Aucun poste en " + monthLabel(currentMonth)
    : n + (n > 1 ? " jours" : " jour") + " en " + monthLabel(currentMonth);
}

function applySelection() {
  document.body.classList.toggle("has-sel", !!selected);
  document.querySelectorAll("tbody tr[data-names]").forEach(tr => {
    const hit = !!selected && tr.dataset.names.split("|").includes(selected);
    tr.classList.toggle("sel-hit", hit);
    tr.querySelectorAll(".poste[data-names]").forEach(ch =>
      ch.classList.toggle("sel-chip", !!selected && ch.dataset.names.split("|").includes(selected)));
  });
  document.querySelectorAll(".hours-table tbody tr").forEach(tr =>
    tr.classList.toggle("sel-hours-row", !!selected && tr.dataset.key === selected));
  document.getElementById("clearSel").hidden = !selected;
  document.getElementById("docSelect").value = selected;
  updateCount();
  try { history.replaceState(null, "", selected ? "#med=" + encodeURIComponent(selected) : location.pathname + location.search); } catch (e) {}
}

(function initSelector() {
  const sel = document.getElementById("docSelect");
  const byLabel = (a, b) => PEOPLE[a].label.localeCompare(PEOPLE[b].label, "fr");
  const groups = [
    ["Équipe médicale", Object.keys(PEOPLE).filter(k => PEOPLE[k].main).sort(byLabel)],
    ["Internes / autres", Object.keys(PEOPLE).filter(k => !PEOPLE[k].main).sort(byLabel)],
  ];
  groups.forEach(([title, keys]) => {
    if (!keys.length) return;
    const og = document.createElement("optgroup");
    og.label = title;
    keys.forEach(k => { const o = document.createElement("option"); o.value = k; o.textContent = PEOPLE[k].label; og.appendChild(o); });
    sel.appendChild(og);
  });
  sel.addEventListener("change", () => { selected = sel.value; applySelection(); });
  document.getElementById("clearSel").addEventListener("click", () => { selected = ""; applySelection(); });
  const m = location.hash.match(/^#med=(.+)$/);
  if (m && PEOPLE[decodeURIComponent(m[1])]) selected = decodeURIComponent(m[1]);
  applySelection();
})();
</script>
</body>
</html>
"""


def main():
    if len(sys.argv) < 2:
        print("Usage: python3 generate_planning_v2.py <fichier.xlsx> [dossier_sortie]")
        sys.exit(1)
    xlsx_path = sys.argv[1]
    out_dir = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(".")
    out_dir.mkdir(parents=True, exist_ok=True)

    data = build_data(xlsx_path)
    html = HTML_TEMPLATE
    html = html.replace("__DATA_JSON__", json.dumps(data, ensure_ascii=False))
    html = html.replace("__HOURS_HEADERS_JSON__", json.dumps([], ensure_ascii=False))
    html = html.replace("__PEOPLE_JSON__", json.dumps(build_people(data), ensure_ascii=False))
    html = html.replace("__GENERATED_AT__", datetime.now().strftime("%d/%m/%Y %H:%M"))

    out_file = out_dir / "index.html"
    out_file.write_text(html, encoding="utf-8")
    print(f"OK -> {out_file}")


if __name__ == "__main__":
    main()
