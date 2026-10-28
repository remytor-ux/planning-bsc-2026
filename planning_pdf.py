#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Génère un PDF du planning de chaque médecin pour chaque quadrimestre
(Q1 = janvier-avril, Q2 = mai-août, Q3 = septembre-décembre).

Utilisé par generate_planning.py : les PDF sont recréés à chaque construction du
site, donc toujours à jour avec le dernier fichier Excel.
"""

import unicodedata
from datetime import datetime
from pathlib import Path
from xml.sax.saxutils import escape

MOIS_ORDRE = [
    "JANVIER", "FEVRIER", "MARS", "AVRIL", "MAI", "JUIN",
    "JUILLET", "AOUT", "SEPTEMBRE", "OCTOBRE", "NOVEMBRE", "DECEMBRE",
]
MOIS_FR = {
    "JANVIER": "Janvier", "FEVRIER": "Février", "MARS": "Mars", "AVRIL": "Avril",
    "MAI": "Mai", "JUIN": "Juin", "JUILLET": "Juillet", "AOUT": "Août",
    "SEPTEMBRE": "Septembre", "OCTOBRE": "Octobre", "NOVEMBRE": "Novembre",
    "DECEMBRE": "Décembre",
}
QUAD_MONTHS = {"Q1": MOIS_ORDRE[0:4], "Q2": MOIS_ORDRE[4:8], "Q3": MOIS_ORDRE[8:12]}
QUAD_PERIODE = {"Q1": "Janvier – Avril", "Q2": "Mai – Août", "Q3": "Septembre – Décembre"}
QUAD_COURT = {"Q1": "Janv–Avr", "Q2": "Mai–Août", "Q3": "Sept–Déc"}
QUAD_RECAP_SHEET = {"Q1": "AVRIL", "Q2": "AOUT", "Q3": "DECEMBRE"}

NAVY = "#1F3A5F"
AMBER = "#C99A3B"
GREY = "#6B7280"


def _slug(text):
    t = unicodedata.normalize("NFD", str(text))
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return "".join(c if c.isalnum() else "_" for c in t).strip("_").upper()


def _poste_label(col):
    group = (col.get("group") or "").strip()
    label = col["label"]
    if group and group.upper() != "AUTRES":
        return f"{group} · {label}"
    return label


def _entries(data, key, quad):
    """Liste par mois des cases où la personne est citée : (jour, date, poste, texte de la case)."""
    out = []
    for month in QUAD_MONTHS[quad]:
        if month not in data:
            continue
        rows = []
        columns = data[month]["columns"]
        for day in data[month]["days"]:
            for col in columns:
                cell = day["cells"].get(col["col"])
                if cell and cell.get("val") and key in cell.get("names", []):
                    rows.append((day.get("jour") or "", day.get("date") or "", _poste_label(col), cell["val"]))
        out.append((month, rows))
    return out


def _summary(data, key, quad):
    """Heures et CA du quadrimestre, lus dans le récap de l'Excel (dernier mois du quad)."""
    sheet = QUAD_RECAP_SHEET[quad]
    rc = (data.get(sheet) or {}).get("recap")
    if not rc:
        return None, None
    hours = next((r for r in rc["hours"] if r.get("key") == key), None)
    ca = next((r for r in rc["ca"] if r.get("key") == key), None)
    ca_quad = None
    if ca and rc["quad"] == quad and ca.get("quads"):
        ca_quad = ca["quads"][-1]
    return hours, ca_quad


def build_pdfs(data, people, out_dir, year, src_name, logo_path=None):
    """Crée out_dir/pdf/<Qn>/Planning_<NOM>_<Qn>_<année>.pdf.
    Retourne {clé_personne: {Qn: {"url", "court"}}} pour les liens de la page."""
    try:
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_LEFT
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.lib.units import mm
        from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    except ImportError:
        print("reportlab absent : les PDF ne sont pas générés.")
        return {}

    navy, amber, grey = colors.HexColor(NAVY), colors.HexColor(AMBER), colors.HexColor(GREY)
    st_title = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=16, leading=19, textColor=navy)
    st_sub = ParagraphStyle("s", fontName="Helvetica", fontSize=10.5, leading=14, textColor=colors.HexColor("#1D2430"))
    st_month = ParagraphStyle("m", fontName="Helvetica-Bold", fontSize=11.5, leading=14, textColor=navy,
                              spaceBefore=10, spaceAfter=4)
    st_cell = ParagraphStyle("c", fontName="Helvetica", fontSize=9, leading=11, alignment=TA_LEFT)
    st_cell_b = ParagraphStyle("cb", parent=st_cell, fontName="Helvetica-Bold")
    st_head = ParagraphStyle("h", parent=st_cell, fontName="Helvetica-Bold", textColor=colors.white)
    st_note = ParagraphStyle("n", fontName="Helvetica-Oblique", fontSize=9, leading=11, textColor=grey)

    generated = datetime.now().strftime("%d/%m/%Y %H:%M")
    index = {}

    def footer(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(grey)
        canvas.drawString(15 * mm, 9 * mm,
                          f"Généré le {generated} depuis {src_name} — planning susceptible d'évoluer : "
                          f"consulter le site pour la dernière version.")
        canvas.drawRightString(A4[0] - 15 * mm, 9 * mm, f"Page {doc.page}")
        canvas.restoreState()

    for key in sorted(k for k, v in people.items() if v.get("main")):
        label = people[key]["label"]
        for quad in ("Q1", "Q2", "Q3"):
            per_month = _entries(data, key, quad)
            if not any(rows for _, rows in per_month):
                continue  # rien à imprimer pour ce quadrimestre

            rel = f"pdf/{quad}/Planning_{_slug(label)}_{quad}_{year}.pdf"
            path = Path(out_dir) / rel
            path.parent.mkdir(parents=True, exist_ok=True)

            doc = SimpleDocTemplate(str(path), pagesize=A4, leftMargin=15 * mm, rightMargin=15 * mm,
                                    topMargin=14 * mm, bottomMargin=16 * mm,
                                    title=f"Planning {label} — {quad} {year}",
                                    author="CH Bagnols-sur-Cèze — Urgences · SMUR · UHCD")
            story = []

            # En-tête : logo + titre
            title_block = [Paragraph(f"Planning — {escape(label.title())}", st_title),
                           Paragraph(f"{quad} {year} ({QUAD_PERIODE[quad]})", st_sub)]
            if logo_path and Path(logo_path).exists():
                w = 62 * mm
                img = Image(str(logo_path), width=w, height=w * 154 / 640)
                head = Table([[img, title_block]], colWidths=[68 * mm, None])
            else:
                head = Table([[title_block]], colWidths=[None])
            head.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                                      ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                      ("LINEBELOW", (0, 0), (-1, 0), 2, amber),
                                      ("BOTTOMPADDING", (0, 0), (-1, 0), 8)]))
            story += [head, Spacer(1, 6)]

            # Récapitulatif du quadrimestre (valeurs du fichier Excel)
            hours, ca_quad = _summary(data, key, quad)
            parts = []
            if hours:
                def fmt(v):
                    return "" if v is None else f"{v:g}".replace(".", ",")
                parts.append(f"Heures du quadrimestre : à faire {fmt(hours['a_faire'])} h · "
                             f"fait {fmt(hours['fait'])} h · delta {fmt(hours['delta'])} h")
            if ca_quad is not None:
                parts.append(f"CA / AT pris sur le quadrimestre : {f'{ca_quad:g}'.replace('.', ',')} j")
            if parts:
                box = Table([[Paragraph("<br/>".join(escape(p) for p in parts), st_cell)]], colWidths=[None])
                box.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F3F5F9")),
                                         ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#DADFE5")),
                                         ("LEFTPADDING", (0, 0), (-1, -1), 8), ("TOPPADDING", (0, 0), (-1, -1), 6),
                                         ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
                story += [box]

            # Un tableau par mois
            for month, rows in per_month:
                nb_days = len({r[1] for r in rows})
                story.append(Paragraph(
                    f"{MOIS_FR[month]} {year} — {nb_days} jour{'s' if nb_days > 1 else ''}", st_month))
                if not rows:
                    story.append(Paragraph("Aucun poste ce mois-ci.", st_note))
                    continue
                table_rows = [[Paragraph(h, st_head) for h in ("Jour", "Date", "Poste", "Case du planning")]]
                style = [("BACKGROUND", (0, 0), (-1, 0), navy),
                         ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                         ("LINEBELOW", (0, 0), (-1, -1), 0.4, colors.HexColor("#DADFE5")),
                         ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]
                for i, (jour, date, poste, val) in enumerate(rows, start=1):
                    date_full = f"{date}/{year}" if date else ""
                    table_rows.append([Paragraph(escape(jour), st_cell_b), Paragraph(date_full, st_cell),
                                       Paragraph(escape(poste), st_cell), Paragraph(escape(val), st_cell)])
                    if jour in ("Samedi", "Dimanche"):
                        style.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#EFEFF3")))
                t = Table(table_rows, colWidths=[24 * mm, 26 * mm, 62 * mm, None], repeatRows=1)
                t.setStyle(TableStyle(style))
                story.append(t)

            doc.build(story, onFirstPage=footer, onLaterPages=footer)
            index.setdefault(key, {})[quad] = {"url": rel, "court": QUAD_COURT[quad]}

    print(f"PDF : {sum(len(v) for v in index.values())} fichiers générés")
    return index
