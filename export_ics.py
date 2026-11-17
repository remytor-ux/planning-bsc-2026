"""Export des plannings individuels au format iCalendar (.ics).

Un fichier par médecin et par année (ics/<nom>.ics) : les médecins peuvent s'y abonner depuis
l'agenda de leur téléphone (Apple Calendrier, Google Agenda, Outlook). Le fichier est régénéré à
chaque mise en ligne du planning ; l'agenda abonné se met à jour tout seul.
"""
import re
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

try:
    from zoneinfo import ZoneInfo
    PARIS = ZoneInfo("Europe/Paris")
except Exception:  # pragma: no cover
    PARIS = None

# code compact -> (titre, heure début, heure fin ou None). None = événement « journée entière ».
SHIFTS = {
    "a": ("Ligne 1 · Jour", (9, 0), (19, 0)),
    "b": ("Ligne 1 · Nuit", (19, 0), (9, 0)),   # fin le lendemain
    "c": ("Garde 12 h (9h–21h)", (9, 0), (21, 0)),
    "d": ("Ligne 2 · Jour", (9, 0), (19, 0)),
    "e": ("Ligne 2 · Nuit", (19, 0), (9, 0)),
    "f": ("FFI / Interne · Jour", None, None),
    "g": ("FFI / Interne · Nuit", None, None),
    "h": ("FFI / Interne 24h", None, None),
    "i": ("UHCD", None, None),
    "j": ("USC", None, None),
    "k": ("RUF", None, None),
    "l": ("DPI / CDP", None, None),
    "m": ("Congés (CA)", None, None),
    "n": ("Indisponible", None, None),
    "o": ("Formation", None, None),
    "p": ("Arrêt de travail", None, None),
}
MERGE = set("mnop")  # absences : les jours consécutifs forment un seul événement

# Cases du planning qui contiennent un prénom : rattachées au médecin (cf. application Outils Médecin)
ALIAS = {"sylvie": "cauvy", "chloe": "orioli", "bouba": "kantche", "maria": "rachidi", "lionel": "boulat",
         "stephane": "tkatchouk", "richard": "morel", "adrien": "fontaine", "alice": "brugiroux"}


def slug(name):
    return re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")


def _esc(s):
    return s.replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")


def _fold(line):
    out, b = [], line.encode("utf-8")
    while len(b) > 74:
        cut = 74
        while (b[cut] & 0xC0) == 0x80:
            cut -= 1
        out.append(b[:cut].decode("utf-8"))
        b = b[cut:]
        b = b" " + b
    out.append(b.decode("utf-8"))
    return "\r\n".join(out)


def _utc(d, hm, plus_days=0):
    dt = datetime(d.year, d.month, d.day, hm[0], hm[1]) + timedelta(days=plus_days)
    dt = dt.replace(tzinfo=PARIS) if PARIS else dt
    return dt.astimezone(timezone.utc).strftime("%Y%m%dT%H%M%SZ")


def _parse(compact_data):
    """{nom: "AAMMJJ<codes> ..."} -> {nom canonique: {date: set(codes)}}"""
    people = {}
    for name, s in compact_data.items():
        canon = ALIAS.get(name, name)
        for tok in s.split():
            d = date(2000 + int(tok[0:2]), int(tok[2:4]), int(tok[4:6]))
            people.setdefault(canon, {}).setdefault(d, set()).update(tok[6:])
    return people


def _events(days):
    ev = []  # (code, début, fin incluse)
    for code in sorted({c for v in days.values() for c in v}):
        ds = sorted(d for d, v in days.items() if code in v)
        if code in MERGE:
            start = prev = None
            for d in ds:
                if prev is not None and (d - prev).days == 1:
                    prev = d
                    continue
                if start:
                    ev.append((code, start, prev))
                start = prev = d
            if start:
                ev.append((code, start, prev))
        else:
            ev.extend((code, d, d) for d in ds)
    return sorted(ev, key=lambda e: (e[1], e[0]))


def build_ics(name, days, year, stamp):
    label = name.upper()
    lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//CH Bagnols//Planning BSC//FR", "CALSCALE:GREGORIAN",
             "METHOD:PUBLISH", "X-WR-CALNAME:" + _esc("Gardes %s %s" % (label, year)),
             "X-WR-TIMEZONE:Europe/Paris", "REFRESH-INTERVAL;VALUE=DURATION:PT1H", "X-PUBLISHED-TTL:PT1H"]
    for code, d0, d1 in _events(days):
        title, st, en = SHIFTS[code]
        lines += ["BEGIN:VEVENT", "UID:%s-%s-%s@planning-bsc" % (slug(name), d0.strftime("%Y%m%d"), code),
                  "DTSTAMP:" + stamp, "SUMMARY:" + _esc(title)]
        if st:
            over = en <= st
            lines += ["DTSTART:" + _utc(d0, st), "DTEND:" + _utc(d0, en, 1 if over else 0)]
        else:
            lines += ["DTSTART;VALUE=DATE:" + d0.strftime("%Y%m%d"),
                      "DTEND;VALUE=DATE:" + (d1 + timedelta(days=1)).strftime("%Y%m%d"), "TRANSP:TRANSPARENT"]
        lines.append("END:VEVENT")
    lines.append("END:VCALENDAR")
    return "\r\n".join(_fold(l) for l in lines) + "\r\n"


def export_ics(compact_data, year, out_dir):
    """Écrit out_dir/ics/<nom>.ics pour chaque personne ; retourne la liste des fichiers."""
    folder = Path(out_dir) / "ics"
    folder.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    written = []
    for name, days in sorted(_parse(compact_data).items()):
        (folder / (slug(name) + ".ics")).write_text(build_ics(name, days, year, stamp), encoding="utf-8", newline="")
        written.append(slug(name))
    return written
