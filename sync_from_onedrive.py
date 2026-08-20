#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Télécharge le fichier planning depuis le lien de partage OneDrive et l'enregistre
dans data/BSC_Planning_2026V2.xlsx. Conçu pour tourner dans un GitHub Action
programmé (cron), sans authentification Microsoft : on exploite simplement le
lien de partage public "Consultation" tel qu'obtenu depuis OneDrive.

Usage :
    python3 sync_from_onedrive.py "<lien_1drv.ms>" data/BSC_Planning_2026V2.xlsx
"""

import sys
import requests

USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/120.0 Safari/537.36")


def download_onedrive_file(share_url: str) -> bytes:
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT})

    # Premier appel : résout le lien court 1drv.ms et établit les cookies de session.
    r1 = session.get(share_url, allow_redirects=True, timeout=30)
    r1.raise_for_status()

    if not r1.history:
        raise RuntimeError("Le lien OneDrive ne redirige pas comme attendu (a-t-il changé ?).")

    # On récupère l'URL du tout premier saut de redirection (avant la page de
    # visualisation Doc.aspx), et on lui ajoute le paramètre de téléchargement direct.
    first_hop_url = r1.history[0].headers["Location"]
    sep = "&" if "?" in first_hop_url else "?"
    download_url = f"{first_hop_url}{sep}download=1"

    r2 = session.get(download_url, timeout=60)
    r2.raise_for_status()

    content_type = r2.headers.get("Content-Type", "")
    if "spreadsheet" not in content_type and "octet-stream" not in content_type:
        raise RuntimeError(
            f"Réponse inattendue (Content-Type: {content_type}). "
            "Le lien OneDrive a peut-être expiré ou ses permissions ont changé."
        )
    return r2.content


def main():
    if len(sys.argv) < 3:
        print("Usage: python3 sync_from_onedrive.py <lien_partage> <chemin_sortie.xlsx>")
        sys.exit(1)
    share_url = sys.argv[1]
    out_path = sys.argv[2]

    content = download_onedrive_file(share_url)
    with open(out_path, "wb") as f:
        f.write(content)
    print(f"OK -> {out_path} ({len(content)} octets)")


if __name__ == "__main__":
    main()
