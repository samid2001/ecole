"""Classe la messagerie des entreprises de l'onglet « Écartées (site trouvé) ».

Usage :
    python3 classer_messagerie.py <original.xlsx> <sortie.xlsx> [--sans-dns]

Le fichier d'origine n'est jamais modifié : il est lu, puis une copie est écrite.
Ajoute trois colonnes : Domaine (H), Messagerie (I) et Profil (J, formule).

Messagerie : « Microsoft 365 » si un MX contient mail.protection.outlook.com,
« Google » si un MX contient google.com, « Autre » sinon.
Profil : 1 = Microsoft 365, 2 = autre ou Google, vide si pas de domaine ou non vérifié.
"""
import sys
from copy import copy
from urllib.parse import urlparse

import dns.resolver
import openpyxl
from openpyxl.utils import get_column_letter

SHEET = "Écartées (site trouvé)"
NO_DOMAIN = "Pas de domaine"
NOT_CHECKED = "Non vérifié"


def domain_of(url):
    if not url or not str(url).startswith("http"):
        return None
    host = (urlparse(str(url)).hostname or "").lower()
    return host[4:] if host.startswith("www.") else host


def classify(domain, resolver):
    try:
        exchanges = [str(a.exchange).lower() for a in resolver.resolve(domain, "MX")]
    except Exception:
        # Pas de MX, NXDOMAIN, timeout : rangé en « Autre », comme le prévoit la règle.
        return "Autre"
    joined = " ".join(exchanges)
    if "mail.protection.outlook.com" in joined:
        return "Microsoft 365"
    if "google.com" in joined:
        return "Google"
    return "Autre"


def main():
    src, dst = sys.argv[1], sys.argv[2]
    sans_dns = "--sans-dns" in sys.argv

    wb = openpyxl.load_workbook(src)  # lecture seule de l'original
    ws = wb[SHEET]
    resolver = dns.resolver.Resolver()
    resolver.lifetime = 8

    header_src = ws.cell(1, 7)
    new_headers = ["Domaine", "Messagerie", "Profil"]
    for i, title in enumerate(new_headers):
        cell = ws.cell(1, 8 + i, title)
        cell.font = copy(header_src.font)
        cell.fill = copy(header_src.fill)
        cell.alignment = copy(header_src.alignment)
        cell.border = copy(header_src.border)
    for col, width in ((8, 34), (9, 16), (10, 10)):
        ws.column_dimensions[get_column_letter(col)].width = width

    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 1).value is None:
            continue
        domain = domain_of(ws.cell(r, 6).value)
        if domain is None:
            messagerie = NO_DOMAIN
        elif sans_dns:
            messagerie = NOT_CHECKED
        else:
            messagerie = classify(domain, resolver)

        values = [domain or "—", messagerie,
                  f'=IF(OR(I{r}="{NO_DOMAIN}",I{r}="{NOT_CHECKED}"),"",IF(I{r}="Microsoft 365",1,2))']
        for i, value in enumerate(values):
            cell = ws.cell(r, 8 + i, value)
            base = ws.cell(r, 7)
            cell.font = copy(base.font)
            cell.border = copy(base.border)
            cell.alignment = copy(base.alignment)
            cell.number_format = base.number_format
            if i == 1:
                print(r, ws.cell(r, 1).value, "|", domain, "|", messagerie)

    last = ws.max_row
    ws.auto_filter.ref = f"A1:J{last}"
    notes = [
        "Notes",
        "Domaine : hôte du site trouvé, sans « www. ». Les sites hébergés sur un sous-domaine (ex. wixsite.com) donnent le MX de l'hébergeur.",
        "Messagerie : « Autre » est aussi attribué quand le domaine n'a pas de MX ou ne répond pas : à vérifier à la main.",
        "Profil : 1 = Microsoft 365 ; 2 = autre ou Google (hors Microsoft 365). Vide = pas de domaine ou non vérifié.",
    ]
    for i, text in enumerate(notes):
        ws.cell(last + 2 + i, 1, text).font = copy(header_src.font) if i == 0 else copy(ws.cell(2, 1).font)

    wb.save(dst)
    print("Écrit :", dst)


if __name__ == "__main__":
    main()
