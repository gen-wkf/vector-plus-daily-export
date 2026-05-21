#!/usr/bin/env python3
"""
Script pour convertir un fichier JSON contenant un array en fichiers CSV et XLSX.
"""

import json
import csv
from datetime import datetime
from pathlib import Path
from typing import Any, List, Dict, Optional
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO


def get_nested_value(obj: Dict, path: str) -> Any:
    """
    Récupère une valeur imbriquée dans un dictionnaire en utilisant une notation pointée.
    Ex: "application_data/livraison/date_premiere_livraison"
    """
    keys = path.split('/')
    current = obj
    for key in keys:
        if isinstance(current, dict):
            current = current.get(key)
        else:
            return None
    return current


def extract_date_part(datetime_str: Optional[str]) -> Optional[str]:
    """
    Extrait la partie date d'une chaîne datetime au format YYYY-MM-DD.
    """
    if not datetime_str:
        return None
    try:
        # Essayer différents formats
        for fmt in ['%Y-%m-%dT%H:%M:%S', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d']:
            try:
                dt = datetime.strptime(datetime_str, fmt)
                return dt.strftime('%Y-%m-%d')
            except ValueError:
                continue
        return datetime_str[:10] if len(datetime_str) >= 10 else None
    except Exception:
        return None


def extract_time_part(datetime_str: Optional[str]) -> Optional[str]:
    """
    Extrait la partie heure d'une chaîne datetime au format HH:mm.
    """
    if not datetime_str:
        return None
    try:
        # Essayer différents formats
        for fmt in ['%Y-%m-%dT%H:%M:%S', '%Y-%m-%d %H:%M:%S']:
            try:
                dt = datetime.strptime(datetime_str, fmt)
                return dt.strftime('%H:%M')
            except ValueError:
                continue
        # Essayer extraction directe
        if 'T' in datetime_str:
            time_part = datetime_str.split('T')[1]
            return time_part[:5]
        elif ' ' in datetime_str:
            time_part = datetime_str.split(' ')[1]
            return time_part[:5]
        return None
    except Exception:
        return None


def find_participant_by_role(participants: List[Dict], role_code: str) -> Optional[str]:
    """
    Trouve un participant par code de rôle et retourne sa raison sociale.
    """
    if not isinstance(participants, list):
        return None

    for participant in participants:
        if isinstance(participant, dict):
            role = participant.get('role', {})
            if isinstance(role, dict):
                if role.get('code') == role_code or role.get('libelle') == role_code:
                    societe = participant.get('societe', {})
                    if isinstance(societe, dict):
                        return societe.get('raison_sociale')
    return None


def find_client_final(participants: List[Dict]) -> Optional[str]:
    """
    Trouve le client final (MOA > CATEGORIE > None).
    """
    if not isinstance(participants, list):
        return None

    # Chercher MOA d'abord
    for participant in participants:
        if isinstance(participant, dict):
            role = participant.get('role', {})
            if isinstance(role, dict) and role.get('code') == 'moa':
                societe = participant.get('societe', {})
                if isinstance(societe, dict):
                    return societe.get('raison_sociale')

    # Puis CATEGORIE
    for participant in participants:
        if isinstance(participant, dict):
            role = participant.get('role', {})
            if isinstance(role, dict) and role.get('code') == 'categorie':
                societe = participant.get('societe', {})
                if isinstance(societe, dict):
                    return societe.get('raison_sociale')

    return None


def find_entreprise_generale(participants: List[Dict]) -> Optional[str]:
    """
    Trouve l'entreprise générale (entreprise_generale > "Entreprise autre" > None).
    """
    if not isinstance(participants, list):
        return None

    # Chercher entreprise_generale d'abord
    for participant in participants:
        if isinstance(participant, dict):
            role = participant.get('role', {})
            if isinstance(role, dict) and role.get('code') == 'entreprise_generale':
                societe = participant.get('societe', {})
                if isinstance(societe, dict):
                    return societe.get('raison_sociale')

    # Puis libelle "Entreprise autre"
    for participant in participants:
        if isinstance(participant, dict):
            role = participant.get('role', {})
            if isinstance(role, dict) and role.get('libelle') == 'Entreprise autre':
                societe = participant.get('societe', {})
                if isinstance(societe, dict):
                    return societe.get('raison_sociale')

    return None


def extract_natures_projet(natures: Any) -> Optional[str]:
    """
    Convertit natures_projet en string lisible.
    """
    if isinstance(natures, list):
        return ', '.join(str(n) for n in natures if n)
    elif isinstance(natures, str):
        return natures
    return None


def extract_lots_description(lots: List[Dict]) -> Optional[str]:
    """
    Crée une liste des objets de lots.
    """
    if not isinstance(lots, list):
        return None

    descriptions = []
    for lot in lots:
        if isinstance(lot, dict):
            objet = lot.get('objet_lot')
            if objet:
                descriptions.append(str(objet))

    return ', '.join(descriptions) if descriptions else None


def coalesce(*values: Any) -> Optional[str]:
    """
    Retourne la première valeur non-None de la liste.
    """
    for value in values:
        if value is not None and value != '':
            return str(value)
    return None


def extract_row_data(market: Dict) -> Dict[str, Any]:
    """
    Extrait une ligne de données complète selon les mappages définis.
    """
    participation = market.get('participation', {})
    societes_intervenantes = participation.get('societes_intervenantes', [])

    calendrier = market.get('calendrier', {})
    date_limite = calendrier.get('date_limite_remise_offres', {})

    dce_data = market.get('dce_data', {})

    qualification = market.get('qualification', {})

    lotissement = market.get('lotissement', {})
    lots = lotissement.get('lots', [])

    # Gérer localisation qui peut contenir site_execution comme array
    localisation = market.get('localisation', {})
    site_execution = None
    if isinstance(localisation, dict):
        site_exec_data = localisation.get('site_execution')
        if isinstance(site_exec_data, list) and len(site_exec_data) > 0:
            site_execution_obj = site_exec_data[0]
            if isinstance(site_execution_obj, dict):
                site_execution = site_execution_obj.get('libelle')
        elif isinstance(site_exec_data, dict):
            site_execution = site_exec_data.get('libelle')

    return {
        'Region': market.get('Région'),
        'Agence': market.get('Agence'),
        'Nom de l\'AO': market.get('objet_marche'),
        'Lieu des travaux': site_execution,
        'Etablissement SMAC': '',  # À remplir manuellement
        'Source de l\'AO': 'Vecteur Plus',
        'Date de Saisie': extract_date_part(get_nested_value(market, 'application_data/livraison/date_premiere_livraison')),
        'SI Rectificatif': get_nested_value(market, 'application_data/livraison/motif'),
        'Nom du Client final': find_client_final(societes_intervenantes),
        'Date limite remise offre': extract_date_part(date_limite.get('date')),
        'Heure limite de remise offre': extract_time_part(date_limite.get('date')),
        'Nom de l\'architecte': find_participant_by_role(societes_intervenantes, 'architecte'),
        'Nom de l\'economiste': find_participant_by_role(societes_intervenantes, 'economiste'),
        'Nom de l\'entreprise generale': find_entreprise_generale(societes_intervenantes),
        'Public/Prive': get_nested_value(qualification, 'type_procedure/libelle'),
        'Nature du projet': extract_natures_projet(get_nested_value(qualification, 'natures_projet')),
        'Visite de site': coalesce(
            get_nested_value(qualification, 'renseignements_complementaires'),
            # Ajouter d'autres champs si nécessaire
        ),
        'Lien vers AO': coalesce(
            get_nested_value(qualification, 'dce_url'),
            dce_data.get('url')
        ),
        'Technique': get_nested_value(qualification, 'renseignements_techniques'),
        'Description technique': extract_lots_description(lots),
    }


def create_xlsx_file(xlsx_file: str, columns: List[str], rows: List[Dict[str, Any]]) -> None:
    """
    Crée un fichier XLSX sans dépendance externe en utilisant zipfile et XML.
    """
    def col_letter(n: int) -> str:
        """Convertir un numéro de colonne en lettre (1=A, 2=B, etc.)"""
        result = ""
        while n > 0:
            n -= 1
            result = chr(65 + n % 26) + result
            n //= 26
        return result

    def cell_ref(row: int, col: int) -> str:
        """Créer une référence de cellule (ex: A1)"""
        return f"{col_letter(col)}{row}"

    # Créer le contenu du fichier sheet1.xml
    def create_sheet_xml() -> str:
        xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        xml += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
        xml += '<sheetData>\n'

        # En-têtes
        xml += '<row r="1">'
        for col_num, column_title in enumerate(columns, 1):
            xml += f'<c r="{cell_ref(1, col_num)}" t="inlineStr" s="1"><is><t>{escape_xml(column_title)}</t></is></c>'
        xml += '</row>\n'

        # Données
        for row_num, row_data in enumerate(rows, 2):
            xml += f'<row r="{row_num}">'
            for col_num, column_title in enumerate(columns, 1):
                value = row_data.get(column_title)
                if value is not None:
                    xml += f'<c r="{cell_ref(row_num, col_num)}" t="inlineStr"><is><t>{escape_xml(str(value))}</t></is></c>'
                else:
                    xml += f'<c r="{cell_ref(row_num, col_num)}" />'
            xml += '</row>\n'

        xml += '</sheetData>\n'
        xml += '</worksheet>'
        return xml

    def escape_xml(text: str) -> str:
        """Échapper les caractères spéciaux XML"""
        text = text.replace("&", "&amp;")
        text = text.replace("<", "&lt;")
        text = text.replace(">", "&gt;")
        text = text.replace('"', "&quot;")
        text = text.replace("'", "&apos;")
        return text

    # Créer les fichiers internes du ZIP
    files = {}

    # [Content_Types].xml
    files['[Content_Types].xml'] = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>'''

    # _rels/.rels
    files['_rels/.rels'] = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>'''

    # xl/_rels/workbook.xml.rels
    files['xl/_rels/workbook.xml.rels'] = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>'''

    # xl/workbook.xml
    files['xl/workbook.xml'] = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="Données" sheetId="1" r:id="rId1"/>
</sheets>
</workbook>'''

    # xl/styles.xml
    files['xl/styles.xml'] = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2">
<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>
<font><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/><b val="true"/></font>
</fonts>
<fills count="2">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
</fills>
<borders count="1">
<border><left/><right/><top/><bottom/><diagonal/></border>
</borders>
<cellStyleXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
</cellStyleXfs>
<cellXfs count="2">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0"><patternFill patternType="solid" fgColor="FF4472C4"/></xf>
</cellXfs>
</styleSheet>'''

    # xl/theme/theme1.xml (minimal)
    files['xl/theme/theme1.xml'] = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
<a:themeElements/>
</a:theme>'''

    # docProps/core.xml
    files['docProps/core.xml'] = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:creator>Python Script</dc:creator>
<cp:lastModifiedBy>Python Script</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>'''

    # xl/worksheets/sheet1.xml
    files['xl/worksheets/sheet1.xml'] = create_sheet_xml()

    # Créer le ZIP (fichier XLSX)
    with zipfile.ZipFile(xlsx_file, 'w', zipfile.ZIP_DEFLATED) as zf:
        for filename, content in files.items():
            zf.writestr(filename, content)


def read_csv_rows(csv_file: str) -> List[Dict[str, Any]]:
    if not Path(csv_file).exists():
        return []
    csv.field_size_limit(10 * 1024 * 1024)
    with open(csv_file, 'r', encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        return [row for row in reader]


def append_csv_rows(csv_file: str, columns: List[str], rows: List[Dict[str, Any]]) -> None:
    file_exists = Path(csv_file).exists()
    with open(csv_file, 'a' if file_exists else 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=columns)
        if not file_exists:
            writer.writeheader()
        writer.writerows(rows)


def convert_json_to_csv_and_xlsx(json_file: str, output_prefix: str = 'output') -> None:
    """
    Convertit un fichier JSON en fichiers CSV et XLSX.

    Args:
        json_file: Chemin du fichier JSON source
        output_prefix: Préfixe pour les fichiers de sortie
    """
    # Lire le fichier JSON
    with open(json_file, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # S'assurer que c'est un array
    if not isinstance(data, list):
        raise ValueError("Le fichier JSON doit contenir un array")

    # Définir les colonnes dans l'ordre souhaité
    columns = [
        'Region',
        'Agence',
        'Nom de l\'AO',
        'Lieu des travaux',
        'Etablissement SMAC',
        'Source de l\'AO',
        'Date de Saisie',
        'SI Rectificatif',
        'Nom du Client final',
        'Date limite remise offre',
        'Heure limite de remise offre',
        'Nom de l\'architecte',
        'Nom de l\'economiste',
        'Nom de l\'entreprise generale',
        'Public/Prive',
        'Nature du projet',
        'Visite de site',
        'Lien vers AO',
        'Technique',
        'Description technique',
    ]

    # Extraire les données
    rows = []
    for market in data:
        try:
            row = extract_row_data(market)
            rows.append(row)
        except Exception as e:
            print(f"Erreur lors du traitement d'un élément: {e}")
            continue

    # Créer / mettre à jour le fichier CSV
    csv_file = f"{output_prefix}.csv"
    append_csv_rows(csv_file, columns, rows)
    print(f"✓ Fichier CSV mis à jour: {csv_file}")

    # Reconstruire le XLSX à partir du CSV cumulatif
    xlsx_file = f"{output_prefix}.xlsx"
    all_rows = read_csv_rows(csv_file)
    create_xlsx_file(xlsx_file, columns, all_rows)
    print(f"✓ Fichier XLSX mis à jour: {xlsx_file}")

    print(f"\nRésumé: {len(rows)} enregistrements ajoutés, {len(all_rows)} enregistrements totaux")


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python convert_json_to_table.py <json_file> [output_prefix]")
        print("\nExemple:")
        print("  python convert_json_to_table.py responses/all_markets.json output")
        sys.exit(1)

    json_file = sys.argv[1]
    output_prefix = sys.argv[2] if len(sys.argv) > 2 else "output"

    if not Path(json_file).exists():
        print(f"Erreur: Le fichier {json_file} n'existe pas")
        sys.exit(1)

    try:
        convert_json_to_csv_and_xlsx(json_file, output_prefix)
    except Exception as e:
        print(f"Erreur: {e}")
        sys.exit(1)

