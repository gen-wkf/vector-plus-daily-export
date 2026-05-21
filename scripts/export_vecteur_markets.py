#!/usr/bin/env python3
"""Export Vecteur Plus market details for all client/application pairs.

The script:
1. Logs in to fetch token + clients/applications
2. Fetches market IDs from `/marches/search` for a page interval (default: 1..6)
3. Fetches each market detail from `/marches/{marcheId}`
4. Exports results to JSON, CSV, and XLSX
"""

from __future__ import annotations

import argparse
from copy import deepcopy
import csv
from datetime import date, timedelta
import json
import re
import socket
import sys
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any, Iterable
from urllib import error, request


SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent

DEFAULT_HTTP_FILE = PROJECT_ROOT / "app.http"
DEFAULT_OUTPUT_PREFIX = PROJECT_ROOT / "responses/all_markets"
DEFAULT_PAGE_START = 1
DEFAULT_PAGE_END = 6
DEFAULT_PAGE_SIZE = 100
DEFAULT_TIMEOUT_SECONDS = 60
DEFAULT_WORKERS = 8
DEFAULT_PROGRESS_EVERY = 200

DEFAULT_FILTERS = {
    "date_modification": {
        "from": ""
    },
    "sort": {
        "selected": {
            "code": "date_livraison",
            "order": "desc",
        }
    },
    "filter_groups": [
        {
            "name": "resultat",
            "filters": [
                {"name": "SalesForce", "value": False},
                {"name": "CRM", "value": False},
                {"name": "En attente", "value": False},
                {"name": "Hubspot", "value": False},
            ],
        }
    ]
}


class APIRequestError(RuntimeError):
    """Represents API/network/JSON errors for a single HTTP request."""

    def __init__(self, message: str, *, status: int | None = None, body: str | None = None):
        super().__init__(message)
        self.status = status
        self.body = body


@dataclass(frozen=True)
class MarketRef:
    client_id: int
    client_name: str
    application_id: int
    application_name: str
    page: int
    marche_result_id: str
    marche_result_marche_id: Any


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Fetch market details for all client/application pairs and export JSON/CSV/XLSX files."
        )
    )
    parser.add_argument("--http-file", type=Path, default=DEFAULT_HTTP_FILE)
    parser.add_argument("--base-url", type=str, default=None)
    parser.add_argument("--password", type=str, default=None, help="Vecteur Plus login password.")
    parser.add_argument("--token", type=str, default=None, help="Override token returned by /login.")
    parser.add_argument("--output-prefix", type=Path, default=DEFAULT_OUTPUT_PREFIX)
    parser.add_argument("--page-start", type=int, default=DEFAULT_PAGE_START)
    parser.add_argument(
        "--max-pages",
        type=int,
        default=None,
        help="Maximum number of pages to fetch per app. If omitted, follows 'has_next' from API.",
    )
    parser.add_argument("--page-size", type=int, default=DEFAULT_PAGE_SIZE)
    parser.add_argument("--workers", type=int, default=DEFAULT_WORKERS)
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT_SECONDS)
    parser.add_argument(
        "--progress-every",
        type=int,
        default=DEFAULT_PROGRESS_EVERY,
        help="Log progress every N market detail calls.",
    )
    parser.add_argument(
        "--without-facets",
        action="store_true",
        help="Do not send default facets block in the /marches/search payload.",
    )
    parser.add_argument(
        "--client-ids",
        type=str,
        default="",
        help="Comma-separated client IDs to include. Empty = all.",
    )
    parser.add_argument(
        "--application-ids",
        type=str,
        default="",
        help="Comma-separated application IDs to include. Empty = all.",
    )
    parser.add_argument(
        "--smac",
        action="store_true",
        help="Export in SMAC format with specific column mapping and transformations.",
    )
    return parser.parse_args()


def parse_id_filter(raw: str) -> set[int] | None:
    cleaned = raw.strip()
    if not cleaned:
        return None
    ids: set[int] = set()
    for part in cleaned.split(","):
        part = part.strip()
        if not part:
            continue
        try:
            ids.add(int(part))
        except ValueError as exc:
            raise ValueError(f"Invalid ID '{part}' in filter '{raw}'.") from exc
    return ids


def read_http_variable(path: Path, name: str) -> str | None:
    pattern = re.compile(rf"@{re.escape(name)}\s*=\s*(\S+)")
    try:
        content = path.read_text(encoding="utf-8")
    except OSError:
        return None

    match = pattern.search(content)
    return match.group(1).strip() if match else None


def read_http_login_password(path: Path) -> str | None:
    try:
        content = path.read_text(encoding="utf-8")
    except OSError:
        return None

    match = re.search(r'"password"\s*:\s*"([^"]+)"', content)
    return match.group(1) if match else None


def resolve_base_url(args: argparse.Namespace) -> str:
    if args.base_url:
        return args.base_url.rstrip("/")

    base_url = read_http_variable(args.http_file, "base_url_vecteur_plus")
    if base_url:
        return base_url.rstrip("/")

    if not args.http_file.exists():
        raise RuntimeError(
            f"Unable to read {args.http_file} and no --base-url was provided."
        )
    raise RuntimeError(
        f"Could not extract @base_url_vecteur_plus from {args.http_file}. "
        "Pass --base-url explicitly."
    )


def request_json(
    *,
    url: str,
    method: str,
    timeout: int,
    token: str | None = None,
    payload: Any | None = None,
    wait_before_request: bool = True,
) -> Any:
    if wait_before_request:
        time.sleep(5)
    headers = {"Accept": "application/json"}
    if token is not None:
        headers["Authorization"] = f"Bearer {token}"
    data = None
    if payload is not None:
        headers["Content-Type"] = "application/json"
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")

    req = request.Request(url=url, data=data, headers=headers, method=method)
    try:
        with request.urlopen(req, timeout=timeout) as response:
            raw = response.read().decode("utf-8")
    except (TimeoutError, socket.timeout) as exc:
        raise APIRequestError(
            f"Timeout for {method} {url} after {timeout}s"
        ) from exc
    except error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        raise APIRequestError(
            f"HTTP {exc.code} for {method} {url}", status=exc.code, body=body
        ) from exc
    except error.URLError as exc:
        raise APIRequestError(f"Network error for {method} {url}: {exc}") from exc

    if not raw.strip():
        return {}
    try:
        return json.loads(raw)
    except json.JSONDecodeError as exc:
        snippet = raw[:300]
        raise APIRequestError(
            f"Invalid JSON for {method} {url}. Body starts with: {snippet}"
        ) from exc


def login(base_url: str, password: str, timeout: int) -> dict[str, Any]:
    response = request_json(
        url=f"{base_url}/login",
        method="POST",
        timeout=timeout,
        payload={"password": password},
        wait_before_request=False,
    )
    if not isinstance(response, dict):
        raise APIRequestError("Login response is not a JSON object.")
    return response


def build_search_payload(page: int, size: int, include_facets: bool) -> dict[str, Any]:
    payload: dict[str, Any] = {"query": {"page": page, "size": size}}
    if include_facets:
        facets = deepcopy(DEFAULT_FILTERS)
        yesterday = date.today() - timedelta(days=1)
        facets["date_modification"]["from"] = f"{yesterday:%Y-%m-%d}T00:00:00.000"
        payload["facets"] = facets
    return payload


def collect_market_refs(
    *,
    base_url: str,
    token: str,
    clients: Iterable[dict[str, Any]],
    page_start: int,
    max_pages: int | None,
    page_size: int,
    include_facets: bool,
    timeout: int,
    client_filter: set[int] | None,
    application_filter: set[int] | None,
) -> tuple[list[MarketRef], list[dict[str, Any]]]:
    refs: list[MarketRef] = []
    errors: list[dict[str, Any]] = []
    seen: set[tuple[int, int, str]] = set()

    for client in clients:
        client_id = client.get("id")
        if not isinstance(client_id, int):
            continue
        if client_filter and client_id not in client_filter:
            continue

        client_name = str(client.get("name", ""))
        applications = client.get("applications") or []
        if not isinstance(applications, list):
            continue

        for app in applications:
            application_id = app.get("id")
            if not isinstance(application_id, int):
                continue
            if application_filter and application_id not in application_filter:
                continue

            application_name = str(app.get("name", ""))

            page = page_start
            pages_fetched = 0
            while True:
                if max_pages is not None and pages_fetched >= max_pages:
                    break

                print(f"Client: {client_id} - App: {application_id} - Collecting market IDs at page {page}...", file=sys.stderr)
                url = f"{base_url}/{client_id}/{application_id}/marches/search"
                payload = build_search_payload(page, page_size, include_facets)
                try:
                    response = request_json(
                        url=url,
                        method="PUT",
                        token=token,
                        timeout=timeout,
                        payload=payload,
                    )
                except APIRequestError as exc:
                    if exc.status == 401:
                        print(f"Error: Unauthorized access (401) for {client_name}/{application_name}. The token might be expired.", file=sys.stderr)
                    errors.append(
                        {
                            "stage": "search",
                            "client_id": client_id,
                            "client_name": client_name,
                            "application_id": application_id,
                            "application_name": application_name,
                            "page": page,
                            "url": url,
                            "status": exc.status,
                            "message": str(exc),
                            "body": (exc.body or "")[:1200],
                        }
                    )
                    break # Stop pagination for this app on error

                results = response.get("results") if isinstance(response, dict) else None
                if not isinstance(results, list):
                    errors.append(
                        {
                            "stage": "search",
                            "client_id": client_id,
                            "client_name": client_name,
                            "application_id": application_id,
                            "application_name": application_name,
                            "page": page,
                            "url": url,
                            "status": None,
                            "message": "Response has no 'results' list.",
                            "body": json.dumps(response, ensure_ascii=False)[:1200],
                        }
                    )
                    break

                for result in results:
                    if not isinstance(result, dict):
                        continue
                    marche_result_id = str(result.get("id", "")).strip()
                    if not marche_result_id:
                        continue
                    key = (client_id, application_id, marche_result_id)
                    if key in seen:
                        continue
                    seen.add(key)
                    refs.append(
                        MarketRef(
                            client_id=client_id,
                            client_name=client_name,
                            application_id=application_id,
                            application_name=application_name,
                            page=page,
                            marche_result_id=marche_result_id,
                            marche_result_marche_id=result.get("marche_id"),
                        )
                    )

                pages_fetched += 1
                has_next = response.get("has_next")
                if has_next is True: # and page < 2:
                    page += 1
                else:
                    break

    return refs, errors


def fetch_market_detail(
    *,
    base_url: str,
    token: str,
    timeout: int,
    market_ref: MarketRef,
) -> tuple[dict[str, Any] | None, dict[str, Any] | None]:
    url = (
        f"{base_url}/{market_ref.client_id}/{market_ref.application_id}/marches/"
        f"{market_ref.marche_result_id}"
    )
    try:
        detail = request_json(url=url, method="GET", token=token, timeout=timeout)
    except APIRequestError as exc:
        return None, {
            "stage": "detail",
            "client_id": market_ref.client_id,
            "client_name": market_ref.client_name,
            "application_id": market_ref.application_id,
            "application_name": market_ref.application_name,
            "page": market_ref.page,
            "marche_result_id": market_ref.marche_result_id,
            "marche_result_marche_id": market_ref.marche_result_marche_id,
            "url": url,
            "status": exc.status,
            "message": str(exc),
            "body": (exc.body or "")[:1200],
        }

    if not isinstance(detail, dict):
        return None, {
            "stage": "detail",
            "client_id": market_ref.client_id,
            "client_name": market_ref.client_name,
            "application_id": market_ref.application_id,
            "application_name": market_ref.application_name,
            "page": market_ref.page,
            "marche_result_id": market_ref.marche_result_id,
            "marche_result_marche_id": market_ref.marche_result_marche_id,
            "url": url,
            "status": None,
            "message": "Detail payload is not an object.",
            "body": json.dumps(detail, ensure_ascii=False)[:1200],
        }

    detail["Region"] = market_ref.client_name
    detail["Agence"] = market_ref.application_name
    detail["_meta"] = {
        "client_id": market_ref.client_id,
        "client_name": market_ref.client_name,
        "application_id": market_ref.application_id,
        "application_name": market_ref.application_name,
        "page": market_ref.page,
        "marche_result_id": market_ref.marche_result_id,
        "marche_result_marche_id": market_ref.marche_result_marche_id,
        "source_url": url,
    }
    return detail, None


def format_smac_date(iso_str: str | None) -> str:
    if not iso_str:
        return ""
    match = re.search(r"\d{4}-\d{2}-\d{2}", str(iso_str))
    return match.group(0) if match else ""


def format_smac_time(iso_str: str | None) -> str:
    if not iso_str:
        return ""
    text = str(iso_str)
    match = re.search(r"(?:T|\s)(\d{2}):(\d{2})", text)
    if not match:
        match = re.search(r"(\d{2}):(\d{2})", text)
    return f"{match.group(1)}:{match.group(2)}" if match else ""


def to_string_array_cell(values: list[str]) -> str:
    if not values:
        return ""
    return json.dumps(values, ensure_ascii=False)


def extract_values_as_strings(value: Any, *, item_key: str | None = None) -> list[str]:
    items = value if isinstance(value, list) else [value]
    extracted: list[str] = []
    for item in items:
        candidate = item
        if isinstance(item, dict):
            if item_key:
                candidate = item.get(item_key)
            else:
                candidate = item.get("libelle", item.get("label", item.get("name", item.get("code"))))
        if candidate is None:
            continue
        text = str(candidate).strip()
        if text:
            extracted.append(text)
    return extracted


def get_societes_intervenantes(market: dict[str, Any]) -> list[dict[str, Any]]:
    participation = market.get("participation")
    if not isinstance(participation, dict):
        return []
    societes = participation.get("societes_intervenantes")
    if not isinstance(societes, list):
        return []
    return [item for item in societes if isinstance(item, dict)]


def get_societe_by_role_code(market: dict[str, Any], role_code: str) -> str:
    expected = role_code.strip().lower()
    for participant in get_societes_intervenantes(market):
        role = participant.get("role")
        if not isinstance(role, dict):
            continue
        code = str(role.get("code", "")).strip().lower()
        if code != expected:
            continue
        societe = participant.get("societe")
        if not isinstance(societe, dict):
            continue
        raison_sociale = str(societe.get("raison_sociale", "")).strip()
        if raison_sociale:
            return raison_sociale
    return ""


def get_societe_by_role_libelle(market: dict[str, Any], role_libelle: str) -> str:
    expected = role_libelle.strip().lower()
    for participant in get_societes_intervenantes(market):
        role = participant.get("role")
        if not isinstance(role, dict):
            continue
        libelle = str(role.get("libelle", "")).strip().lower()
        if libelle != expected:
            continue
        societe = participant.get("societe")
        if not isinstance(societe, dict):
            continue
        raison_sociale = str(societe.get("raison_sociale", "")).strip()
        if raison_sociale:
            return raison_sociale
    return ""


def first_non_empty(*values: Any) -> str:
    for value in values:
        if value is None:
            continue
        text = str(value).strip()
        if text:
            return text
    return ""


def transform_to_smac(details: list[dict[str, Any]]) -> list[dict[str, Any]]:
    transformed: list[dict[str, Any]] = []
    for m in details:
        meta = m.get("_meta", {})
        if not isinstance(meta, dict):
            meta = {}

        qualification = m.get("qualification", {})
        if not isinstance(qualification, dict):
            qualification = {}

        application_data = m.get("application_data", {})
        if not isinstance(application_data, dict):
            application_data = {}

        livraison = application_data.get("livraison", {})
        if not isinstance(livraison, dict):
            livraison = {}

        calendrier = m.get("calendrier", {})
        if not isinstance(calendrier, dict):
            calendrier = {}

        date_limite = calendrier.get("date_limite_remise_offres", {})
        if not isinstance(date_limite, dict):
            date_limite = {}

        localisation = m.get("localisation", {})
        if not isinstance(localisation, dict):
            localisation = {}

        lotissement = m.get("lotissement", {})
        if not isinstance(lotissement, dict):
            lotissement = {}

        row: dict[str, Any] = {}
        row["Région"] = m.get(
            "Région",
            m.get("Region", m.get("client_name", meta.get("client_name", ""))),
        )
        row["Agence"] = m.get(
            "Agence",
            m.get("application_name", meta.get("application_name", "")),
        )
        row["Nom de l'AO"] = m.get("objet_marche", "")
        row["Lieu des travaux"] = to_string_array_cell(
            extract_values_as_strings(localisation.get("site_execution"), item_key="libelle")
        )
        row["Etablissement SMAC"] = ""
        row["Source de l'AO"] = "Vecteur Plus"
        row["Date de Saisie"] = format_smac_date(livraison.get("date_premiere_livraison"))
        row["SI Rectificatif"] = livraison.get("motif", "")

        client_final = get_societe_by_role_code(m, "moa")
        if not client_final:
            client_final = get_societe_by_role_code(m, "categorie")
        row["Nom du Client final"] = client_final

        date_limite_iso = date_limite.get("date")
        row["Date limite remise offre"] = format_smac_date(date_limite_iso)
        row["Heure limite de remise offre"] = format_smac_time(date_limite_iso)
        row["Nom de l'architecte"] = get_societe_by_role_code(m, "architecte")
        row["Nom de l'économiste"] = get_societe_by_role_code(m, "economiste")

        entreprise_generale = get_societe_by_role_code(m, "entreprise_generale")
        if not entreprise_generale:
            entreprise_generale = get_societe_by_role_libelle(m, "Entreprise autre")
        row["Nom de l'entreprise générale"] = entreprise_generale

        row["Public/Privé"] = (
            qualification.get("type_procedure", {}).get("libelle", "")
            if isinstance(qualification.get("type_procedure"), dict)
            else ""
        )
        row["Nature du projet"] = to_string_array_cell(
            extract_values_as_strings(qualification.get("natures_projet"))
        )
        row["Visite de site"] = first_non_empty(
            qualification.get("renseignements_complementaires"),
            qualification.get("re_renseignements_complementaires"),
        )

        dce_data = m.get("dce_data", {})
        if not isinstance(dce_data, dict):
            dce_data = {}
        row["Lien vers AO"] = first_non_empty(qualification.get("dce_url"), dce_data.get("url"))
        row["Technique"] = qualification.get("renseignements_techniques", "")

        row["Description technique"] = to_string_array_cell(
            extract_values_as_strings(lotissement.get("lots"), item_key="objet_lot")
        )

        transformed.append(row)
    return transformed


def flatten_record(
    value: Any, *, prefix: str = "", output: dict[str, Any] | None = None
) -> dict[str, Any]:
    if output is None:
        output = {}

    if isinstance(value, dict):
        if not value and prefix:
            output[prefix] = "{}"
        for key, nested in value.items():
            next_prefix = f"{prefix}.{key}" if prefix else str(key)
            flatten_record(nested, prefix=next_prefix, output=output)
        return output

    if isinstance(value, list):
        output[prefix] = json.dumps(value, ensure_ascii=False)
        return output

    output[prefix] = "" if value is None else value
    return output


def to_cell_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    return str(value)


def write_json_file(path: Path, records: list[dict[str, Any]]) -> None:
    with path.open("w", encoding="utf-8") as file:
        json.dump(records, file, ensure_ascii=False, indent=2)


def write_csv_file(path: Path, columns: list[str], rows: list[dict[str, Any]]) -> None:
    with path.open("w", encoding="utf-8", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=columns)
        writer.writeheader()
        for row in rows:
            writer.writerow({column: to_cell_text(row.get(column, "")) for column in columns})


def excel_column_name(index_1_based: int) -> str:
    name = ""
    n = index_1_based
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        name = chr(ord("A") + remainder) + name
    return name


def remove_invalid_xml_chars(text: str) -> str:
    def is_valid(char: str) -> bool:
        code = ord(char)
        return (
            code in (0x09, 0x0A, 0x0D)
            or 0x20 <= code <= 0xD7FF
            or 0xE000 <= code <= 0xFFFD
            or 0x10000 <= code <= 0x10FFFF
        )

    return "".join(char for char in text if is_valid(char))


def xml_escape(text: str) -> str:
    sanitized = remove_invalid_xml_chars(text)
    return (
        sanitized.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def make_inline_string_cell(ref: str, value: str) -> str:
    escaped = xml_escape(value)
    preserve = ""
    if value and (value[0].isspace() or value[-1].isspace()):
        preserve = ' xml:space="preserve"'
    return f'<c r="{ref}" t="inlineStr"><is><t{preserve}>{escaped}</t></is></c>'


def write_xlsx_file(path: Path, columns: list[str], rows: list[dict[str, Any]]) -> None:
    with TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        sheet_xml = temp_path / "sheet1.xml"

        with sheet_xml.open("w", encoding="utf-8") as sheet:
            sheet.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
            sheet.write(
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            )
            sheet.write("<sheetData>")

            row_number = 1
            header_cells: list[str] = []
            for col_idx, column in enumerate(columns, start=1):
                ref = f"{excel_column_name(col_idx)}{row_number}"
                header_cells.append(make_inline_string_cell(ref, column))
            sheet.write(f'<row r="{row_number}">{"".join(header_cells)}</row>')

            for source_row in rows:
                row_number += 1
                row_cells: list[str] = []
                for col_idx, column in enumerate(columns, start=1):
                    ref = f"{excel_column_name(col_idx)}{row_number}"
                    value = to_cell_text(source_row.get(column, ""))
                    row_cells.append(make_inline_string_cell(ref, value))
                sheet.write(f'<row r="{row_number}">{"".join(row_cells)}</row>')

            sheet.write("</sheetData></worksheet>")

        workbook_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            "<sheets><sheet name=\"markets\" sheetId=\"1\" r:id=\"rId1\"/></sheets>"
            "</workbook>"
        )
        workbook_rels_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            'Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
            'Target="styles.xml"/>'
            "</Relationships>"
        )
        root_rels_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="xl/workbook.xml"/>'
            "</Relationships>"
        )
        content_types_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/styles.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            "</Types>"
        )
        styles_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
            '<borders count="1"><border/></borders>'
            '<cellStyleXfs count="1"><xf/></cellStyleXfs>'
            '<cellXfs count="1"><xf xfId="0"/></cellXfs>'
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            "</styleSheet>"
        )

        with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("[Content_Types].xml", content_types_xml)
            archive.writestr("_rels/.rels", root_rels_xml)
            archive.writestr("xl/workbook.xml", workbook_xml)
            archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
            archive.writestr("xl/styles.xml", styles_xml)
            archive.write(sheet_xml, "xl/worksheets/sheet1.xml")


def order_columns(columns: set[str]) -> list[str]:
    preferred = [
        "Region",
        "Région",
        "Agence",
        "_meta.client_id",
        "_meta.client_name",
        "_meta.application_id",
        "_meta.application_name",
        "_meta.page",
        "_meta.marche_result_id",
        "_meta.marche_result_marche_id",
        "id",
        "marche_id",
        "objet_marche",
        "reference",
        "type_marche",
        "date_creation",
        "date_modification",
    ]
    remaining = sorted(column for column in columns if column not in preferred)
    ordered: list[str] = [column for column in preferred if column in columns]
    ordered.extend(remaining)
    return ordered


def main() -> int:
    args = parse_args()

    if args.page_start < 1:
        print(
            f"Invalid page start: {args.page_start}",
            file=sys.stderr,
        )
        return 2

    if args.max_pages is not None and args.max_pages < 1:
        print("max-pages must be >= 1", file=sys.stderr)
        return 2

    if args.workers < 1:
        print("workers must be >= 1", file=sys.stderr)
        return 2

    try:
        client_filter = parse_id_filter(args.client_ids)
        application_filter = parse_id_filter(args.application_ids)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 2

    try:
        base_url = resolve_base_url(args)
    except RuntimeError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    password = args.password or read_http_login_password(args.http_file)
    if not password:
        print(
            f"No login password found. Provide --password or add it to {args.http_file}.",
            file=sys.stderr,
        )
        return 1

    print(f"Logging in to {base_url}/login...", file=sys.stderr)
    try:
        auth_data = login(base_url, password, args.timeout)
    except APIRequestError as exc:
        print(f"Login failed: {exc}", file=sys.stderr)
        if exc.body:
            print(exc.body[:1200], file=sys.stderr)
        return 1

    token = args.token or auth_data.get("token")
    if not isinstance(token, str) or not token:
        print("Login response does not contain a token.", file=sys.stderr)
        return 1

    clients = auth_data.get("clients")
    if not isinstance(clients, list):
        print("Login response does not contain a 'clients' array.", file=sys.stderr)
        return 1

    include_facets = not args.without_facets
    print(
        f"Collecting market IDs from {base_url} starting at page {args.page_start}...",
        file=sys.stderr,
    )
    market_refs, errors = collect_market_refs(
        base_url=base_url,
        token=token,
        clients=clients,
        page_start=args.page_start,
        max_pages=args.max_pages,
        page_size=args.page_size,
        include_facets=include_facets,
        timeout=args.timeout,
        client_filter=client_filter,
        application_filter=application_filter,
    )
    print(f"Search references collected: {len(market_refs)}", file=sys.stderr)

    details: list[dict[str, Any]] = []
    total_refs = len(market_refs)
    if total_refs:
        print(f"Fetching {total_refs} market details with {args.workers} workers...", file=sys.stderr)
        processed = 0
        with ThreadPoolExecutor(max_workers=args.workers) as executor:
            futures = [
                executor.submit(
                    fetch_market_detail,
                    base_url=base_url,
                    token=token,
                    timeout=args.timeout,
                    market_ref=market_ref,
                )
                for market_ref in market_refs
            ]
            for future in as_completed(futures):
                detail, maybe_error = future.result()
                processed += 1
                if detail is not None:
                    details.append(detail)
                if maybe_error is not None:
                    errors.append(maybe_error)
                if processed % args.progress_every == 0 or processed == total_refs:
                    print(f"Detail progress: {processed}/{total_refs}", file=sys.stderr)

    print(f"Details fetched successfully: {len(details)}", file=sys.stderr)
    print(f"Errors captured: {len(errors)}", file=sys.stderr)

    if args.smac:
        smac_details = transform_to_smac(details)
        if smac_details:
            columns = list(smac_details[0].keys())
        else:
            columns = []
        flattened_rows = smac_details
    else:
        flattened_rows: list[dict[str, Any]] = []
        all_columns: set[str] = set()
        for detail in details:
            flat = flatten_record(detail)
            flattened_rows.append(flat)
            all_columns.update(flat.keys())
        columns = order_columns(all_columns)

    output_prefix = args.output_prefix
    output_prefix.parent.mkdir(parents=True, exist_ok=True)
    json_file = Path(f"{output_prefix}.json")
    csv_file = Path(f"{output_prefix}.csv")
    xlsx_file = Path(f"{output_prefix}.xlsx")
    errors_file = Path(f"{output_prefix}_errors.json")

    write_json_file(json_file, details)
    write_csv_file(csv_file, columns, flattened_rows)
    write_xlsx_file(xlsx_file, columns, flattened_rows)
    if errors:
        write_json_file(errors_file, errors)

    print(f"Wrote {json_file}", file=sys.stderr)
    print(f"Wrote {csv_file}", file=sys.stderr)
    print(f"Wrote {xlsx_file}", file=sys.stderr)
    if errors:
        print(f"Wrote {errors_file}", file=sys.stderr)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
