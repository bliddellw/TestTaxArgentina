#!/usr/bin/env python3
"""
fetch_ar_pit_bands.py

Fetches Argentina individual income tax bands from PwC Tax Summaries,
transforms them into marginal band rates, and appends the rows to an
Excel table stored in the authenticated user's OneDrive via Microsoft
Graph API.

Usage:
    python fetch_ar_pit_bands.py [--dry-run]

Environment variables (required unless --dry-run):
    MS_TENANT_ID      Azure AD tenant ID
    MS_CLIENT_ID      App registration client ID
    MS_CLIENT_SECRET  App registration client secret

Optional environment variables:
    ONEDRIVE_FILE_PATH  Path in OneDrive where the workbook is stored
                        (default: /TaxRates/Argentina_PIT_Bands.xlsx)
"""

import argparse
import json
import logging
import os
import re
import sys
from datetime import datetime, timezone

import requests
from bs4 import BeautifulSoup

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

PWC_URL = (
    "https://taxsummaries.pwc.com/argentina/individual/taxes-on-personal-income"
)

DEFAULT_ONEDRIVE_PATH = "/TaxRates/Argentina_PIT_Bands.xlsx"
WORKSHEET_NAME = "Bands"
TABLE_NAME = "BandsTable"
TABLE_COLUMNS = ["snapshot_date", "band_min", "band_max", "band_rate"]

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_TOKEN_URL = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%dT%H:%M:%SZ",
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Scraping
# ---------------------------------------------------------------------------


def fetch_html(url: str) -> str:
    """Download the PwC Tax Summaries page and return its HTML."""
    log.info("Fetching %s", url)
    resp = requests.get(url, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
    resp.raise_for_status()
    log.info("Received %d bytes", len(resp.content))
    return resp.text


def _parse_number(text: str):
    """Return an integer from a string like '1,749,902', or None."""
    cleaned = text.strip().replace(",", "").replace("\xa0", "").replace(" ", "")
    if not cleaned or cleaned in ("-", "—", ""):
        return None
    try:
        return int(cleaned)
    except ValueError:
        try:
            return int(float(cleaned))
        except ValueError:
            return None


def _parse_percent(text: str):
    """Return a decimal fraction (e.g. 0.05) from a string like '5%'."""
    cleaned = text.strip().replace("%", "").replace(",", ".").strip()
    if not cleaned or cleaned in ("-", "—"):
        return 0.0
    try:
        return float(cleaned) / 100.0
    except ValueError:
        return 0.0


def _find_pit_table(soup: BeautifulSoup):
    """
    Locate the PIT band table on the PwC page.

    The target table has headers resembling:
        Over (ARS) | Not over (ARS) | Tax on column 1 | % on excess
    Returns the <table> element or raises RuntimeError.
    """
    for table in soup.find_all("table"):
        headers = [
            th.get_text(strip=True).lower()
            for th in table.find_all("th")
        ]
        header_text = " ".join(headers)
        # Accept the table if it contains the key header tokens
        if "over" in header_text and "not over" in header_text and "%" in header_text:
            log.info("Found PIT table with headers: %s", headers)
            return table

    # Fallback: look for a table inside a section that mentions "tax rates"
    for heading in soup.find_all(["h2", "h3", "h4"]):
        if "tax rate" in heading.get_text(strip=True).lower():
            sibling = heading.find_next("table")
            if sibling:
                log.info(
                    "Found table via heading '%s'", heading.get_text(strip=True)
                )
                return sibling

    raise RuntimeError(
        "Could not find the Argentina PIT bands table on the PwC page. "
        "The page structure may have changed. "
        "Please check: " + PWC_URL
    )


def parse_pit_bands(html: str) -> list[dict]:
    """
    Parse the HTML and return a list of raw band dicts:
        {over, not_over, tax_on_col1, pct_on_excess}
    Sorted ascending by `over`.
    """
    soup = BeautifulSoup(html, "html.parser")
    table = _find_pit_table(soup)

    rows = table.find_all("tr")
    bands = []
    for row in rows:
        cells = row.find_all(["td", "th"])
        if not cells or any(
            c.name == "th" for c in cells
        ):
            # Skip header rows
            continue
        texts = [c.get_text(strip=True) for c in cells]
        if len(texts) < 4:
            continue

        over_val = _parse_number(texts[0])
        if over_val is None:
            # Likely a header-like row embedded in tbody
            continue

        not_over_text = texts[1].lower()
        if "and on" in not_over_text or not_over_text in ("", "-", "—"):
            not_over_val = None
        else:
            not_over_val = _parse_number(texts[1])

        tax_on_col1 = _parse_number(texts[2]) if texts[2].strip() else None
        pct_on_excess = _parse_percent(texts[3])

        bands.append(
            {
                "over": over_val,
                "not_over": not_over_val,
                "tax_on_col1": tax_on_col1,
                "pct_on_excess": pct_on_excess,
            }
        )

    if not bands:
        raise RuntimeError(
            "Parsed zero rows from the PIT table. "
            "The page structure may have changed."
        )

    bands.sort(key=lambda b: b["over"])
    log.info("Parsed %d raw band rows", len(bands))
    return bands


def transform_to_marginal_rates(
    raw_bands: list[dict], snapshot_date: str
) -> list[dict]:
    """
    Convert raw PwC bands (tax-on-excess) into marginal band rows.

    PwC semantics:
        row[i].pct_on_excess  = rate applied to income above row[i].over
                                 up to row[i].not_over

    Desired output semantics:
        band_rate for band [min, max) = pct_on_excess of the *previous* row
        (The first band's rate is always 0%.)

    Returns list of dicts with keys: snapshot_date, band_min, band_max, band_rate
    """
    output = []
    for i, band in enumerate(raw_bands):
        if i == 0:
            marginal_rate = 0.0
        else:
            marginal_rate = raw_bands[i - 1]["pct_on_excess"]

        output.append(
            {
                "snapshot_date": snapshot_date,
                "band_min": band["over"],
                "band_max": band["not_over"],  # None for the top band
                "band_rate": marginal_rate,
            }
        )

    log.info("Transformed to %d marginal-rate rows", len(output))
    return output


# ---------------------------------------------------------------------------
# Microsoft Graph helpers
# ---------------------------------------------------------------------------


def get_access_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    """Obtain an OAuth2 client-credentials token from Entra ID."""
    url = GRAPH_TOKEN_URL.format(tenant=tenant_id)
    resp = requests.post(
        url,
        data={
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
        },
        timeout=30,
    )
    resp.raise_for_status()
    token = resp.json()["access_token"]
    log.info("Acquired access token (tenant=%s, client=%s)", tenant_id, client_id)
    return token


def _graph_headers(token: str) -> dict:
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }


def _ensure_folder(token: str, drive_id: str, folder_path: str):
    """Create the folder hierarchy in OneDrive if it does not exist."""
    # folder_path like  /TaxRates
    parts = [p for p in folder_path.split("/") if p]
    current = ""
    for part in parts:
        parent_ref = f"/root:/{current}" if current else "/root"
        current = f"{current}/{part}" if current else part
        check_url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{current}"
        r = requests.get(check_url, headers=_graph_headers(token), timeout=30)
        if r.status_code == 404:
            # Create folder
            if current.count("/") == 0:
                create_url = f"{GRAPH_BASE}/drives/{drive_id}/root/children"
            else:
                parent_path = "/".join(current.split("/")[:-1])
                create_url = (
                    f"{GRAPH_BASE}/drives/{drive_id}/root:/{parent_path}:/children"
                )
            body = {"name": part, "folder": {}, "@microsoft.graph.conflictBehavior": "rename"}
            cr = requests.post(
                create_url, headers=_graph_headers(token), json=body, timeout=30
            )
            cr.raise_for_status()
            log.info("Created OneDrive folder: %s", current)
        else:
            r.raise_for_status()


def _get_or_create_workbook(token: str, drive_id: str, file_path: str) -> str:
    """
    Return the OneDrive item ID for the workbook at file_path,
    creating an empty .xlsx if it does not exist.
    """
    # file_path like /TaxRates/Argentina_PIT_Bands.xlsx
    encoded = file_path.lstrip("/")
    check_url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{encoded}"
    r = requests.get(check_url, headers=_graph_headers(token), timeout=30)

    if r.status_code == 200:
        item_id = r.json()["id"]
        log.info("Workbook already exists (item_id=%s)", item_id)
        return item_id

    if r.status_code != 404:
        r.raise_for_status()

    # Create parent folder
    folder_path = "/".join(file_path.split("/")[:-1])
    if folder_path and folder_path != "/":
        _ensure_folder(token, drive_id, folder_path)

    # Upload a minimal valid xlsx (base64 of an empty workbook)
    # We use openpyxl to generate the bytes in-memory
    import io
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = WORKSHEET_NAME
    # Write header row
    ws.append(TABLE_COLUMNS)
    buf = io.BytesIO()
    wb.save(buf)
    xlsx_bytes = buf.getvalue()

    upload_url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{encoded}:/content"
    up_headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }
    ur = requests.put(upload_url, headers=up_headers, data=xlsx_bytes, timeout=60)
    ur.raise_for_status()
    item_id = ur.json()["id"]
    log.info("Uploaded new workbook (item_id=%s)", item_id)
    return item_id


def _ensure_worksheet_and_table(token: str, drive_id: str, item_id: str):
    """
    Ensure worksheet WORKSHEET_NAME and table TABLE_NAME exist in the workbook.
    Creates them if absent.
    """
    ws_url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/workbook/worksheets"
    ws_resp = requests.get(ws_url, headers=_graph_headers(token), timeout=30)
    ws_resp.raise_for_status()
    existing_sheets = [w["name"] for w in ws_resp.json().get("value", [])]

    if WORKSHEET_NAME not in existing_sheets:
        add_resp = requests.post(
            ws_url,
            headers=_graph_headers(token),
            json={"name": WORKSHEET_NAME},
            timeout=30,
        )
        add_resp.raise_for_status()
        log.info("Created worksheet '%s'", WORKSHEET_NAME)

    # Check for existing table
    tbl_url = (
        f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets/{WORKSHEET_NAME}/tables"
    )
    tbl_resp = requests.get(tbl_url, headers=_graph_headers(token), timeout=30)
    tbl_resp.raise_for_status()
    existing_tables = [t["name"] for t in tbl_resp.json().get("value", [])]

    if TABLE_NAME not in existing_tables:
        # Write header to A1 first so we can create a table over A1:D1
        range_url = (
            f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}"
            f"/workbook/worksheets/{WORKSHEET_NAME}/range(address='A1:D1')"
        )
        rng_resp = requests.patch(
            range_url,
            headers=_graph_headers(token),
            json={"values": [TABLE_COLUMNS]},
            timeout=30,
        )
        rng_resp.raise_for_status()

        # Create the table
        create_tbl = requests.post(
            tbl_url,
            headers=_graph_headers(token),
            json={"address": "A1:D1", "hasHeaders": True},
            timeout=30,
        )
        create_tbl.raise_for_status()
        tbl_id = create_tbl.json().get("id") or create_tbl.json().get("name")

        # Rename the table
        rename_url = (
            f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}"
            f"/workbook/worksheets/{WORKSHEET_NAME}/tables/{tbl_id}"
        )
        rename_resp = requests.patch(
            rename_url,
            headers=_graph_headers(token),
            json={"name": TABLE_NAME},
            timeout=30,
        )
        rename_resp.raise_for_status()
        log.info("Created table '%s'", TABLE_NAME)


def _get_drive_id(token: str) -> str:
    """Return the drive ID for the authenticated user's OneDrive."""
    url = f"{GRAPH_BASE}/me/drive"
    r = requests.get(url, headers=_graph_headers(token), timeout=30)
    r.raise_for_status()
    drive_id = r.json()["id"]
    log.info("OneDrive drive_id=%s", drive_id)
    return drive_id


def append_rows_to_table(token: str, file_path: str, rows: list[dict]):
    """
    Append rows to the BandsTable in the OneDrive workbook.
    Creates the workbook / worksheet / table if they do not exist.
    """
    drive_id = _get_drive_id(token)
    item_id = _get_or_create_workbook(token, drive_id, file_path)
    _ensure_worksheet_and_table(token, drive_id, item_id)

    # Build the rows payload [[col1, col2, col3, col4], ...]
    values = []
    for row in rows:
        values.append(
            [
                row["snapshot_date"],
                row["band_min"],
                row["band_max"] if row["band_max"] is not None else "",
                row["band_rate"],
            ]
        )

    append_url = (
        f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}"
        f"/workbook/tables/{TABLE_NAME}/rows/add"
    )
    payload = {"values": values}
    resp = requests.post(
        append_url,
        headers=_graph_headers(token),
        json=payload,
        timeout=60,
    )
    resp.raise_for_status()
    log.info("Appended %d rows to %s!%s", len(rows), file_path, TABLE_NAME)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main():
    parser = argparse.ArgumentParser(
        description="Fetch Argentina PIT bands from PwC and write to OneDrive Excel."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Parse and print data without writing to OneDrive.",
    )
    args = parser.parse_args()

    snapshot_date = datetime.now(tz=timezone.utc).strftime("%Y-%m-%d")
    log.info("Snapshot date: %s", snapshot_date)

    # --- Scrape ---
    html = fetch_html(PWC_URL)
    raw_bands = parse_pit_bands(html)
    output_rows = transform_to_marginal_rates(raw_bands, snapshot_date)

    if args.dry_run:
        print("\n--- DRY RUN OUTPUT ---")
        print(f"{'snapshot_date':<14}  {'band_min':>12}  {'band_max':>12}  {'band_rate':>10}")
        print("-" * 56)
        for row in output_rows:
            bmax = str(row["band_max"]) if row["band_max"] is not None else ""
            print(
                f"{row['snapshot_date']:<14}  {row['band_min']:>12}  {bmax:>12}  {row['band_rate']:>10.4f}"
            )
        print(f"\nTotal rows: {len(output_rows)}")
        return

    # --- Write to OneDrive ---
    tenant_id = os.environ.get("MS_TENANT_ID", "")
    client_id = os.environ.get("MS_CLIENT_ID", "")
    client_secret = os.environ.get("MS_CLIENT_SECRET", "")

    missing = [
        name
        for name, val in [
            ("MS_TENANT_ID", tenant_id),
            ("MS_CLIENT_ID", client_id),
            ("MS_CLIENT_SECRET", client_secret),
        ]
        if not val
    ]
    if missing:
        log.error(
            "Missing required environment variables: %s. "
            "Use --dry-run to skip OneDrive integration.",
            ", ".join(missing),
        )
        sys.exit(1)

    file_path = os.environ.get("ONEDRIVE_FILE_PATH", DEFAULT_ONEDRIVE_PATH)
    token = get_access_token(tenant_id, client_id, client_secret)
    append_rows_to_table(token, file_path, output_rows)
    log.info("Done.")


if __name__ == "__main__":
    main()
