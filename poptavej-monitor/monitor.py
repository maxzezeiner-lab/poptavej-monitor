"""
Poptavej.cz public tender ("verejna zakazka") monitor.

Scrapes a filtered search results page on poptavej.cz, keeps only tenders
above a value threshold, diffs them against a persisted "seen" state file,
and emails an Excel report of newly-found tenders (or a short status email
if nothing new was found, or a failure alert if the scrape breaks).

Politeness / low-footprint notes (see README.md for the full rationale):
  - One sequential pass through the target search pages per run.
  - Detail pages are only fetched for tenders that are BOTH above the value
    threshold AND not already in the seen-state (i.e. genuinely new).
  - Small random delays between requests, no concurrency.
  - On a failed request: wait ~30-60s, retry once, then give up.
"""

import json
import os
import random
import re
import smtplib
import sys
import time
import traceback
from datetime import date, datetime, timedelta, timezone
from email.message import EmailMessage
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

BASE_URL = "https://www.poptavej.cz"
TARGET_URL = (
    "https://www.poptavej.cz/verejne-zakazky"
    "?filters%5Bkategorie%5D%5B0%5D=17"
    "&filters%5Bkategorie%5D%5B1%5D=1106"
    "&filters%5Bkategorie%5D%5B2%5D=10"
)

VALUE_THRESHOLD_CZK = 10_000_000

# Listings are sorted newest-first. Once a fetched page's oldest row is older
# than this many days, we stop paginating -- anything further back was
# already covered by a previous run. Kept deliberately larger than the
# 3-day run cadence as a safety margin (clock drift, a slow day, etc.).
# This also keeps a from-scratch first run bounded, instead of walking the
# site's full multi-thousand-page history.
LOOKBACK_DAYS = 6

# Hard safety cap regardless of dates, in case date parsing ever misbehaves.
MAX_PAGES = 60

REQUEST_TIMEOUT_S = 20
RETRY_DELAY_RANGE_S = (30, 60)
POLITE_DELAY_RANGE_S = (2, 5)

STATE_FILE = Path(__file__).parent / "state" / "seen_tenders.json"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "cs-CZ,cs;q=0.9,en;q=0.8",
    "Upgrade-Insecure-Requests": "1",
}

# ---------------------------------------------------------------------------
# TODO (future): authenticated session
#
# Once you have a poptavej.cz login, this is where a session login step
# would go, e.g.:
#
#     session.post(BASE_URL + "/prihlaseni", data={
#         "email": os.environ["POPTAVEJ_LOGIN_EMAIL"],
#         "password": os.environ["POPTAVEJ_LOGIN_PASSWORD"],
#         ...
#     })
#
# Logging in unlocks additional fields (e.g. contact info on the "Kontakt
# na zadavatele" panel of the detail page). Not implemented yet -- do this
# once credentials exist, and extend fetch_detail() to pull the new fields.
# ---------------------------------------------------------------------------


class ScrapeError(Exception):
    """Raised whenever the page doesn't look like what we expect."""


def build_session():
    session = requests.Session()
    session.headers.update(HEADERS)
    return session


def polite_sleep(delay_range=POLITE_DELAY_RANGE_S):
    time.sleep(random.uniform(*delay_range))


def fetch(url, session):
    resp = session.get(url, timeout=REQUEST_TIMEOUT_S)
    resp.raise_for_status()
    return resp.text


def fetch_with_retry(url, session):
    try:
        return fetch(url, session)
    except Exception:
        time.sleep(random.uniform(*RETRY_DELAY_RANGE_S))
        try:
            return fetch(url, session)
        except Exception as second_err:
            raise ScrapeError(f"Failed to fetch {url} (after one retry): {second_err}") from second_err


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------

def parse_czech_value(text):
    """'12 500 000 Kc' -> 12500000. 'neurceno' (or anything with no digits) -> None."""
    digits = re.sub(r"[^\d]", "", text)
    return int(digits) if digits else None


def parse_listing_date(text, today):
    text = text.strip()
    if re.match(r"^Dnes\s+\d{1,2}:\d{2}$", text):
        return today
    if re.match(r"^Včera\s+\d{1,2}:\d{2}$", text):  # "Vcera" (yesterday) - defensive, not observed live
        return today - timedelta(days=1)
    m = re.match(r"^(\d{1,2})\.(\d{1,2})\.(\d{4})$", text)
    if m:
        d, mo, y = (int(g) for g in m.groups())
        try:
            return date(y, mo, d)
        except ValueError:
            return None
    return None


def parse_listing_page(html, page_url):
    soup = BeautifulSoup(html, "html.parser")
    container = soup.select_one("div.procurement-list")
    if container is None:
        raise ScrapeError(
            f"Could not find 'div.procurement-list' on {page_url} - "
            "page structure may have changed."
        )

    rows = []
    for row in container.select("div.row.procurement"):
        name_a = row.select_one(".col.nazev a")
        if name_a is None or not name_a.get("href"):
            raise ScrapeError(
                f"A listing row on {page_url} is missing its name/link - "
                "page structure may have changed."
            )
        name = name_a.get_text(strip=True)
        href = name_a["href"]
        detail_url = urljoin(BASE_URL, href)

        id_match = re.search(r"/verejna-zakazka/([^/]+)/", href)
        # Fall back to the full detail URL as the unique key if the URL
        # pattern ever changes shape (still unique and stable per tender).
        tender_id = id_match.group(1) if id_match else detail_url

        date_el = row.select_one(".col.date")
        date_text = date_el.get_text(strip=True) if date_el else ""

        value_el = row.select_one(".col.cena")
        value_text = value_el.get_text(strip=True) if value_el else ""

        rows.append(
            {
                "tender_id": tender_id,
                "name": name,
                "detail_url": detail_url,
                "date_text": date_text,
                "value_number": parse_czech_value(value_text),
            }
        )
    return rows


def fetch_summary(detail_url, session):
    html = fetch_with_retry(detail_url, session)
    soup = BeautifulSoup(html, "html.parser")
    popis = soup.select_one("p.popis")
    if popis is None:
        raise ScrapeError(
            f"Could not find description ('p.popis') on {detail_url} - "
            "page structure may have changed."
        )
    return popis.get_text(separator="\n", strip=True)


# ---------------------------------------------------------------------------
# Pagination
# ---------------------------------------------------------------------------

def scrape_all_pages(session):
    today = datetime.now().date()
    all_rows = []

    for page_num in range(1, MAX_PAGES + 1):
        url = TARGET_URL if page_num == 1 else f"{TARGET_URL}&page={page_num}"
        if page_num > 1:
            polite_sleep()
        html = fetch_with_retry(url, session)
        rows = parse_listing_page(html, url)

        if page_num == 1 and not rows:
            raise ScrapeError(
                "Zero listings found on page 1 of the search results - "
                "the site may be unreachable, blocking us, or the page "
                "structure has changed."
            )
        if not rows:
            break  # ran past the last page

        all_rows.extend(rows)

        parsed_dates = [parse_listing_date(r["date_text"], today) for r in rows]
        known_dates = [d for d in parsed_dates if d is not None]
        if known_dates:
            oldest_on_page = min(known_dates)
            if (today - oldest_on_page).days > LOOKBACK_DAYS:
                break

    return all_rows


# ---------------------------------------------------------------------------
# State file
# ---------------------------------------------------------------------------

def load_state():
    if not STATE_FILE.exists():
        return {}
    with open(STATE_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def save_state(state):
    STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2, sort_keys=True)
        f.write("\n")


# ---------------------------------------------------------------------------
# Excel output
# ---------------------------------------------------------------------------

def build_excel(new_items, found_at, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Nove zakazky"
    headers = [
        "Cislo zakazky",
        "Nazev",
        "Popis",
        "Predpokladana hodnota (Kc)",
        "Odkaz",
        "Nalezeno (datum/cas)",
    ]
    ws.append(headers)
    for item in new_items:
        ws.append(
            [
                item["tender_id"],
                item["name"],
                item.get("summary", ""),
                item["value_number"],
                item["detail_url"],
                found_at,
            ]
        )
    widths = [16, 45, 60, 22, 55, 20]
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = width
    wb.save(output_path)
    return output_path


# ---------------------------------------------------------------------------
# Email
# ---------------------------------------------------------------------------

def send_email(subject, body, attachment_path=None):
    sender = os.environ["GMAIL_ADDRESS"]
    password = os.environ["GMAIL_APP_PASSWORD"]
    # Easy to change later: just update the EMAIL_TO secret/env var.
    recipient = os.environ.get("EMAIL_TO", sender)

    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.set_content(body)

    if attachment_path:
        with open(attachment_path, "rb") as f:
            data = f.read()
        msg.add_attachment(
            data,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=Path(attachment_path).name,
        )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(sender, password)
        smtp.send_message(msg)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def run():
    session = build_session()
    rows = scrape_all_pages(session)

    # Dedupe by tender_id (pagination boundaries could in theory repeat a row).
    qualifying = {}
    for r in rows:
        if r["value_number"] is not None and r["value_number"] > VALUE_THRESHOLD_CZK:
            qualifying[r["tender_id"]] = r

    state = load_state()
    new_items = [r for tid, r in qualifying.items() if tid not in state]

    for item in new_items:
        polite_sleep()
        item["summary"] = fetch_summary(item["detail_url"], session)

    now = datetime.now(timezone.utc)
    now_iso = now.isoformat()
    found_at_display = now.strftime("%Y-%m-%d %H:%M UTC")

    for tid, r in qualifying.items():
        if tid not in state:
            state[tid] = {"first_seen_at": now_iso, "name": r["name"]}
    save_state(state)

    if new_items:
        output_path = Path(__file__).parent / "new_tenders.xlsx"
        build_excel(new_items, found_at_display, output_path)
        subject = f"Poptavej.cz monitor: {len(new_items)} new public tender(s) found (>10M CZK)"
        lines = [f"- {r['name']} ({r['tender_id']}): {r['value_number']:,} Kc".replace(",", " ") for r in new_items]
        body = (
            f"Found {len(new_items)} new public tender(s) above 10,000,000 CZK "
            f"on the monitored poptavej.cz search page.\n\n"
            + "\n".join(lines)
            + "\n\nFull details (including summary/description) are in the attached Excel file."
        )
        send_email(subject, body, attachment_path=str(output_path))
    else:
        send_email(
            "Poptavej.cz monitor: no new tenders found",
            f"Run completed at {found_at_display}.\n"
            f"No new public tenders above 10,000,000 CZK were found.\n"
            f"Currently tracking {len(state)} tender(s) in total in the seen-state file.\n\n"
            "(This email confirms the job ran; it is not an error.)",
        )


def main():
    try:
        run()
    except Exception as e:
        error_text = f"{type(e).__name__}: {e}\n\n{traceback.format_exc()}"
        print(error_text, file=sys.stderr)
        try:
            send_email(
                "Poptavej.cz monitor FAILED",
                "The scrape/monitor run failed with an error and did not "
                "complete normally. No data email was sent for this run.\n\n"
                f"Error details:\n\n{error_text}",
            )
        except Exception as email_err:
            print(f"Additionally failed to send the failure alert email: {email_err}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
