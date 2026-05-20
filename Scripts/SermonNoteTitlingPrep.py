# -*- coding: utf-8 -*-
"""
Pull the most recent "Worship Preview" email from the configured inbox and
create a OneNote page summarising the upcoming Sunday sermon.

Output page title (all three lines packed into OneNote's title slot):
    mm/dd/YYYY
    <Scripture book chapter:verses>
    <Sermon Title>

Auth: delegated device-code flow against personal Microsoft accounts.
First local run will print a microsoft.com/devicelogin code; subsequent
runs reuse the cached refresh token at .msal_cache.json (gitignored).

Local invocation:
    python Scripts/SermonNoteTitlingPrep.py            # real email + post
    python Scripts/SermonNoteTitlingPrep.py --test     # canned data, no IMAP
    python Scripts/SermonNoteTitlingPrep.py --dry-run  # parse only, no post

GitHub Actions invocation runs the no-flag form and reads secrets from env.
For CI, pre-seed .msal_cache.json by base64-encoding the local cache after a
successful interactive run and writing it from a GH secret before the script
runs.
"""

import argparse
import email
import imaplib
import os
import sys
from datetime import date, datetime, timedelta
from email.header import decode_header, make_header
from email.utils import mktime_tz, parsedate_tz
from pathlib import Path

import requests
from msal import PublicClientApplication, SerializableTokenCache

try:
    from dotenv import load_dotenv

    load_dotenv()
except ImportError:
    pass  # CI provides env vars directly; .env loader is local-only convenience


SUBJECT_PREFIX = "Worship Preview:"
SCOPES = ["Notes.ReadWrite"]
CACHE_PATH = Path(__file__).resolve().parent.parent / ".msal_cache.json"

CANNED_TEST_DATA = (
    date.today(),
    "Test 1:1-5",
    "Test page from SermonNoteTitlingPrep --test",
)


def _env(name: str) -> str:
    value = os.environ.get(name)
    if not value:
        sys.exit(f"Missing required environment variable: {name}")
    return value


def _decode(header_value: str) -> str:
    return str(make_header(decode_header(header_value or "")))


def parse_subject(subject: str) -> tuple[str, str]:
    """'Worship Preview: 1 John 3:1-10 | Children of God' -> ('1 John 3:1-10', 'Children of God')."""
    if not subject.startswith(SUBJECT_PREFIX):
        raise ValueError(f"Subject is not a worship preview: {subject!r}")
    body = subject[len(SUBJECT_PREFIX):].strip()
    passage, title = body.split("|", 1)
    return passage.strip(), title.strip()


def fetch_latest_worship_preview(
    server: str, port: int, address: str, password: str, sender: str
):
    """Return (sermon_date, passage, title) for the most recent matching email,
    or None if no matching email exists. The email arrives Saturday; the
    sermon is the following Sunday, so we add one day to the email date.
    """
    with imaplib.IMAP4_SSL(server, port) as imap:
        imap.login(address, password)
        imap.select("INBOX")
        typ, data = imap.search(
            None, "FROM", f'"{sender}"', "SUBJECT", f'"{SUBJECT_PREFIX}"'
        )
        if typ != "OK" or not data or not data[0]:
            return None
        latest_uid = data[0].split()[-1]
        _, msg_data = imap.fetch(latest_uid, "(RFC822)")
        message = email.message_from_bytes(msg_data[0][1])

    subject = _decode(message["Subject"])
    passage, title = parse_subject(subject)

    date_tuple = parsedate_tz(message["Date"])
    email_dt = datetime.fromtimestamp(mktime_tz(date_tuple))
    sermon_date = email_dt.date() + timedelta(days=1)
    return sermon_date, passage, title


def build_page_title(sermon_date, passage: str, title: str) -> str:
    """Pack date / passage / title into a single page-title string (newline-
    separated). OneNote preserves the newlines so the three values render as
    stacked lines in the page-title slot rather than in the body."""
    return f"{sermon_date:%m/%d/%Y}\n{passage}\n{title}"


def build_page_html(page_title: str) -> str:
    return (
        "<!DOCTYPE html>"
        "<html>"
        f"<head><title>{page_title}</title></head>"
        "<body></body>"
        "</html>"
    )


def acquire_access_token(client_id: str, tenant_id: str) -> str:
    """Delegated auth via device code flow. Caches the refresh token at
    CACHE_PATH so subsequent runs are silent until the refresh token expires.
    """
    cache = SerializableTokenCache()
    if CACHE_PATH.exists():
        cache.deserialize(CACHE_PATH.read_text())

    app = PublicClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        token_cache=cache,
    )

    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            sys.exit(f"Device flow init failed: {flow}")
        print(flow["message"], flush=True)
        result = app.acquire_token_by_device_flow(flow)

    if cache.has_state_changed:
        CACHE_PATH.write_text(cache.serialize())

    access_token = result.get("access_token")
    if not access_token:
        sys.exit(f"Auth failed: {result.get('error_description', result)}")
    return access_token


def post_to_onenote(html: str, access_token: str, section_id: str) -> str:
    """POST the page and return the new page id."""
    response = requests.post(
        f"https://graph.microsoft.com/v1.0/me/onenote/sections/{section_id}/pages",
        headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/xhtml+xml",
        },
        data=html.encode("utf-8"),
        timeout=30,
    )
    response.raise_for_status()
    return response.json().get("id", "")


def verify_page_created(
    access_token: str, section_id: str, expected_title: str
) -> bool:
    """Query the section for its most recent page and confirm its title
    matches what we just posted. Returns True on match."""
    response = requests.get(
        f"https://graph.microsoft.com/v1.0/me/onenote/sections/{section_id}/pages",
        headers={"Authorization": f"Bearer {access_token}"},
        params={"$top": 1, "$orderby": "createdDateTime desc"},
        timeout=30,
    )
    response.raise_for_status()
    pages = response.json().get("value", [])
    if not pages:
        return False
    # OneNote may collapse newlines in the returned title; compare on the
    # whitespace-normalised form so cosmetic differences don't fail the check.
    actual = " ".join(pages[0].get("title", "").split())
    expected = " ".join(expected_title.split())
    return actual == expected


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[1])
    parser.add_argument(
        "--test",
        action="store_true",
        help="Skip IMAP; post a canned page so you can verify OneNote auth end-to-end.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Parse the email but skip the OneNote post.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if args.test:
        preview = CANNED_TEST_DATA
        print("Running in --test mode with canned data.")
    else:
        preview = fetch_latest_worship_preview(
            server=_env("EMAIL_SERVER"),
            port=int(_env("EMAIL_PORT")),
            address=_env("RECEIVING_ADDRESS"),
            password=_env("EMAIL_PASSWORD"),
            sender=_env("INCOMING_EMAIL"),
        )
        if preview is None:
            print("No worship preview email found.")
            return

    sermon_date, passage, title = preview
    print(f"Preview: {sermon_date:%m/%d/%Y} | {passage} | {title}")

    if args.dry_run:
        print("--dry-run set; skipping OneNote post.")
        return

    page_title = build_page_title(sermon_date, passage, title)
    html = build_page_html(page_title)

    section_id = _env("SECTION_ID")
    access_token = acquire_access_token(
        client_id=_env("CLIENT_ID"),
        tenant_id=_env("TENANT_ID"),
    )
    page_id = post_to_onenote(html, access_token, section_id)
    print(f"OneNote page created: id={page_id or '<unknown>'}")

    if verify_page_created(access_token, section_id, page_title):
        print("Verification: most recent page in section matches expected title.")
    else:
        sys.exit("Verification failed: most recent page does not match the posted title.")


if __name__ == "__main__":
    main()
