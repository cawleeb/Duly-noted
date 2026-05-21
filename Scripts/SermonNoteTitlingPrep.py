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
    python Scripts/SermonNoteTitlingPrep.py                 # real email + post (same as cron)
    python Scripts/SermonNoteTitlingPrep.py --test          # alias for manual on-demand runs
    python Scripts/SermonNoteTitlingPrep.py --dry-run       # scrape + parse, skip OneNote post
    python Scripts/SermonNoteTitlingPrep.py --list-sections # auth and list every section the account can see

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
from datetime import date
from email.header import decode_header, make_header
from pathlib import Path

import requests
from msal import PublicClientApplication, SerializableTokenCache

try:
    from dotenv import load_dotenv

    load_dotenv()
except ImportError:
    pass  # CI provides env vars directly; .env loader is local-only convenience


PREVIEW_MARKER = "Worship Preview:"
SCOPES = ["Notes.ReadWrite"]
CACHE_PATH = Path(__file__).resolve().parent.parent / ".msal_cache.json"


def _env(name: str) -> str:
    value = os.environ.get(name)
    if not value:
        sys.exit(f"Missing required environment variable: {name}")
    return value


def _decode(header_value: str) -> str:
    return str(make_header(decode_header(header_value or "")))


def parse_preview_line(line: str) -> tuple[str, str]:
    """'Worship Preview: 1 John 3:1-10 | Children of God' -> ('1 John 3:1-10', 'Children of God').
    Tolerant of any leading text before 'Worship Preview:' (case-insensitive)."""
    lower = line.lower()
    idx = lower.find(PREVIEW_MARKER.lower())
    if idx < 0 or "|" not in line[idx:]:
        raise ValueError(f"Line is not a worship preview entry: {line!r}")
    body = line[idx + len(PREVIEW_MARKER):].strip()
    passage, title = body.split("|", 1)
    return passage.strip(), title.strip()


def _extract_body_text(message) -> str:
    """Pull plain-text body parts out of an email message."""
    parts = []
    for part in (message.walk() if message.is_multipart() else [message]):
        if part.get_content_type() != "text/plain":
            continue
        payload = part.get_payload(decode=True)
        if payload:
            charset = part.get_content_charset() or "utf-8"
            parts.append(payload.decode(charset, errors="replace"))
    return "\n".join(parts)


def find_preview_line(subject: str, body_text: str) -> str | None:
    """Locate the 'Worship Preview: ... | ...' line in either subject or body."""
    marker = PREVIEW_MARKER.lower()
    candidates = [subject] + body_text.splitlines()
    for line in candidates:
        if marker in line.lower() and "|" in line.split(":", 1)[-1]:
            return line.strip()
    return None


def _extract_email_address(sender: str) -> str:
    """'Gospel Grace Church <info@gospelgrace.com>' -> 'info@gospelgrace.com'.
    Returns the input unchanged if no angle-bracket form is present."""
    if "<" in sender and ">" in sender:
        return sender.split("<", 1)[1].split(">", 1)[0].strip()
    return sender.strip()


def fetch_latest_worship_preview(
    server: str, port: int, address: str, password: str, sender: str,
    mailbox: str = "INBOX",
):
    """Return (sermon_date, passage, title) for the most recent matching email,
    or None if no matching email exists. sermon_date is the runtime date —
    the production cron fires Sunday 04:00 UTC (= Saturday evening Mountain),
    so the runner's date.today() resolves to Sunday at execution time.

    mailbox defaults to INBOX but can be set (via IMAP_MAILBOX env var) to
    "[Gmail]/All Mail" so archived/filtered messages are still found."""
    sender_address = _extract_email_address(sender)
    with imaplib.IMAP4_SSL(server, port) as imap:
        imap.login(address, password)
        typ, _ = imap.select(f'"{mailbox}"', readonly=True)
        if typ != "OK":
            sys.exit(f"Could not open mailbox {mailbox!r}.")
        typ, data = imap.search(
            None, "FROM", f'"{sender_address}"', "TEXT", f'"{PREVIEW_MARKER}"'
        )
        if typ != "OK":
            sys.exit(f"IMAP search failed in {mailbox!r}.")
        uids = data[0].split() if data and data[0] else []
        print(f"IMAP search [{mailbox}] FROM={sender_address!r} TEXT={PREVIEW_MARKER!r}: {len(uids)} match(es)")

        for uid in reversed(uids):
            _, msg_data = imap.fetch(uid, "(RFC822)")
            message = email.message_from_bytes(msg_data[0][1])
            subject = _decode(message["Subject"])
            body_text = _extract_body_text(message)
            line = find_preview_line(subject, body_text)
            if line is None:
                print(f"  uid={uid.decode()}: no parseable preview line; trying next.")
                continue
            try:
                passage, title = parse_preview_line(line)
            except ValueError as exc:
                print(f"  uid={uid.decode()}: {exc}; trying next.")
                continue
            print(f"Found preview line in uid={uid.decode()}: {line!r}")
            return date.today(), passage, title

    return None


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


def list_sections(access_token: str) -> None:
    """Print the signed-in account and every notebook/section it can see.
    Use this to confirm which account the cached token belongs to and to
    copy the correct SECTION_ID into .env."""
    me = requests.get(
        "https://graph.microsoft.com/v1.0/me",
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=30,
    ).json()
    print(f"Signed in as: {me.get('userPrincipalName') or me.get('mail') or me}")

    resp = requests.get(
        "https://graph.microsoft.com/v1.0/me/onenote/sections",
        headers={"Authorization": f"Bearer {access_token}"},
        params={"$select": "id,displayName,parentNotebook", "$expand": "parentNotebook($select=displayName)"},
        timeout=30,
    )
    resp.raise_for_status()
    sections = resp.json().get("value", [])
    if not sections:
        print("(no sections found for this account)")
        return
    for s in sections:
        notebook = (s.get("parentNotebook") or {}).get("displayName", "?")
        print(f"  [{notebook}] {s['displayName']}\n      id={s['id']}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[1])
    parser.add_argument(
        "--test",
        action="store_true",
        help="Manual on-demand run: scrape the most recent worship email, format, and post (same pipeline as the scheduled cron). Use for testing outside of the Saturday schedule.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Parse the email but skip the OneNote post.",
    )
    parser.add_argument(
        "--list-sections",
        action="store_true",
        help="Auth, print the signed-in account and all accessible OneNote sections, then exit.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if args.list_sections:
        access_token = acquire_access_token(
            client_id=_env("CLIENT_ID"),
            tenant_id=_env("TENANT_ID"),
        )
        list_sections(access_token)
        return

    if args.test:
        print("Running in --test mode (manual on-demand, real email).")
    preview = fetch_latest_worship_preview(
        server=_env("EMAIL_SERVER"),
        port=int(_env("EMAIL_PORT")),
        address=_env("RECEIVING_ADDRESS"),
        password=_env("EMAIL_PASSWORD"),
        sender=_env("INCOMING_EMAIL"),
        mailbox=os.environ.get("IMAP_MAILBOX", "INBOX"),
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
