#!/usr/bin/env python3
"""
ms-outlook.py
==============

A single-file MCP (Model Context Protocol) server giving an LLM read-only
access to a locally installed *classic* Microsoft Outlook client (mail +
calendar) on Windows, via COM automation.

Designed for an airgapped Windows endpoint where Outlook is installed, running,
and logged into an on-premises Exchange profile. This script makes NO network
calls; all access is local COM to the already-authenticated Outlook process.

Transport: newline-delimited JSON-RPC 2.0 over stdio (standard MCP stdio
transport, which the VSCode Continue extension speaks).

DEPENDENCY
----------
Requires pywin32 (win32com.client + pythoncom) - the ONLY non-stdlib
dependency. Install via your pip proxy:  pip install pywin32

REQUIREMENTS
------------
- Classic Win32 Outlook (NOT "New Outlook", which has no COM support).
- Outlook installed, running, and logged into a profile.

TOOLS EXPOSED (all read-only)
-----------------------------
- outlook_list_recent_emails : recent Inbox messages
- outlook_search_emails      : search Inbox by subject / sender
- outlook_get_email          : full body of one message by EntryID
- outlook_get_calendar       : calendar events in a date range (recurring expanded)
- outlook_list_sent_emails   : messages you SENT, in a date range (e.g. "what did I do last week")
- outlook_search_recent      : all mail across Inbox/Archive/Sent in a date range (configurable)
- outlook_list_folders       : list every mail folder across all stores (to configure the above)

==============================================================================
CONFIGURATION  -  all editable settings live in the "USER CONFIGURATION" block
just below this docstring. Edit them there; nothing else needs changing.
==============================================================================

1. BLACKLIST_TERMS  (content compliance filter)
   Any email or calendar item whose content contains a blacklisted term is
   WITHHELD - never sent to the AI. Use it for classified / protectively-marked
   material that may not lawfully be processed by the AI.
   - list / search : blocked items are omitted; a count of withheld items is
                     shown (no subject/sender of a blocked item is ever revealed).
   - get_email     : a blocked message returns a generic refusal, not content.
   - calendar      : blocked events are omitted; a withheld count is shown.
   - The matched term is NEVER shown to the AI (that would leak the marking); it
     is logged to STDERR only, for local audit.
   - FAIL-SAFE: if an item's subject or body cannot be read (so it cannot be
     cleared), the item is treated as BLOCKED.
   Seed the list with PLAIN words (PROTECTED, SECRET, CABINET). In "word" mode
   these are still caught inside bracketed markings like [SEC=PROTECTED]. Do NOT
   add very common words (e.g. "OFFICIAL") unless you mean to block almost
   everything. Replace the seeded examples with your real classification scheme.

2. BLACKLIST_MATCH_MODE
   - "word"      : whole-term match. "SECRET" is caught inside "[SEC=SECRET]"
                   but NOT inside "secretary". Best for plain words. (default)
   - "substring" : matches anywhere. Use only when a term itself contains
                   punctuation (e.g. "[SEC=PROTECTED]") that word mode misses.
   Matching is always case-insensitive.

3. RESTRICT_DATE_FORMAT  (LOCALE-SENSITIVE - read this if calendar returns nothing)
   Outlook's Restrict() date filter formats dates per the machine's regional
   settings. Default is US (MM/DD/YYYY). If outlook_get_calendar returns ZERO
   events on an Australian-locale machine, switch to a day-first format - the
   alternatives are listed inline in the config block.

4. Tunable caps (MAX_BODY_CHARS / CALENDAR_HARD_CAP / SEARCH_SCAN_CAP)
   Safety/size limits. Lower MAX_BODY_CHARS if your local model has a small
   context window. The other two are guard rails you can usually leave as-is.

5. SEARCH_ALL_FOLDERS  (which folders the combined outlook_search_recent covers)
   A list of folder NAMES matched across every store in the profile (main
   mailbox, online archive, mounted PST). Default: Inbox, Sent Items, Archive.
   "Archive" is ambiguous - it may be a mailbox folder, an online archive, or a
   PST - so run outlook_list_folders first to see what actually exists and set
   this list to the real names. Note: matching is by name across ALL stores, so
   if you have shared mailboxes with same-named folders they may be included;
   each result is labelled with its store/folder so you can see the source.

EXTERNAL BLACKLIST FILE (optional)
----------------------------------
Instead of (or in addition to) editing BLACKLIST_TERMS, supply extra terms in a
file via --blacklist-file. One term per line; '#' starts a comment. File terms
are ADDED to the built-in list (they never reduce it). Example file contents:

    # outlook-blacklist.txt  - classification terms to withhold from the AI
    PROTECTED
    SECRET
    TOP SECRET
    CABINET
    CABINET-IN-CONFIDENCE

CONTINUE config.yaml ENTRY (copy/paste, adjust paths)
-----------------------------------------------------
    mcpServers:
      - name: outlook
        command: python
        args:
          - C:\\path\\to\\outlook_mcp.py
          - --blacklist-file          # optional
          - C:\\config\\outlook-blacklist.txt
        env:
          PYTHONUTF8: "1"

USAGE / TESTING
---------------
- As an MCP server (normal mode): launched by the MCP client. Run with no
  arguments (optionally --blacklist-file).
- Connectivity check (run manually on the endpoint before wiring it in):

      python outlook_mcp.py --check

  Connects to Outlook and prints mailbox diagnostics + blacklist status to
  stderr, then exits.

IMPORTANT (stdio-on-Windows pitfalls)
-------------------------------------
- ALL diagnostic output goes to stderr. Anything on stdout that is not a
  JSON-RPC message corrupts the protocol stream.
- Set PYTHONUTF8=1 in the launching environment (see config.yaml above) so
  stdout is UTF-8 and Unicode subjects do not crash on cp1252.
"""

import os
import re
import sys
import json
import argparse
import datetime
import traceback


# ============================================================================
# USER CONFIGURATION  -  EDIT THIS BLOCK  (see the docstring above for detail)
# ============================================================================

# --- 1. Content blacklist: terms that cause an item to be withheld from the AI.
#        Seeded with example AU-style markings; REPLACE with your real scheme.
BLACKLIST_TERMS = [
    "PROTECTED",
    "SECRET",
    "TOP SECRET",
    "CABINET",
    "CABINET-IN-CONFIDENCE",
]

# --- 2. How blacklist terms are matched: "word" (default) or "substring".
BLACKLIST_MATCH_MODE = "word"

# --- 3. Calendar date filter format (LOCALE-SENSITIVE). If outlook_get_calendar
#        returns nothing on an AU-locale machine, switch to one of the alternatives.
RESTRICT_DATE_FORMAT = "%m/%d/%Y %I:%M %p"          # US:  MM/DD/YYYY hh:mm AM/PM (default)
# RESTRICT_DATE_FORMAT = "%d/%m/%Y %I:%M %p"        # AU:  DD/MM/YYYY hh:mm AM/PM   <- try this
# RESTRICT_DATE_FORMAT = "%d/%m/%Y %H:%M"           # AU 24-hour, no AM/PM          <- or this

# --- 4. Tunable caps.
MAX_BODY_CHARS = 20000      # truncate very long email bodies for the LLM
CALENDAR_HARD_CAP = 1000    # ceiling on calendar items iterated (anti-runaway)
SEARCH_SCAN_CAP = 500       # ceiling on raw search hits scanned

# --- 5. Folders included in the combined outlook_search_recent tool. Matched by
#        folder NAME (case-insensitive) across EVERY mailbox/store in the profile,
#        including an online (In-Place) archive or a mounted archive.pst. Run the
#        outlook_list_folders tool to see the exact names available, then edit this
#        list to match your setup. (The separate outlook_list_sent_emails tool is
#        unaffected by this list.)
SEARCH_ALL_FOLDERS = ["Inbox", "Sent Items", "Archive"]

# ============================================================================
# END USER CONFIGURATION  -  you should not need to edit below this line
# ============================================================================


# ---------------------------------------------------------------------------
# Stream setup: force UTF-8 so non-ASCII subjects/bodies cannot crash output.
# ---------------------------------------------------------------------------
for _stream in ("stdin", "stdout"):
    try:
        getattr(sys, _stream).reconfigure(encoding="utf-8")
    except Exception:
        pass


def log(message):
    """Write a diagnostic line to stderr ONLY. Never touch stdout here."""
    print(message, file=sys.stderr, flush=True)


# ---------------------------------------------------------------------------
# pywin32 import. If it is missing the server cannot function, so fail loudly.
# ---------------------------------------------------------------------------
try:
    import pythoncom
    import win32com.client
except ImportError:
    log("FATAL: pywin32 is not installed. Run:  pip install pywin32")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Internal constants (not user configuration)
# ---------------------------------------------------------------------------
OL_FOLDER_INBOX = 6
OL_FOLDER_SENT = 5
OL_FOLDER_CALENDAR = 9
OL_CLASS_MAIL = 43  # olMail

# Ceiling on Sent Items scanned when listing by date. Items are sorted
# newest-first and iteration stops as soon as it passes the start of the
# window, so a recent window is reached immediately; this only bounds
# pathological queries for very old windows in a huge Sent folder.
SENT_SCAN_CAP = 5000

# PR_TRANSPORT_MESSAGE_HEADERS (Unicode) - full internet headers, which carry
# the authoritative protective marking (e.g. X-Protective-Marking).
PROP_TRANSPORT_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"

# Compiled at startup by build_blacklist(); None means no filtering is active.
_BLACKLIST_RE = None


# ---------------------------------------------------------------------------
# Blacklist construction and scanning
# ---------------------------------------------------------------------------

def build_blacklist(extra_terms=None):
    """
    Compile BLACKLIST_TERMS (plus any extra_terms) into a single regex and set
    the module-level _BLACKLIST_RE. Logs the active status to stderr.
    """
    global _BLACKLIST_RE

    terms = list(BLACKLIST_TERMS)
    if extra_terms:
        terms.extend(extra_terms)

    # Clean, de-duplicate (case-insensitively), drop blanks.
    cleaned = []
    seen = set()
    for term in terms:
        term = (term or "").strip()
        key = term.lower()
        if term and key not in seen:
            seen.add(key)
            cleaned.append(term)

    if not cleaned:
        _BLACKLIST_RE = None
        log("WARNING: content blacklist is EMPTY - NO compliance filtering is active.")
        return

    escaped = [re.escape(term) for term in cleaned]
    if BLACKLIST_MATCH_MODE == "substring":
        pattern = "(?:" + "|".join(escaped) + ")"
    else:
        # \b on each side: whole-term match. Catches "SECRET" inside "[SEC=SECRET]"
        # (brackets/equals are non-word chars) but not "secretary".
        pattern = r"\b(?:" + "|".join(escaped) + r")\b"

    _BLACKLIST_RE = re.compile(pattern, re.IGNORECASE)
    log("Content blacklist ACTIVE: {0} term(s), mode='{1}'.".format(
        len(cleaned), BLACKLIST_MATCH_MODE))


def blacklisted_match(text):
    """Return the matched blacklisted term if `text` contains one, else None."""
    if _BLACKLIST_RE is None or not text:
        return None
    found = _BLACKLIST_RE.search(text)
    return found.group(0) if found else None


def load_blacklist_file(path):
    """Read extra blacklist terms from a file (one per line; '#' = comment)."""
    terms = []
    with open(path, "r", encoding="utf-8", errors="replace") as handle:
        for line in handle:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            terms.append(line)
    return terms


def email_block_reason(item):
    """
    Decide whether a mail item must be withheld from the AI.

    Returns the matched term (str) if BLOCKED, else None. Fail-safe: if Subject
    or Body cannot be read (so the item cannot be cleared), treat as BLOCKED.
    """
    if _BLACKLIST_RE is None:
        return None

    parts = []
    # Essential fields. If these cannot be read, we cannot verify the item is
    # clean, so we block it.
    try:
        parts.append(item.Subject or "")
    except Exception:
        return "<unreadable subject>"
    try:
        parts.append(item.Body or "")
    except Exception:
        return "<unreadable body>"

    # Supplementary fields - best-effort; a read failure just skips the field.
    for getter in (lambda: item.SenderName,
                   lambda: item.To,
                   lambda: item.CC,
                   lambda: item.Categories):
        try:
            value = getter()
            if value:
                parts.append(str(value))
        except Exception:
            pass

    # Authoritative protective marking lives in the transport headers.
    try:
        headers = item.PropertyAccessor.GetProperty(PROP_TRANSPORT_HEADERS)
        if headers:
            parts.append(str(headers))
    except Exception:
        pass

    return blacklisted_match("\n".join(parts))


def appointment_block_reason(item):
    """As email_block_reason, for a calendar appointment. Fail-safe on Subject/Body."""
    if _BLACKLIST_RE is None:
        return None

    parts = []
    try:
        parts.append(item.Subject or "")
    except Exception:
        return "<unreadable subject>"
    try:
        parts.append(item.Body or "")
    except Exception:
        return "<unreadable body>"

    for getter in (lambda: item.Location,
                   lambda: item.Organizer,
                   lambda: item.Categories):
        try:
            value = getter()
            if value:
                parts.append(str(value))
        except Exception:
            pass

    return blacklisted_match("\n".join(parts))


# ---------------------------------------------------------------------------
# Outlook connection (lazy, cached, with reconnect-on-failure)
# ---------------------------------------------------------------------------
_namespace = None


def get_namespace():
    """Return a cached MAPI namespace, connecting to the running Outlook on first use."""
    global _namespace
    if _namespace is None:
        app = win32com.client.Dispatch("Outlook.Application")
        _namespace = app.GetNamespace("MAPI")
    return _namespace


def reset_namespace():
    """Drop the cached namespace so the next call reconnects from scratch."""
    global _namespace
    _namespace = None


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def fmt_dt(value):
    """Format a COM/pywintypes datetime as a readable string; fall back to str()."""
    try:
        return value.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return str(value)


def sender_smtp(item):
    """Best-effort SMTP address of a sender; degrades to '' rather than prompting."""
    try:
        addr = item.SenderEmailAddress or ""
    except Exception:
        return ""
    if addr.upper().startswith("/O="):
        try:
            exch = item.Sender.GetExchangeUser()
            if exch is not None:
                return exch.PrimarySmtpAddress or addr
        except Exception:
            pass
    return addr


def parse_date(value, fallback):
    """Parse a YYYY-MM-DD string into a date; return fallback if value is empty."""
    if not value:
        return fallback
    return datetime.datetime.strptime(value, "%Y-%m-%d").date()


def _com_to_naive(value):
    """Convert a COM/pywintypes datetime to a naive Python datetime, or None."""
    try:
        return datetime.datetime(
            value.year, value.month, value.day,
            value.hour, value.minute, value.second,
        )
    except Exception:
        return None


def item_received(item):
    """The item's ReceivedTime as a naive datetime, or None."""
    try:
        return _com_to_naive(item.ReceivedTime)
    except Exception:
        return None


def item_best_datetime(item):
    """Best available date for an item: ReceivedTime, else SentOn, else None."""
    dt = item_received(item)
    if dt is None:
        try:
            dt = _com_to_naive(item.SentOn)
        except Exception:
            dt = None
    return dt


def _walk_folders(folder, prefix, out, depth=0):
    """Recursively collect (path, folder_object) pairs beneath `folder`."""
    if depth > 20:  # guard against pathological nesting / loops
        return
    try:
        subfolders = folder.Folders
    except Exception:
        return
    for sub in subfolders:
        try:
            name = sub.Name
        except Exception:
            continue
        path = prefix + "/" + name
        out.append((path, sub))
        _walk_folders(sub, path, out, depth + 1)


def all_mail_folders():
    """
    Return a list of (path, folder_object) for every folder across every store
    in the profile (main mailbox, online archive, mounted PSTs).
    """
    ns = get_namespace()
    out = []
    try:
        stores = ns.Stores
    except Exception:
        stores = None
    if stores is None:
        return out
    for store in stores:
        try:
            root = store.GetRootFolder()
            store_name = store.DisplayName
        except Exception:
            continue
        _walk_folders(root, store_name, out)
    return out


def find_target_folders(names_lower):
    """Return (label, folder) for folders whose LEAF name matches names_lower."""
    matched = []
    for path, folder in all_mail_folders():
        leaf = path.rsplit("/", 1)[-1].lower()
        if leaf in names_lower:
            matched.append((path, folder))
    return matched


# ---------------------------------------------------------------------------
# Tool implementations (each returns a human-readable text string)
# ---------------------------------------------------------------------------

def tool_list_recent_emails(args):
    count = int(args.get("count", 10))
    unread_only = bool(args.get("unread_only", False))

    ns = get_namespace()
    inbox = ns.GetDefaultFolder(OL_FOLDER_INBOX)
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)  # True = descending (newest first)

    lines = []
    withheld = 0
    for item in items:
        if len(lines) >= count:
            break
        try:
            if item.Class != OL_CLASS_MAIL:
                continue
            if unread_only and not item.UnRead:
                continue
        except Exception:
            continue

        # Compliance filter: withhold blocked messages entirely.
        reason = email_block_reason(item)
        if reason:
            withheld += 1
            log("Withheld an Inbox message (blacklist match: {0}).".format(reason))
            continue

        try:
            flag = "UNREAD" if item.UnRead else "read"
            lines.append(
                "- [{flag}] {received} | {sender}\n"
                "    Subject : {subject}\n"
                "    EntryID : {eid}".format(
                    flag=flag,
                    received=fmt_dt(item.ReceivedTime),
                    sender=(item.SenderName or "(unknown sender)"),
                    subject=(item.Subject or "(no subject)"),
                    eid=item.EntryID,
                )
            )
        except Exception:
            continue

    note = ""
    if withheld:
        note = "\n\n[{0} message(s) withheld by the content blacklist and not shown.]".format(withheld)

    if not lines:
        if withheld:
            return "No viewable messages. {0} message(s) were withheld by the content blacklist.".format(withheld)
        return "No matching messages found in the Inbox."
    header = "Showing {n} message(s) from the Inbox (newest first):".format(n=len(lines))
    return header + "\n" + "\n".join(lines) + note


def tool_search_emails(args):
    query = (args.get("query") or "").strip()
    if not query:
        return "Error: 'query' is required."
    count = int(args.get("count", 10))

    ns = get_namespace()
    inbox = ns.GetDefaultFolder(OL_FOLDER_INBOX)
    items = inbox.Items

    # DASL @SQL filter using LIKE '%...%'. String filters are NOT locale-sensitive.
    safe = query.replace("'", "''")
    dasl = (
        '@SQL=' +
        '"urn:schemas:httpmail:subject" LIKE \'%' + safe + '%\'' +
        ' OR ' +
        '"urn:schemas:httpmail:fromname" LIKE \'%' + safe + '%\''
    )

    try:
        matches = items.Restrict(dasl)
    except Exception as exc:
        return "Error: Outlook rejected the search filter ({0}).".format(exc)

    # Collect matches (bounded), then sort newest-first in Python.
    collected = []
    scanned = 0
    for item in matches:
        if scanned >= SEARCH_SCAN_CAP:
            break
        scanned += 1
        try:
            if item.Class != OL_CLASS_MAIL:
                continue
            collected.append(item)
        except Exception:
            continue

    def _received(it):
        try:
            return it.ReceivedTime
        except Exception:
            return None
    collected.sort(key=lambda it: (_received(it) is not None, _received(it)), reverse=True)

    lines = []
    withheld = 0
    for item in collected:
        if len(lines) >= count:
            break
        # Compliance filter.
        reason = email_block_reason(item)
        if reason:
            withheld += 1
            log("Withheld a search hit (blacklist match: {0}).".format(reason))
            continue
        try:
            lines.append(
                "- {received} | {sender}\n"
                "    Subject : {subject}\n"
                "    EntryID : {eid}".format(
                    received=fmt_dt(item.ReceivedTime),
                    sender=(item.SenderName or "(unknown sender)"),
                    subject=(item.Subject or "(no subject)"),
                    eid=item.EntryID,
                )
            )
        except Exception:
            continue

    note = ""
    if withheld:
        note = "\n\n[{0} matching message(s) withheld by the content blacklist.]".format(withheld)

    if not lines:
        if withheld:
            return "No viewable matches for '{0}'. {1} match(es) were withheld by the content blacklist.".format(query, withheld)
        return "No Inbox messages matched '{0}' (searched subject and sender name).".format(query)
    header = "Found {n} message(s) matching '{q}' (newest first):".format(n=len(lines), q=query)
    return header + "\n" + "\n".join(lines) + note


def tool_get_email(args):
    entry_id = (args.get("entry_id") or "").strip()
    if not entry_id:
        return "Error: 'entry_id' is required (get it from a list/search result)."

    ns = get_namespace()
    try:
        item = ns.GetItemFromID(entry_id)
    except Exception:
        return "Error: no message found for that EntryID (it may have moved or been deleted)."

    # Compliance filter: refuse blocked messages with a GENERIC message that does
    # not reveal the matched term or any classified content.
    reason = email_block_reason(item)
    if reason:
        log("Refused get_email (blacklist match: {0}) for EntryID {1}.".format(reason, entry_id))
        return ("This message cannot be displayed: its content is withheld under the "
                "content blacklist (classification / compliance policy).")

    try:
        body = item.Body or ""
    except Exception:
        return ("This message cannot be displayed: its body could not be read and "
                "therefore cannot be cleared for display.")
    truncated_note = ""
    if len(body) > MAX_BODY_CHARS:
        body = body[:MAX_BODY_CHARS]
        truncated_note = "\n\n[... body truncated at {0} characters ...]".format(MAX_BODY_CHARS)

    def safe(getter, default=""):
        try:
            return getter() or default
        except Exception:
            return default

    smtp = sender_smtp(item)
    sender_line = safe(lambda: item.SenderName, "(unknown)")
    if smtp:
        sender_line += " <{0}>".format(smtp)

    parts = [
        "Subject : {0}".format(safe(lambda: item.Subject, "(no subject)")),
        "From    : {0}".format(sender_line),
        "To      : {0}".format(safe(lambda: item.To)),
        "CC      : {0}".format(safe(lambda: item.CC)),
        "Received: {0}".format(fmt_dt(safe(lambda: item.ReceivedTime))),
        "",
        body + truncated_note,
    ]
    return "\n".join(parts)


def tool_list_sent_emails(args):
    """List messages from the Sent Items folder within a date range (newest first)."""
    today = datetime.date.today()
    try:
        start_date = parse_date(args.get("start_date"), today - datetime.timedelta(days=7))
        end_date = parse_date(args.get("end_date"), today)
    except ValueError:
        return "Error: dates must be in YYYY-MM-DD format."
    if end_date < start_date:
        return "Error: end_date is before start_date."
    max_results = int(args.get("max_results", 100))

    start_dt = datetime.datetime.combine(start_date, datetime.time(0, 0, 0))
    end_dt = datetime.datetime.combine(end_date, datetime.time(23, 59, 59))

    ns = get_namespace()
    sent = ns.GetDefaultFolder(OL_FOLDER_SENT)
    items = sent.Items
    items.Sort("[SentOn]", True)  # True = descending (newest first)

    lines = []
    withheld = 0
    iterated = 0
    for item in items:
        if len(lines) >= max_results or iterated >= SENT_SCAN_CAP:
            break
        iterated += 1
        try:
            if item.Class != OL_CLASS_MAIL:
                continue
            sent_on = item.SentOn
            # Represent the COM time as a naive datetime for a day-range comparison.
            sent_naive = datetime.datetime(
                sent_on.year, sent_on.month, sent_on.day,
                sent_on.hour, sent_on.minute, sent_on.second,
            )
        except Exception:
            continue

        # Sorted newest-first: skip anything newer than the window; once we reach
        # an item older than the window, every later item is older too, so stop.
        if sent_naive > end_dt:
            continue
        if sent_naive < start_dt:
            break

        # The content blacklist applies to sent mail as well.
        reason = email_block_reason(item)
        if reason:
            withheld += 1
            log("Withheld a sent message (blacklist match: {0}).".format(reason))
            continue

        try:
            lines.append(
                "- {sent} | To: {to}\n"
                "    Subject : {subject}\n"
                "    EntryID : {eid}".format(
                    sent=fmt_dt(item.SentOn),
                    to=(item.To or "(no recipient)"),
                    subject=(item.Subject or "(no subject)"),
                    eid=item.EntryID,
                )
            )
        except Exception:
            continue

    note = ""
    if withheld:
        note = "\n\n[{0} sent message(s) withheld by the content blacklist.]".format(withheld)
    if not lines:
        if withheld:
            return "No viewable sent messages between {0} and {1}. {2} withheld by the content blacklist.".format(
                start_date, end_date, withheld)
        return "No sent messages between {0} and {1}.".format(start_date, end_date)
    header = "Sent messages {0} to {1} ({2} shown, newest first):".format(start_date, end_date, len(lines))
    return header + "\n" + "\n".join(lines) + note


def tool_list_folders(args):
    """List every mail folder across all stores, with item counts, for discovery/config."""
    max_lines = int(args.get("max_results", 300))
    folders = all_mail_folders()
    if not folders:
        return "No folders found (could not enumerate Outlook stores)."
    lines = []
    for path, folder in folders:
        if len(lines) >= max_lines:
            lines.append("... (list truncated; raise max_results to see more)")
            break
        try:
            count = folder.Items.Count
        except Exception:
            count = "?"
        lines.append("- {0}  (items: {1})".format(path, count))
    return "Outlook folders (store/path, with item counts):\n" + "\n".join(lines)


def tool_search_recent(args):
    """List mail across the configured folders (Inbox/Archive/Sent) within a date range."""
    today = datetime.date.today()
    try:
        start_date = parse_date(args.get("start_date"), today - datetime.timedelta(days=7))
        end_date = parse_date(args.get("end_date"), today)
    except ValueError:
        return "Error: dates must be in YYYY-MM-DD format."
    if end_date < start_date:
        return "Error: end_date is before start_date."
    query = (args.get("query") or "").strip().lower()
    max_results = int(args.get("max_results", 100))

    start_dt = datetime.datetime.combine(start_date, datetime.time(0, 0, 0))
    end_dt = datetime.datetime.combine(end_date, datetime.time(23, 59, 59))

    names_lower = set(name.lower() for name in SEARCH_ALL_FOLDERS)
    targets = find_target_folders(names_lower)
    if not targets:
        return ("No folders matched {0}. Run outlook_list_folders to see the folder "
                "names available, then adjust SEARCH_ALL_FOLDERS.".format(SEARCH_ALL_FOLDERS))

    collected = []  # tuples of (naive_datetime, folder_label, item)
    withheld = 0
    for label, folder in targets:
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)  # newest first (see break-early note below)
            sortable = True
        except Exception:
            items = folder.Items
            sortable = False

        scanned = 0
        for item in items:
            if scanned >= SENT_SCAN_CAP:
                break
            scanned += 1
            try:
                if item.Class != OL_CLASS_MAIL:
                    continue
            except Exception:
                continue

            received = item_received(item)
            when = item_best_datetime(item)
            if when is None:
                continue
            if when > end_dt:
                continue
            if when < start_dt:
                # Sorted by ReceivedTime desc: if the SORT KEY is past the window,
                # every later item is older too, so stop. Otherwise just skip (the
                # item's usable date came from a fallback field, so ordering is not
                # guaranteed and we must keep scanning).
                if sortable and received is not None and received < start_dt:
                    break
                continue

            # Optional free-text filter on subject / sender / recipient.
            if query:
                def _low(getter):
                    try:
                        return (getter() or "").lower()
                    except Exception:
                        return ""
                if (query not in _low(lambda: item.Subject)
                        and query not in _low(lambda: item.SenderName)
                        and query not in _low(lambda: item.To)):
                    continue

            # Content blacklist applies here too.
            reason = email_block_reason(item)
            if reason:
                withheld += 1
                log("Withheld an item from {0} (blacklist match: {1}).".format(label, reason))
                continue

            collected.append((when, label, item))

    collected.sort(key=lambda row: row[0], reverse=True)

    lines = []
    for when, label, item in collected[:max_results]:
        try:
            subject = item.Subject or "(no subject)"
        except Exception:
            subject = "(no subject)"
        who = ""
        try:
            who = item.SenderName or ""
        except Exception:
            who = ""
        if not who:  # sent items have no sender name; show recipient instead
            try:
                who = "To: " + (item.To or "")
            except Exception:
                who = ""
        try:
            eid = item.EntryID
        except Exception:
            eid = ""
        lines.append(
            "- {when} | {folder} | {who}\n"
            "    Subject : {subject}\n"
            "    EntryID : {eid}".format(
                when=when.strftime("%Y-%m-%d %H:%M"),
                folder=label,
                who=(who or "(unknown)"),
                subject=subject,
                eid=eid,
            )
        )

    note = ""
    if withheld:
        note += "\n\n[{0} item(s) withheld by the content blacklist.]".format(withheld)
    if len(collected) > max_results:
        note += "\n[Showing first {0} of {1} matches; narrow the dates or add a query.]".format(
            max_results, len(collected))

    if not lines:
        if withheld:
            return "No viewable emails between {0} and {1}. {2} withheld by the content blacklist.".format(
                start_date, end_date, withheld)
        return "No emails between {0} and {1} across {2}.".format(
            start_date, end_date, ", ".join(SEARCH_ALL_FOLDERS))
    header = "Emails {0} to {1} across {2} ({3} shown, newest first):".format(
        start_date, end_date, ", ".join(SEARCH_ALL_FOLDERS), len(lines))
    return header + "\n" + "\n".join(lines) + note


def tool_get_calendar(args):
    today = datetime.date.today()
    try:
        start_date = parse_date(args.get("start_date"), today)
        end_date = parse_date(args.get("end_date"), today + datetime.timedelta(days=7))
    except ValueError:
        return "Error: dates must be in YYYY-MM-DD format."

    if end_date < start_date:
        return "Error: end_date is before start_date."

    max_results = int(args.get("max_results", 50))

    start_dt = datetime.datetime.combine(start_date, datetime.time(0, 0))
    end_dt = datetime.datetime.combine(end_date, datetime.time(23, 59))

    ns = get_namespace()
    cal = ns.GetDefaultFolder(OL_FOLDER_CALENDAR)
    items = cal.Items

    # Order matters for expanding recurring appointments: Sort -> IncludeRecurrences -> Restrict.
    items.Sort("[Start]")
    items.IncludeRecurrences = True

    restriction = (
        "[Start] >= '" + start_dt.strftime(RESTRICT_DATE_FORMAT) + "' AND "
        "[Start] <= '" + end_dt.strftime(RESTRICT_DATE_FORMAT) + "'"
    )
    try:
        restricted = items.Restrict(restriction)
    except Exception as exc:
        return "Error: Outlook rejected the calendar filter ({0}).".format(exc)

    lines = []
    withheld = 0
    iterated = 0
    for item in restricted:
        if iterated >= CALENDAR_HARD_CAP or len(lines) >= max_results:
            break
        iterated += 1

        # Compliance filter.
        reason = appointment_block_reason(item)
        if reason:
            withheld += 1
            log("Withheld a calendar item (blacklist match: {0}).".format(reason))
            continue

        try:
            all_day = bool(item.AllDayEvent)
            when = "{start} -> {end}".format(start=fmt_dt(item.Start), end=fmt_dt(item.End))
            if all_day:
                when = "{0} (all day)".format(item.Start.strftime("%Y-%m-%d"))
            recur = " [recurring]" if bool(item.IsRecurring) else ""
            location = ""
            try:
                location = item.Location or ""
            except Exception:
                pass
            organizer = ""
            try:
                organizer = item.Organizer or ""
            except Exception:
                pass
            lines.append(
                "- {when}{recur}\n"
                "    Subject  : {subject}\n"
                "    Location : {loc}\n"
                "    Organizer: {org}".format(
                    when=when,
                    recur=recur,
                    subject=(item.Subject or "(no subject)"),
                    loc=(location or "-"),
                    org=(organizer or "-"),
                )
            )
        except Exception:
            continue

    note = ""
    if withheld:
        note = "\n\n[{0} event(s) withheld by the content blacklist and not shown.]".format(withheld)

    if not lines:
        if withheld:
            return "No viewable events between {0} and {1}. {2} event(s) were withheld by the content blacklist.".format(
                start_date, end_date, withheld)
        return "No calendar events between {0} and {1}.".format(start_date, end_date)
    header = "Calendar events {0} to {1} ({2} shown):".format(start_date, end_date, len(lines))
    return header + "\n" + "\n".join(lines) + note


# ---------------------------------------------------------------------------
# MCP tool registry
# ---------------------------------------------------------------------------

TOOLS = [
    {
        "name": "outlook_list_recent_emails",
        "description": (
            "List the most recent emails in the Outlook Inbox, newest first. "
            "Returns subject, sender, received time, read/unread status, and an "
            "EntryID for each. Use the EntryID with outlook_get_email to read the "
            "full message body. Some messages may be withheld by a content policy."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "count": {"type": "integer", "description": "How many messages to return (default 10)."},
                "unread_only": {"type": "boolean", "description": "If true, only unread messages (default false)."},
            },
        },
    },
    {
        "name": "outlook_search_emails",
        "description": (
            "Search the Outlook Inbox for messages whose subject OR sender name "
            "contains the query text. Returns subject, sender, received time, and "
            "an EntryID for each match. Some matches may be withheld by a content policy."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "query": {"type": "string", "description": "Text to look for in the subject or sender name."},
                "count": {"type": "integer", "description": "Maximum matches to return, newest first (default 10)."},
            },
            "required": ["query"],
        },
    },
    {
        "name": "outlook_get_email",
        "description": (
            "Retrieve the full details and plain-text body of a single email, "
            "identified by the EntryID from a list or search result. The message "
            "may be withheld if it is blocked by a content policy."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "entry_id": {"type": "string", "description": "The EntryID of the message to read."},
            },
            "required": ["entry_id"],
        },
    },
    {
        "name": "outlook_get_calendar",
        "description": (
            "List calendar events between two dates (inclusive). Recurring meetings "
            "are expanded into individual occurrences. Dates are YYYY-MM-DD; defaults "
            "to today through 7 days ahead. Some events may be withheld by a content policy."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "start_date": {"type": "string", "description": "Start date, YYYY-MM-DD (default: today)."},
                "end_date": {"type": "string", "description": "End date, YYYY-MM-DD (default: 7 days from today)."},
                "max_results": {"type": "integer", "description": "Maximum events to return (default 50)."},
            },
        },
    },
    {
        "name": "outlook_list_sent_emails",
        "description": (
            "List messages you SENT, within a date range (inclusive), newest first. "
            "Returns sent time, recipients, and subject for each, plus an EntryID for "
            "reading the full body with outlook_get_email. Useful for reviewing what "
            "you did over a period, e.g. 'what did I send last week'. Dates are "
            "YYYY-MM-DD; defaults to the last 7 days. Some messages may be withheld by "
            "a content policy."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "start_date": {"type": "string", "description": "Start date, YYYY-MM-DD (default: 7 days ago)."},
                "end_date": {"type": "string", "description": "End date, YYYY-MM-DD (default: today)."},
                "max_results": {"type": "integer", "description": "Maximum messages to return (default 100)."},
            },
        },
    },
    {
        "name": "outlook_search_recent",
        "description": (
            "List emails across MULTIPLE folders at once (by default Inbox, Sent Items, "
            "and Archive) within a date range, newest first, merged into one list. Each "
            "result is labelled with the folder it came from. Optionally filter by text "
            "in the subject, sender, or recipient. Use this for 'all my email over the "
            "last week' style questions. Dates are YYYY-MM-DD; defaults to the last 7 "
            "days. Some items may be withheld by a content policy."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "start_date": {"type": "string", "description": "Start date, YYYY-MM-DD (default: 7 days ago)."},
                "end_date": {"type": "string", "description": "End date, YYYY-MM-DD (default: today)."},
                "query": {"type": "string", "description": "Optional text to match in subject, sender, or recipient."},
                "max_results": {"type": "integer", "description": "Maximum emails to return (default 100)."},
            },
        },
    },
    {
        "name": "outlook_list_folders",
        "description": (
            "List every mail folder across all Outlook stores (main mailbox, online "
            "archive, mounted PSTs), with item counts. Use this to discover the exact "
            "folder names available - especially to see what your 'Archive' really is - "
            "so the combined search can be pointed at the right folders."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "max_results": {"type": "integer", "description": "Maximum folders to list (default 300)."},
            },
        },
    },
]

TOOL_DISPATCH = {
    "outlook_list_recent_emails": tool_list_recent_emails,
    "outlook_search_emails": tool_search_emails,
    "outlook_get_email": tool_get_email,
    "outlook_get_calendar": tool_get_calendar,
    "outlook_list_sent_emails": tool_list_sent_emails,
    "outlook_search_recent": tool_search_recent,
    "outlook_list_folders": tool_list_folders,
}


# ---------------------------------------------------------------------------
# JSON-RPC / MCP plumbing
# ---------------------------------------------------------------------------

PROTOCOL_VERSION_DEFAULT = "2024-11-05"
SERVER_INFO = {"name": "outlook-mcp", "version": "1.3.0"}


def rpc_result(req_id, result):
    return {"jsonrpc": "2.0", "id": req_id, "result": result}


def rpc_error(req_id, code, message):
    return {"jsonrpc": "2.0", "id": req_id, "error": {"code": code, "message": message}}


def text_content(text, is_error=False):
    return {"content": [{"type": "text", "text": text}], "isError": is_error}


def handle_request(req):
    """Process one JSON-RPC request. Return a response dict, or None for notifications."""
    method = req.get("method")
    req_id = req.get("id")
    is_notification = "id" not in req

    if method == "initialize":
        params = req.get("params") or {}
        proto = params.get("protocolVersion", PROTOCOL_VERSION_DEFAULT)
        return rpc_result(req_id, {
            "protocolVersion": proto,
            "capabilities": {"tools": {}},
            "serverInfo": SERVER_INFO,
        })

    if method == "notifications/initialized":
        return None

    if method == "ping":
        return rpc_result(req_id, {})

    if method == "tools/list":
        return rpc_result(req_id, {"tools": TOOLS})

    if method == "tools/call":
        params = req.get("params") or {}
        name = params.get("name")
        arguments = params.get("arguments") or {}
        func = TOOL_DISPATCH.get(name)
        if func is None:
            return rpc_result(req_id, text_content("Unknown tool: {0}".format(name), is_error=True))
        try:
            output = func(arguments)
            return rpc_result(req_id, text_content(output, is_error=False))
        except Exception as exc:
            log("Tool '{0}' failed:\n{1}".format(name, traceback.format_exc()))
            reset_namespace()  # force a fresh Outlook connection next call
            return rpc_result(req_id, text_content("Outlook tool error: {0}".format(exc), is_error=True))

    if is_notification:
        return None
    return rpc_error(req_id, -32601, "Method not found: {0}".format(method))


def run_server():
    """Main stdio loop: read newline-delimited JSON-RPC, dispatch, respond."""
    pythoncom.CoInitialize()
    log("outlook-mcp server started (stdio). Waiting for requests...")
    try:
        for raw_line in sys.stdin:
            line = raw_line.strip()
            if not line:
                continue
            try:
                request = json.loads(line)
            except json.JSONDecodeError:
                log("Ignoring malformed JSON line.")
                continue
            response = handle_request(request)
            if response is not None:
                sys.stdout.write(json.dumps(response) + "\n")
                sys.stdout.flush()
    except KeyboardInterrupt:
        pass
    finally:
        pythoncom.CoUninitialize()
        log("outlook-mcp server stopped.")


def run_check():
    """Connect to Outlook, print diagnostics + blacklist status to stderr, then exit."""
    pythoncom.CoInitialize()
    try:
        ns = get_namespace()
        log("Connected to Outlook MAPI namespace.")
        try:
            log("Current user        : {0}".format(ns.CurrentUser))
        except Exception as exc:
            log("Could not read CurrentUser: {0}".format(exc))

        inbox = ns.GetDefaultFolder(OL_FOLDER_INBOX)
        log("Inbox folder        : {0}".format(inbox.Name))
        try:
            log("Inbox item count    : {0}".format(inbox.Items.Count))
        except Exception as exc:
            log("Could not count Inbox items: {0}".format(exc))

        cal = ns.GetDefaultFolder(OL_FOLDER_CALENDAR)
        log("Calendar folder     : {0}".format(cal.Name))
        log("CHECK OK - Outlook COM link is working.")
        return 0
    except Exception:
        log("CHECK FAILED:\n{0}".format(traceback.format_exc()))
        return 1
    finally:
        pythoncom.CoUninitialize()


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Read-only MCP server exposing local Outlook mail and calendar via COM, "
            "with a content blacklist that withholds classified/marked items. With "
            "no check flag it runs as an stdio MCP server."
        )
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="Connect to Outlook, print diagnostics + blacklist status to stderr, then exit.",
    )
    parser.add_argument(
        "--blacklist-file",
        help="Path to a file of EXTRA blacklist terms (one per line; '#' for comments). "
             "Terms are added to the built-in list.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version="outlook-mcp {0}".format(SERVER_INFO["version"]),
    )
    args = parser.parse_args()

    # Load any external blacklist terms and compile the filter BEFORE serving.
    extra_terms = []
    if args.blacklist_file:
        if not os.path.isfile(args.blacklist_file):
            log("FATAL: --blacklist-file not found: {0}".format(args.blacklist_file))
            sys.exit(2)
        try:
            extra_terms = load_blacklist_file(args.blacklist_file)
        except Exception as exc:
            log("FATAL: could not read --blacklist-file: {0}".format(exc))
            sys.exit(2)
    build_blacklist(extra_terms)

    if args.check:
        sys.exit(run_check())
    run_server()


if __name__ == "__main__":
    main()
