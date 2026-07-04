#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
msword_mcp.py - A single-file MCP (Model Context Protocol) stdio server that
gives an AI agent read/search/edit/generate access to Word .docx files.

It follows a simple open -> edit -> save workflow (msword_open ... msword_save),
with the .docx engine provided by the Python library `python-docx`.

WHAT IT CAN DO
    - Open a .docx into an in-memory session, keyed by session_id.
    - Read full content (linear text, or a structured block list).
    - Search text (body paragraphs and table cells).
    - Replace text, correctly handling matches that span multiple runs.
    - Record replacements/deletions as Word TRACKED CHANGES (revisions) with an
      author and the current date, then list / accept-all / reject-all them.
    - Append paragraphs, headings and tables.
    - Apply simple paragraph formatting (bold/italic/underline/alignment/style).
    - Save in place, or save-as to a new path.

TRACKED CHANGES - SCOPE AND VERIFICATION
    Scope is text-level only: tracked insertions and deletions of text (a
    tracked "replace" is a delete of the old text plus an insert of the new).
    NOT supported, by design (these are the fragile OOXML cases): tracked
    paragraph-mark changes (splitting/merging paragraphs as a revision),
    tracked moves, and tracked formatting-only changes.

    Verified before delivery: correct <w:ins>/<w:del>/<w:delText> structure with
    unique ids + author + date; cross-run matches; accept-all and reject-all
    round-trips; and an independent round-trip through LibreOffice, which parsed
    and preserved the revisions. The ONE thing not verifiable off your network is
    the exact Microsoft Word review UI - open one test file in Word on your
    endpoint and confirm the changes appear under Review and that Accept/Reject
    behave as expected before relying on it for anything important.

WHAT IT CANNOT DO
    - Comment threads (no reliable python-docx API).
    - Tracked paragraph-mark / move / formatting revisions (see scope above).

=============================================================================
 DEPENDENCIES  (this is the ONE place this project leaves the standard library)
=============================================================================
    python-docx   (import name: docx)   -> the .docx engine
      +- lxml            (compiled C extension, pulled in by python-docx)
      +- typing_extensions

    AIRGAPPED WINDOWS INSTALL (sideload wheels, same pattern as pymupdf):
      1. On an internet-connected box, download wheels for your EXACT
         interpreter. lxml is compiled, so the wheel must match Python
         version + architecture, e.g. for CPython 3.11 64-bit:
             lxml-<ver>-cp311-cp311-win_amd64.whl
         python-docx and typing_extensions are pure-Python (any wheel works).
         A reliable way to grab the correct set in one go, run ON the target
         Python version if you can:
             python -m pip download python-docx -d .\wheels
      2. Transfer the .\wheels folder to the airgapped endpoint.
      3. Install with the SAME interpreter Continue will launch (use -m pip so
         the interpreter and pip cannot drift apart):
             "C:\path\to\python.exe" -m pip install --no-index ^
                 --find-links .\wheels python-docx
      4. Confirm the interpreter can see it:
             "C:\path\to\python.exe" -m docx  (no error = importable)

    If your corporate mirror proxies PyPI you may be able to skip the manual
    download and just run:  python -m pip install python-docx
    lxml wheel/interpreter mismatch is the #1 failure here; this server logs
    sys.executable and both versions at startup so a mismatch is obvious.

=============================================================================
 VALIDATE BEFORE WIRING IN  (single-transfer sanity check)
=============================================================================
    Run the built-in self-test. It creates a temp .docx, opens/edits/saves/
    reopens it, and prints PASS/FAIL. No arguments, no network, no side files
    left behind:
        "C:\path\to\python.exe" msword_mcp.py --check

    Expected tail of output on success:
        [check] round-trip: PASS
        [check] ALL CHECKS PASSED

=============================================================================
 WIRE INTO CONTINUE  (config.yaml)
=============================================================================
    Add under mcpServers. Use the SAME python.exe you installed the wheels
    with, and set PYTHONUTF8 so Windows cp1252 cannot corrupt the stdio JSON
    stream. After editing config.yaml, run "Developer: Reload Window" rather
    than toggling the server (avoids the "already connected to transport" bug).

        mcpServers:
          - name: msword-py
            command: C:\path\to\python.exe
            args:
              - C:\path\to\msword_mcp.py
              - --author
              - Matt
            env:
              PYTHONUTF8: "1"

    The --author value is stamped on every tracked change. Omit the two --author
    lines to fall back to the TRACKED_CHANGE_AUTHOR config constant below.

    (If your Continue build uses the older command/args-in-one style, match
    whatever your existing working Python MCP servers use - the launch shape
    is identical to them.)

=============================================================================
 PROTOCOL / TRANSPORT NOTES
=============================================================================
    - MCP stdio transport = newline-delimited JSON-RPC 2.0 on stdin/stdout.
    - stdout is SACRED: only JSON-RPC messages go there, one per line. Every
      diagnostic goes to stderr via log(). Any stray print() to stdout would
      corrupt the stream.
    - Streams are reconfigured to UTF-8 in-script as a belt-and-braces measure
      alongside PYTHONUTF8.

Author's assumptions (flagged per the airgap "a caveat is cheaper than a
failed transfer" rule):
    - A "reasonably modern" python-docx (>= 0.8.11; tested here on 1.2.0).
    - Editing scope for search/replace is the document BODY paragraphs and
      TABLE cells. Headers/footers are intentionally NOT edited (predictable,
      and they rarely carry the target content). Say the word if you need them.
    - When a replaced match spans runs with different formatting, the inserted
      text takes the formatting of the run where the match STARTS. This is the
      standard, unavoidable trade-off for run-aware replacement.
"""

# =============================================================================
# CONFIGURATION  (all user-editable settings live here, nothing scattered below)
# =============================================================================
SERVER_NAME = "msword-py"          # advertised to the MCP client
SERVER_VERSION = "1.0.0"
PROTOCOL_VERSION_FALLBACK = "2024-11-05"  # used if the client sends none

# Optional path sandbox. If set to a directory string, the server will refuse
# to open or save any file outside that directory tree (belt-and-braces for a
# compliance-sensitive environment). Leave as None to allow any path.
#   e.g. DOCUMENT_ROOT = r"C:\Users\you\Documents\ai_docs"
DOCUMENT_ROOT = None

MAX_SESSIONS = 32                    # guard against runaway open() calls
SEARCH_CONTEXT_CHARS = 40            # chars of context either side of a match

# Default author name stamped on tracked changes (w:author). Override per
# launch with:  --author "Matt"  in Continue's args: block. The date on each
# change is always the current date, computed at edit time.
TRACKED_CHANGE_AUTHOR = "AI Assistant (Continue)"
# =============================================================================

import sys
import os
import json
import uuid
import copy
import argparse
import traceback
import datetime

# --- Make stdio UTF-8 regardless of the Windows console codepage. -----------
# Belt-and-braces alongside PYTHONUTF8=1 in the Continue env: block.
for _stream in ("stdin", "stdout", "stderr"):
    try:
        getattr(sys, _stream).reconfigure(encoding="utf-8")
    except Exception:
        # Older/edge interpreters may not support reconfigure; PYTHONUTF8 covers it.
        pass


def log(msg):
    """All diagnostics go to stderr ONLY. stdout is reserved for JSON-RPC."""
    try:
        sys.stderr.write("[{}] {}\n".format(SERVER_NAME, msg))
        sys.stderr.flush()
    except Exception:
        pass


# --- Import the engine, failing loudly and specifically on mismatch. --------
try:
    import docx  # python-docx
    from docx.document import Document as _DocumentClass
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl
    from docx.table import Table, _Cell
    from docx.text.paragraph import Paragraph
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.opc.exceptions import PackageNotFoundError
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
except Exception as _imp_err:  # pragma: no cover - exercised only on a broken install
    sys.stderr.write(
        "[{}] FATAL: could not import python-docx.\n"
        "        Interpreter : {}\n"
        "        Error       : {}\n"
        "        Fix: install python-docx into THIS interpreter "
        "(see the docstring's airgapped install steps).\n".format(
            SERVER_NAME, sys.executable, _imp_err
        )
    )
    sys.exit(1)


def _versions():
    """Best-effort version strings for startup diagnostics."""
    try:
        from importlib.metadata import version
        docx_ver = version("python-docx")
    except Exception:
        docx_ver = "unknown"
    try:
        from lxml import etree
        lxml_ver = etree.__version__
    except Exception:
        lxml_ver = "unknown"
    return docx_ver, lxml_ver


# =============================================================================
# SESSION STATE
# =============================================================================
# session_id -> {"path": str, "doc": docx Document, "opened_at": iso str}
SESSIONS = {}

# Author stamped on tracked changes. Seeded from config, optionally replaced
# by the --author launch flag in main().
AUTHOR = TRACKED_CHANGE_AUTHOR

# XML namespace literal for the reserved 'xml' prefix (qn() does not map it).
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


class ToolError(Exception):
    """Raised by a tool handler to return a clean isError result to the client."""


def _now_iso():
    return datetime.datetime.now().isoformat(timespec="seconds")


def _require(args, key):
    """Fetch a required argument or raise a clear ToolError."""
    if key not in args or args[key] is None:
        raise ToolError("Missing required argument: '{}'".format(key))
    return args[key]


def _resolve_path(path):
    """Expand + absolutise a path, and enforce DOCUMENT_ROOT if configured."""
    if not isinstance(path, str) or not path.strip():
        raise ToolError("Path must be a non-empty string")
    rp = os.path.abspath(os.path.expanduser(path))
    if DOCUMENT_ROOT:
        root = os.path.abspath(os.path.expanduser(DOCUMENT_ROOT))
        try:
            common = os.path.commonpath([root, rp])
        except ValueError:
            # Different drives on Windows raise ValueError from commonpath.
            common = None
        if common != root:
            raise ToolError(
                "Path is outside the permitted DOCUMENT_ROOT and was refused."
            )
    return rp


def _get_session(args):
    """Return the session dict for a required session_id argument."""
    sid = _require(args, "session_id")
    session = SESSIONS.get(sid)
    if session is None:
        raise ToolError(
            "Unknown session_id '{}'. Call msword_open first.".format(sid)
        )
    return session


# =============================================================================
# DOCUMENT HELPERS
# =============================================================================
def _iter_block_items(parent):
    """
    Yield Paragraph and Table objects in true document order.

    python-docx exposes document.paragraphs and document.tables as SEPARATE
    flat lists, losing their interleaving. This walks the underlying XML body
    so blocks come back in the order they actually appear.
    """
    if isinstance(parent, _DocumentClass):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ToolError("Unsupported parent for block iteration")
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def _all_editable_paragraphs(doc):
    """
    Every paragraph we are willing to search/replace within: body paragraphs
    plus paragraphs inside every table cell (recursively for nested tables).
    Order is not guaranteed to be strict document order, which is fine because
    replace only needs to visit each paragraph once.
    """
    def _cell_paragraphs(cell):
        for p in cell.paragraphs:
            yield p
        for tbl in cell.tables:  # nested tables
            for row in tbl.rows:
                for c in row.cells:
                    for p in _cell_paragraphs(c):
                        yield p

    for p in doc.paragraphs:
        yield p
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in _cell_paragraphs(cell):
                    yield p


def _replace_in_paragraph(paragraph, find, replace, remaining):
    """
    Replace occurrences of `find` with `replace` inside a single paragraph,
    correctly handling matches that span multiple runs.

    `remaining` caps how many replacements are still allowed (None = no cap).
    Returns the number of replacements made in this paragraph.

    Method: on each pass, concatenate all run texts, locate the first match in
    the combined text, then rewrite each run so the matched characters are
    removed from the runs they fell in and the replacement text is inserted
    into the run where the match STARTS. That start run's formatting is what
    the replacement inherits (the documented trade-off for cross-run matches).
    """
    made = 0
    while True:
        if remaining is not None and made >= remaining:
            break
        runs = paragraph.runs
        if not runs:
            break
        texts = [r.text for r in runs]
        full = "".join(texts)
        idx = full.find(find)
        if idx == -1:
            break
        end = idx + len(find)

        # Absolute start offset of each run within the combined text.
        starts = []
        acc = 0
        for t in texts:
            starts.append(acc)
            acc += len(t)

        for i, run in enumerate(runs):
            rstart = starts[i]
            rtext = texts[i]
            rend = rstart + len(rtext)
            # Kept slice to the LEFT of the match within this run.
            left = rtext[: max(0, min(rend, idx) - rstart)] if rstart < idx else ""
            # Kept slice to the RIGHT of the match within this run.
            right = rtext[max(0, end - rstart):] if rend > end else ""
            if rstart <= idx < rend:
                # This run contains the match start -> insertion happens here.
                run.text = left + replace + right
            else:
                run.text = left + right
        made += 1
    return made


# -----------------------------------------------------------------------------
# TRACKED CHANGES (revisions)
#
# Tracked changes are not a python-docx feature or a document flag - they are
# specific OOXML elements inside the paragraph:
#   insertion : the run(s) are wrapped in  <w:ins w:id w:author w:date> ... </w:ins>
#   deletion  : the run(s) are wrapped in  <w:del ...>, and each run's <w:t>
#               becomes <w:delText> (so the text is stored but shown struck out)
#   a replace : is simply a deletion of the old text plus an insertion of the new.
# We build these elements directly via lxml (OxmlElement/qn), which python-docx
# exposes on every element. Scope is text-level only: no tracked paragraph-mark,
# move, or formatting revisions (those are the fragile cases and are excluded).
# -----------------------------------------------------------------------------
def _today_iso():
    """Current date as an OOXML xs:dateTime (time zeroed, UTC 'Z')."""
    return datetime.date.today().strftime("%Y-%m-%dT00:00:00Z")


def _make_rev_id_allocator(doc):
    """
    Return a callable yielding revision ids unique within the document. Seeds
    above the highest existing w:ins/w:del id so we never collide with revisions
    already in the file.
    """
    max_id = 0
    ins_tag, del_tag, id_attr = qn("w:ins"), qn("w:del"), qn("w:id")
    for el in doc.element.iter():
        if el.tag in (ins_tag, del_tag):
            v = el.get(id_attr)
            if v and v.lstrip("-").isdigit():
                max_id = max(max_id, int(v))
    state = {"n": max_id}

    def _next():
        state["n"] += 1
        return state["n"]

    return _next


def _normal_run_el(rpr, text):
    """Build a plain <w:r> with the given text, copying run properties (rPr)."""
    r = OxmlElement("w:r")
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = OxmlElement("w:t")
    t.set(_XML_SPACE, "preserve")  # keep leading/trailing spaces intact
    t.text = text
    r.append(t)
    return r


def _ins_el(rpr, text, rev_id, author, date):
    """Build an <w:ins> wrapping a run that carries `text` (an insertion)."""
    ins = OxmlElement("w:ins")
    ins.set(qn("w:id"), str(rev_id))
    ins.set(qn("w:author"), author)
    ins.set(qn("w:date"), date)
    ins.append(_normal_run_el(rpr, text))
    return ins


def _del_el(rpr, text, rev_id, author, date):
    """Build a <w:del> wrapping a run whose text is stored as <w:delText>."""
    dele = OxmlElement("w:del")
    dele.set(qn("w:id"), str(rev_id))
    dele.set(qn("w:author"), author)
    dele.set(qn("w:date"), date)
    r = OxmlElement("w:r")
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    dt = OxmlElement("w:delText")
    dt.set(_XML_SPACE, "preserve")
    dt.text = text
    r.append(dt)
    dele.append(r)
    return dele


def _replace_element(old_el, new_els):
    """Replace old_el in its parent with the ordered list new_els."""
    parent = old_el.getparent()
    idx = parent.index(old_el)
    for j, ne in enumerate(new_els):
        parent.insert(idx + j, ne)
    parent.remove(old_el)


def _tracked_replace_in_paragraph(paragraph, find, replace, remaining, author, date, next_id):
    """
    Tracked-changes equivalent of _replace_in_paragraph. For each match: delete
    the matched text (wrapped in <w:del>) and, if `replace` is non-empty, insert
    the new text (wrapped in <w:ins>) at the match position. Handles matches that
    span multiple runs. Returns the number of matches processed.

    Termination note: once matched text is wrapped in <w:del>/<w:ins> it is no
    longer a direct child <w:r> of the paragraph, so paragraph.runs (which sees
    only top-level runs) will not re-match it. Each pass therefore advances to
    the next untracked occurrence, or stops.
    """
    made = 0
    while True:
        if remaining is not None and made >= remaining:
            break
        runs = paragraph.runs
        if not runs:
            break
        texts = [r.text for r in runs]
        full = "".join(texts)
        idx = full.find(find)
        if idx == -1:
            break
        end = idx + len(find)

        starts = []
        acc = 0
        for t in texts:
            starts.append(acc)
            acc += len(t)

        # Runs overlapping the match [idx, end).
        overlapping = [
            i for i, t in enumerate(texts)
            if not (starts[i] + len(t) <= idx or starts[i] >= end)
        ]
        first = overlapping[0]

        for i in overlapping:
            r_el = runs[i]._r
            rpr = r_el.find(qn("w:rPr"))  # deep-copied inside builders before removal
            text = texts[i]
            rs = starts[i]
            re_ = rs + len(text)
            a = max(rs, idx)
            b = min(re_, end)
            before = text[: a - rs]
            inside = text[a - rs: b - rs]
            after = text[b - rs:]

            new_els = []
            if before:
                new_els.append(_normal_run_el(rpr, before))
            if i == first and replace:
                # Insertion of the new text takes the formatting of the run
                # where the match starts (documented cross-run trade-off).
                new_els.append(_ins_el(rpr, replace, next_id(), author, date))
            if inside:
                new_els.append(_del_el(rpr, inside, next_id(), author, date))
            if after:
                new_els.append(_normal_run_el(rpr, after))

            if new_els:
                _replace_element(r_el, new_els)
            else:
                r_el.getparent().remove(r_el)
        made += 1
    return made


def _iter_revisions(doc):
    """Yield (kind, element) for every top-level insertion/deletion revision."""
    ins_tag, del_tag = qn("w:ins"), qn("w:del")
    for el in doc.element.iter():
        if el.tag == ins_tag:
            yield "insertion", el
        elif el.tag == del_tag:
            yield "deletion", el


def _revision_text(el):
    """Concatenate the visible text of a revision element (w:t or w:delText)."""
    parts = []
    for t in el.iter(qn("w:t")):
        parts.append(t.text or "")
    for t in el.iter(qn("w:delText")):
        parts.append(t.text or "")
    return "".join(parts)


def _accept_all(doc):
    """Accept every tracked change: keep inserted text, drop deleted text."""
    ins_count = del_count = 0
    for el in list(doc.element.iter(qn("w:ins"))):
        parent = el.getparent()
        idx = parent.index(el)
        for child in list(el):          # promote the inserted runs to normal
            parent.insert(idx, child)
            idx += 1
        parent.remove(el)
        ins_count += 1
    for el in list(doc.element.iter(qn("w:del"))):
        el.getparent().remove(el)       # deleted content disappears
        del_count += 1
    return ins_count, del_count


def _reject_all(doc):
    """Reject every tracked change: drop inserted text, restore deleted text."""
    ins_count = del_count = 0
    for el in list(doc.element.iter(qn("w:ins"))):
        el.getparent().remove(el)       # inserted content disappears
        ins_count += 1
    for el in list(doc.element.iter(qn("w:del"))):
        parent = el.getparent()
        idx = parent.index(el)
        for child in list(el):          # restore runs, delText -> t
            for dt in child.findall(qn("w:delText")):
                t = OxmlElement("w:t")
                t.set(_XML_SPACE, "preserve")
                t.text = dt.text
                child.replace(dt, t)
            parent.insert(idx, child)
            idx += 1
        parent.remove(el)
        del_count += 1
    return ins_count, del_count


def _render_linear_text(doc):
    """Full document as plain text, tables flattened to ' | '-joined rows."""
    lines = []
    for block in _iter_block_items(doc):
        if isinstance(block, Paragraph):
            lines.append(block.text)
        elif isinstance(block, Table):
            for row in block.rows:
                lines.append(" | ".join(cell.text for cell in row.cells))
    return "\n".join(lines)


def _render_structured(doc):
    """
    Ordered block list. Paragraphs carry a para_index into doc.paragraphs and
    tables carry a table_index into doc.tables, so the model can address them
    in later calls (set_paragraph_format, get_tables, etc.).
    """
    blocks = []
    p_i = 0
    t_i = 0
    for b_i, block in enumerate(_iter_block_items(doc)):
        if isinstance(block, Paragraph):
            blocks.append({
                "block_index": b_i,
                "type": "paragraph",
                "para_index": p_i,
                "style": block.style.name if block.style is not None else None,
                "text": block.text,
            })
            p_i += 1
        elif isinstance(block, Table):
            blocks.append({
                "block_index": b_i,
                "type": "table",
                "table_index": t_i,
                "rows": len(block.rows),
                "cols": len(block.columns),
            })
            t_i += 1
    return blocks


# =============================================================================
# TOOL HANDLERS  (each takes an `args` dict, returns a JSON-serialisable object)
# =============================================================================
def tool_open(args):
    path = _resolve_path(_require(args, "path"))
    if not os.path.isfile(path):
        raise ToolError("File not found: {}".format(path))
    if not path.lower().endswith(".docx"):
        raise ToolError("Only .docx files are supported (got: {})".format(path))
    if len(SESSIONS) >= MAX_SESSIONS:
        raise ToolError(
            "Too many open sessions ({}). Close some with msword_close.".format(
                MAX_SESSIONS
            )
        )
    try:
        doc = docx.Document(path)
    except PackageNotFoundError:
        raise ToolError(
            "Not a valid .docx file (corrupt, or not really Office Open XML)."
        )
    except PermissionError:
        raise ToolError("Permission denied opening: {}".format(path))
    sid = uuid.uuid4().hex[:12]
    SESSIONS[sid] = {"path": path, "doc": doc, "opened_at": _now_iso()}
    log("opened {} as session {}".format(path, sid))
    return {
        "session_id": sid,
        "path": path,
        "paragraphs": len(doc.paragraphs),
        "tables": len(doc.tables),
    }


def tool_close(args):
    sid = _require(args, "session_id")
    if sid not in SESSIONS:
        raise ToolError("Unknown session_id '{}'.".format(sid))
    SESSIONS.pop(sid)
    log("closed session {}".format(sid))
    return {"closed": sid, "note": "Any unsaved changes were discarded."}


def tool_list_sessions(args):
    out = []
    for sid, s in SESSIONS.items():
        out.append({
            "session_id": sid,
            "path": s["path"],
            "opened_at": s["opened_at"],
            "paragraphs": len(s["doc"].paragraphs),
            "tables": len(s["doc"].tables),
        })
    return {"sessions": out, "count": len(out)}


def tool_get_content(args):
    session = _get_session(args)
    doc = session["doc"]
    mode = args.get("mode", "text")
    if mode == "text":
        return {"mode": "text", "content": _render_linear_text(doc)}
    elif mode == "structured":
        return {"mode": "structured", "blocks": _render_structured(doc)}
    else:
        raise ToolError("mode must be 'text' or 'structured' (got '{}')".format(mode))


def tool_search(args):
    session = _get_session(args)
    doc = session["doc"]
    query = _require(args, "query")
    if query == "":
        raise ToolError("query must not be empty")
    case_sensitive = bool(args.get("case_sensitive", False))
    max_results = args.get("max_results")
    if max_results is not None:
        max_results = int(max_results)

    needle = query if case_sensitive else query.lower()

    def _hits(text):
        hay = text if case_sensitive else text.lower()
        positions = []
        start = 0
        while True:
            i = hay.find(needle, start)
            if i == -1:
                break
            positions.append(i)
            start = i + len(needle)
        return positions

    def _snippet(text, i):
        a = max(0, i - SEARCH_CONTEXT_CHARS)
        b = min(len(text), i + len(query) + SEARCH_CONTEXT_CHARS)
        prefix = "..." if a > 0 else ""
        suffix = "..." if b < len(text) else ""
        return prefix + text[a:b] + suffix

    matches = []

    # Body paragraphs (addressable by para_index).
    for p_i, p in enumerate(doc.paragraphs):
        for pos in _hits(p.text):
            matches.append({
                "location": "body_paragraph",
                "para_index": p_i,
                "char_offset": pos,
                "snippet": _snippet(p.text, pos),
            })
            if max_results and len(matches) >= max_results:
                return {"query": query, "matches": matches, "truncated": True}

    # Table cells (addressable by table/row/col).
    for t_i, tbl in enumerate(doc.tables):
        for r_i, row in enumerate(tbl.rows):
            for c_i, cell in enumerate(row.cells):
                for pos in _hits(cell.text):
                    matches.append({
                        "location": "table_cell",
                        "table_index": t_i,
                        "row": r_i,
                        "col": c_i,
                        "char_offset": pos,
                        "snippet": _snippet(cell.text, pos),
                    })
                    if max_results and len(matches) >= max_results:
                        return {"query": query, "matches": matches, "truncated": True}

    return {"query": query, "matches": matches, "truncated": False}


def tool_replace_text(args):
    session = _get_session(args)
    doc = session["doc"]
    find = _require(args, "find")
    if find == "":
        raise ToolError("'find' must not be empty")
    replace = args.get("replace", "")
    count = args.get("count")  # max total replacements; None = all
    if count is not None:
        count = int(count)
        if count < 0:
            raise ToolError("'count' must be >= 0")

    track = bool(args.get("track_changes", False))
    author = args.get("author") or AUTHOR
    date = _today_iso()
    next_id = _make_rev_id_allocator(doc) if track else None

    total = 0
    for paragraph in _all_editable_paragraphs(doc):
        remaining = None if count is None else max(0, count - total)
        if remaining == 0:
            break
        if track:
            total += _tracked_replace_in_paragraph(
                paragraph, find, replace, remaining, author, date, next_id
            )
        else:
            total += _replace_in_paragraph(paragraph, find, replace, remaining)

    result = {"find": find, "replace": replace, "replacements": total,
              "tracked": track}
    if track:
        result["author"] = author
        result["date"] = date
    return result


def tool_add_paragraph(args):
    session = _get_session(args)
    doc = session["doc"]
    text = args.get("text", "")
    style = args.get("style")
    try:
        if style:
            doc.add_paragraph(text, style=style)
        else:
            doc.add_paragraph(text)
    except KeyError:
        raise ToolError("Unknown paragraph style: '{}'".format(style))
    return {"para_index": len(doc.paragraphs) - 1, "text": text}


def tool_add_heading(args):
    session = _get_session(args)
    doc = session["doc"]
    text = args.get("text", "")
    level = int(args.get("level", 1))
    if not 0 <= level <= 9:
        raise ToolError("heading level must be between 0 (Title) and 9")
    try:
        doc.add_heading(text, level=level)
    except KeyError:
        raise ToolError(
            "Heading style for level {} is not defined in this document.".format(level)
        )
    return {"para_index": len(doc.paragraphs) - 1, "level": level, "text": text}


def tool_add_table(args):
    session = _get_session(args)
    doc = session["doc"]
    data = args.get("data")
    style = args.get("style")  # e.g. "Table Grid"

    if data is not None:
        if not isinstance(data, list) or not data or not all(
            isinstance(row, list) for row in data
        ):
            raise ToolError("'data' must be a non-empty list of row lists")
        rows = len(data)
        cols = max(len(r) for r in data)
    else:
        rows = int(_require(args, "rows"))
        cols = int(_require(args, "cols"))
        if rows < 1 or cols < 1:
            raise ToolError("rows and cols must both be >= 1")

    table = doc.add_table(rows=rows, cols=cols)
    if style:
        try:
            table.style = style
        except KeyError:
            raise ToolError("Unknown table style: '{}'".format(style))
    if data is not None:
        for r_i, row in enumerate(data):
            for c_i in range(cols):
                value = row[c_i] if c_i < len(row) else ""
                table.rows[r_i].cells[c_i].text = "" if value is None else str(value)

    return {"table_index": len(doc.tables) - 1, "rows": rows, "cols": cols}


def tool_get_tables(args):
    session = _get_session(args)
    doc = session["doc"]
    table_index = args.get("table_index")

    def _dump(tbl):
        return [[cell.text for cell in row.cells] for row in tbl.rows]

    if table_index is not None:
        table_index = int(table_index)
        if not 0 <= table_index < len(doc.tables):
            raise ToolError("table_index out of range (0..{})".format(len(doc.tables) - 1))
        return {"table_index": table_index, "cells": _dump(doc.tables[table_index])}

    return {"tables": [
        {"table_index": i, "cells": _dump(t)} for i, t in enumerate(doc.tables)
    ]}


# Accept both US and Australian spellings of "centre".
_ALIGN_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "centre": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def tool_set_paragraph_format(args):
    session = _get_session(args)
    doc = session["doc"]
    para_index = int(_require(args, "para_index"))
    if not 0 <= para_index < len(doc.paragraphs):
        raise ToolError(
            "para_index out of range (0..{}).".format(len(doc.paragraphs) - 1)
        )
    paragraph = doc.paragraphs[para_index]
    applied = {}

    # Character formatting is applied to every run in the paragraph.
    for attr in ("bold", "italic", "underline"):
        if attr in args and args[attr] is not None:
            value = bool(args[attr])
            for run in paragraph.runs:
                setattr(run, attr, value)
            applied[attr] = value

    if "alignment" in args and args["alignment"] is not None:
        key = str(args["alignment"]).lower()
        if key not in _ALIGN_MAP:
            raise ToolError(
                "alignment must be one of: {}".format(", ".join(sorted(_ALIGN_MAP)))
            )
        paragraph.alignment = _ALIGN_MAP[key]
        applied["alignment"] = key

    if "style" in args and args["style"] is not None:
        try:
            paragraph.style = args["style"]
        except KeyError:
            raise ToolError("Unknown paragraph style: '{}'".format(args["style"]))
        applied["style"] = args["style"]

    if not applied:
        raise ToolError(
            "No formatting supplied. Provide one or more of: "
            "bold, italic, underline, alignment, style."
        )
    return {"para_index": para_index, "applied": applied}


def tool_save(args):
    session = _get_session(args)
    doc = session["doc"]
    save_as = args.get("path")
    if save_as:
        target = _resolve_path(save_as)
        if not target.lower().endswith(".docx"):
            raise ToolError("Save path must end in .docx")
        parent = os.path.dirname(target)
        if parent and not os.path.isdir(parent):
            raise ToolError("Destination folder does not exist: {}".format(parent))
    else:
        target = session["path"]

    try:
        doc.save(target)
    except PermissionError:
        raise ToolError(
            "Permission denied saving to {} (is it open in Word?).".format(target)
        )
    except OSError as e:
        raise ToolError("Could not save to {}: {}".format(target, e))

    if save_as:
        session["path"] = target  # rebind session to the new file
    log("saved session {} -> {}".format(args.get("session_id"), target))
    return {"saved": target}


def tool_list_changes(args):
    session = _get_session(args)
    doc = session["doc"]
    changes = []
    for kind, el in _iter_revisions(doc):
        changes.append({
            "type": kind,
            "id": el.get(qn("w:id")),
            "author": el.get(qn("w:author")),
            "date": el.get(qn("w:date")),
            "text": _revision_text(el),
        })
    return {"changes": changes, "count": len(changes)}


def tool_accept_all_changes(args):
    session = _get_session(args)
    ins_count, del_count = _accept_all(session["doc"])
    return {"accepted_insertions": ins_count, "accepted_deletions": del_count}


def tool_reject_all_changes(args):
    session = _get_session(args)
    ins_count, del_count = _reject_all(session["doc"])
    return {"rejected_insertions": ins_count, "rejected_deletions": del_count}


# =============================================================================
# TOOL REGISTRY  (name -> (handler, JSON-Schema inputSchema, description))
# =============================================================================
TOOLS = [
    {
        "name": "msword_open",
        "description": "Open a .docx file into an in-memory session and return a session_id used by all other tools.",
        "handler": tool_open,
        "inputSchema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "Absolute or ~ path to a .docx file."}
            },
            "required": ["path"],
        },
    },
    {
        "name": "msword_close",
        "description": "Close a session and free its document. Unsaved changes are discarded.",
        "handler": tool_close,
        "inputSchema": {
            "type": "object",
            "properties": {"session_id": {"type": "string"}},
            "required": ["session_id"],
        },
    },
    {
        "name": "msword_list_sessions",
        "description": "List all currently open document sessions.",
        "handler": tool_list_sessions,
        "inputSchema": {"type": "object", "properties": {}},
    },
    {
        "name": "msword_get_content",
        "description": "Read a document. mode='text' returns linear plain text (tables flattened); mode='structured' returns an ordered block list with para_index/table_index for addressing.",
        "handler": tool_get_content,
        "inputSchema": {
            "type": "object",
            "properties": {
                "session_id": {"type": "string"},
                "mode": {"type": "string", "enum": ["text", "structured"], "default": "text"},
            },
            "required": ["session_id"],
        },
    },
    {
        "name": "msword_search",
        "description": "Find text in body paragraphs and table cells. Returns match locations and context snippets.",
        "handler": tool_search,
        "inputSchema": {
            "type": "object",
            "properties": {
                "session_id": {"type": "string"},
                "query": {"type": "string"},
                "case_sensitive": {"type": "boolean", "default": False},
                "max_results": {"type": "integer", "description": "Optional cap on results."},
            },
            "required": ["session_id", "query"],
        },
    },
    {
        "name": "msword_replace_text",
        "description": "Replace text across body paragraphs and table cells. Handles matches spanning multiple runs. Set track_changes=true to record each replacement as a Word tracked change (deletion of old text + insertion of new) instead of editing silently. An empty 'replace' with track_changes=true is a tracked deletion.",
        "handler": tool_replace_text,
        "inputSchema": {
            "type": "object",
            "properties": {
                "session_id": {"type": "string"},
                "find": {"type": "string"},
                "replace": {"type": "string", "default": ""},
                "count": {"type": "integer", "description": "Optional max total replacements; omit to replace all."},
                "track_changes": {"type": "boolean", "default": False, "description": "Record edits as Word tracked changes."},
                "author": {"type": "string", "description": "Override the tracked-change author for this call (defaults to the server's --author)."},
            },
            "required": ["session_id", "find"],
        },
    },
    {
        "name": "msword_add_paragraph",
        "description": "Append a paragraph to the end of the document. Optional paragraph style name.",
        "handler": tool_add_paragraph,
        "inputSchema": {
            "type": "object",
            "properties": {
                "session_id": {"type": "string"},
                "text": {"type": "string", "default": ""},
                "style": {"type": "string", "description": "Optional style, e.g. 'Normal', 'Quote'."},
            },
            "required": ["session_id"],
        },
    },
    {
        "name": "msword_add_heading",
        "description": "Append a heading. level 0 = Title, 1..9 = Heading 1..9.",
        "handler": tool_add_heading,
        "inputSchema": {
            "type": "object",
            "properties": {
                "session_id": {"type": "string"},
                "text": {"type": "string", "default": ""},
                "level": {"type": "integer", "minimum": 0, "maximum": 9, "default": 1},
            },
            "required": ["session_id"],
        },
    },
    {
        "name": "msword_add_table",
        "description": "Append a table. Either supply 'data' (list of row lists) to fill it, or 'rows' and 'cols' for an empty grid. Optional table style e.g. 'Table Grid'.",
        "handler": tool_add_table,
        "inputSchema": {
            "type": "object",
            "properties": {
                "session_id": {"type": "string"},
                "data": {
                    "type": "array",
                    "items": {"type": "array", "items": {"type": "string"}},
                    "description": "Row-major cell values.",
                },
                "rows": {"type": "integer", "minimum": 1},
                "cols": {"type": "integer", "minimum": 1},
                "style": {"type": "string"},
            },
            "required": ["session_id"],
        },
    },
    {
        "name": "msword_get_tables",
        "description": "Return table contents as arrays of cell text. Omit table_index for all tables.",
        "handler": tool_get_tables,
        "inputSchema": {
            "type": "object",
            "properties": {
                "session_id": {"type": "string"},
                "table_index": {"type": "integer"},
            },
            "required": ["session_id"],
        },
    },
    {
        "name": "msword_set_paragraph_format",
        "description": "Apply simple formatting to a body paragraph (addressed by para_index from get_content structured). Any of: bold, italic, underline, alignment (left/center/centre/right/justify), style.",
        "handler": tool_set_paragraph_format,
        "inputSchema": {
            "type": "object",
            "properties": {
                "session_id": {"type": "string"},
                "para_index": {"type": "integer"},
                "bold": {"type": "boolean"},
                "italic": {"type": "boolean"},
                "underline": {"type": "boolean"},
                "alignment": {"type": "string", "enum": ["left", "center", "centre", "right", "justify"]},
                "style": {"type": "string"},
            },
            "required": ["session_id", "para_index"],
        },
    },
    {
        "name": "msword_save",
        "description": "Save the session's document. Omit 'path' to save in place, or supply a .docx 'path' to save-as (the session then tracks the new file).",
        "handler": tool_save,
        "inputSchema": {
            "type": "object",
            "properties": {
                "session_id": {"type": "string"},
                "path": {"type": "string", "description": "Optional save-as path ending in .docx."},
            },
            "required": ["session_id"],
        },
    },
    {
        "name": "msword_list_changes",
        "description": "List all tracked changes (insertions and deletions) currently in the document, with author, date and text.",
        "handler": tool_list_changes,
        "inputSchema": {
            "type": "object",
            "properties": {"session_id": {"type": "string"}},
            "required": ["session_id"],
        },
    },
    {
        "name": "msword_accept_all_changes",
        "description": "Accept every tracked change: inserted text is kept as normal text and deleted text is removed.",
        "handler": tool_accept_all_changes,
        "inputSchema": {
            "type": "object",
            "properties": {"session_id": {"type": "string"}},
            "required": ["session_id"],
        },
    },
    {
        "name": "msword_reject_all_changes",
        "description": "Reject every tracked change: inserted text is removed and deleted text is restored.",
        "handler": tool_reject_all_changes,
        "inputSchema": {
            "type": "object",
            "properties": {"session_id": {"type": "string"}},
            "required": ["session_id"],
        },
    },
]

_TOOL_BY_NAME = {t["name"]: t for t in TOOLS}


# =============================================================================
# JSON-RPC / MCP DISPATCH
# =============================================================================
def _jsonrpc_result(msg_id, result):
    return {"jsonrpc": "2.0", "id": msg_id, "result": result}


def _jsonrpc_error(msg_id, code, message):
    return {"jsonrpc": "2.0", "id": msg_id, "error": {"code": code, "message": message}}


def handle_message(msg):
    """
    Process one decoded JSON-RPC message. Returns a response dict, or None for
    notifications (which must not be answered).
    """
    method = msg.get("method")
    msg_id = msg.get("id")
    is_notification = "id" not in msg
    params = msg.get("params") or {}

    if method == "initialize":
        client_proto = params.get("protocolVersion") or PROTOCOL_VERSION_FALLBACK
        return _jsonrpc_result(msg_id, {
            "protocolVersion": client_proto,
            "capabilities": {"tools": {}},
            "serverInfo": {"name": SERVER_NAME, "version": SERVER_VERSION},
        })

    if method in ("notifications/initialized", "initialized"):
        return None  # notification, no response

    if method == "ping":
        return _jsonrpc_result(msg_id, {})

    if method == "tools/list":
        listed = [
            {"name": t["name"], "description": t["description"], "inputSchema": t["inputSchema"]}
            for t in TOOLS
        ]
        return _jsonrpc_result(msg_id, {"tools": listed})

    if method == "tools/call":
        name = params.get("name")
        arguments = params.get("arguments") or {}
        tool = _TOOL_BY_NAME.get(name)
        if tool is None:
            return _jsonrpc_result(msg_id, {
                "content": [{"type": "text", "text": "Unknown tool: {}".format(name)}],
                "isError": True,
            })
        try:
            result = tool["handler"](arguments)
            text = json.dumps(result, indent=2, ensure_ascii=False)
            return _jsonrpc_result(msg_id, {
                "content": [{"type": "text", "text": text}],
                "isError": False,
            })
        except ToolError as e:
            return _jsonrpc_result(msg_id, {
                "content": [{"type": "text", "text": "Error: {}".format(e)}],
                "isError": True,
            })
        except Exception as e:  # unexpected - log full trace to stderr, return summary
            log("UNEXPECTED in tool {}:\n{}".format(name, traceback.format_exc()))
            return _jsonrpc_result(msg_id, {
                "content": [{"type": "text", "text": "Internal error: {}".format(e)}],
                "isError": True,
            })

    # Unknown method.
    if is_notification:
        return None
    return _jsonrpc_error(msg_id, -32601, "Method not found: {}".format(method))


def serve():
    """Main stdio loop: read newline-delimited JSON-RPC, dispatch, reply."""
    docx_ver, lxml_ver = _versions()
    log("starting")
    log("interpreter: {}".format(sys.executable))
    log("python-docx {} / lxml {}".format(docx_ver, lxml_ver))
    log("tracked-change author: {}".format(AUTHOR))
    if DOCUMENT_ROOT:
        log("path sandbox DOCUMENT_ROOT = {}".format(DOCUMENT_ROOT))

    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            msg = json.loads(line)
        except json.JSONDecodeError:
            # Cannot know the id; emit a parse error with null id per spec.
            sys.stdout.write(json.dumps(_jsonrpc_error(None, -32700, "Parse error")) + "\n")
            sys.stdout.flush()
            continue

        try:
            response = handle_message(msg)
        except Exception as e:  # last-resort guard so the server never dies
            log("FATAL in handle_message:\n{}".format(traceback.format_exc()))
            response = _jsonrpc_error(msg.get("id"), -32603, "Internal error: {}".format(e))

        if response is not None:
            sys.stdout.write(json.dumps(response, ensure_ascii=False) + "\n")
            sys.stdout.flush()

    log("stdin closed, exiting")


# =============================================================================
# SELF-TEST  (--check): offline round-trip so a single transfer can be validated
# =============================================================================
def run_check():
    import tempfile
    docx_ver, lxml_ver = _versions()
    print("[check] interpreter : {}".format(sys.executable))
    print("[check] python-docx : {}".format(docx_ver))
    print("[check] lxml        : {}".format(lxml_ver))

    tmpdir = tempfile.mkdtemp(prefix="msword_check_")
    path = os.path.join(tmpdir, "roundtrip.docx")
    ok = True
    try:
        # Build a document exercising each editing path.
        d = docx.Document()
        d.add_heading("Title Here", level=0)
        d.add_paragraph("The quick brown fox jumps over the lazy dog.")
        d.add_table(rows=2, cols=2)
        d.tables[0].rows[0].cells[0].text = "old value"
        d.save(path)

        # Reopen through the tool layer.
        opened = tool_open({"path": path})
        sid = opened["session_id"]
        assert opened["paragraphs"] >= 2, "expected >=2 paragraphs"
        assert opened["tables"] == 1, "expected 1 table"

        # Search should find the fox sentence.
        s = tool_search({"session_id": sid, "query": "brown fox"})
        assert s["matches"], "search found nothing"

        # Replace across body + table cell.
        r1 = tool_replace_text({"session_id": sid, "find": "quick", "replace": "slow"})
        assert r1["replacements"] == 1, "body replace count wrong"
        r2 = tool_replace_text({"session_id": sid, "find": "old value", "replace": "new value"})
        assert r2["replacements"] == 1, "table-cell replace count wrong"

        # Append + format.
        tool_add_paragraph({"session_id": sid, "text": "Appended paragraph."})
        pi = tool_get_content({"session_id": sid, "mode": "structured"})["blocks"]
        last_para = max(b["para_index"] for b in pi if b["type"] == "paragraph")
        tool_set_paragraph_format({"session_id": sid, "para_index": last_para,
                                   "bold": True, "alignment": "centre"})

        # Save-as and reopen to confirm persistence of the edits.
        path2 = os.path.join(tmpdir, "roundtrip2.docx")
        tool_save({"session_id": sid, "path": path2})
        tool_close({"session_id": sid})

        reopened = tool_open({"path": path2})
        text = tool_get_content({"session_id": reopened["session_id"], "mode": "text"})["content"]
        assert "slow brown fox" in text, "body edit did not persist"
        assert "new value" in text, "table edit did not persist"
        assert "Appended paragraph." in text, "append did not persist"
        tool_close({"session_id": reopened["session_id"]})

        print("[check] round-trip: PASS")

        # --- Tracked-changes round-trip -------------------------------------
        tpath = os.path.join(tmpdir, "tracked.docx")
        td = docx.Document()
        td.add_paragraph("The report is DRAFT and confidential.")
        td.save(tpath)

        ts = tool_open({"path": tpath})
        tsid = ts["session_id"]
        rt = tool_replace_text({"session_id": tsid, "find": "DRAFT",
                                "replace": "FINAL", "track_changes": True,
                                "author": "Self Test"})
        assert rt["replacements"] == 1 and rt["tracked"], "tracked replace failed"

        listed = tool_list_changes({"session_id": tsid})
        kinds = sorted(c["type"] for c in listed["changes"])
        assert kinds == ["deletion", "insertion"], \
            "expected one insertion + one deletion, got {}".format(kinds)
        assert all(c["author"] == "Self Test" for c in listed["changes"]), \
            "author not stamped"

        tpath2 = os.path.join(tmpdir, "tracked2.docx")
        tool_save({"session_id": tsid, "path": tpath2})
        tool_close({"session_id": tsid})

        # Accept path: FINAL stays, DRAFT gone.
        a = tool_open({"path": tpath2})
        tool_accept_all_changes({"session_id": a["session_id"]})
        atext = tool_get_content({"session_id": a["session_id"], "mode": "text"})["content"]
        assert "FINAL" in atext and "DRAFT" not in atext, "accept-all wrong"
        tool_close({"session_id": a["session_id"]})

        # Reject path: DRAFT restored, FINAL gone.
        r = tool_open({"path": tpath2})
        tool_reject_all_changes({"session_id": r["session_id"]})
        rtext = tool_get_content({"session_id": r["session_id"], "mode": "text"})["content"]
        assert "DRAFT" in rtext and "FINAL" not in rtext, "reject-all wrong"
        tool_close({"session_id": r["session_id"]})

        print("[check] tracked-changes: PASS")
    except Exception as e:
        ok = False
        print("[check] round-trip: FAIL -> {}".format(e))
        traceback.print_exc()
    finally:
        # Clean up temp files; leave nothing behind.
        try:
            for f in os.listdir(tmpdir):
                os.remove(os.path.join(tmpdir, f))
            os.rmdir(tmpdir)
        except Exception:
            pass

    if ok:
        print("[check] ALL CHECKS PASSED")
        return 0
    print("[check] CHECKS FAILED")
    return 1


def main():
    parser = argparse.ArgumentParser(
        description="MS Word (.docx) python-docx MCP stdio server. "
                    "With no arguments it runs as an MCP server on stdin/stdout."
    )
    parser.add_argument(
        "--check", action="store_true",
        help="Run an offline open/edit/save/reopen self-test and exit."
    )
    parser.add_argument(
        "--author", default=None, metavar="NAME",
        help="Author name stamped on tracked changes "
             "(default: the TRACKED_CHANGE_AUTHOR config value)."
    )
    args = parser.parse_args()

    global AUTHOR
    if args.author:
        AUTHOR = args.author

    if args.check:
        sys.exit(run_check())

    try:
        serve()
    except KeyboardInterrupt:
        log("interrupted, exiting")


if __name__ == "__main__":
    main()
