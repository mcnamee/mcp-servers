#!/usr/bin/env python3
r"""
pdf2md_mcp.py - Self-contained MCP (Model Context Protocol) stdio server for
Continue that converts PDFs in a folder to Markdown, including tables (both
bordered and borderless).

=============================================================================
 DEPENDENCIES
=============================================================================
Beyond the Python standard library, this needs two packages (the second pulls
in table/layout handling):

    pip install pymupdf pymupdf4llm

If your network's pip proxy needs to be named explicitly:

    pip install --proxy http://YOUR-PROXY:PORT pymupdf pymupdf4llm

Notes:
  * pymupdf4llm depends on pymupdf (PyMuPDF), so installing it usually pulls
    pymupdf in automatically; both are listed for clarity.
  * INTERPRETER MATCH (common gotcha): install into the SAME Python that
    Continue launches this server with. If you install with one "pip" but
    Continue runs a different Python, the server reports the package as
    missing even though "pip show pymupdf4llm" lists it. The reliable way is
    to use ONE explicit interpreter for both steps:
        "C:\path\to\python.exe" -m pip install pymupdf pymupdf4llm
    and set that same full path as "command" in the config below (instead of
    a bare "python"). Verify with:
        "C:\path\to\python.exe" -c "import pymupdf4llm, fitz; print(pymupdf4llm.__file__)"
    On startup this server logs its own interpreter path (sys.executable) to
    stderr - compare it against where the packages are installed.
  * OCR (for scanned, image-only PDFs) additionally needs Tesseract installed
    on the machine. It is NOT required for normal text PDFs or their tables.
    Without Tesseract, text PDFs still convert; scanned pages simply produce
    no text and are reported as a per-file failure.

=============================================================================
 CONTINUE CONFIGURATION  (paste into the "mcpServers" block of config.yaml /
 config.json, adjusting the paths)
=============================================================================
    "pdf2md": {
      "command": "python",
      "args": [
        "C:\\path\\to\\pdf2md_mcp.py",
        "--input-dir",  "C:\\Reference\\PDFs",
        "--output-dir", "C:\\Reference\\Markdown"
      ],
      "env": { "PYTHONUTF8": "1" }
    }

  * Backslashes in JSON paths must be doubled (or use forward slashes).
  * Prefer the FULL path to a specific python.exe for "command" (e.g.
    "C:\\Python311\\python.exe") rather than a bare "python", so Continue
    launches the exact interpreter that has pymupdf4llm installed.
  * Add  "--recursive"  to the args list to also process sub-folders
    (the sub-folder structure is mirrored under the output folder).
  * PYTHONUTF8=1 in the env block prevents stdout encoding crashes on Windows.
  * MCP tools are only exposed to the model in Continue's AGENT mode.
  * After editing the config, run  "Developer: Reload Window"  (a full reload,
    not the MCP toggle) to avoid the "already connected to transport" bug.

=============================================================================
 MANUAL TESTING (PowerShell, on the target box, outside Continue)
=============================================================================
Feed newline-delimited JSON-RPC on stdin. Use a here-string (NOT an inline
echo with nested JSON, which silently drops the "arguments" object):

    $msgs = @'
    {"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05"}}
    {"jsonrpc":"2.0","id":2,"method":"tools/list"}
    {"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"convert_all_pdfs","arguments":{}}}
    '@
    $msgs | python C:\path\to\pdf2md_mcp.py --input-dir "C:\Reference\PDFs" --output-dir "C:\Reference\Markdown"

Expected: three JSON lines on stdout (initialize result, tool list, conversion
summary). Diagnostics appear on stderr and never on stdout.

To test the single-file fuzzy tool, replace the third line with:
    {"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"convert_pdf_to_markdown","arguments":{"query":"procurement policy"}}}

A quick Python conversion check (no MCP) can be done with:
    python -c "import pdf2md_mcp as m; print(m.convert_one(r'C:\Reference\PDFs\Some File.pdf'))"

=============================================================================
 TOOLS EXPOSED
=============================================================================
1. convert_all_pdfs
   Convert every PDF in the input folder to Markdown in the output folder.
   PDFs whose .md already exists are SKIPPED unless force=true. Backs the
   request "convert all pdfs to markdown".

2. convert_pdf_to_markdown
   Convert a single PDF identified by a rough name (e.g. "procurement
   policy"). Refuses to guess when several PDFs match equally well.

=============================================================================
 CONVERSION BEHAVIOUR
=============================================================================
  * Extraction uses pymupdf4llm, which handles reading order, headings, and
    TABLES - both bordered (ruling lines) and borderless (column alignment).
  * Post-processing then:
      - rejoins list markers that land on their own line (the "1." on one
        line, content on the next" problem) with their content;
      - collapses any run of 2+ blank lines into a single blank line;
      - strips trailing whitespace from every line.
  * Table rows are never altered by the list-marker fix.

=============================================================================
 TRANSPORT NOTES (MCP over stdio)
=============================================================================
  * stdout carries ONLY newline-delimited JSON-RPC 2.0. Nothing else.
  * All diagnostics go to stderr.
  * PyMuPDF/MuPDF and pymupdf4llm emit advisory text (and occasionally raw
    bytes) to the C-level stdout on import and during conversion. Those calls
    are wrapped so OS-level stdout (fd 1) is redirected away for their
    duration, keeping the JSON-RPC stream clean.
  * Outgoing JSON uses ensure_ascii=True, so the stdout stream stays pure
    ASCII regardless of document content.
"""

import argparse
import contextlib
import difflib
import json
import os
import re
import sys
import traceback
from pathlib import Path


# --------------------------------------------------------------------------
# OS-level stdout suppression (protects the JSON-RPC channel)
# --------------------------------------------------------------------------
@contextlib.contextmanager
def _suppress_fd_stdout():
    """Redirect OS-level stdout (fd 1) to the null device for the duration of
    the block. This catches C-level prints from MuPDF / pymupdf4llm that a
    Python-level redirect would miss."""
    saved_fd = os.dup(1)
    devnull_fd = os.open(os.devnull, os.O_WRONLY)
    try:
        os.dup2(devnull_fd, 1)
        yield
    finally:
        os.dup2(saved_fd, 1)
        os.close(saved_fd)
        os.close(devnull_fd)


# Import the heavy libraries with stdout suppressed (banner safety). Each is
# imported separately so the error names the EXACT module that is missing -
# the usual cause is an interpreter mismatch (the package was pip-installed
# into a different Python than the one Continue launches this server with).
fitz = None
pymupdf4llm = None
IMPORT_ERROR = None
try:
    with _suppress_fd_stdout():
        import fitz  # PyMuPDF
except Exception as _exc:  # noqa: BLE001 - surfaced lazily at call time
    IMPORT_ERROR = "PyMuPDF (fitz) failed to import: {!r}".format(_exc)
if IMPORT_ERROR is None:
    try:
        with _suppress_fd_stdout():
            import pymupdf4llm
    except Exception as _exc:  # noqa: BLE001 - surfaced lazily at call time
        IMPORT_ERROR = "pymupdf4llm failed to import: {!r}".format(_exc)


# --------------------------------------------------------------------------
# Configuration constants
# --------------------------------------------------------------------------
SERVER_NAME = "pdf2md-mcp"
SERVER_VERSION = "3.0.0"
DEFAULT_PROTOCOL_VERSION = "2024-11-05"

# Table detection strategy passed to pymupdf4llm. "lines_strict" is the
# library default and detected both bordered and borderless tables in testing.
# Alternatives if a particular document misbehaves: "lines" or "text".
TABLE_STRATEGY = "lines_strict"

# Fuzzy matching (single-file tool).
MIN_RATIO = 0.40
AMBIGUITY_DELTA = 0.05

# Cap Markdown inlined into a single tool result (full text is always written
# to the .md file regardless).
MAX_INLINE_CHARS = 20000


# Marker-only line patterns. Bare "-" and "*" are deliberately excluded to
# avoid merging stray separators; numbered / parenthesised / lettered /
# bulleted markers are covered.
MARKER_ONLY_RE = re.compile(
    r"^\s*(?:"
    r"\(\d{1,3}\)"                  # (1)
    r"|\d{1,3}[.)]"                 # 1.  1)
    r"|\([a-zA-Z]\)"                # (a)
    r"|[a-zA-Z][.)]"                # a.  a)
    r"|[ivxlcdmIVXLCDM]{1,5}[.)]"   # ii.  iv)
    r"|[\u2022\u25aa\u25e6\u00b7\u2023]"  # bullets
    r")\s*$"
)
LEADING_BULLET_RE = re.compile(r"^\s*[\u2022\u25aa\u25e6\u00b7\u2023]\s+")


# --------------------------------------------------------------------------
# Logging (stderr only)
# --------------------------------------------------------------------------
def log(message):
    print("[{}] {}".format(SERVER_NAME, message), file=sys.stderr, flush=True)


# --------------------------------------------------------------------------
# Markdown post-processing
# --------------------------------------------------------------------------
def _normalise_marker(marker):
    """Turn a detected list marker into a tidy Markdown form."""
    m = marker.strip()
    digits = re.match(r"^\(?(\d{1,3})\)?[.)]?$", m)
    if digits:
        return digits.group(1) + "."
    if re.match(r"^[\u2022\u25aa\u25e6\u00b7\u2023]$", m):
        return "-"
    return m  # letters / roman numerals kept as-is


def _is_table_block(block):
    """True if every non-blank line in the block is a Markdown table row."""
    lines = [ln for ln in block.splitlines() if ln.strip()]
    return bool(lines) and all(ln.lstrip().startswith("|") for ln in lines)


def _is_heading_block(block):
    return block.lstrip().startswith("#")


def _tidy_block(block):
    """Light tidy for non-table blocks (normalise a leading unicode bullet)."""
    if _is_table_block(block):
        return block
    return LEADING_BULLET_RE.sub("- ", block)


def _postprocess(markdown):
    """Apply the local fixes on top of pymupdf4llm's output:
    strip trailing whitespace, rejoin orphan list markers, collapse runs of
    blank lines to a single blank line."""
    # 1) Strip trailing whitespace per line.
    markdown = "\n".join(line.rstrip() for line in markdown.split("\n"))

    # 2) Split into blocks on one-or-more blank lines (tables stay intact,
    #    since pymupdf4llm emits their rows on consecutive lines).
    blocks = [b for b in re.split(r"\n[ \t]*\n+", markdown.strip("\n")) if b.strip()]

    # 3) Merge a marker-only block with the (plain) block that follows it.
    merged = []
    i = 0
    n = len(blocks)
    while i < n:
        block = blocks[i].strip()
        if MARKER_ONLY_RE.match(block) and i + 1 < n:
            nxt = blocks[i + 1].strip()
            if (
                not _is_table_block(nxt)
                and not _is_heading_block(nxt)
                and not MARKER_ONLY_RE.match(nxt)
            ):
                content = " ".join(nxt.split("\n")).strip()
                merged.append(_normalise_marker(block) + " " + content)
                i += 2
                continue
        merged.append(blocks[i].strip("\n"))
        i += 1

    # 4) Tidy and reassemble with exactly one blank line between blocks.
    merged = [_tidy_block(b) for b in merged]
    markdown = "\n\n".join(merged)
    markdown = re.sub(r"\n{3,}", "\n\n", markdown)  # 2+ blank lines -> 1
    return markdown.strip() + "\n"


def convert_pdf_to_markdown_text(pdf_path):
    """Convert a PDF file to a Markdown string. Raises on failure
    (file errors, encrypted PDF, parse errors)."""
    # Light open first, for a clean error on encrypted PDFs.
    doc = fitz.open(pdf_path)
    try:
        if doc.needs_pass:
            raise ValueError("PDF is password protected")
    finally:
        doc.close()

    # Extract with pymupdf4llm (tables + headings + reading order). Suppress
    # OS-level stdout so its advisory output cannot corrupt the protocol.
    with _suppress_fd_stdout():
        raw_markdown = pymupdf4llm.to_markdown(str(pdf_path), table_strategy=TABLE_STRATEGY)

    return _postprocess(raw_markdown)


def convert_one(pdf_path):
    """Convert one PDF. Returns (ok, markdown, error_text)."""
    if fitz is None or pymupdf4llm is None:
        return (
            False,
            "",
            "Conversion library unavailable: {}\nInterpreter running this server: {}\n"
            "(If pip says the package is installed but you see this, the package "
            "is in a different Python than the one above - install into that "
            "interpreter, or point Continue's 'command' at the interpreter that "
            "has it.)".format(IMPORT_ERROR, sys.executable),
        )
    try:
        markdown = convert_pdf_to_markdown_text(str(pdf_path))
    except FileNotFoundError:
        return (False, "", "File not found")
    except PermissionError:
        return (False, "", "Permission denied reading the PDF")
    except Exception as exc:  # noqa: BLE001 - report cause to the caller
        return (False, "", "{}: {}".format(type(exc).__name__, exc))

    if not markdown.strip():
        return (False, "", "No extractable text (possibly a scanned/image-only PDF needing OCR)")
    return (True, markdown, "")


# --------------------------------------------------------------------------
# Filesystem helpers
# --------------------------------------------------------------------------
def find_pdfs(folder, recursive):
    """Sorted list of PDF Paths in folder (case-insensitive on suffix)."""
    base = Path(folder)
    iterator = base.rglob("*") if recursive else base.iterdir()
    return sorted(f for f in iterator if f.is_file() and f.suffix.lower() == ".pdf")


def md_output_path(pdf_path, input_dir, output_dir):
    """Map an input PDF to its .md path in the output folder, preserving any
    sub-folder structure (relevant when --recursive is used)."""
    rel = Path(pdf_path).resolve().relative_to(Path(input_dir).resolve())
    return Path(output_dir) / rel.with_suffix(".md")


def write_markdown(markdown, out_path):
    """Write Markdown to out_path (creating parent folders). newline='' keeps
    the bytes faithful (no CRLF rewriting); UTF-8 preserves non-ASCII."""
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8", newline="") as fh:
        fh.write(markdown)


def format_file_list(paths, limit=25):
    names = [p.name for p in paths[:limit]]
    extra = len(paths) - len(names)
    listing = "\n".join("  - {}".format(n) for n in names)
    if extra > 0:
        listing += "\n  ... and {} more".format(extra)
    return listing


# --------------------------------------------------------------------------
# Fuzzy matching (single-file tool)
# --------------------------------------------------------------------------
def normalise(text):
    text = text.lower()
    for ch in ("_", "-", ".", "(", ")", "[", "]", ",", "&", "+"):
        text = text.replace(ch, " ")
    return " ".join(text.split())


def score_name(query_norm, name_norm):
    q_tokens = query_norm.split()
    n_tokens = set(name_norm.split())
    contained = (sum(1 for t in q_tokens if t in n_tokens) / len(q_tokens)) if q_tokens else 0.0
    ratio = difflib.SequenceMatcher(None, query_norm, name_norm).ratio()
    return (contained, ratio)


def match_pdfs(query, pdfs):
    query_norm = normalise(query)
    scored = [(p, score_name(query_norm, normalise(p.stem))) for p in pdfs]
    scored.sort(key=lambda item: item[1], reverse=True)
    return scored


# --------------------------------------------------------------------------
# Tool implementations
# --------------------------------------------------------------------------
def do_convert_all(config, force=False):
    """Batch convert. Returns (text, is_error)."""
    input_dir = config["input_dir"]
    output_dir = config["output_dir"]
    recursive = config["recursive"]

    if not os.path.isdir(input_dir):
        return ("Configured input folder does not exist: {}".format(input_dir), True)

    try:
        pdfs = find_pdfs(input_dir, recursive)
    except OSError as exc:
        return ("Could not read input folder {}: {}".format(input_dir, exc), True)

    if not pdfs:
        return ("No PDF files found in {}.".format(input_dir), True)

    converted, skipped, failed = [], [], []
    for pdf in pdfs:
        out_path = md_output_path(pdf, input_dir, output_dir)
        if out_path.exists() and not force:
            skipped.append(pdf.name)
            continue
        ok, markdown, err = convert_one(pdf)
        if not ok:
            failed.append((pdf.name, err))
            continue
        try:
            write_markdown(markdown, out_path)
        except OSError as exc:
            failed.append((pdf.name, "write failed: {}".format(exc)))
            continue
        converted.append(pdf.name)

    parts = [
        "Batch conversion of {} PDF(s) in {}".format(len(pdfs), input_dir),
        "Output folder: {}".format(output_dir),
        "",
        "Converted: {}".format(len(converted)),
    ]
    for name in converted:
        parts.append("  + {}".format(name))
    parts.append("Skipped (already have .md): {}".format(len(skipped)))
    if skipped:
        parts.append("  ({})".format(", ".join(skipped) if len(skipped) <= 20
                                      else "{} files".format(len(skipped))))
    parts.append("Failed: {}".format(len(failed)))
    for name, reason in failed:
        parts.append("  ! {}: {}".format(name, reason))

    # Only an error if nothing converted AND nothing deliberately skipped.
    is_error = (not converted) and bool(failed) and (not skipped)
    return ("\n".join(parts), is_error)


def do_convert_one_fuzzy(query, config):
    """Single fuzzy convert. Returns (text, is_error)."""
    input_dir = config["input_dir"]
    output_dir = config["output_dir"]
    recursive = config["recursive"]

    if not os.path.isdir(input_dir):
        return ("Configured input folder does not exist: {}".format(input_dir), True)

    try:
        pdfs = find_pdfs(input_dir, recursive)
    except OSError as exc:
        return ("Could not read input folder {}: {}".format(input_dir, exc), True)

    if not pdfs:
        return ("No PDF files found in {}.".format(input_dir), True)

    ranked = match_pdfs(query, pdfs)
    best_path, (best_contained, best_ratio) = ranked[0]

    if best_contained == 0.0 and best_ratio < MIN_RATIO:
        return (
            "No PDF closely matching '{}' was found in {}.\nAvailable PDFs:\n{}".format(
                query, input_dir, format_file_list(pdfs)
            ),
            True,
        )

    tied = []
    for path, (contained, ratio) in ranked:
        if contained == best_contained and (best_ratio - ratio) <= AMBIGUITY_DELTA:
            tied.append(path)
        else:
            break
    if len(tied) > 1:
        return (
            "'{}' matches several PDFs about equally well. Please be more "
            "specific. Candidates:\n{}".format(query, format_file_list(tied)),
            True,
        )

    log("Converting (fuzzy): {}".format(best_path))
    ok, markdown, err = convert_one(best_path)
    if not ok:
        return ("Failed to convert '{}': {}".format(best_path.name, err), True)

    out_path = md_output_path(best_path, input_dir, output_dir)
    try:
        write_markdown(markdown, out_path)
    except OSError as exc:
        return ("Converted OK but could not write output {}: {}".format(out_path, exc), True)

    total = len(markdown)
    if total > MAX_INLINE_CHARS:
        preview = markdown[:MAX_INLINE_CHARS]
        note = "\n\n... [preview truncated: first {} of {} characters; full Markdown is in the file above]".format(
            MAX_INLINE_CHARS, total
        )
    else:
        preview = markdown
        note = ""

    header = (
        "Converted '{name}' to Markdown.\nSource: {src}\n"
        "Markdown written to: {out} ({n} characters)\n\n--- Markdown ---\n".format(
            name=best_path.name, src=best_path, out=out_path, n=total
        )
    )
    return (header + preview + note, False)


# --------------------------------------------------------------------------
# JSON-RPC plumbing
# --------------------------------------------------------------------------
TOOLS = [
    {
        "name": "convert_all_pdfs",
        "description": (
            "Convert every PDF in the configured input folder to Markdown, "
            "writing .md files into the configured output folder. PDFs that "
            "already have a matching .md in the output folder are skipped. "
            "Use this for requests like 'convert all pdfs to markdown'. Set "
            "force=true to reconvert even when the .md already exists."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "force": {
                    "type": "boolean",
                    "description": "Reconvert PDFs even if their .md already exists. Default false.",
                }
            },
            "required": [],
        },
    },
    {
        "name": "convert_pdf_to_markdown",
        "description": (
            "Convert a single PDF, identified by a rough or approximate name, "
            "to Markdown. Example query: 'procurement policy'. The .md is "
            "written to the configured output folder. If several PDFs match "
            "equally well, it lists them and asks for a more specific name."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "query": {
                    "type": "string",
                    "description": "Rough name/keywords of the PDF, e.g. 'procurement policy'.",
                }
            },
            "required": ["query"],
        },
    },
]


def make_result(msg_id, result):
    return {"jsonrpc": "2.0", "id": msg_id, "result": result}


def make_error(msg_id, code, message):
    return {"jsonrpc": "2.0", "id": msg_id, "error": {"code": code, "message": message}}


def send(message):
    sys.stdout.write(json.dumps(message, ensure_ascii=True) + "\n")
    sys.stdout.flush()


def handle_initialize(params):
    client_version = params.get("protocolVersion") or DEFAULT_PROTOCOL_VERSION
    return {
        "protocolVersion": client_version,
        "capabilities": {"tools": {}},
        "serverInfo": {"name": SERVER_NAME, "version": SERVER_VERSION},
    }


def handle_tools_call(params, config):
    name = params.get("name")
    arguments = params.get("arguments") or {}

    if name == "convert_all_pdfs":
        force = bool(arguments.get("force", False))
        text, is_error = do_convert_all(config, force=force)
    elif name == "convert_pdf_to_markdown":
        query = arguments.get("query")
        if not query or not str(query).strip():
            return {
                "content": [{"type": "text", "text": "Please provide a 'query' (a rough PDF name)."}],
                "isError": True,
            }
        text, is_error = do_convert_one_fuzzy(str(query).strip(), config)
    else:
        return {
            "content": [{"type": "text", "text": "Unknown tool: {}".format(name)}],
            "isError": True,
        }

    return {"content": [{"type": "text", "text": text}], "isError": is_error}


def handle_message(msg, config):
    if not isinstance(msg, dict):
        return
    method = msg.get("method")
    msg_id = msg.get("id")
    is_notification = "id" not in msg
    if method is None:
        return

    try:
        if method == "initialize":
            if not is_notification:
                send(make_result(msg_id, handle_initialize(msg.get("params") or {})))
        elif method == "tools/list":
            if not is_notification:
                send(make_result(msg_id, {"tools": TOOLS}))
        elif method == "tools/call":
            result = handle_tools_call(msg.get("params") or {}, config)
            if not is_notification:
                send(make_result(msg_id, result))
        elif method == "ping":
            if not is_notification:
                send(make_result(msg_id, {}))
        elif method.startswith("notifications/"):
            pass
        else:
            if not is_notification:
                send(make_error(msg_id, -32601, "Method not found: {}".format(method)))
    except Exception as exc:  # noqa: BLE001 - keep the loop alive
        log("Error handling '{}': {}".format(method, traceback.format_exc()))
        if not is_notification:
            send(make_error(msg_id, -32603, "Internal error: {}".format(exc)))


def parse_args(argv):
    parser = argparse.ArgumentParser(
        description="Self-contained MCP stdio server: convert PDFs in a folder to Markdown (with tables)."
    )
    parser.add_argument("--input-dir", required=True, help="Folder containing the source PDFs.")
    parser.add_argument("--output-dir", required=True, help="Folder to write .md files into.")
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Search sub-folders of --input-dir (sub-folder structure is mirrored in the output).",
    )
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(sys.argv[1:] if argv is None else argv)

    # Reserve stdout strictly for JSON-RPC; newline='\n' stops Windows CRLF
    # rewriting that would corrupt the line-delimited protocol.
    try:
        sys.stdout.reconfigure(encoding="utf-8", newline="\n")
        sys.stdin.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

    config = {
        "input_dir": args.input_dir,
        "output_dir": args.output_dir,
        "recursive": args.recursive,
    }
    log("interpreter (sys.executable): {}".format(sys.executable))
    log("started (input_dir={!r}, output_dir={!r}, recursive={})".format(
        args.input_dir, args.output_dir, args.recursive))
    if fitz is None or pymupdf4llm is None:
        log("WARNING: {}".format(IMPORT_ERROR))
        log("WARNING: the package is missing from the interpreter above. "
            "Install it into THIS interpreter (\"{}\" -m pip install pymupdf pymupdf4llm) "
            "or set Continue's 'command' to the interpreter that has it.".format(sys.executable))

    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            msg = json.loads(line)
        except json.JSONDecodeError as exc:
            log("Ignoring non-JSON line: {}".format(exc))
            continue
        handle_message(msg, config)

    log("stdin closed, exiting.")


if __name__ == "__main__":
    main()
