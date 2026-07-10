#!/usr/bin/env python3
"""
knowledge-base-semantic.py
==========================

A single-file MCP (Model Context Protocol) server that gives an LLM read-only,
MEANING-BASED access to a folder of your own markdown notes / policies /
reference documents. Point it at a folder, and the agent can answer questions
like "can I extend my work trip by two days, pay for my own weekend
accommodation, and fly back Monday?" by finding the relevant part of your
travel policy and reasoning over it - without you naming the file.

Transport: newline-delimited JSON-RPC 2.0 over stdio (the standard MCP stdio
transport, and what the VSCode Continue extension speaks).

DEPENDENCY
----------
Standard library only. No pip installs required.

HOW THE "SEMANTIC" PART WORKS  (read this - it explains the design)
-------------------------------------------------------------------
There are only three ways to do meaning-based retrieval:
  1. a LOCAL embedding model (needs a model file shipped to the machine);
  2. a REMOTE embeddings API (needs network access + an API key);
  3. let the LANGUAGE MODEL YOU ALREADY TALK TO do the semantic matching.

This server takes route 3. MCP has no way for a server to ask the client for
embeddings, and Continue does not support MCP "sampling" (server-initiated LLM
calls), so the server cannot invoke your model directly. Instead it leans on
the fact that the model IS the agent calling these tools: it hands the agent a
compact, structure-aware MAP of the knowledge base (every document + its
section headings + a one-line lead per section), the agent uses its own
understanding to pick - by meaning, not keywords - the document and section
that answers the question, and then it reads exactly that section in full.

Net effect: semantic retrieval, using the model you connect to, with NO local
embedding model, NO network calls, and NO extra dependencies.

TOOLS EXPOSED (all read-only)
-----------------------------
- kb_list    : list every document (with its title). Quick inventory.
- kb_outline : the MAP - every document's heading tree (+ a lead line per
               section). Call this FIRST for a topic question; choose the
               relevant document/section by meaning, then kb_read it.
- kb_search  : keyword search across all documents for a fast lookup or when
               the outline is too large; reports the matching section heading.
- kb_read    : read ONE document in full, OR just one named section of it.
               Accepts a loose name/topic and resolves it to the best file.

CONFIGURATION
-------------
The knowledge-base folder is supplied via --docs-dir (preferred) or the
KB_DOCS_DIR environment variable. In Continue's config.yaml:

    mcpServers:
      - name: kb
        command: C:\\path\\to\\python.exe
        args:
          - C:\\path\\to\\knowledge-base-semantic.py
          - --docs-dir
          - C:\\Users\\me\\knowledge-base
        env:
          PYTHONUTF8: "1"

You can run this ALONGSIDE the keyword-only knowledge-base.py (different server
name + folder) if you want both.

USAGE
-----
- As an MCP server (normal mode): launched by the MCP client with --docs-dir
  (or KB_DOCS_DIR set).
- Connectivity check (run manually first):

      python knowledge-base-semantic.py --docs-dir C:\\path\\to\\kb --check

  Lists the folder's documents (and a heading count each) to stderr, then exits.

NOTES
-----
- ALL diagnostic output goes to stderr; stdout carries only JSON-RPC.
- Set PYTHONUTF8=1 in the launching environment so non-ASCII content does not
  crash on the default Windows cp1252 codec.
- Path traversal is blocked: a requested document is always resolved INSIDE the
  configured folder; requests pointing outside it are refused.
- Headings are parsed from ATX markdown (lines starting with #..######), with
  fenced code blocks ignored so a "# comment" in a code sample is not mistaken
  for a heading.
"""

import os
import re
import sys
import json
import difflib
import argparse
import traceback


# ---------------------------------------------------------------------------
# Stream setup: force UTF-8 so non-ASCII content cannot crash output.
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
# Configuration constants
# ---------------------------------------------------------------------------

# Which file types count as knowledge-base documents.
DOC_EXTENSIONS = {".md", ".markdown", ".txt"}

# Cap on the size of a single document (or section) returned in full.
MAX_DOC_CHARS = 80000

# Cap on the whole-knowledge-base outline. If exceeded, the outline is
# truncated with a note steering the agent to narrow with kb_search or to
# request a single document by name.
OUTLINE_MAX_CHARS = 30000

# Deepest heading level shown in the outline by default (1=#, 2=##, 3=###).
OUTLINE_MAX_LEVEL = 3

# Max characters of the "lead" line shown under each heading in the outline.
LEAD_CHARS = 140

# Default number of results for a keyword search.
DEFAULT_SEARCH_RESULTS = 5

# Words that only frame a request and carry no topic meaning. Removed before
# matching so "review the procurement policy document" matches the same files
# as "procurement policy". Domain words like "policy" are deliberately kept.
STOPWORDS = {
    "the", "a", "an", "this", "that", "these", "those",
    "document", "documents", "doc", "docs", "file", "files", "section", "sections",
    "please", "review", "read", "find", "get", "show", "use", "tell",
    "reference", "references", "material", "materials",
    "about", "regarding", "on", "for", "of", "and", "with",
    "our", "my", "me", "i", "to", "in", "is", "it", "can",
}

# Set in main() once the folder is validated. Always an absolute, real path.
DOCS_DIR = None

# ATX heading, allowing up to 3 leading spaces and trailing '#'s (per CommonMark).
HEADING_RE = re.compile(r"^ {0,3}(#{1,6})\s+(.+?)\s*#*\s*$")


# ---------------------------------------------------------------------------
# Filesystem helpers
# ---------------------------------------------------------------------------

def list_documents():
    """
    Return a sorted list of absolute paths to all knowledge-base documents.
    Anything that RESOLVES outside the configured folder is excluded (e.g. a
    file symlink pointing elsewhere), so a link dropped into the folder cannot
    expose files beyond it. os.walk already declines to follow directory
    symlinks (followlinks=False is its default).
    """
    found = []
    for root, _dirs, files in os.walk(DOCS_DIR):
        for filename in files:
            if os.path.splitext(filename)[1].lower() in DOC_EXTENSIONS:
                full = os.path.join(root, filename)
                if not is_within(full, DOCS_DIR):
                    log("Excluded (resolves outside the docs folder): {0}".format(full))
                    continue
                found.append(full)
    found.sort()
    return found


def to_rel(path):
    """Relative path from DOCS_DIR, using forward slashes for stable display."""
    return os.path.relpath(path, DOCS_DIR).replace("\\", "/")


def is_within(path, base):
    """
    True if `path` resolves to a location inside `base`. Guards against path
    traversal (e.g. '..\\..\\secrets.txt') and cross-drive escapes on Windows.
    """
    try:
        real_path = os.path.realpath(path)
        real_base = os.path.realpath(base)
        # commonpath raises ValueError across different drives -> treated as outside.
        return os.path.commonpath([real_path, real_base]) == real_base
    except Exception:
        return False


def read_text(path):
    """Read a document as text (untruncated)."""
    with open(path, "r", encoding="utf-8", errors="replace") as handle:
        return handle.read()


def truncate(text, limit=MAX_DOC_CHARS):
    """Truncate very long text with a visible note."""
    if len(text) > limit:
        return text[:limit] + "\n\n[... truncated at {0} characters ...]".format(limit)
    return text


def doc_title(path):
    """
    A short human-readable title for a document: the first markdown heading if
    present, otherwise the first non-empty line, otherwise empty.
    """
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as handle:
            for _ in range(40):  # only inspect the top of the file
                line = handle.readline()
                if not line:
                    break
                stripped = line.strip()
                if stripped.startswith("#"):
                    return stripped.lstrip("#").strip()[:100]
                if stripped:
                    return stripped[:100]
    except Exception:
        pass
    return ""


# ---------------------------------------------------------------------------
# Markdown structure helpers (headings, leads, section extraction)
# ---------------------------------------------------------------------------

def parse_headings(content):
    """
    Return [{level, text, line}, ...] for every ATX heading in `content`, in
    document order. Lines inside fenced code blocks are ignored so a '#'
    comment in a code sample is not mistaken for a heading.
    """
    heads = []
    in_fence = False
    fence = None
    for idx, raw in enumerate(content.splitlines()):
        stripped = raw.strip()
        if stripped.startswith("```") or stripped.startswith("~~~"):
            marker = stripped[:3]
            if not in_fence:
                in_fence, fence = True, marker
            elif stripped.startswith(fence):
                in_fence, fence = False, None
            continue
        if in_fence:
            continue
        match = HEADING_RE.match(raw)
        if match:
            heads.append({"level": len(match.group(1)), "text": match.group(2).strip(), "line": idx})
    return heads


def lead_after(lines, idx):
    """First non-empty prose line shortly after heading line `idx`, trimmed."""
    for j in range(idx + 1, min(idx + 10, len(lines))):
        stripped = lines[j].strip()
        if not stripped:
            continue
        if stripped.startswith("```") or stripped.startswith("~~~"):
            return ""
        if HEADING_RE.match(lines[j]):
            return ""
        return re.sub(r"\s+", " ", stripped)[:LEAD_CHARS]
    return ""


def section_for_line(heads, line_idx):
    """Text of the nearest heading at or before `line_idx`, or '' if none."""
    current = ""
    for head in heads:
        if head["line"] <= line_idx:
            current = head["text"]
        else:
            break
    return current


def extract_section(content, section_query):
    """
    Resolve `section_query` to one heading and return its slice of the document.

    Returns (status, payload):
      - ("ok", (heading_text, section_text))
      - ("ambiguous", [heading_text, ...])   several headings match equally
      - ("notfound", [heading_text, ...])    no good match; here are the headings
    """
    heads = parse_headings(content)
    if not heads:
        return "notfound", []

    lines = content.splitlines()
    query = section_query.strip().lower()
    query_tokens = set(meaningful_tokens(section_query))

    scored = []
    for head in heads:
        htext = head["text"].lower()
        htokens = set(meaningful_tokens(head["text"]))
        overlap = len(query_tokens & htokens)
        ratio = difflib.SequenceMatcher(None, query, htext).ratio()
        exact = 3.0 if htext == query else 0.0
        contains = 1.0 if (query and (query in htext or htext in query)) else 0.0
        scored.append((exact + contains + overlap + ratio, ratio, head))

    scored.sort(key=lambda item: (item[0], item[1]), reverse=True)
    best_score = scored[0][0]
    if best_score < 0.6:
        return "notfound", [head["text"] for head in heads]

    # Ambiguous only if a runner-up is genuinely close AND names a different heading.
    close = [item for item in scored if best_score - item[0] < 0.2 and item[0] >= 0.6]
    distinct = []
    for _score, _ratio, head in close:
        if head["text"] not in distinct:
            distinct.append(head["text"])
    if len(distinct) > 1:
        return "ambiguous", distinct[:8]

    best = scored[0][2]
    start, level = best["line"], best["level"]
    end = len(lines)
    for head in heads:
        if head["line"] > start and head["level"] <= level:
            end = head["line"]
            break
    section_text = "\n".join(lines[start:end]).strip()
    return "ok", (best["text"], section_text)


# ---------------------------------------------------------------------------
# Text / matching helpers
# ---------------------------------------------------------------------------

def tokenise(text):
    """Lowercase and split on any non-alphanumeric run."""
    return [tok for tok in re.split(r"[^a-z0-9]+", text.lower()) if tok]


def meaningful_tokens(text):
    """Topic-bearing tokens: lowercase, framing stopwords and 1-char tokens removed."""
    return [tok for tok in tokenise(text) if tok not in STOPWORDS and len(tok) > 1]


def path_tokens(path):
    """Set of tokens from a document's relative path, excluding its extension."""
    rel_no_ext = os.path.splitext(to_rel(path))[0]
    return set(tokenise(rel_no_ext))


def make_snippet(content, query_tokens):
    """First line of content containing any query token, trimmed for display."""
    for line in content.splitlines():
        low = line.lower()
        if any(tok in low for tok in query_tokens):
            snippet = line.strip()
            if snippet:
                return snippet[:240] + ("..." if len(snippet) > 240 else "")
    return "(matched on filename)"


# ---------------------------------------------------------------------------
# Document resolution (loose name -> a specific file)
# ---------------------------------------------------------------------------

def try_direct_path(name):
    """
    If `name` looks like an actual path/filename under DOCS_DIR, return its
    absolute path. Tries the name as-is and with each known extension appended.
    Returns None if no safe, existing file matches.
    """
    candidate = name.replace("/", os.sep).replace("\\", os.sep)
    attempts = [candidate]
    if not os.path.splitext(candidate)[1]:
        attempts.extend(candidate + ext for ext in DOC_EXTENSIONS)
    for attempt in attempts:
        full = os.path.normpath(os.path.join(DOCS_DIR, attempt))
        if is_within(full, DOCS_DIR) and os.path.isfile(full):
            return full
    return None


def resolve_document(name):
    """
    Resolve a loose name/topic to document path(s).

    Returns (paths, unambiguous):
      - ([path], True)        : one clear match -> safe to read in full
      - ([p1, p2, ...], False): several plausible matches -> ask to disambiguate
      - ([], False)           : nothing matched
    """
    direct = try_direct_path(name)
    if direct is not None:
        return [direct], True

    query = meaningful_tokens(name)
    documents = list_documents()
    if not documents:
        return [], False

    query_str = " ".join(query)
    scored = []
    for path in documents:
        overlap = sum(1 for tok in query if tok in path_tokens(path))
        stem = os.path.splitext(os.path.basename(path))[0]
        stem_words = stem.replace("-", " ").replace("_", " ").lower()
        ratio = difflib.SequenceMatcher(None, query_str, stem_words).ratio()
        scored.append((overlap, ratio, path))

    scored.sort(key=lambda item: (item[0], item[1]), reverse=True)
    best_overlap, best_ratio, _best_path = scored[0]

    if best_overlap == 0 and best_ratio < 0.5:
        return [], False

    top = [item for item in scored if item[0] == best_overlap]
    if len(top) == 1:
        return [top[0][2]], True

    top.sort(key=lambda item: item[1], reverse=True)
    if top[0][1] - top[1][1] >= 0.15:
        return [top[0][2]], True

    return [item[2] for item in top[:5]], False


# ---------------------------------------------------------------------------
# Tool implementations (each returns a human-readable text string)
# ---------------------------------------------------------------------------

def tool_kb_list(_args):
    documents = list_documents()
    if not documents:
        return "No documents found in {0}.".format(DOCS_DIR)
    lines = []
    for path in documents:
        title = doc_title(path)
        rel = to_rel(path)
        lines.append("- {0} : {1}".format(rel, title) if title else "- {0}".format(rel))
    return "Knowledge base documents in {0} ({1} found):\n{2}".format(
        DOCS_DIR, len(documents), "\n".join(lines)
    )


def tool_kb_outline(args):
    name = (args.get("name") or "").strip()
    try:
        max_level = int(args.get("max_level", OUTLINE_MAX_LEVEL))
    except (TypeError, ValueError):
        max_level = OUTLINE_MAX_LEVEL
    max_level = max(1, min(6, max_level))
    include_leads = bool(args.get("include_leads", True))

    if name:
        paths, unambiguous = resolve_document(name)
        if not paths:
            return "No document matched '{0}'. Use kb_list to see what is available.".format(name)
        if not unambiguous:
            listing = "\n".join("- " + to_rel(p) for p in paths)
            return "'{0}' matched several documents. Name one specifically:\n{1}".format(name, listing)
        documents = [paths[0]]
    else:
        documents = list_documents()

    if not documents:
        return "No documents found in {0}.".format(DOCS_DIR)

    blocks = []
    total = 0
    truncated = False
    for path in documents:
        try:
            content = read_text(path)
        except Exception:
            continue
        lines = content.splitlines()
        heads = [h for h in parse_headings(content) if h["level"] <= max_level]

        block = ["**{0}**".format(to_rel(path))]
        if not heads:
            block.append("  (no headings)")
        for head in heads:
            indent = "  " * (head["level"] - 1)
            entry = "{indent}- {text}".format(indent=indent, text=head["text"])
            if include_leads:
                lead = lead_after(lines, head["line"])
                if lead:
                    entry += "  —  {0}".format(lead)
            block.append(entry)

        chunk = "\n".join(block)
        if blocks and total + len(chunk) > OUTLINE_MAX_CHARS:
            truncated = True
            break
        blocks.append(chunk)
        total += len(chunk)

    intro = (
        "Knowledge base outline ({0} document(s)). Choose - by MEANING, not just keywords - "
        "the document and section that answers the question, then call kb_read(name, section) "
        "to read that section in full.\n\n".format(len(blocks))
    )
    note = ""
    if truncated:
        note = ("\n\n[... outline truncated at {0} chars. Narrow with kb_search, or call "
                "kb_outline with a single document 'name'. ...]".format(OUTLINE_MAX_CHARS))
    return intro + "\n\n".join(blocks) + note


def tool_kb_search(args):
    query = (args.get("query") or "").strip()
    if not query:
        return "Error: 'query' is required."
    try:
        max_results = int(args.get("max_results", DEFAULT_SEARCH_RESULTS))
    except (TypeError, ValueError):
        max_results = DEFAULT_SEARCH_RESULTS

    query_tokens = meaningful_tokens(query)
    if not query_tokens:
        return "Error: the query had no searchable terms after removing common words."

    results = []
    for path in list_documents():
        path_overlap = sum(1 for tok in query_tokens if tok in path_tokens(path))
        try:
            content = read_text(path)
        except Exception:
            continue  # an unreadable file should not sink the whole search
        lowered = content.lower()
        content_hits = sum(lowered.count(tok) for tok in query_tokens)
        if path_overlap == 0 and content_hits == 0:
            continue

        # Locate the first matching line and the heading it sits under, so the
        # agent can jump straight to kb_read(name, section).
        section = ""
        lines = content.splitlines()
        heads = parse_headings(content)
        for line_idx, line in enumerate(lines):
            low = line.lower()
            if any(tok in low for tok in query_tokens):
                section = section_for_line(heads, line_idx)
                break

        relevance = path_overlap * 100 + content_hits
        results.append((relevance, path, content_hits, path_overlap, section,
                        make_snippet(content, query_tokens)))

    if not results:
        return "No documents matched '{0}'.".format(query)

    results.sort(key=lambda item: item[0], reverse=True)
    lines_out = []
    for _relevance, path, hits, overlap, section, snippet in results[:max_results]:
        section_note = "  ·  section: {0}".format(section) if section else ""
        lines_out.append(
            "- {rel}{sec}  (filename match: {fm}, content hits: {ch})\n    {snip}".format(
                rel=to_rel(path),
                sec=section_note,
                fm=("yes" if overlap else "no"),
                ch=hits,
                snip=snippet,
            )
        )
    return (
        "Found {n} document(s) for '{q}', most relevant first.\n"
        "Read one with kb_read(name[, section]).\n{body}".format(
            n=len(lines_out), q=query, body="\n".join(lines_out)
        )
    )


def tool_kb_read(args):
    name = (args.get("name") or "").strip()
    if not name:
        return "Error: 'name' is required (a document name or topic)."
    section = (args.get("section") or "").strip()

    paths, unambiguous = resolve_document(name)
    if not paths:
        return "No document matched '{0}'. Use kb_list to see what is available.".format(name)
    if not unambiguous:
        listing = "\n".join("- " + to_rel(p) for p in paths)
        return ("'{0}' matched several documents. Call kb_read again naming one "
                "specifically:\n{1}".format(name, listing))

    path = paths[0]
    try:
        content = read_text(path)
    except Exception as exc:
        return "Error: could not read '{0}' ({1}).".format(to_rel(path), exc)

    if not section:
        return "Document: {0}\n\n{1}".format(to_rel(path), truncate(content))

    status, payload = extract_section(content, section)
    if status == "ok":
        heading, section_text = payload
        return "Document: {0}\nSection : {1}\n\n{2}".format(
            to_rel(path), heading, truncate(section_text))
    if status == "ambiguous":
        listing = "\n".join("- " + text for text in payload)
        return ("Several sections in {0} match '{1}'. Call kb_read again with the exact "
                "heading:\n{2}".format(to_rel(path), section, listing))
    # notfound
    if payload:
        listing = "\n".join("- " + text for text in payload)
        return ("No section matching '{0}' in {1}. Available headings:\n{2}\n\n"
                "(Call kb_read without 'section' to read the whole document.)".format(
                    section, to_rel(path), listing))
    return ("{0} has no headings to select a section from. Call kb_read without "
            "'section' to read the whole document.".format(to_rel(path)))


# ---------------------------------------------------------------------------
# MCP tool registry
# ---------------------------------------------------------------------------

TOOLS = [
    {
        "name": "kb_list",
        "description": (
            "List every document in the knowledge base, with its title. Use this "
            "to see what exists. For a topic/question, prefer kb_outline."
        ),
        "inputSchema": {"type": "object", "properties": {}},
    },
    {
        "name": "kb_outline",
        "description": (
            "Return a structural MAP of the knowledge base: each document and its "
            "section headings, with a one-line lead under each. CALL THIS FIRST for "
            "a topic or policy question. Then use your own understanding to pick - by "
            "MEANING, not keyword overlap - the document and section that answers it "
            "(e.g. a question about extending a trip and paying for your own weekend "
            "accommodation maps to a travel policy's trip-extension / personal-travel "
            "section), and read it with kb_read(name, section)."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "name": {
                    "type": "string",
                    "description": (
                        "Optional: outline only ONE document (a filename or loose "
                        "name). Omit to map the whole knowledge base."
                    ),
                },
                "max_level": {
                    "type": "integer",
                    "description": "Deepest heading level to show, 1-6 (default 3).",
                },
                "include_leads": {
                    "type": "boolean",
                    "description": "Show a one-line lead under each heading (default true).",
                },
            },
        },
    },
    {
        "name": "kb_search",
        "description": (
            "Keyword search across all documents; returns a ranked list with the "
            "matching section heading and a snippet. Use this as a fast lookup, or to "
            "narrow a large knowledge base before kb_outline / kb_read. For meaning-based "
            "questions where the wording may not match the document's, prefer kb_outline."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "query": {"type": "string", "description": "Words to look for, e.g. 'travel accommodation'."},
                "max_results": {"type": "integer", "description": "Maximum documents to return (default 5)."},
            },
            "required": ["query"],
        },
    },
    {
        "name": "kb_read",
        "description": (
            "Read the FULL text of ONE document, or just ONE section of it. Pass "
            "'section' with a heading (exact or loose) to read only that section - "
            "the most token-efficient way to pull the relevant clause after kb_outline "
            "or kb_search. Omit 'section' to read the whole document. 'name' accepts a "
            "filename, a relative path, or a descriptive name like 'travel policy'."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "name": {
                    "type": "string",
                    "description": "The document to read: a filename, relative path, or descriptive name.",
                },
                "section": {
                    "type": "string",
                    "description": (
                        "Optional heading of the section to read (from kb_outline / "
                        "kb_search). Matched loosely; includes its subsections."
                    ),
                },
            },
            "required": ["name"],
        },
    },
]

TOOL_DISPATCH = {
    "kb_list": tool_kb_list,
    "kb_outline": tool_kb_outline,
    "kb_search": tool_kb_search,
    "kb_read": tool_kb_read,
}


# ---------------------------------------------------------------------------
# JSON-RPC / MCP plumbing
# ---------------------------------------------------------------------------

PROTOCOL_VERSION_DEFAULT = "2024-11-05"
SERVER_INFO = {"name": "knowledge-base-semantic", "version": "1.0.1"}

# Sent in the initialize response so a capable client can prime the model on how
# to use these tools for meaning-based retrieval.
SERVER_INSTRUCTIONS = (
    "This server exposes a personal knowledge base of markdown documents. For a "
    "topic or policy question, first call kb_outline to see every document and its "
    "sections, then use your own understanding to choose - by meaning, not keyword "
    "overlap - the document and section that answers the question, and read it with "
    "kb_read(name, section). Use kb_search for a quick keyword lookup, and kb_list "
    "for a plain inventory. All access is read-only."
)


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
            "instructions": SERVER_INSTRUCTIONS,
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
            return rpc_result(req_id, text_content("Knowledge-base tool error: {0}".format(exc), is_error=True))

    if is_notification:
        return None
    return rpc_error(req_id, -32601, "Method not found: {0}".format(method))


def run_server():
    """Main stdio loop: read newline-delimited JSON-RPC, dispatch, respond."""
    log("knowledge-base-semantic server started (stdio). Folder: {0}".format(DOCS_DIR))
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
        log("knowledge-base-semantic server stopped.")


def run_check():
    """List the configured folder's documents (+ heading counts) to stderr, then exit."""
    documents = list_documents()
    log("Knowledge base folder: {0}".format(DOCS_DIR))
    log("Documents found      : {0}".format(len(documents)))
    for path in documents[:50]:
        try:
            heads = len(parse_headings(read_text(path)))
        except Exception:
            heads = "?"
        log("  - {0}  ({1} headings)".format(to_rel(path), heads))
    if len(documents) > 50:
        log("  ... and {0} more".format(len(documents) - 50))
    log("CHECK OK" if documents else "CHECK OK (folder is readable but empty)")
    return 0


def main():
    global DOCS_DIR

    parser = argparse.ArgumentParser(
        description=(
            "Read-only MCP server for meaning-based retrieval over a folder of local "
            "markdown documents. The connected model does the semantic selection over "
            "a structural map the server provides - no local embedding model, no network. "
            "With no check flag it runs as an stdio MCP server."
        )
    )
    parser.add_argument(
        "--docs-dir",
        default=os.environ.get("KB_DOCS_DIR"),
        help="Folder containing knowledge-base documents. Falls back to the "
             "KB_DOCS_DIR environment variable.",
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="List the folder's documents (and heading counts) to stderr, then exit.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version="knowledge-base-semantic {0}".format(SERVER_INFO["version"]),
    )
    args = parser.parse_args()

    if not args.docs_dir:
        log("FATAL: no knowledge-base folder set. Pass --docs-dir or set KB_DOCS_DIR.")
        sys.exit(2)
    if not os.path.isdir(args.docs_dir):
        log("FATAL: knowledge-base folder does not exist or is not a directory: {0}".format(args.docs_dir))
        sys.exit(2)

    DOCS_DIR = os.path.realpath(args.docs_dir)

    if args.check:
        sys.exit(run_check())
    run_server()


if __name__ == "__main__":
    main()
