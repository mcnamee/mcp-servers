#!/usr/bin/env python3
"""
knowledge-base.py
================

A single-file MCP (Model Context Protocol) server that gives an LLM
read-only access to a folder of local reference documents (markdown / text).
Drop files (markdown files) into a reference folder, point this server at that
folder, and the agent can find and read them by natural-language topic or by a
specific document name.

Transport: newline-delimited JSON-RPC 2.0 over stdio (the standard MCP stdio
transport, and what the VSCode Continue extension speaks).

DEPENDENCY
----------
Standard library only. No pip installs required.

WHAT IT READS
-------------
Files with extensions .md, .markdown, or .txt, found anywhere under the
configured folder (subfolders are searched too).

TOOLS EXPOSED (all read-only)
-----------------------------
- reference_list   : list every available reference document (+ its title)
- reference_search : search ACROSS all documents for a topic; returns a ranked
                     list with snippets. Use for "material about <topic>".
- reference_get    : read ONE specific document in full. Accepts a loose name
                     or topic and resolves it to the best-matching file.
                     Use for "review the <named> document".

CONFIGURATION
-------------
The reference folder is supplied via --docs-dir (preferred) or the
REFERENCE_DOCS_DIR environment variable. In Continue's config.yaml:

    mcpServers:
      - name: reference
        command: python
        args:
          - C:\\path\\to\\reference_mcp.py
          - --docs-dir
          - C:\\reference-docs
        env:
          PYTHONUTF8: "1"

USAGE
-----
- As an MCP server (normal mode): launched by the MCP client. Run with the
  --docs-dir argument (or REFERENCE_DOCS_DIR set).
- Connectivity check (run manually on the endpoint first):

      python reference_mcp.py --docs-dir C:\\reference-docs --check

  Lists the folder contents to stderr and exits, so you can confirm the right
  files are visible in a single transfer cycle.

NOTES
-----
- ALL diagnostic output goes to stderr; stdout carries only JSON-RPC.
- Set PYTHONUTF8=1 in the launching environment so non-ASCII content does not
  crash on the default Windows cp1252 codec.
- Path traversal is blocked: a requested document is always resolved INSIDE
  the configured folder; requests pointing outside it are refused.
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

# Which file types count as reference documents.
DOC_EXTENSIONS = {".md", ".markdown", ".txt"}

# Cap on the size of a single document returned in full. Policy documents can
# be long; if your local model has a small context window you may want to lower
# this and rely on reference_search to point at the relevant document instead.
MAX_DOC_CHARS = 80000

# Default number of results for a topic search.
DEFAULT_SEARCH_RESULTS = 5

# Words that only frame a request and carry no topic meaning. Removed before
# matching so "review the procurement policy document" matches the same files
# as "procurement policy". Domain words like "policy" are deliberately kept.
STOPWORDS = {
    "the", "a", "an", "this", "that", "these", "those",
    "document", "documents", "doc", "docs", "file", "files",
    "please", "review", "read", "find", "get", "show", "use",
    "reference", "references", "material", "materials",
    "about", "regarding", "on", "for", "of", "and", "with",
    "our", "my", "me", "i", "to", "in", "is", "it",
}

# Set in main() once the folder is validated. Always an absolute, real path.
DOCS_DIR = None


# ---------------------------------------------------------------------------
# Filesystem helpers
# ---------------------------------------------------------------------------

def list_documents():
    """Return a sorted list of absolute paths to all reference documents."""
    found = []
    for root, _dirs, files in os.walk(DOCS_DIR):
        for filename in files:
            if os.path.splitext(filename)[1].lower() in DOC_EXTENSIONS:
                found.append(os.path.join(root, filename))
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


def read_doc(path):
    """Read a document as text, truncating very long files with a note."""
    with open(path, "r", encoding="utf-8", errors="replace") as handle:
        text = handle.read()
    if len(text) > MAX_DOC_CHARS:
        text = text[:MAX_DOC_CHARS] + (
            "\n\n[... document truncated at {0} characters ...]".format(MAX_DOC_CHARS)
        )
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
#
# This deliberately hides filename complexity from the model: the model can say
# "the procurement policy document" and the tool figures out it means
# policy-procurement.md. Mirrors the title-to-ID resolution pattern.
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
    # 1) Exact path / filename wins immediately.
    direct = try_direct_path(name)
    if direct is not None:
        return [direct], True

    # 2) Score every document by token overlap, with a fuzzy ratio tiebreaker.
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

    # Best by token overlap first, then fuzzy ratio.
    scored.sort(key=lambda item: (item[0], item[1]), reverse=True)
    best_overlap, best_ratio, _best_path = scored[0]

    # Nothing meaningful matched.
    if best_overlap == 0 and best_ratio < 0.5:
        return [], False

    # All documents sharing the top overlap score are the contenders.
    top = [item for item in scored if item[0] == best_overlap]
    if len(top) == 1:
        return [top[0][2]], True

    # Tie on overlap: accept the front-runner only if its fuzzy ratio is clearly ahead.
    top.sort(key=lambda item: item[1], reverse=True)
    if top[0][1] - top[1][1] >= 0.15:
        return [top[0][2]], True

    # Genuinely ambiguous: hand back the candidates.
    return [item[2] for item in top[:5]], False


# ---------------------------------------------------------------------------
# Tool implementations (each returns a human-readable text string)
# ---------------------------------------------------------------------------

def tool_reference_list(_args):
    documents = list_documents()
    if not documents:
        return "No reference documents found in {0}.".format(DOCS_DIR)
    lines = []
    for path in documents:
        title = doc_title(path)
        rel = to_rel(path)
        lines.append("- {0} : {1}".format(rel, title) if title else "- {0}".format(rel))
    return "Reference documents in {0} ({1} found):\n{2}".format(
        DOCS_DIR, len(documents), "\n".join(lines)
    )


def tool_reference_search(args):
    query = (args.get("query") or "").strip()
    if not query:
        return "Error: 'query' is required."
    max_results = int(args.get("max_results", DEFAULT_SEARCH_RESULTS))

    query_tokens = meaningful_tokens(query)
    if not query_tokens:
        return "Error: the query had no searchable terms after removing common words."

    results = []
    for path in list_documents():
        path_overlap = sum(1 for tok in query_tokens if tok in path_tokens(path))
        try:
            with open(path, "r", encoding="utf-8", errors="replace") as handle:
                content = handle.read()
        except Exception:
            continue  # unreadable file should not sink the whole search
        lowered = content.lower()
        content_hits = sum(lowered.count(tok) for tok in query_tokens)
        if path_overlap == 0 and content_hits == 0:
            continue
        # Filename matches weigh far more heavily than scattered body mentions.
        relevance = path_overlap * 100 + content_hits
        results.append((relevance, path, content_hits, path_overlap, make_snippet(content, query_tokens)))

    if not results:
        return "No reference documents matched '{0}'.".format(query)

    results.sort(key=lambda item: item[0], reverse=True)
    lines = []
    for _relevance, path, hits, overlap, snippet in results[:max_results]:
        lines.append(
            "- {rel}  (filename match: {fm}, content hits: {ch})\n    {snip}".format(
                rel=to_rel(path),
                fm=("yes" if overlap else "no"),
                ch=hits,
                snip=snippet,
            )
        )
    return (
        "Found {n} matching document(s) for '{q}', most relevant first.\n"
        "Use reference_get with a document path to read one in full.\n{body}".format(
            n=len(lines), q=query, body="\n".join(lines)
        )
    )


def tool_reference_get(args):
    name = (args.get("name") or "").strip()
    if not name:
        return "Error: 'name' is required (a document name or topic)."

    paths, unambiguous = resolve_document(name)
    if not paths:
        return (
            "No reference document matched '{0}'. "
            "Use reference_list to see what is available.".format(name)
        )
    if not unambiguous:
        listing = "\n".join("- " + to_rel(p) for p in paths)
        return (
            "'{0}' matched several documents. Call reference_get again naming one "
            "specifically:\n{1}".format(name, listing)
        )

    path = paths[0]
    try:
        content = read_doc(path)
    except Exception as exc:
        return "Error: could not read '{0}' ({1}).".format(to_rel(path), exc)
    return "Document: {0}\n\n{1}".format(to_rel(path), content)


# ---------------------------------------------------------------------------
# MCP tool registry
# ---------------------------------------------------------------------------

TOOLS = [
    {
        "name": "reference_list",
        "description": (
            "List every available reference document, with its title. Use this "
            "to discover what reference material exists before searching or reading."
        ),
        "inputSchema": {"type": "object", "properties": {}},
    },
    {
        "name": "reference_search",
        "description": (
            "Search ACROSS all reference documents for a topic and return a "
            "ranked list of relevant documents with short snippets. Use this when "
            "the user asks for material 'about' or 'on' a subject, or when you are "
            "unsure which document is relevant (it may return several). Follow up "
            "with reference_get to read a chosen document in full."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "query": {
                    "type": "string",
                    "description": "The topic to search for, e.g. 'procurement policy'.",
                },
                "max_results": {
                    "type": "integer",
                    "description": "Maximum documents to return (default 5).",
                },
            },
            "required": ["query"],
        },
    },
    {
        "name": "reference_get",
        "description": (
            "Read the FULL text of ONE specific reference document. Use this when "
            "the user refers to a particular named document, e.g. 'the procurement "
            "policy document'. Accepts a loose name or topic and resolves it to the "
            "best-matching file. If several documents match equally, it returns the "
            "candidates so you can call again naming one specifically."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "name": {
                    "type": "string",
                    "description": (
                        "The document to read: a filename, a relative path, or a "
                        "descriptive name such as 'procurement policy'."
                    ),
                },
            },
            "required": ["name"],
        },
    },
]

TOOL_DISPATCH = {
    "reference_list": tool_reference_list,
    "reference_search": tool_reference_search,
    "reference_get": tool_reference_get,
}


# ---------------------------------------------------------------------------
# JSON-RPC / MCP plumbing
# ---------------------------------------------------------------------------

PROTOCOL_VERSION_DEFAULT = "2024-11-05"
SERVER_INFO = {"name": "reference-mcp", "version": "1.0.0"}


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
            return rpc_result(req_id, text_content("Reference tool error: {0}".format(exc), is_error=True))

    if is_notification:
        return None
    return rpc_error(req_id, -32601, "Method not found: {0}".format(method))


def run_server():
    """Main stdio loop: read newline-delimited JSON-RPC, dispatch, respond."""
    log("reference-mcp server started (stdio). Folder: {0}".format(DOCS_DIR))
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
        log("reference-mcp server stopped.")


def run_check():
    """List the configured folder's documents to stderr, then exit."""
    documents = list_documents()
    log("Reference folder: {0}".format(DOCS_DIR))
    log("Documents found : {0}".format(len(documents)))
    for path in documents[:50]:
        log("  - {0}".format(to_rel(path)))
    if len(documents) > 50:
        log("  ... and {0} more".format(len(documents) - 50))
    log("CHECK OK" if documents else "CHECK OK (folder is readable but empty)")
    return 0


def main():
    global DOCS_DIR

    parser = argparse.ArgumentParser(
        description=(
            "Read-only MCP server exposing a folder of local reference documents "
            "(.md/.markdown/.txt). With no check flag it runs as an stdio MCP server."
        )
    )
    parser.add_argument(
        "--docs-dir",
        default=os.environ.get("REFERENCE_DOCS_DIR"),
        help="Folder containing reference documents. Falls back to the "
             "REFERENCE_DOCS_DIR environment variable.",
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="List the folder's documents to stderr, then exit.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version="reference-mcp {0}".format(SERVER_INFO["version"]),
    )
    args = parser.parse_args()

    # Validate the folder before doing anything else; fail loudly on stderr.
    if not args.docs_dir:
        log("FATAL: no reference folder set. Pass --docs-dir or set REFERENCE_DOCS_DIR.")
        sys.exit(2)
    if not os.path.isdir(args.docs_dir):
        log("FATAL: reference folder does not exist or is not a directory: {0}".format(args.docs_dir))
        sys.exit(2)

    DOCS_DIR = os.path.realpath(args.docs_dir)

    if args.check:
        sys.exit(run_check())
    run_server()


if __name__ == "__main__":
    main()
