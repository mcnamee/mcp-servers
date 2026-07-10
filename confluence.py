#!/usr/bin/env python3
"""
confluence_mcp.py - A single-file MCP (Model Context Protocol) server for
querying Confluence Data Center (tested against the 9.x v1 REST API) using
only the Python 3 standard library.

It speaks MCP over stdio (newline-delimited JSON-RPC 2.0), which is what the
Continue VSCode extension launches for a `type: stdio` server. No third-party
packages are required.

Tools exposed (read-only / query):
  - confluence_search           : free-text search for pages
  - confluence_search_cql       : advanced search using raw CQL
  - confluence_get_page         : fetch one page by numeric ID (with body text)
  - confluence_get_page_by_title: fetch one page by exact title + space key

Configuration is read from environment variables (the natural fit for
Continue's `env:` block); non-secret settings can be overridden by
command-line arguments. CREDENTIALS ARE ENV-VAR ONLY - there are no
--token/--user/--password flags, because command-line arguments are visible
to other local users in process listings:

  CONFLUENCE_BASE_URL   e.g. https://confluence.internal.example.com
                        (include any context path, no trailing slash)
  CONFLUENCE_TOKEN      Personal Access Token (preferred; sent as Bearer)
  CONFLUENCE_USER       username   } basic-auth fallback if no token is given
  CONFLUENCE_PASSWORD   password   }
  CONFLUENCE_VERIFY_SSL "false" to disable TLS verification (default: verify)
  CONFLUENCE_CA_CERT    path to a PEM CA bundle for an internal CA
  CONFLUENCE_TIMEOUT    request timeout in seconds (default: 30)
  CONFLUENCE_MAX_BODY   truncate page bodies to N chars (0 = unlimited, default 0)
                        (this limit applies only to the text returned to the
                        model; files saved to CONFLUENCE_KB_DIR are never truncated)
  CONFLUENCE_KB_DIR     if set, every page that is read is also saved as a
                        Markdown file into this folder, for feeding a local RAG
                        knowledge base. Files are named 'Confluence - <title>.md'
                        and overwritten if they already exist. Leave unset to
                        disable saving.

Diagnostic output goes ONLY to stderr. stdout is reserved for the JSON-RPC
stream - writing anything else there would corrupt the protocol.
"""

import argparse
import base64
import datetime
import html.parser
import json
import os
import re
import ssl
import sys
import urllib.error
import urllib.parse
import urllib.request

SERVER_NAME = "confluence-mcp"
SERVER_VERSION = "1.2.0"
# Protocol version we default to if the client does not send one. We echo the
# client's requested version when possible (see handle_initialize) so that we
# stay compatible with whatever the host negotiated.
DEFAULT_PROTOCOL_VERSION = "2024-11-05"

# JSON-RPC error codes (subset we use)
PARSE_ERROR = -32700
INVALID_REQUEST = -32600
METHOD_NOT_FOUND = -32601
INTERNAL_ERROR = -32603


def log(*args):
    """Write a diagnostic line to stderr (never stdout)."""
    print("[confluence-mcp]", *args, file=sys.stderr, flush=True)


# ---------------------------------------------------------------------------
# HTML -> plain text
# ---------------------------------------------------------------------------
class _TextExtractor(html.parser.HTMLParser):
    """
    Minimal HTML/XHTML to plain-text converter.

    Confluence "storage format" bodies are XHTML with extra <ac:...> macro
    tags. We don't try to interpret macros; we just keep the readable text and
    insert line breaks around block-level elements so the result is legible.
    convert_charrefs=True (the default) means entities like &amp; are decoded
    for us and arrive via handle_data.
    """

    _BLOCK_TAGS = {
        "p", "br", "div", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6",
        "table", "ul", "ol", "blockquote", "pre", "section", "header",
        "footer", "article",
    }
    _SKIP_TAGS = {"script", "style"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self._parts = []
        self._skip_depth = 0

    def handle_starttag(self, tag, attrs):
        if tag in self._SKIP_TAGS:
            self._skip_depth += 1
        if tag in self._BLOCK_TAGS:
            self._parts.append("\n")

    def handle_endtag(self, tag):
        if tag in self._SKIP_TAGS and self._skip_depth > 0:
            self._skip_depth -= 1
        if tag in self._BLOCK_TAGS:
            self._parts.append("\n")

    def handle_data(self, data):
        if self._skip_depth == 0:
            self._parts.append(data)

    def get_text(self):
        text = "".join(self._parts)
        # Collapse runs of blank lines and trim trailing spaces per line.
        lines = [ln.rstrip() for ln in text.splitlines()]
        out = []
        blank = False
        for ln in lines:
            if ln.strip() == "":
                if not blank:
                    out.append("")
                blank = True
            else:
                out.append(ln)
                blank = False
        return "\n".join(out).strip()


def html_to_text(raw):
    """Convert an HTML/XHTML string to plain text, defensively."""
    if not raw:
        return ""
    parser = _TextExtractor()
    try:
        parser.feed(raw)
        parser.close()
        return parser.get_text()
    except Exception:
        # If parsing somehow fails, fall back to returning the raw string
        # rather than losing the content entirely.
        return raw


class _MarkdownExtractor(html.parser.HTMLParser):
    """
    Convert Confluence storage-format XHTML into reasonable Markdown.

    This is a best-effort converter aimed at RAG ingestion, not a pixel-perfect
    renderer. It handles the common structural elements (headings, paragraphs,
    lists, bold/italic, links, inline code, code blocks, block quotes, rules and
    tables). Confluence macros (<ac:...>) are not interpreted, but their inner
    text - including code-macro CDATA bodies - is preserved. Text is not
    Markdown-escaped, so the occasional literal '*' may look like emphasis; that
    is a deliberate trade-off to keep the captured text faithful for search.
    """

    _SKIP_TAGS = {"script", "style"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []
        self._skip_depth = 0
        self._in_pre = 0
        self._list_stack = []      # 'ul' / 'ol' per nesting level
        self._ol_counters = []     # running item number per ordered-list level
        self._href_stack = []      # href per currently-open <a>
        # Table buffering (only the outermost table is rendered as a grid)
        self._table_depth = 0
        self._rows = None          # list of cell-lists for the current table
        self._row = None           # current row (list of cell strings)
        self._cell = None          # buffer for the current cell, or None

    def _emit(self, s):
        # Route text either into the current table cell or the main output.
        if self._cell is not None:
            self._cell.append(s)
        else:
            self.parts.append(s)

    def handle_starttag(self, tag, attrs):
        if tag in self._SKIP_TAGS:
            self._skip_depth += 1
            return
        if self._skip_depth:
            return
        if tag == "br":
            self._emit("\n" if self._in_pre else "  \n")
        elif tag == "p":
            self._emit("\n\n")
        elif tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
            self._emit("\n\n" + "#" * int(tag[1]) + " ")
        elif tag in ("strong", "b") and not self._in_pre:
            self._emit("**")
        elif tag in ("em", "i") and not self._in_pre:
            self._emit("*")
        elif tag == "code" and not self._in_pre:
            self._emit("`")
        elif tag == "pre":
            self._in_pre += 1
            self._emit("\n\n```\n")
        elif tag == "blockquote":
            self._emit("\n\n> ")
        elif tag == "hr":
            self._emit("\n\n---\n\n")
        elif tag == "a":
            href = ""
            for key, val in attrs:
                if key == "href":
                    href = val or ""
            self._href_stack.append(href)
            self._emit("[")
        elif tag == "ul":
            self._list_stack.append("ul")
        elif tag == "ol":
            self._list_stack.append("ol")
            self._ol_counters.append(0)
        elif tag == "li":
            indent = "  " * max(0, len(self._list_stack) - 1)
            if self._list_stack and self._list_stack[-1] == "ol":
                self._ol_counters[-1] += 1
                marker = "{}. ".format(self._ol_counters[-1])
            else:
                marker = "- "
            self._emit("\n" + indent + marker)
        elif tag == "table":
            self._table_depth += 1
            if self._table_depth == 1:
                self._rows = []
        elif tag == "tr":
            if self._rows is not None:
                self._row = []
        elif tag in ("td", "th"):
            if self._row is not None:
                self._cell = []

    def handle_endtag(self, tag):
        if tag in self._SKIP_TAGS:
            if self._skip_depth:
                self._skip_depth -= 1
            return
        if self._skip_depth:
            return
        if tag == "p":
            self._emit("\n\n")
        elif tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
            self._emit("\n\n")
        elif tag in ("strong", "b") and not self._in_pre:
            self._emit("**")
        elif tag in ("em", "i") and not self._in_pre:
            self._emit("*")
        elif tag == "code" and not self._in_pre:
            self._emit("`")
        elif tag == "pre":
            if self._in_pre:
                self._in_pre -= 1
            self._emit("\n```\n\n")
        elif tag == "blockquote":
            self._emit("\n\n")
        elif tag == "a":
            href = self._href_stack.pop() if self._href_stack else ""
            self._emit("]({})".format(href))
        elif tag in ("ul", "ol"):
            if self._list_stack:
                if self._list_stack.pop() == "ol" and self._ol_counters:
                    self._ol_counters.pop()
            self._emit("\n")
        elif tag in ("td", "th"):
            if self._cell is not None and self._row is not None:
                # Markdown cells are single-line: flatten and escape pipes.
                cell_text = " ".join("".join(self._cell).split())
                self._row.append(cell_text.replace("|", "\\|"))
                self._cell = None
        elif tag == "tr":
            if self._row is not None and self._rows is not None:
                self._rows.append(self._row)
                self._row = None
        elif tag == "table":
            if self._table_depth == 1 and self._rows is not None:
                self._emit_table(self._rows)
                self._rows = None
            if self._table_depth:
                self._table_depth -= 1

    def handle_data(self, data):
        if self._skip_depth:
            return
        if self._in_pre:
            self._emit(data)
            return
        if data.strip() == "":
            # Whitespace-only node between tags: keep a single separating space
            # rather than injecting blank lines.
            if data:
                self._emit(" ")
            return
        # Collapse embedded newlines so wrapped source doesn't break paragraphs.
        self._emit(data.replace("\r", " ").replace("\n", " "))

    def unknown_decl(self, data):
        # Capture CDATA content, e.g. Confluence code-macro bodies.
        if self._skip_depth:
            return
        if data.startswith("CDATA["):
            inner = data[6:]
            if inner.endswith("]"):
                inner = inner[:-1]
            self._emit(inner)

    def _emit_table(self, rows):
        if not rows:
            return
        ncols = max((len(r) for r in rows), default=0)
        if ncols == 0:
            return

        def fmt(cells):
            padded = list(cells) + [""] * (ncols - len(cells))
            return "| " + " | ".join(padded) + " |"

        out = [fmt(rows[0]), "| " + " | ".join(["---"] * ncols) + " |"]
        out.extend(fmt(r) for r in rows[1:])
        self.parts.append("\n\n" + "\n".join(out) + "\n\n")

    def get_markdown(self):
        text = "".join(self.parts)
        # Trim trailing spaces and collapse runs of blank lines to a single one.
        lines = [ln.rstrip() for ln in text.split("\n")]
        out = []
        blank = 0
        for ln in lines:
            if ln.strip() == "":
                blank += 1
                if blank <= 1:
                    out.append("")
            else:
                blank = 0
                out.append(ln)
        return "\n".join(out).strip() + "\n"


def html_to_markdown(raw):
    """Convert an HTML/XHTML string to Markdown, falling back to plain text."""
    if not raw:
        return ""
    parser = _MarkdownExtractor()
    try:
        parser.feed(raw)
        parser.close()
        return parser.get_markdown()
    except Exception:
        # Never lose content: fall back to the plain-text extractor.
        return html_to_text(raw)


def safe_filename(name, max_len=150):
    """
    Turn a page title into a filesystem-safe filename component (no extension).

    Strips characters that are illegal on Windows (< > : " / \\ | ? * and control
    chars), collapses whitespace, removes trailing dots/spaces (also illegal on
    Windows), and caps the length. Returns 'untitled' if nothing usable is left.
    """
    if not name:
        return "untitled"
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', " ", name)
    cleaned = " ".join(cleaned.split()).strip(" .")
    if len(cleaned) > max_len:
        cleaned = cleaned[:max_len].rstrip(" .")
    return cleaned or "untitled"


def cql_quote(value):
    """
    Escape a string for safe inclusion inside a double-quoted CQL literal.
    Backslashes and double quotes must be escaped. This prevents a value
    containing a quote from breaking out of the literal.
    """
    return value.replace("\\", "\\\\").replace('"', '\\"')


# ---------------------------------------------------------------------------
# Confluence client
# ---------------------------------------------------------------------------
class ConfluenceError(Exception):
    """Raised for any failure talking to Confluence; message is user-facing."""


class ConfluenceClient:
    def __init__(self, base_url, token=None, user=None, password=None,
                 verify_ssl=True, ca_cert=None, timeout=30, max_body=0,
                 kb_dir=None):
        if not base_url:
            raise ValueError("base_url is required")
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self.max_body = max_body
        # Folder to mirror pages into as Markdown; None/empty disables saving.
        self.kb_dir = kb_dir or None

        # Build auth header. Prefer a Personal Access Token (Bearer) if given.
        self.headers = {"Accept": "application/json"}
        if token:
            self.headers["Authorization"] = "Bearer " + token
        elif user is not None and password is not None:
            raw = "{}:{}".format(user, password).encode("utf-8")
            self.headers["Authorization"] = "Basic " + base64.b64encode(raw).decode("ascii")
        else:
            raise ValueError(
                "No credentials: set CONFLUENCE_TOKEN, or CONFLUENCE_USER and "
                "CONFLUENCE_PASSWORD."
            )

        # Build the TLS context. Only relevant for https:// URLs; ignored for
        # plain http. A custom CA bundle takes precedence; otherwise we either
        # verify normally or, if explicitly asked, disable verification.
        if ca_cert:
            self.ssl_context = ssl.create_default_context(cafile=ca_cert)
        elif not verify_ssl:
            self.ssl_context = ssl.create_default_context()
            self.ssl_context.check_hostname = False
            self.ssl_context.verify_mode = ssl.CERT_NONE
        else:
            self.ssl_context = ssl.create_default_context()

    def _get(self, path, params=None):
        """Perform a GET against the REST API and return parsed JSON."""
        url = self.base_url + path
        if params:
            # urlencode percent-encodes values (including CQL special chars).
            url = url + "?" + urllib.parse.urlencode(params)
        req = urllib.request.Request(url, headers=self.headers, method="GET")
        try:
            with urllib.request.urlopen(req, timeout=self.timeout,
                                        context=self.ssl_context) as resp:
                body = resp.read()
        except urllib.error.HTTPError as e:
            # Try to surface Confluence's error message from the response body.
            detail = ""
            try:
                detail = e.read().decode("utf-8", "replace")[:500]
            except Exception:
                pass
            raise ConfluenceError(
                "HTTP {} from Confluence for {}{}".format(
                    e.code, url, (": " + detail) if detail else ""
                )
            )
        except urllib.error.URLError as e:
            raise ConfluenceError(
                "Could not reach Confluence at {} ({}). Check the base URL, "
                "network reachability and TLS settings.".format(url, e.reason)
            )
        except ssl.SSLError as e:
            raise ConfluenceError(
                "TLS error talking to Confluence ({}). For an internal CA, set "
                "CONFLUENCE_CA_CERT, or CONFLUENCE_VERIFY_SSL=false to disable "
                "verification.".format(e)
            )
        try:
            return json.loads(body.decode("utf-8"))
        except (ValueError, UnicodeDecodeError) as e:
            raise ConfluenceError("Confluence returned a non-JSON response: {}".format(e))

    def _abs_link(self, data, link):
        """Build an absolute web URL from a result's webui link."""
        if not link:
            return ""
        base = ""
        links = data.get("_links") if isinstance(data, dict) else None
        if isinstance(links, dict):
            base = links.get("base") or ""
        if not base:
            base = self.base_url
        return base.rstrip("/") + link

    def search(self, cql, limit):
        """Run a CQL query against the content search endpoint."""
        params = {
            "cql": cql,
            "limit": limit,
            "expand": "space,version",
        }
        data = self._get("/rest/api/content/search", params)
        results = data.get("results", []) or []
        lines = []
        for item in results:
            space = (item.get("space") or {}).get("key", "?")
            title = item.get("title", "(untitled)")
            cid = item.get("id", "?")
            ctype = item.get("type", "content")
            link = self._abs_link(
                data, (item.get("_links") or {}).get("webui", "")
            )
            lines.append(
                "- id={id}  type={type}  space={space}\n  title: {title}\n  url: {url}".format(
                    id=cid, type=ctype, space=space, title=title, url=link
                )
            )
        header = "Found {} result(s) for CQL: {}".format(len(lines), cql)
        if not lines:
            return header + "\n(no matching content)"
        return header + "\n\n" + "\n\n".join(lines)

    def _render_page(self, page):
        """Format a single content object (with body.storage) as text."""
        title = page.get("title", "(untitled)")
        cid = page.get("id", "?")
        ctype = page.get("type", "content")
        space = (page.get("space") or {}).get("key", "?")
        version = (page.get("version") or {}).get("number", "?")
        link = self._abs_link(page, (page.get("_links") or {}).get("webui", ""))
        storage = (((page.get("body") or {}).get("storage") or {}).get("value")) or ""
        text = html_to_text(storage)
        truncated_note = ""
        if self.max_body and len(text) > self.max_body:
            text = text[: self.max_body]
            truncated_note = "\n\n[...body truncated to {} characters...]".format(self.max_body)
        meta = (
            "Title: {title}\nID: {id}\nType: {type}\nSpace: {space}\n"
            "Version: {version}\nURL: {url}\n\n--- Content ---\n".format(
                title=title, id=cid, type=ctype, space=space,
                version=version, url=link,
            )
        )
        rendered = meta + (text if text else "(this page has no readable body content)") + truncated_note

        # If a knowledge-base folder is configured, mirror the page to Markdown.
        # This is a side effect of reading a page; it must never break the tool,
        # so any failure is reported but swallowed.
        if self.kb_dir:
            try:
                path = self._save_to_kb(title, link, space, version, storage)
                log("saved page to knowledge base: {}".format(path))
                rendered += "\n\n[Saved to knowledge base: {}]".format(path)
            except OSError as e:
                log("knowledge-base save failed: {}".format(e))
                rendered += "\n\n[Knowledge-base save FAILED: {}]".format(e)
        return rendered

    def _save_to_kb(self, title, link, space, version, storage):
        """
        Write the page to '<kb_dir>/Confluence - <title>.md', overwriting any
        existing file. Returns the path written; raises OSError on failure.

        The FULL body is always saved (the CONFLUENCE_MAX_BODY limit only trims
        what is returned to the model, not what is stored for RAG).
        """
        md_body = html_to_markdown(storage)
        stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        header = (
            "# {title}\n\n"
            "- Source: {url}\n"
            "- Space: {space}\n"
            "- Version: {version}\n"
            "- Fetched: {stamp}\n\n"
            "---\n\n"
        ).format(title=title, url=link or "(unknown)", space=space,
                 version=version, stamp=stamp)
        content = header + (md_body if md_body else "(no readable body content)\n")

        # Create the folder if needed, then write. newline="\n" keeps endings
        # consistent and avoids CRLF doubling on Windows.
        os.makedirs(self.kb_dir, exist_ok=True)
        filename = "Confluence - " + safe_filename(title) + ".md"
        path = os.path.join(self.kb_dir, filename)
        with open(path, "w", encoding="utf-8", newline="\n") as fh:
            fh.write(content)
        return path

    def get_page(self, page_id):
        if page_id is None or str(page_id).strip() == "":
            raise ConfluenceError("'page_id' is required")
        page_id = str(page_id).strip()
        params = {"expand": "body.storage,version,space"}
        page = self._get("/rest/api/content/" + urllib.parse.quote(page_id, safe=""), params)
        return self._render_page(page)

    def get_page_by_title(self, title, space):
        if not title or not space:
            raise ConfluenceError("Both 'title' and 'space' are required")
        params = {
            "title": title,
            "spaceKey": space,
            "expand": "body.storage,version,space",
            "limit": 1,
        }
        data = self._get("/rest/api/content", params)
        results = data.get("results", []) or []
        if not results:
            return "No page titled {!r} found in space {!r}.".format(title, space)
        return self._render_page(results[0])

    def resolve_page_id(self, title, space):
        """
        Look up a single page's numeric ID from its exact title + space key.
        Returns the ID string, or raises ConfluenceError if not found.
        """
        if not title or not space:
            raise ConfluenceError("Both 'parent_title' and 'space' are required "
                                  "when 'parent_id' is not given")
        params = {"title": title, "spaceKey": space, "limit": 1}
        data = self._get("/rest/api/content", params)
        results = data.get("results", []) or []
        if not results:
            raise ConfluenceError(
                "No page titled {!r} found in space {!r}.".format(title, space)
            )
        return str(results[0].get("id"))

    def list_pages_under(self, parent_id, direct_only=False,
                         modified_within_days=None, limit=25):
        """
        List pages beneath a parent page. Builds the CQL internally so the
        caller never has to know CQL.
          - direct_only=True  -> only immediate children (CQL 'parent')
          - direct_only=False -> all descendants at any depth (CQL 'ancestor')
          - modified_within_days -> optionally restrict to pages changed in the
            last N days (CQL 'lastmodified >= now("-Nd")').
        """
        parent_id = str(parent_id).strip()
        if not parent_id:
            raise ConfluenceError("A parent page could not be identified")
        if not parent_id.isdigit():
            # Enforced so the unquoted embed below cannot inject CQL.
            raise ConfluenceError(
                "'parent_id' must be a numeric content ID (got {!r}). Use "
                "'parent_title' plus 'space' if you only know the title.".format(parent_id)
            )
        field = "parent" if direct_only else "ancestor"
        # parent_id is validated as numeric above, so it is safe to embed unquoted.
        clauses = ["{} = {}".format(field, parent_id), "type = page"]
        if modified_within_days is not None:
            try:
                days = int(modified_within_days)
            except (TypeError, ValueError):
                raise ConfluenceError("'modified_within_days' must be a whole number")
            if days > 0:
                clauses.append('lastmodified >= now("-{}d")'.format(days))
        cql = " AND ".join(clauses) + " ORDER BY lastmodified DESC"
        return self.search(cql, limit)


# ---------------------------------------------------------------------------
# Tool definitions and dispatch
# ---------------------------------------------------------------------------
def tool_definitions():
    """Return the list advertised via tools/list (JSON-Schema input specs)."""
    return [
        {
            "name": "confluence_search",
            "description": (
                "Search Confluence pages by free text. Returns matching pages "
                "with their numeric ID, title, space key and URL. Use "
                "'confluence_get_page' afterwards to read a page's full content. "
                "Optionally restrict to a single space by its space key."
            ),
            "inputSchema": {
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "Free-text search terms.",
                    },
                    "space": {
                        "type": "string",
                        "description": "Optional space key to restrict the search (e.g. 'DOCS').",
                    },
                    "limit": {
                        "type": "integer",
                        "description": "Maximum number of results (1-50, default 25).",
                    },
                },
                "required": ["query"],
            },
        },
        {
            "name": "confluence_search_cql",
            "description": (
                "Search Confluence using a raw CQL (Confluence Query Language) "
                "query for advanced filtering. Examples: "
                "'space = \"DOCS\" AND type = page', "
                "'text ~ \"release notes\" AND lastModified >= now(\"-30d\")', "
                "'title ~ \"runbook\"'. Returns matching content with IDs, "
                "titles, space keys and URLs."
            ),
            "inputSchema": {
                "type": "object",
                "properties": {
                    "cql": {
                        "type": "string",
                        "description": "A valid CQL query string.",
                    },
                    "limit": {
                        "type": "integer",
                        "description": "Maximum number of results (1-50, default 25).",
                    },
                },
                "required": ["cql"],
            },
        },
        {
            "name": "confluence_get_page",
            "description": (
                "Retrieve a single Confluence page by its numeric page ID. "
                "Returns the title, space, version, URL and the page body "
                "converted to plain text."
            ),
            "inputSchema": {
                "type": "object",
                "properties": {
                    "page_id": {
                        "type": "string",
                        "description": "The numeric Confluence content ID, e.g. '393217'.",
                    },
                },
                "required": ["page_id"],
            },
        },
        {
            "name": "confluence_get_page_by_title",
            "description": (
                "Retrieve a Confluence page by its exact title within a given "
                "space. Returns the same details as confluence_get_page."
            ),
            "inputSchema": {
                "type": "object",
                "properties": {
                    "title": {
                        "type": "string",
                        "description": "The exact page title.",
                    },
                    "space": {
                        "type": "string",
                        "description": "The space key the page lives in (e.g. 'DOCS').",
                    },
                },
                "required": ["title", "space"],
            },
        },
        {
            "name": "confluence_list_pages_under",
            "description": (
                "List pages located beneath a parent page in the page tree - "
                "use this for requests like 'pages under X' or 'child pages of "
                "X'. You do NOT need to write CQL; just identify the parent by "
                "its numeric 'parent_id', or by 'parent_title' plus 'space'. "
                "Set 'direct_only' to true for immediate children only, or "
                "leave it false to include all nested descendants. Optionally "
                "set 'modified_within_days' to only return pages changed "
                "recently (e.g. 30 for the past month). Returns IDs, titles, "
                "space keys and URLs, newest first; follow up with "
                "'confluence_get_page' to read each one."
            ),
            "inputSchema": {
                "type": "object",
                "properties": {
                    "parent_id": {
                        "type": "string",
                        "description": "Numeric ID of the parent page (preferred if known).",
                    },
                    "parent_title": {
                        "type": "string",
                        "description": "Exact title of the parent page (needs 'space' too).",
                    },
                    "space": {
                        "type": "string",
                        "description": "Space key of the parent page (used with 'parent_title').",
                    },
                    "direct_only": {
                        "type": "boolean",
                        "description": "True = immediate children only; false = all descendants. Default false.",
                    },
                    "modified_within_days": {
                        "type": "integer",
                        "description": "Only include pages modified within this many days (optional).",
                    },
                    "limit": {
                        "type": "integer",
                        "description": "Maximum number of results (1-50, default 25).",
                    },
                },
            },
        },
    ]


def clamp_limit(value, default=25, lo=1, hi=50):
    """Coerce a user-supplied limit into a sane integer range."""
    try:
        n = int(value)
    except (TypeError, ValueError):
        return default
    return max(lo, min(hi, n))


def call_tool(client, name, arguments):
    """
    Execute a named tool. Returns the text payload on success.
    Raises ConfluenceError (or ValueError) on a tool-domain failure, which the
    caller reports back as an MCP tool error (isError=true).
    """
    arguments = arguments or {}
    if name == "confluence_search":
        query = arguments.get("query")
        if not query:
            raise ConfluenceError("'query' is required")
        limit = clamp_limit(arguments.get("limit"))
        cql = 'text ~ "{}"'.format(cql_quote(str(query)))
        space = arguments.get("space")
        if space:
            cql += ' AND space = "{}"'.format(cql_quote(str(space)))
        return client.search(cql, limit)

    if name == "confluence_search_cql":
        cql = arguments.get("cql")
        if not cql:
            raise ConfluenceError("'cql' is required")
        limit = clamp_limit(arguments.get("limit"))
        return client.search(str(cql), limit)

    if name == "confluence_get_page":
        return client.get_page(arguments.get("page_id"))

    if name == "confluence_get_page_by_title":
        return client.get_page_by_title(
            arguments.get("title"), arguments.get("space")
        )

    if name == "confluence_list_pages_under":
        parent_id = arguments.get("parent_id")
        if not parent_id:
            # No explicit ID: resolve it from the parent's title + space.
            parent_id = client.resolve_page_id(
                arguments.get("parent_title"), arguments.get("space")
            )
        limit = clamp_limit(arguments.get("limit"))
        return client.list_pages_under(
            parent_id,
            direct_only=bool(arguments.get("direct_only", False)),
            modified_within_days=arguments.get("modified_within_days"),
            limit=limit,
        )

    raise ConfluenceError("Unknown tool: {}".format(name))


# ---------------------------------------------------------------------------
# JSON-RPC / MCP plumbing
# ---------------------------------------------------------------------------
def make_result(msg_id, result):
    return {"jsonrpc": "2.0", "id": msg_id, "result": result}


def make_error(msg_id, code, message):
    return {"jsonrpc": "2.0", "id": msg_id, "error": {"code": code, "message": message}}


def handle_initialize(params):
    # Echo the client's protocol version when it sends one, for compatibility.
    requested = ""
    if isinstance(params, dict):
        requested = params.get("protocolVersion") or ""
    protocol = requested if isinstance(requested, str) and requested else DEFAULT_PROTOCOL_VERSION
    return {
        "protocolVersion": protocol,
        "capabilities": {"tools": {}},
        "serverInfo": {"name": SERVER_NAME, "version": SERVER_VERSION},
    }


def handle_message(client, msg):
    """
    Process one JSON-RPC message object.
    Returns a response dict, or None for notifications (which get no reply).
    """
    if not isinstance(msg, dict):
        return make_error(None, INVALID_REQUEST, "Invalid Request: not an object")

    method = msg.get("method")
    msg_id = msg.get("id")
    is_request = "id" in msg  # notifications have no id and get no response

    if not isinstance(method, str):
        return make_error(msg_id, INVALID_REQUEST, "Missing method") if is_request else None

    # --- lifecycle / housekeeping ---
    if method == "initialize":
        return make_result(msg_id, handle_initialize(msg.get("params")))

    if method == "ping":
        return make_result(msg_id, {})

    if method.startswith("notifications/"):
        # e.g. notifications/initialized, notifications/cancelled - just ignore.
        return None

    # --- tools ---
    if method == "tools/list":
        return make_result(msg_id, {"tools": tool_definitions()})

    if method == "tools/call":
        params = msg.get("params") or {}
        name = params.get("name")
        arguments = params.get("arguments") or {}
        if not name:
            # report as a tool error so the agent can recover
            return make_result(msg_id, {
                "content": [{"type": "text", "text": "Error: no tool name supplied."}],
                "isError": True,
            })
        try:
            text = call_tool(client, name, arguments)
            return make_result(msg_id, {
                "content": [{"type": "text", "text": text}],
                "isError": False,
            })
        except (ConfluenceError, ValueError) as e:
            log("tool '{}' failed: {}".format(name, e))
            return make_result(msg_id, {
                "content": [{"type": "text", "text": "Error: {}".format(e)}],
                "isError": True,
            })
        except Exception as e:  # never let a tool crash the whole server
            log("unexpected error in tool '{}': {}".format(name, e))
            return make_result(msg_id, {
                "content": [{"type": "text", "text": "Unexpected error: {}".format(e)}],
                "isError": True,
            })

    # --- unknown method ---
    if is_request:
        return make_error(msg_id, METHOD_NOT_FOUND, "Method not found: {}".format(method))
    return None


def serve(client):
    """
    Main stdio loop. MCP stdio framing is newline-delimited JSON: one JSON-RPC
    message per line, no embedded newlines, responses written the same way.
    """
    log("server started; waiting for JSON-RPC on stdin")
    stdin = sys.stdin
    while True:
        line = stdin.readline()
        if line == "":
            break  # EOF: the client closed the pipe
        line = line.strip()
        if not line:
            continue
        try:
            incoming = json.loads(line)
        except ValueError:
            _write(make_error(None, PARSE_ERROR, "Parse error: invalid JSON"))
            continue

        # JSON-RPC permits a batch (array) of messages. MCP's newer revisions
        # dropped batching, but we handle it defensively.
        if isinstance(incoming, list):
            responses = []
            for item in incoming:
                resp = handle_message(client, item)
                if resp is not None:
                    responses.append(resp)
            if responses:
                _write(responses)
        else:
            resp = handle_message(client, incoming)
            if resp is not None:
                _write(resp)

    log("stdin closed; shutting down")


def _write(obj):
    """
    Write a single JSON value as one line to stdout, then flush.

    ensure_ascii=True keeps the output pure ASCII (non-ASCII characters become
    \\uXXXX escapes, which are valid JSON). This is critical on Windows, where
    stdout defaults to a legacy code page (e.g. cp1252) that cannot encode many
    characters found in Confluence page bodies - writing them raw would raise
    UnicodeEncodeError and kill the server. main() also forces the streams to
    UTF-8 as a second layer of defence.
    """
    try:
        sys.stdout.write(json.dumps(obj, ensure_ascii=True) + "\n")
        sys.stdout.flush()
    except (BrokenPipeError, OSError):
        # The client closed the pipe; nothing useful we can do but stop.
        raise SystemExit(0)


# ---------------------------------------------------------------------------
# Entry point / configuration
# ---------------------------------------------------------------------------
def env_bool(name, default=True):
    val = os.environ.get(name)
    if val is None:
        return default
    return val.strip().lower() not in ("0", "false", "no", "off", "")


def build_arg_parser():
    p = argparse.ArgumentParser(
        description="MCP server for querying Confluence Data Center (stdio transport).",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    p.add_argument("--base-url", default=os.environ.get("CONFLUENCE_BASE_URL"),
                   help="Confluence base URL incl. any context path, no trailing slash "
                        "(env CONFLUENCE_BASE_URL).")
    # SECURITY: credentials are deliberately env-var ONLY (CONFLUENCE_TOKEN, or
    # CONFLUENCE_USER + CONFLUENCE_PASSWORD). Command-line arguments are visible
    # to other local users in process listings, so no --token/--user/--password
    # flags are offered.
    p.add_argument("--ca-cert", default=os.environ.get("CONFLUENCE_CA_CERT"),
                   help="Path to a PEM CA bundle for an internal CA "
                        "(env CONFLUENCE_CA_CERT).")
    p.add_argument("--insecure", action="store_true",
                   default=not env_bool("CONFLUENCE_VERIFY_SSL", True),
                   help="Disable TLS certificate verification "
                        "(env CONFLUENCE_VERIFY_SSL=false).")
    p.add_argument("--timeout", type=int,
                   default=int(os.environ.get("CONFLUENCE_TIMEOUT", "30")),
                   help="HTTP request timeout in seconds (env CONFLUENCE_TIMEOUT).")
    p.add_argument("--max-body", type=int,
                   default=int(os.environ.get("CONFLUENCE_MAX_BODY", "0")),
                   help="Truncate page bodies to N characters, 0 = unlimited "
                        "(env CONFLUENCE_MAX_BODY). Applies to returned text "
                        "only, not to saved knowledge-base files.")
    p.add_argument("--kb-dir", default=os.environ.get("CONFLUENCE_KB_DIR"),
                   help="If set, every page read is also saved as a Markdown "
                        "file into this folder for a local RAG knowledge base "
                        "(env CONFLUENCE_KB_DIR). Files are named "
                        "'Confluence - <title>.md' and overwritten each time.")
    return p


def main(argv=None):
    # Force the JSON-RPC streams to UTF-8. On Windows the console/pipe encoding
    # defaults to a legacy code page that cannot represent many characters in
    # Confluence content; without this, reading or writing such characters can
    # crash the server and the client then reports "not connected".
    for stream in (sys.stdin, sys.stdout):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            # reconfigure is unavailable (e.g. a redirected non-text stream);
            # _write still emits ASCII-only JSON, so output stays safe anyway.
            pass

    args = build_arg_parser().parse_args(argv)

    # Credentials come from the environment ONLY (never argv - see build_arg_parser).
    token = os.environ.get("CONFLUENCE_TOKEN")
    user = os.environ.get("CONFLUENCE_USER")
    password = os.environ.get("CONFLUENCE_PASSWORD")

    if not args.base_url:
        log("FATAL: no base URL. Set CONFLUENCE_BASE_URL or pass --base-url.")
        return 2
    if not token and not (user and password):
        log("FATAL: no credentials. Set the CONFLUENCE_TOKEN environment "
            "variable, or CONFLUENCE_USER and CONFLUENCE_PASSWORD.")
        return 2

    try:
        client = ConfluenceClient(
            base_url=args.base_url,
            token=token,
            user=user,
            password=password,
            verify_ssl=not args.insecure,
            ca_cert=args.ca_cert,
            timeout=args.timeout,
            max_body=args.max_body,
            kb_dir=args.kb_dir,
        )
    except (ValueError, ssl.SSLError, OSError) as e:
        log("FATAL: could not initialise client: {}".format(e))
        return 2

    if args.insecure:
        log("WARNING: TLS verification is disabled (--insecure).")
    if client.kb_dir:
        log("knowledge-base mirroring enabled -> {}".format(client.kb_dir))
    log("configured for base URL {}".format(client.base_url))

    try:
        serve(client)
    except KeyboardInterrupt:
        pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
