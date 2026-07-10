#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
ms-excel.py -- Read-only Excel (.xlsx) MCP server for VSCode Continue.

PURPOSE
    A single-file, standard-library-only MCP (Model Context Protocol) stdio
    server that lets a local model query and ANALYSE Excel workbooks. It is
    strictly READ-ONLY: it never writes, never opens Excel, and never touches
    the network. It parses .xlsx directly (a .xlsx file is just a ZIP of XML),
    so no third-party packages (openpyxl, pandas, pywin32) are required.

    Tools exposed:
      * excel_list_workbooks   -- list .xlsx files in the configured folder
      * excel_list_sheets      -- sheet names + declared dimensions for a workbook
      * excel_get_headers      -- the first (header) row of a sheet
      * excel_read_range       -- read cells from a sheet, optionally an A1 range
      * excel_search           -- find cells whose value matches a query
      * excel_column_stats     -- summary statistics for one column

    Design mirrors the other servers in this suite: stdout is reserved for
    JSON-RPC only, all diagnostics go to stderr, config lives in one fenced
    block below, workbook names are resolved fuzzily, and a --check flag
    validates the environment before wiring into Continue.

SUPPORTED / NOT SUPPORTED
    * Supported: .xlsx (and macro-enabled .xlsm, same underlying format).
    * NOT supported: legacy .xls (old binary BIFF format), .xlsb (binary),
      and password-encrypted workbooks. These are detected and reported
      clearly rather than mis-parsed.
    * Formula cells: the LAST VALUE CACHED BY EXCEL is returned (the same
      value you would see on screen). Formulas are not re-evaluated, so a
      workbook saved by a tool that did not cache values may show blanks for
      formula cells. This is inherent to reading without a calc engine.
    * Dates: Excel stores dates as serial numbers. Cells whose STYLE marks
      them as a date/time are converted to ISO-8601 text. Dates before
      1900-03-01 may be off by one day due to Excel's historical 1900
      leap-year bug; this affects almost no real enterprise data.

REQUIREMENTS
    * Python 3.8+ (standard library only). No pip install required.

CONFIGURATION
    Edit the CONFIG block directly below the imports. The workbook folder is
    REQUIRED - set WORKBOOK_FOLDER there, or supply --folder / the
    EXCEL_WORKBOOK_FOLDER environment variable at launch; the server refuses
    to start without one and only ever reads files inside it (symlinks that
    resolve outside the folder are excluded). Other settings can also be
    overridden per-run via command-line flags (see --help).

STANDALONE TESTING (before wiring into Continue)
    1) Environment / config sanity check (prints interpreter + folder state):
         python ms-excel.py --check

    2) List available workbooks without starting the server loop:
         python ms-excel.py --list

    3) Drive the JSON-RPC protocol by hand. On Windows PowerShell, create a
       file "probe.txt" with these three lines (each a complete JSON object
       on ONE line) and pipe it in:

         {"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"probe","version":"0"}}}
         {"jsonrpc":"2.0","method":"notifications/initialized"}
         {"jsonrpc":"2.0","id":2,"method":"tools/list"}

       Then:
         Get-Content probe.txt | python ms-excel.py
       You should see three JSON lines back on stdout (the third listing the
       tools). Any diagnostic text appears on stderr and does NOT corrupt the
       protocol stream.

    4) Call a tool by hand (adjust the workbook/sheet names):
         {"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"excel_list_sheets","arguments":{"workbook":"budget"}}}

INTEGRATION WITH CONTINUE (config.yaml)
    Add under mcpServers. Use the SAME Python interpreter you tested with
    (match interpreter paths exactly to avoid module-mismatch surprises),
    and keep PYTHONUTF8=1 so Windows cp1252 does not corrupt output.

      mcpServers:
        - name: excel
          command: C:\path\to\python.exe
          args:
            - C:\path\to\ms-excel.py
            - --folder
            - C:\path\to\your\workbooks
          env:
            PYTHONUTF8: "1"

    After editing config.yaml, run "Developer: Reload Window" in VSCode rather
    than toggling the server, to avoid the "already connected to transport"
    reconnection bug.

PROTOCOL NOTE
    Transport is newline-delimited JSON-RPC 2.0 over stdio (one JSON object
    per line, no embedded newlines), which is the MCP stdio convention.
    Advertised protocolVersion is "2024-11-05".
"""

import sys
import os
import io
import json
import argparse
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta

# ===========================================================================
# CONFIG  -- edit these values, or override with command-line flags.
# ===========================================================================

# REQUIRED: folder containing the .xlsx workbooks the model is allowed to
# read. The server only ever reads files inside this folder (symlinks that
# resolve outside it are excluded) and REFUSES TO START without one. Set it
# here, or at launch with --folder or the EXCEL_WORKBOOK_FOLDER environment
# variable (which take priority over this constant).
#   e.g. WORKBOOK_FOLDER = r"C:\Users\me\Documents\workbooks"
WORKBOOK_FOLDER = None

# File extensions treated as readable workbooks (lower-case, incl. dot).
ALLOWED_EXTENSIONS = (".xlsx", ".xlsm")

# Safety caps so a huge sheet cannot flood the model's context window.
MAX_ROWS_PER_READ = 200      # max rows returned by excel_read_range in one call
MAX_COLS_PER_READ = 64       # max columns returned per row
MAX_SEARCH_HITS = 100        # max matches returned by excel_search
MAX_CELL_TEXT_LEN = 500      # long cell text is truncated to this many chars

# Server identity reported to the client.
SERVER_NAME = "excel-readonly"
SERVER_VERSION = "1.1.0"
PROTOCOL_VERSION = "2024-11-05"

# ===========================================================================
# End of CONFIG
# ===========================================================================


# --- stdout/stderr hygiene --------------------------------------------------
# stdout MUST carry only JSON-RPC. Reconfigure both streams to UTF-8 so that
# Windows' default cp1252 codec cannot raise on non-ASCII content. We use
# line buffering on stdout so each JSON reply is flushed promptly.
try:
    sys.stdout.reconfigure(encoding="utf-8", newline="\n")
    sys.stderr.reconfigure(encoding="utf-8")
except AttributeError:
    # Python < 3.7 fallback (not expected on 3.8+, kept defensive).
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", newline="\n")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")


def log(*parts):
    """Diagnostics to stderr ONLY. Never write diagnostics to stdout."""
    print("[excel_mcp]", *parts, file=sys.stderr, flush=True)


# ===========================================================================
# XLSX parsing (pure standard library)
# ===========================================================================

# OOXML uses XML namespaces; we work in local names to stay namespace-agnostic.
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Built-in numFmt ids that always denote a date/time (per the OOXML spec).
_BUILTIN_DATE_FMT_IDS = {14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47}


def _lname(tag):
    """Return the local part of a namespaced XML tag ('{ns}row' -> 'row')."""
    return tag.rsplit("}", 1)[-1]


def _attr_local(elem, local):
    """Fetch an attribute by its local name, ignoring any namespace prefix."""
    for key, val in elem.attrib.items():
        if key.rsplit("}", 1)[-1] == local:
            return val
    return None


def _col_letters_to_index(letters):
    """'A' -> 0, 'B' -> 1, 'Z' -> 25, 'AA' -> 26 ..."""
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
    return idx - 1


def _col_index_to_letters(idx):
    """0 -> 'A', 25 -> 'Z', 26 -> 'AA' ..."""
    idx += 1
    out = []
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        out.append(chr(ord("A") + rem))
    return "".join(reversed(out))


def _split_cell_ref(ref):
    """'B12' -> ('B', 12). Returns (col_letters, row_number)."""
    col = []
    i = 0
    while i < len(ref) and ref[i].isalpha():
        col.append(ref[i])
        i += 1
    row_part = ref[i:]
    row = int(row_part) if row_part.isdigit() else None
    return "".join(col), row


def _looks_like_date_format(code):
    """
    Heuristic: does a custom number-format code represent a date/time?
    Strip quoted literals, escaped chars, and [colour]/[condition] brackets,
    then look for any date/time token (y m d h s). 'm' is ambiguous
    (month vs minute) but its presence still implies a date/time format.
    """
    if not code:
        return False
    cleaned = []
    i = 0
    while i < len(code):
        ch = code[i]
        if ch == '"':                      # skip quoted literal
            i += 1
            while i < len(code) and code[i] != '"':
                i += 1
            i += 1
            continue
        if ch == "\\":                     # skip escaped char
            i += 2
            continue
        if ch == "[":                      # skip [Red], [$-409], [>0] etc.
            i += 1
            while i < len(code) and code[i] != "]":
                i += 1
            i += 1
            continue
        cleaned.append(ch)
        i += 1
    text = "".join(cleaned).lower()
    return any(tok in text for tok in ("y", "m", "d", "h", "s"))


def _serial_to_isoformat(serial, date1904):
    """
    Convert an Excel date serial to an ISO-8601 string.
    See module docstring re: the pre-1900-03-01 leap-year quirk.
    """
    try:
        serial = float(serial)
    except (TypeError, ValueError):
        return None
    if date1904:
        base = datetime(1904, 1, 1)
    else:
        base = datetime(1899, 12, 30)
    dt = base + timedelta(days=serial)
    # If there is no fractional part, it's a pure date; render date-only.
    if abs(serial - round(serial)) < 1e-9 and dt.time() == datetime.min.time():
        return dt.date().isoformat()
    return dt.isoformat(sep=" ")


class WorkbookError(Exception):
    """Raised for unreadable / unsupported / corrupt workbooks."""


class Workbook:
    """
    Minimal read-only .xlsx reader. Loads shared strings, styles, and the
    sheet index up front; sheet data is parsed on demand.
    """

    def __init__(self, path):
        self.path = path
        if not os.path.isfile(path):
            raise WorkbookError("File not found: %s" % path)
        try:
            self._zip = zipfile.ZipFile(path, "r")
        except zipfile.BadZipFile:
            # Legacy .xls and encrypted files are not ZIP containers.
            raise WorkbookError(
                "Not a readable .xlsx. Legacy .xls, .xlsb, or password-"
                "protected files are not supported: %s" % os.path.basename(path)
            )
        names = set(self._zip.namelist())
        if "xl/workbook.xml" not in names:
            raise WorkbookError(
                "Missing xl/workbook.xml -- not a valid .xlsx: %s"
                % os.path.basename(path)
            )
        self.date1904 = False
        self._shared = []            # shared strings (index -> text)
        self._date_style = []        # style index -> bool (is a date format)
        self._sheets = []            # list of {"name", "path"}
        self._load_workbook_index()
        self._load_shared_strings()
        self._load_styles()

    # -- loading helpers ----------------------------------------------------

    def _read_xml(self, member):
        """Parse a member of the zip into an ElementTree root, or None."""
        try:
            data = self._zip.read(member)
        except KeyError:
            return None
        try:
            return ET.fromstring(data)
        except ET.ParseError as exc:
            raise WorkbookError("Corrupt XML in %s: %s" % (member, exc))

    def _load_workbook_index(self):
        root = self._read_xml("xl/workbook.xml")
        if root is None:
            raise WorkbookError("Could not read xl/workbook.xml")

        # Detect the 1904 date system (rare; old Mac-authored files).
        for el in root.iter():
            if _lname(el.tag) == "workbookPr":
                d1904 = _attr_local(el, "date1904")
                if d1904 in ("1", "true", "True"):
                    self.date1904 = True
                break

        # Map relationship id -> target path.
        rels = {}
        rroot = self._read_xml("xl/_rels/workbook.xml.rels")
        if rroot is not None:
            for rel in rroot:
                if _lname(rel.tag) != "Relationship":
                    continue
                rid = rel.get("Id")
                target = rel.get("Target")
                if not rid or not target:
                    continue
                # Normalise target into a full path within the zip.
                if target.startswith("/"):
                    full = target.lstrip("/")
                else:
                    full = "xl/" + target.lstrip("./")
                # Collapse any '..' the rare Target might contain.
                full = os.path.normpath(full).replace(os.sep, "/")
                rels[rid] = full

        # Ordered list of sheets from workbook.xml.
        for el in root.iter():
            if _lname(el.tag) != "sheet":
                continue
            name = el.get("name") or "Sheet"
            rid = _attr_local(el, "id")   # r:id
            spath = rels.get(rid)
            if spath is None:
                # Fallback guess if the relationship is missing.
                log("warning: sheet '%s' has no resolvable relationship" % name)
                continue
            self._sheets.append({"name": name, "path": spath})

        if not self._sheets:
            raise WorkbookError("No worksheets found in %s"
                                % os.path.basename(self.path))

    def _load_shared_strings(self):
        root = self._read_xml("xl/sharedStrings.xml")
        if root is None:
            return
        for si in root:
            if _lname(si.tag) != "si":
                continue
            # Concatenate every <t> descendant (covers plain and rich text).
            parts = [t.text or "" for t in si.iter() if _lname(t.tag) == "t"]
            self._shared.append("".join(parts))

    def _load_styles(self):
        root = self._read_xml("xl/styles.xml")
        if root is None:
            return
        # Custom number formats: numFmtId -> format code.
        custom = {}
        for el in root.iter():
            if _lname(el.tag) == "numFmt":
                fid = el.get("numFmtId")
                code = el.get("formatCode")
                if fid is not None:
                    try:
                        custom[int(fid)] = code or ""
                    except ValueError:
                        pass
        # cellXfs: ordered list; a cell's s="N" indexes into this.
        for el in root.iter():
            if _lname(el.tag) != "cellXfs":
                continue
            for xf in el:
                if _lname(xf.tag) != "xf":
                    continue
                fid_txt = xf.get("numFmtId")
                is_date = False
                if fid_txt is not None:
                    try:
                        fid = int(fid_txt)
                    except ValueError:
                        fid = -1
                    if fid in _BUILTIN_DATE_FMT_IDS:
                        is_date = True
                    elif fid in custom:
                        is_date = _looks_like_date_format(custom[fid])
                self._date_style.append(is_date)
            break

    # -- public API ---------------------------------------------------------

    def sheet_names(self):
        return [s["name"] for s in self._sheets]

    def _resolve_sheet(self, sheet):
        """Resolve a sheet by exact name, case-insensitive name, or 1-based index."""
        if sheet is None:
            return self._sheets[0]
        # Exact match first.
        for s in self._sheets:
            if s["name"] == sheet:
                return s
        # Case-insensitive.
        low = str(sheet).strip().lower()
        for s in self._sheets:
            if s["name"].lower() == low:
                return s
        # 1-based numeric index.
        if str(sheet).isdigit():
            i = int(sheet) - 1
            if 0 <= i < len(self._sheets):
                return self._sheets[i]
        raise WorkbookError(
            "Sheet '%s' not found. Available: %s"
            % (sheet, ", ".join(self.sheet_names()))
        )

    def _cell_value(self, c):
        """Resolve a <c> element to a Python value (str/int/float/bool/None)."""
        ctype = c.get("t")          # cell type
        style = c.get("s")          # style index (for date detection)
        v_el = None
        is_el = None
        for child in c:
            ln = _lname(child.tag)
            if ln == "v":
                v_el = child
            elif ln == "is":
                is_el = child
        v = v_el.text if v_el is not None else None

        if ctype == "s":                       # shared string (v = index)
            try:
                return self._trunc(self._shared[int(v)])
            except (ValueError, IndexError, TypeError):
                return None
        if ctype == "inlineStr":               # inline string in <is>
            if is_el is not None:
                parts = [t.text or "" for t in is_el.iter()
                         if _lname(t.tag) == "t"]
                return self._trunc("".join(parts))
            return None
        if ctype == "str":                     # formula string result
            return self._trunc(v) if v is not None else None
        if ctype == "b":                       # boolean
            return v == "1"
        if ctype == "e":                       # error text, e.g. #DIV/0!
            return self._trunc(v) if v is not None else None

        # Default: numeric (ctype None or "n"). Could be a date by style.
        if v is None or v == "":
            return None
        if style is not None:
            try:
                if self._date_style[int(style)]:
                    iso = _serial_to_isoformat(v, self.date1904)
                    if iso is not None:
                        return iso
            except (ValueError, IndexError):
                pass
        # Plain number: keep ints as ints for clean output.
        try:
            f = float(v)
        except ValueError:
            return self._trunc(v)
        if f.is_integer():
            return int(f)
        return f

    @staticmethod
    def _trunc(text):
        if text is None:
            return None
        text = str(text)
        if len(text) > MAX_CELL_TEXT_LEN:
            return text[:MAX_CELL_TEXT_LEN] + " ...[truncated]"
        return text

    def declared_dimension(self, sheet):
        """Return the sheet's declared used-range ref (e.g. 'A1:H100') or None."""
        s = self._resolve_sheet(sheet)
        root = self._read_xml(s["path"])
        if root is None:
            return None
        for el in root.iter():
            if _lname(el.tag) == "dimension":
                return el.get("ref")
        return None

    def iter_rows(self, sheet, min_row=1, max_row=None, max_col=None):
        """
        Yield (row_number, [values...]) for a sheet. Sparse cells are filled
        with None so column alignment is preserved. Rows are yielded in
        document order; empty trailing cells are trimmed per row.
        """
        s = self._resolve_sheet(sheet)
        root = self._read_xml(s["path"])
        if root is None:
            return
        # Find <sheetData>.
        sheet_data = None
        for el in root:
            if _lname(el.tag) == "sheetData":
                sheet_data = el
                break
        if sheet_data is None:
            return

        for row_el in sheet_data:
            if _lname(row_el.tag) != "row":
                continue
            r_attr = row_el.get("r")
            row_num = int(r_attr) if r_attr and r_attr.isdigit() else None
            if row_num is None:
                continue
            if row_num < min_row:
                continue
            if max_row is not None and row_num > max_row:
                break

            # Place cells by column index into a dict, then flatten.
            cells = {}
            highest = -1
            for c in row_el:
                if _lname(c.tag) != "c":
                    continue
                ref = c.get("r")
                if ref:
                    col_letters, _ = _split_cell_ref(ref)
                    col_idx = _col_letters_to_index(col_letters)
                else:
                    col_idx = highest + 1  # positional fallback
                if max_col is not None and col_idx >= max_col:
                    continue
                cells[col_idx] = self._cell_value(c)
                if col_idx > highest:
                    highest = col_idx

            if highest < 0:
                yield row_num, []
                continue
            width = highest + 1
            values = [cells.get(i) for i in range(width)]
            yield row_num, values

    def close(self):
        try:
            self._zip.close()
        except Exception:
            pass


# ===========================================================================
# Workbook folder resolution
# ===========================================================================

def list_workbook_files(folder):
    """
    Return sorted list of readable workbook file names in the folder.
    Files that RESOLVE outside the folder (e.g. a symlink pointing elsewhere)
    are excluded, so a link dropped into the folder cannot expose workbooks
    beyond it.
    """
    if not folder:
        raise WorkbookError(
            "No workbook folder configured. Pass --folder, set the "
            "EXCEL_WORKBOOK_FOLDER environment variable, or set "
            "WORKBOOK_FOLDER in this file."
        )
    if not os.path.isdir(folder):
        raise WorkbookError("Workbook folder does not exist: %s" % folder)
    real_base = os.path.realpath(folder)
    out = []
    for name in os.listdir(folder):
        if name.startswith("~$"):          # Excel lock/temp files
            continue
        ext = os.path.splitext(name)[1].lower()
        full = os.path.join(folder, name)
        if ext in ALLOWED_EXTENSIONS and os.path.isfile(full):
            try:
                real = os.path.realpath(full)
                contained = os.path.commonpath([real, real_base]) == real_base
            except ValueError:  # different drives on Windows
                contained = False
            if not contained:
                log("excluded (resolves outside the workbook folder): %s" % full)
                continue
            out.append(name)
    return sorted(out)


def resolve_workbook_path(folder, requested):
    """
    Resolve a loosely specified workbook name to a full path.
    Matching order: exact -> exact+ext -> case-insensitive -> substring.
    Rejects any path that escapes the configured folder.
    """
    files = list_workbook_files(folder)
    if not files:
        raise WorkbookError("No workbooks found in %s" % folder)

    req = os.path.basename(str(requested).strip())  # strip any path parts
    req_low = req.lower()

    # Exact filename match.
    for f in files:
        if f == req:
            return os.path.join(folder, f)
    # Match ignoring a missing extension.
    for f in files:
        if os.path.splitext(f)[0].lower() == req_low:
            return os.path.join(folder, f)
    for f in files:
        if f.lower() == req_low:
            return os.path.join(folder, f)
    # Substring (unique) match.
    hits = [f for f in files if req_low in f.lower()]
    if len(hits) == 1:
        return os.path.join(folder, hits[0])
    if len(hits) > 1:
        raise WorkbookError(
            "Workbook '%s' is ambiguous. Candidates: %s"
            % (requested, ", ".join(hits))
        )
    raise WorkbookError(
        "Workbook '%s' not found. Available: %s"
        % (requested, ", ".join(files))
    )


# ===========================================================================
# Tool implementations. Each returns a plain string (rendered to the model).
# ===========================================================================

def _open(folder, workbook):
    return Workbook(resolve_workbook_path(folder, workbook))


def tool_list_workbooks(folder, args):
    files = list_workbook_files(folder)
    if not files:
        return "No workbooks found in %s" % folder
    lines = ["Workbooks in %s:" % folder]
    for f in files:
        full = os.path.join(folder, f)
        try:
            size = os.path.getsize(full)
            lines.append("  - %s (%d bytes)" % (f, size))
        except OSError:
            lines.append("  - %s" % f)
    return "\n".join(lines)


def tool_list_sheets(folder, args):
    wb = _open(folder, args.get("workbook"))
    try:
        lines = ["Workbook: %s" % os.path.basename(wb.path)]
        if wb.date1904:
            lines.append("(uses the 1904 date system)")
        lines.append("Sheets:")
        for name in wb.sheet_names():
            dim = wb.declared_dimension(name)
            dim_txt = " declared range %s" % dim if dim else " (no declared range)"
            lines.append("  - %s:%s" % (name, dim_txt))
        return "\n".join(lines)
    finally:
        wb.close()


def tool_get_headers(folder, args):
    wb = _open(folder, args.get("workbook"))
    try:
        sheet = args.get("sheet")
        header_row = int(args.get("header_row", 1) or 1)
        for row_num, values in wb.iter_rows(sheet, min_row=header_row,
                                             max_row=header_row,
                                             max_col=MAX_COLS_PER_READ):
            cols = []
            for i, val in enumerate(values):
                letter = _col_index_to_letters(i)
                cols.append("%s=%s" % (letter, "" if val is None else val))
            resolved = wb._resolve_sheet(sheet)["name"]
            if not cols:
                return "Header row %d of sheet '%s' is empty." % (header_row, resolved)
            return ("Headers for sheet '%s' (row %d):\n  "
                    % (resolved, header_row)) + "\n  ".join(cols)
        resolved = wb._resolve_sheet(sheet)["name"]
        return "Sheet '%s' has no row %d." % (resolved, header_row)
    finally:
        wb.close()


def _parse_a1_range(rng):
    """
    Parse 'A1:D20' (or 'A1') into (min_row, max_row, min_col, max_col),
    using None for open ends. Returns None if rng is falsy.
    """
    if not rng:
        return None
    rng = str(rng).replace(" ", "")
    if ":" in rng:
        start, end = rng.split(":", 1)
    else:
        start = end = rng
    sc, sr = _split_cell_ref(start)
    ec, er = _split_cell_ref(end)
    min_col = _col_letters_to_index(sc) if sc else None
    max_col = _col_letters_to_index(ec) if ec else None
    if (min_col is not None and max_col is not None and min_col > max_col):
        min_col, max_col = max_col, min_col
    if (sr is not None and er is not None and sr > er):
        sr, er = er, sr
    return sr, er, min_col, max_col


def tool_read_range(folder, args):
    wb = _open(folder, args.get("workbook"))
    try:
        sheet = args.get("sheet")
        parsed = _parse_a1_range(args.get("range"))
        if parsed:
            min_row, max_row, min_col, max_col = parsed
            min_row = min_row or 1
        else:
            min_row, max_row, min_col, max_col = 1, None, None, None

        # Apply safety caps.
        eff_max_col = MAX_COLS_PER_READ
        if max_col is not None:
            eff_max_col = min(MAX_COLS_PER_READ, max_col + 1)
        row_cap = MAX_ROWS_PER_READ

        out_rows = []
        truncated = False
        resolved = wb._resolve_sheet(sheet)["name"]
        for row_num, values in wb.iter_rows(sheet, min_row=min_row,
                                            max_row=max_row,
                                            max_col=eff_max_col):
            # Trim to requested left column if a range was given.
            if min_col is not None:
                values = values[min_col:] if min_col < len(values) else []
                start_letter = _col_index_to_letters(min_col)
            else:
                start_letter = "A"
            cell_texts = []
            for i, val in enumerate(values):
                letter = _col_index_to_letters(
                    (_col_letters_to_index(start_letter) + i))
                cell_texts.append("%s%d=%s"
                                  % (letter, row_num,
                                     "" if val is None else val))
            out_rows.append("  " + " | ".join(cell_texts) if cell_texts
                            else "  (row %d empty)" % row_num)
            if len(out_rows) >= row_cap:
                truncated = True
                break

        if not out_rows:
            return "No data found in sheet '%s' for the requested range." % resolved
        header = "Sheet '%s'" % resolved
        if args.get("range"):
            header += " range %s" % args.get("range")
        if truncated:
            header += " (truncated to %d rows)" % row_cap
        return header + ":\n" + "\n".join(out_rows)
    finally:
        wb.close()


def tool_search(folder, args):
    query = args.get("query")
    if query is None or str(query) == "":
        return "Error: 'query' is required."
    query_low = str(query).lower()
    case_sensitive = bool(args.get("case_sensitive", False))
    target_sheet = args.get("sheet")   # optional: restrict to one sheet

    wb = _open(folder, args.get("workbook"))
    try:
        sheets = ([wb._resolve_sheet(target_sheet)["name"]]
                  if target_sheet else wb.sheet_names())
        hits = []
        for sname in sheets:
            for row_num, values in wb.iter_rows(sname,
                                                max_col=MAX_COLS_PER_READ):
                for i, val in enumerate(values):
                    if val is None:
                        continue
                    hay = str(val)
                    needle = str(query)
                    found = (needle in hay) if case_sensitive \
                        else (query_low in hay.lower())
                    if found:
                        ref = "%s%d" % (_col_index_to_letters(i), row_num)
                        hits.append("  [%s] %s = %s" % (sname, ref, hay))
                        if len(hits) >= MAX_SEARCH_HITS:
                            break
                if len(hits) >= MAX_SEARCH_HITS:
                    break
            if len(hits) >= MAX_SEARCH_HITS:
                break
        if not hits:
            return "No cells matched '%s'." % query
        head = "Found %d match(es) for '%s'" % (len(hits), query)
        if len(hits) >= MAX_SEARCH_HITS:
            head += " (stopped at cap %d)" % MAX_SEARCH_HITS
        return head + ":\n" + "\n".join(hits)
    finally:
        wb.close()


def _resolve_column_index(wb, sheet, column, header_row):
    """
    Turn a column spec into a 0-based index. Accepts a column letter
    ('C'), a 1-based number ('3'), or a header name matched in header_row.
    """
    spec = str(column).strip()

    # 1-based column number, e.g. "3".
    if spec.isdigit():
        return int(spec) - 1

    header_map = _header_map(wb, sheet, header_row)

    # Header name match (case-insensitive) takes priority over letter parsing,
    # so a column literally named "C" still resolves to that header.
    if spec.lower() in header_map:
        return header_map[spec.lower()]

    # Otherwise, only treat it as a column letter if it plausibly IS one
    # (short and all letters). This avoids silently mapping an unmatched
    # header name like "Cost" onto some far-off column.
    if spec.isalpha() and len(spec) <= 3:
        return _col_letters_to_index(spec)

    raise WorkbookError(
        "Column '%s' not found. Provide a column letter (e.g. C), a 1-based "
        "number (e.g. 3), or one of these headers: %s"
        % (column, ", ".join(sorted(header_map.keys())) or "(none)")
    )


def _header_map(wb, sheet, header_row):
    """Map lower-cased header text -> column index for the given header row."""
    mapping = {}
    for _rn, values in wb.iter_rows(sheet, min_row=header_row,
                                    max_row=header_row,
                                    max_col=MAX_COLS_PER_READ):
        for i, val in enumerate(values):
            if val is not None and str(val).strip() != "":
                mapping[str(val).strip().lower()] = i
    return mapping


def tool_column_stats(folder, args):
    if "column" not in args:
        return "Error: 'column' is required (letter, number, or header name)."
    wb = _open(folder, args.get("workbook"))
    try:
        sheet = args.get("sheet")
        header_row = int(args.get("header_row", 1) or 1)
        data_start = int(args.get("data_start_row", header_row + 1)
                         or header_row + 1)
        col_idx = _resolve_column_index(wb, sheet, args["column"], header_row)
        resolved = wb._resolve_sheet(sheet)["name"]

        total = 0
        non_empty = 0
        numbers = []
        distinct = {}
        for row_num, values in wb.iter_rows(sheet, min_row=data_start):
            if col_idx >= len(values):
                total += 1
                continue
            val = values[col_idx]
            total += 1
            if val is None or (isinstance(val, str) and val.strip() == ""):
                continue
            non_empty += 1
            if isinstance(val, bool):
                key = str(val)
            elif isinstance(val, (int, float)):
                numbers.append(float(val))
                key = repr(val)
            else:
                key = str(val)
            distinct[key] = distinct.get(key, 0) + 1

        lines = ["Column stats for '%s' in sheet '%s' (data from row %d):"
                 % (args["column"], resolved, data_start)]
        lines.append("  rows scanned : %d" % total)
        lines.append("  non-empty    : %d" % non_empty)
        lines.append("  empty        : %d" % (total - non_empty))
        lines.append("  distinct     : %d" % len(distinct))

        if numbers:
            n = len(numbers)
            s = sum(numbers)
            mean = s / n
            srt = sorted(numbers)
            mid = n // 2
            median = srt[mid] if n % 2 else (srt[mid - 1] + srt[mid]) / 2.0
            lines.append("  numeric count: %d" % n)
            lines.append("  sum          : %s" % _fmt_num(s))
            lines.append("  mean         : %s" % _fmt_num(mean))
            lines.append("  median       : %s" % _fmt_num(median))
            lines.append("  min          : %s" % _fmt_num(srt[0]))
            lines.append("  max          : %s" % _fmt_num(srt[-1]))
        else:
            lines.append("  (no numeric values found in this column)")

        # Top categorical values (handy for non-numeric columns).
        if distinct and not numbers:
            top = sorted(distinct.items(), key=lambda kv: kv[1], reverse=True)[:10]
            lines.append("  top values   :")
            for k, cnt in top:
                lines.append("      %s (%d)" % (k, cnt))
        return "\n".join(lines)
    finally:
        wb.close()


def _fmt_num(x):
    """Render a float cleanly (drop trailing .0 for whole numbers)."""
    if isinstance(x, float) and x.is_integer():
        return str(int(x))
    return "%.6g" % x if isinstance(x, float) else str(x)


# ===========================================================================
# MCP tool registry (name -> (handler, description, input schema))
# ===========================================================================

TOOLS = {
    "excel_list_workbooks": {
        "handler": tool_list_workbooks,
        "description": "List the Excel workbooks (.xlsx/.xlsm) available to read "
                       "in the configured folder, with file sizes.",
        "schema": {
            "type": "object",
            "properties": {},
        },
    },
    "excel_list_sheets": {
        "handler": tool_list_sheets,
        "description": "List the sheet names and declared used-ranges in a workbook.",
        "schema": {
            "type": "object",
            "properties": {
                "workbook": {"type": "string",
                             "description": "Workbook name (loose match; "
                                            "extension optional)."},
            },
            "required": ["workbook"],
        },
    },
    "excel_get_headers": {
        "handler": tool_get_headers,
        "description": "Return the header row of a sheet, one entry per column, "
                       "with its column letter. Use this before reading data so "
                       "you know which columns exist.",
        "schema": {
            "type": "object",
            "properties": {
                "workbook": {"type": "string"},
                "sheet": {"type": "string",
                          "description": "Sheet name or 1-based index. "
                                         "Defaults to the first sheet."},
                "header_row": {"type": "integer",
                               "description": "Row number of the header "
                                              "(default 1)."},
            },
            "required": ["workbook"],
        },
    },
    "excel_read_range": {
        "handler": tool_read_range,
        "description": "Read cell values from a sheet. Optionally pass an A1 "
                       "range like 'A1:D50'. Output is capped for safety; use a "
                       "range to page through large sheets.",
        "schema": {
            "type": "object",
            "properties": {
                "workbook": {"type": "string"},
                "sheet": {"type": "string",
                          "description": "Sheet name or 1-based index. "
                                         "Defaults to the first sheet."},
                "range": {"type": "string",
                          "description": "Optional A1 range, e.g. 'A1:D50'. "
                                         "Omit to read from the top."},
            },
            "required": ["workbook"],
        },
    },
    "excel_search": {
        "handler": tool_search,
        "description": "Find cells whose value contains a query string. Searches "
                       "all sheets unless 'sheet' is given. Returns cell "
                       "references and values.",
        "schema": {
            "type": "object",
            "properties": {
                "workbook": {"type": "string"},
                "query": {"type": "string",
                          "description": "Text to search for."},
                "sheet": {"type": "string",
                          "description": "Optional: restrict to one sheet."},
                "case_sensitive": {"type": "boolean",
                                   "description": "Default false."},
            },
            "required": ["workbook", "query"],
        },
    },
    "excel_column_stats": {
        "handler": tool_column_stats,
        "description": "Summarise one column: count, sum, mean, median, min, "
                       "max for numeric data, or top values for categorical "
                       "data. Column may be a letter (C), a number (3), or a "
                       "header name.",
        "schema": {
            "type": "object",
            "properties": {
                "workbook": {"type": "string"},
                "column": {"type": "string",
                           "description": "Column letter, 1-based number, or "
                                          "header name."},
                "sheet": {"type": "string",
                          "description": "Sheet name or 1-based index. "
                                         "Defaults to the first sheet."},
                "header_row": {"type": "integer",
                               "description": "Header row number (default 1)."},
                "data_start_row": {"type": "integer",
                                   "description": "First data row "
                                                  "(default header_row + 1)."},
            },
            "required": ["workbook", "column"],
        },
    },
}


# ===========================================================================
# JSON-RPC / MCP stdio server
# ===========================================================================

def _make_result(req_id, result):
    return {"jsonrpc": "2.0", "id": req_id, "result": result}


def _make_error(req_id, code, message):
    return {"jsonrpc": "2.0", "id": req_id,
            "error": {"code": code, "message": message}}


def _tool_text_result(text):
    return {"content": [{"type": "text", "text": text}], "isError": False}


def _tool_error_result(text):
    return {"content": [{"type": "text", "text": text}], "isError": True}


def handle_request(msg, folder):
    """
    Handle one parsed JSON-RPC message. Returns a response dict, or None for
    notifications (which must not be answered).
    """
    method = msg.get("method")
    req_id = msg.get("id")
    is_notification = "id" not in msg

    if method == "initialize":
        result = {
            "protocolVersion": PROTOCOL_VERSION,
            "capabilities": {"tools": {}},
            "serverInfo": {"name": SERVER_NAME, "version": SERVER_VERSION},
        }
        return _make_result(req_id, result)

    if method in ("notifications/initialized", "initialized", "notifications/cancelled"):
        return None  # notification: no response

    if method == "ping":
        return _make_result(req_id, {})

    if method == "tools/list":
        tool_list = []
        for name, spec in TOOLS.items():
            tool_list.append({
                "name": name,
                "description": spec["description"],
                "inputSchema": spec["schema"],
            })
        return _make_result(req_id, {"tools": tool_list})

    if method == "tools/call":
        params = msg.get("params") or {}
        name = params.get("name")
        arguments = params.get("arguments") or {}
        spec = TOOLS.get(name)
        if spec is None:
            return _make_result(
                req_id,
                _tool_error_result("Unknown tool: %s" % name))
        try:
            text = spec["handler"](folder, arguments)
            log("tool '%s' args=%s -> %d chars"
                % (name, json.dumps(arguments, ensure_ascii=False), len(text)))
            return _make_result(req_id, _tool_text_result(text))
        except WorkbookError as exc:
            log("tool '%s' workbook error: %s" % (name, exc))
            return _make_result(req_id, _tool_error_result(str(exc)))
        except Exception as exc:   # never crash the server on one bad call
            log("tool '%s' unexpected error: %r" % (name, exc))
            return _make_result(
                req_id,
                _tool_error_result("Internal error running %s: %s"
                                   % (name, exc)))

    # Unknown method.
    if is_notification:
        return None
    return _make_error(req_id, -32601, "Method not found: %s" % method)


def serve(folder):
    """Main stdio loop: read newline-delimited JSON-RPC, write responses."""
    log("starting; interpreter=%s" % sys.executable)
    log("workbook folder=%s" % folder)
    if not os.path.isdir(folder):
        log("WARNING: workbook folder does not exist yet: %s" % folder)

    for raw in sys.stdin:
        line = raw.strip()
        if not line:
            continue
        try:
            msg = json.loads(line)
        except json.JSONDecodeError as exc:
            log("could not parse line as JSON: %s" % exc)
            # Cannot know the id; emit a parse error with null id.
            sys.stdout.write(json.dumps(
                _make_error(None, -32700, "Parse error")) + "\n")
            sys.stdout.flush()
            continue

        # A batch (list) is technically valid JSON-RPC; handle defensively.
        messages = msg if isinstance(msg, list) else [msg]
        for m in messages:
            if not isinstance(m, dict):
                continue
            try:
                response = handle_request(m, folder)
            except Exception as exc:
                log("fatal handler error: %r" % exc)
                response = _make_error(m.get("id"), -32603,
                                       "Internal error: %s" % exc)
            if response is not None:
                sys.stdout.write(json.dumps(response, ensure_ascii=False) + "\n")
                sys.stdout.flush()

    log("stdin closed; exiting")


# ===========================================================================
# CLI
# ===========================================================================

def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Read-only Excel (.xlsx) MCP server for VSCode Continue.")
    parser.add_argument("--folder",
                        default=os.environ.get("EXCEL_WORKBOOK_FOLDER",
                                               WORKBOOK_FOLDER),
                        help="Folder containing .xlsx/.xlsm workbooks. Falls "
                             "back to the EXCEL_WORKBOOK_FOLDER environment "
                             "variable, then the CONFIG block default.")
    parser.add_argument("--check", action="store_true",
                        help="Print environment/config diagnostics and exit.")
    parser.add_argument("--list", action="store_true",
                        help="List readable workbooks in the folder and exit.")
    args = parser.parse_args(argv)

    folder = args.folder

    if args.check:
        print("excel_mcp environment check")
        print("  python executable : %s" % sys.executable)
        print("  python version    : %s" % sys.version.split()[0])
        print("  workbook folder   : %s" % (folder or "(NOT SET - required)"))
        print("  folder exists     : %s" % (bool(folder) and os.path.isdir(folder)))
        if folder and os.path.isdir(folder):
            try:
                files = list_workbook_files(folder)
                print("  workbooks found   : %d" % len(files))
                for f in files:
                    print("      - %s" % f)
            except WorkbookError as exc:
                print("  error listing     : %s" % exc)
        print("  tools registered  : %d (%s)"
              % (len(TOOLS), ", ".join(TOOLS.keys())))
        return 0

    # The workbook folder is REQUIRED: the server only reads inside it and
    # must not start unconfined.
    if not folder:
        log("FATAL: no workbook folder configured. Pass --folder, set the "
            "EXCEL_WORKBOOK_FOLDER environment variable, or set "
            "WORKBOOK_FOLDER in this file.")
        return 2
    if not os.path.isdir(folder):
        log("FATAL: the configured workbook folder does not exist or is not "
            "a directory: %s" % folder)
        return 2

    if args.list:
        try:
            print(tool_list_workbooks(folder, {}))
        except WorkbookError as exc:
            print("Error: %s" % exc, file=sys.stderr)
            return 1
        return 0

    serve(folder)
    return 0


if __name__ == "__main__":
    sys.exit(main())
