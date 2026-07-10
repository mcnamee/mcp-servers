# mcp-servers
My MCP Servers for Productivity

Each file in this repo is a single-file, stdio-transport MCP server. Most of
them are standard-library only; a few need one extra pip package. Install
into the SAME Python interpreter that your MCP client (e.g. Continue) will
launch the server with.

## Dependencies per server

| Server | pip install | Notes |
|---|---|---|
| `confluence.py` | _none_ | standard library only |
| `knowledge-base.py` | _none_ | standard library only (keyword search) |
| `knowledge-base-semantic.py` | _none_ | standard library only; meaning-based retrieval driven by your connected model (no local embedding model, no network) |
| `ms-excel.py` | _none_ | standard library only (parses .xlsx as a zip of XML) |
| `ms-word.py` | `pip install python-docx` | also pulls in `lxml` (compiled) and `typing_extensions` |
| `ms-outlook.py` | `pip install pywin32` | Windows only (COM automation of classic Outlook) |
| `pdf-to-md.py` | `pip install pymupdf pymupdf4llm` | OCR of scanned PDFs additionally requires Tesseract installed on the machine (not a pip package) |

## Install everything at once

```
pip install python-docx pymupdf pymupdf4llm pywin32
```

(Drop `pywin32` if you're not on Windows / not using `ms-outlook.py`.)

## File access policy

Every server that touches the filesystem is confined to the folder(s) named
in its configuration, and that configuration is **required** — a server will
refuse to start unconfined rather than fall back to "anywhere":

| Server | Confined to | Required setting |
|---|---|---|
| `ms-word.py` | open/save only inside the document root | `--document-root` / `MSWORD_DOCUMENT_ROOT` / `DOCUMENT_ROOT` constant |
| `ms-excel.py` | reads only inside the workbook folder | `--folder` / `EXCEL_WORKBOOK_FOLDER` / `WORKBOOK_FOLDER` constant |
| `knowledge-base.py` | reads only inside the docs folder | `--docs-dir` / `REFERENCE_DOCS_DIR` |
| `knowledge-base-semantic.py` | reads only inside the docs folder | `--docs-dir` / `KB_DOCS_DIR` |
| `pdf-to-md.py` | reads only the input folder, writes only the output folder | `--input-dir` + `--output-dir` (or `PDF2MD_INPUT_DIR`/`PDF2MD_OUTPUT_DIR`) |
| `confluence.py` | no local file access unless `CONFLUENCE_KB_DIR` is set; then writes only inside that folder | n/a (mirroring is optional; unset = no file access) |
| `ms-outlook.py` | no file access (local COM only; reads the optional `--blacklist-file` once at startup) | n/a |

In all cases paths are resolved (symlinks included) before the containment
check, so a symlink dropped inside a configured folder cannot reach files
outside it.

## Notes

- `ms-word.py` and `pdf-to-md.py` log `sys.executable` on startup, so if a
  dependency reports as "missing" even after installing it, check that you
  installed into the same interpreter your MCP client launches the server
  with, e.g.:
  ```
  "C:\path\to\python.exe" -m pip install python-docx pymupdf pymupdf4llm pywin32
  ```
- For airgapped/offline installs, see the docstring at the top of
  `ms-word.py` for the wheel-sideloading steps (same pattern applies to
  `pdf-to-md.py`'s dependencies).

## Installation into Continue (config.yaml)

All servers speak MCP over stdio. Add each one as an entry under the
`mcpServers:` block of Continue's `config.yaml`. Use the full path to the
Python interpreter you installed dependencies into (not a bare `python`),
and set `PYTHONUTF8: "1"` so Windows' default codepage can't corrupt the
stdio JSON stream. After editing `config.yaml`, use VSCode's
"Developer: Reload Window" rather than toggling the server, to avoid a
known "already connected to transport" reconnection bug.

### confluence.py

Configuration is via environment variables (non-secret settings also have
equivalent CLI flags, which take priority over the env vars). **Credentials
are env-var only** — there are deliberately no `--token`/`--user`/`--password`
flags, because command-line arguments are visible to other local users in
process listings.

| Env var | CLI flag | Purpose |
|---|---|---|
| `CONFLUENCE_BASE_URL` | `--base-url` | Base URL incl. any context path, no trailing slash |
| `CONFLUENCE_TOKEN` | _(env only)_ | Personal Access Token, sent as Bearer (preferred over basic auth) |
| `CONFLUENCE_USER` | _(env only)_ | Username for basic auth (fallback if no token) |
| `CONFLUENCE_PASSWORD` | _(env only)_ | Password for basic auth |
| `CONFLUENCE_CA_CERT` | `--ca-cert` | Path to a PEM CA bundle for an internal CA |
| `CONFLUENCE_VERIFY_SSL=false` | `--insecure` | Disable TLS certificate verification |
| `CONFLUENCE_TIMEOUT` | `--timeout` | Request timeout in seconds (default 30) |
| `CONFLUENCE_MAX_BODY` | `--max-body` | Truncate page bodies to N chars, 0 = unlimited (default). Applies only to text returned to the model, not to files saved via `CONFLUENCE_KB_DIR` |
| `CONFLUENCE_KB_DIR` | `--kb-dir` | If set, every page read is also saved as a Markdown file (`Confluence - <title>.md`, overwritten each time) into this folder — handy for feeding a local RAG knowledge base, e.g. alongside `knowledge-base.py` |

```yaml
mcpServers:
  - name: confluence
    command: C:\path\to\python.exe
    args:
      - C:\path\to\confluence.py
      - --max-body
      - "20000"
      - --kb-dir
      - C:\reference-docs\confluence
    env:
      CONFLUENCE_BASE_URL: https://confluence.internal.example.com
      CONFLUENCE_TOKEN: your-personal-access-token
      PYTHONUTF8: "1"
```

### knowledge-base.py

| CLI flag | Purpose |
|---|---|
| `--docs-dir` | Folder of reference docs to expose (`.md`/`.markdown`/`.txt`), searched recursively. Falls back to `REFERENCE_DOCS_DIR` env var |
| `--check` | List the folder's documents to stderr, then exit (no server) |
| `--version` | Print version and exit |

```yaml
mcpServers:
  - name: reference
    command: python
    args:
      - C:\path\to\knowledge-base.py
      - --docs-dir
      - C:\reference-docs
    env:
      PYTHONUTF8: "1"
```

> For meaning-based (rather than keyword) retrieval over the same kind of
> folder, see `knowledge-base-semantic.py` below. You can run both at once
> (different `name:` and folder).

### knowledge-base-semantic.py

Meaning-based retrieval over a folder of your own markdown files (notes,
policies, reference docs). Ask a question in your own words — e.g. *"can I
extend my work trip by two days, pay my own weekend accommodation, and fly
back Monday?"* — and the agent finds the relevant part of your travel policy
and reasons over it, without you naming the file.

**How it stays "semantic" with no embedding model and no network:** MCP has no
way for a server to ask the client for embeddings, and Continue does not
support MCP *sampling* (server-initiated LLM calls), so the server cannot call
your model itself. Instead it hands the agent — which *is* the model you
connect to — a compact, structure-aware **map** of the knowledge base
(`kb_outline`: every document + its section headings + a one-line lead). The
agent picks the right document/section by *meaning*, then reads exactly that
section (`kb_read`). Retrieval by meaning, using the model you already talk to,
with zero dependencies. (If you ever want true vector embeddings, that requires
either a local embedding model or a remote embeddings API — neither of which
this offline-friendly server uses.)

| Tool | Purpose |
|---|---|
| `kb_list` | List every document with its title (inventory) |
| `kb_outline` | The map: every document's heading tree + a lead line per section — call this first for a topic question, then pick by meaning |
| `kb_search` | Keyword search across all docs; reports the matching section heading and a snippet (fast lookup / narrowing a large KB) |
| `kb_read` | Read one document in full, or just one named `section` of it (loose name/heading matching) |

| CLI flag | Purpose |
|---|---|
| `--docs-dir` | Folder of markdown/text docs to expose (`.md`/`.markdown`/`.txt`), searched recursively. Falls back to the `KB_DOCS_DIR` env var |
| `--check` | List the folder's documents (and a heading count each) to stderr, then exit (no server) |
| `--version` | Print version and exit |

```yaml
mcpServers:
  - name: kb
    command: C:\path\to\python.exe
    args:
      - C:\path\to\knowledge-base-semantic.py
      - --docs-dir
      - C:\Users\me\knowledge-base
    env:
      PYTHONUTF8: "1"
```

### ms-excel.py

| CLI flag | Purpose |
|---|---|
| `--folder` | **Required.** Folder of `.xlsx`/`.xlsm` workbooks to expose — the server only reads files inside it and refuses to start without one. Falls back to the `EXCEL_WORKBOOK_FOLDER` env var, then the `WORKBOOK_FOLDER` constant in the file |
| `--check` | Print environment/config diagnostics and exit (no server) |
| `--list` | List readable workbooks in the folder and exit (no server) |

```yaml
mcpServers:
  - name: excel
    command: C:\path\to\python.exe
    args:
      - C:\path\to\ms-excel.py
      - --folder
      - C:\path\to\your\workbooks
    env:
      PYTHONUTF8: "1"
```

### ms-outlook.py

Windows only — requires classic Win32 Outlook (not "New Outlook") installed,
running, and logged into a profile.

| CLI flag | Purpose |
|---|---|
| `--blacklist-file` | Path to a file of extra content-blacklist terms (one per line, `#` for comments), added to the built-in list. Falls back to the `OUTLOOK_BLACKLIST_FILE` env var |
| `--search-folders` | Comma-separated folder names used as the **default** folder set for `outlook_search_recent`, overriding the `SEARCH_ALL_FOLDERS` value in the file (e.g. `"Inbox,Sent Items,Archive"`). A per-call `folders` argument still takes priority. Falls back to the `OUTLOOK_SEARCH_FOLDERS` env var |
| `--require-blacklist` | Fail closed: refuse to start unless the content blacklist has at least one active term, so a missing/empty terms file cannot silently disable the compliance filter. Also via `OUTLOOK_REQUIRE_BLACKLIST=1` or the `REQUIRE_BLACKLIST` constant in the file |
| `--check` | Connect to Outlook, print diagnostics + blacklist status to stderr, then exit (no server) |
| `--version` | Print version and exit |

The content blacklist also applies to **folder names**: folders whose
store/path matches a blacklisted term are withheld from
`outlook_list_folders` and skipped by `outlook_search_recent` (results are
labelled with their folder path, so a marked folder name never appears in
output).

Everything else is configured by editing the `USER CONFIGURATION` block at
the top of `ms-outlook.py` directly (there are no extra CLI flags/env vars
for these):

| Setting | Purpose |
|---|---|
| `BLACKLIST_TERMS` | Built-in list of classification/compliance terms that cause an item to be withheld from the AI entirely |
| `BLACKLIST_MATCH_MODE` | `"word"` (default, whole-term match) or `"substring"` (for terms containing punctuation) |
| `RESTRICT_DATE_FORMAT` | Locale-sensitive date format Outlook expects in its `Restrict()` filter — switch this if `outlook_get_calendar` returns zero events on a non-US-locale machine |
| `MAX_BODY_CHARS` / `CALENDAR_HARD_CAP` / `SEARCH_SCAN_CAP` | Safety caps on body length / items scanned |
| `SEARCH_ALL_FOLDERS` | Folder names (matched across every store) that `outlook_search_recent` searches by default — `["Inbox", "Sent Items", "Archive"]`; use `outlook_list_folders` to see real folder names first. This is only the built-in default: override it at launch with `--search-folders`, or per call by passing a `folders` argument to `outlook_search_recent` |

```yaml
mcpServers:
  - name: outlook
    command: C:\path\to\python.exe
    args:
      - C:\path\to\ms-outlook.py
      - --blacklist-file
      - C:\config\outlook-blacklist.txt
      - --search-folders
      - "Inbox,Sent Items,Archive"
    env:
      PYTHONUTF8: "1"
```

### ms-word.py

| CLI flag | Purpose |
|---|---|
| `--check` | Run an offline open/edit/save/reopen self-test and exit (no server) |
| `--author` | Author name stamped on Word tracked changes. Falls back to the `MSWORD_AUTHOR` env var, then the `TRACKED_CHANGE_AUTHOR` config value in the file. Can also be overridden per-call via the `author` argument on the editing tools (`msword_replace_text`, `msword_set_paragraph_text`, `msword_insert_paragraph`, `msword_delete_paragraph`) |
| `--document-root` | **Required.** Path sandbox: every open/save must be inside this directory tree, and the server refuses to start without one (`--check` is exempt — the self-test sandboxes itself to its own temp folder). Falls back to the `MSWORD_DOCUMENT_ROOT` env var, then the `DOCUMENT_ROOT` config value. This is the only write-capable server in the suite, and the model chooses the open/save paths |

Tracked changes are recorded the way Word itself records them: replacements
are diffed **word-by-word** (only the words that actually change are marked
as deleted/inserted — never "whole paragraph deleted + whole paragraph
reinserted"), and whole-paragraph inserts/deletes include the paragraph mark
so accepting/rejecting adds or removes the paragraph itself. Changes can be
accepted/rejected all at once or individually by id. While changes are
pending, `msword_get_content`/`msword_search` show the final ("No Markup")
view.

```yaml
mcpServers:
  - name: msword-py
    command: C:\path\to\python.exe
    args:
      - C:\path\to\ms-word.py
      - --author
      - Matt
      - --document-root
      - C:\Users\me\Documents\ai_docs
    env:
      PYTHONUTF8: "1"
```

### pdf-to-md.py

| CLI flag | Purpose |
|---|---|
| `--input-dir` | **Required.** Folder containing the source PDFs. Falls back to the `PDF2MD_INPUT_DIR` env var |
| `--output-dir` | **Required.** Folder to write `.md` files into. Falls back to the `PDF2MD_OUTPUT_DIR` env var |
| `--recursive` | Also search sub-folders of `--input-dir` (sub-folder structure is mirrored in the output). Also via `PDF2MD_RECURSIVE=1` |

```yaml
mcpServers:
  - name: pdf2md
    command: C:\path\to\python.exe
    args:
      - C:\path\to\pdf-to-md.py
      - --input-dir
      - C:\Reference\PDFs
      - --output-dir
      - C:\Reference\Markdown
      - --recursive
    env:
      PYTHONUTF8: "1"
```

## Usage examples

These are natural-language prompts you can give an AI agent (e.g. in
Continue's agent mode) once the relevant server is wired in. Each maps to
one or more of the tools the server exposes.

### confluence.py

1. "Search Confluence for our incident response runbook." → `confluence_search`
2. "Find pages in the DOCS space that mention 'release notes' and were updated in the last 30 days." → `confluence_search_cql`
3. "Pull up the full content of Confluence page 393217." → `confluence_get_page`
4. "Open the 'Q3 Roadmap' page in the PROD space and summarise it." → `confluence_get_page_by_title`
5. "List every page under the 'Engineering Handbook' in the DOCS space, direct children only." → `confluence_list_pages_under`
6. "Pull the onboarding runbook into our local knowledge base for offline search." → `confluence_get_page` (or `confluence_get_page_by_title`), automatically mirrored to Markdown when `--kb-dir`/`CONFLUENCE_KB_DIR` is configured, so `knowledge-base.py`'s `reference_search`/`reference_get` can find it afterwards

### knowledge-base.py

1. "What reference documents do we have available?" → `reference_list`
2. "Find any reference material about our procurement policy." → `reference_search`
3. "Read the full expense-reporting policy document and tell me the approval limits." → `reference_get`
4. "Search our reference docs for anything about onboarding, then read whichever one covers IT equipment." → `reference_search` followed by `reference_get`

### knowledge-base-semantic.py

1. "Using my knowledge base, can I extend my work trip by 2 days, pay for my own accommodation for the weekend, and fly back Monday?" → `kb_outline` to map the docs, then `kb_read` (with a `section`) on the travel policy's trip-extension section, and reason over it
2. "What's in my knowledge base?" → `kb_list`
3. "Give me a map of everything in my knowledge base so I can see what topics are covered." → `kb_outline`
4. "Find anything mentioning 'accommodation' or 'per diem'." → `kb_search`
5. "Read just the 'Expense Claims' section of the travel policy." → `kb_read` with `section`
6. "Open the whole onboarding checklist document." → `kb_read` (no `section`)

### ms-excel.py

1. "What Excel workbooks are available for me to look at?" → `excel_list_workbooks`
2. "List the sheets in the 'budget' workbook." → `excel_list_sheets`
3. "What are the column headers on the 'Q3' sheet of the budget workbook?" → `excel_get_headers`
4. "Read rows A1:D50 from the Q3 sheet." → `excel_read_range`
5. "Find every cell in the budget workbook that mentions 'Marketing'." → `excel_search`
6. "Give me the sum, average, min and max of the Revenue column on the Q3 sheet." → `excel_column_stats`

### ms-outlook.py

1. "Show me my 10 most recent unread emails." → `outlook_list_recent_emails`
2. "Search my inbox for anything from 'Jane Smith' about the contract renewal." → `outlook_search_emails`
3. "Open that email from the vendor and summarise the key dates." → `outlook_get_email`
4. "What's on my calendar for the next 7 days?" → `outlook_get_calendar`
5. "What did I send last week?" → `outlook_list_sent_emails`
6. "Find everything about the 'Acme renewal' across my Inbox, Sent Items and Archive from the last month." → `outlook_search_recent`
7. "Search only my 'Projects' and 'Sent Items' folders for anything about the budget review." → `outlook_search_recent` with a `folders` argument overriding the default set
8. "What are my actual Outlook folder names, so I can point the search at the right archive?" → `outlook_list_folders`

### ms-word.py

1. "Open the proposal.docx and show me its full text." → `msword_open` + `msword_get_content`
2. "Find every mention of 'Acme Corp' in the contract and replace it with 'Acme Corporation'." → `msword_search` + `msword_replace_text`
3. "Add a 'Next Steps' heading and a summary paragraph to the end of the report, then save it." → `msword_add_heading` + `msword_add_paragraph` + `msword_save`
4. "Pull out the data from every table in the document as structured rows." → `msword_get_tables`
5. "Add a 3x4 pricing table to the end of the quote document with these values, using the 'Table Grid' style." → `msword_add_table` + `msword_save`
6. "Change 'DRAFT' to 'FINAL' throughout the report as a tracked change so it shows up as a Word revision for review." → `msword_replace_text` with `track_changes=true`
7. "Rewrite the third paragraph to be more concise, showing your edits as tracked changes — only mark the words you actually changed." → `msword_set_paragraph_text` with `track_changes=true` (old vs new text is diffed word-by-word, like editing in Word with Track Changes on)
8. "Add a new paragraph after the introduction as a tracked insertion, so reviewers can reject it if they disagree." → `msword_insert_paragraph` with `track_changes=true`
9. "Delete the whole limitation-of-liability paragraph as a tracked change — struck out, so legal can accept or reject it." → `msword_delete_paragraph` with `track_changes=true`
10. "What tracked changes are currently in this document, and who made them?" → `msword_list_changes`
11. "Accept Jane's two changes in the pricing section but leave everything else pending." → `msword_list_changes` + `msword_accept_changes` with those change ids
12. "Reject just the change that deleted the warranty sentence." → `msword_list_changes` + `msword_reject_changes` with that change id
13. "Accept all the tracked changes in this document now that legal has signed off." → `msword_accept_all_changes`
14. "Reject all the tracked changes and revert this document to its original wording." → `msword_reject_all_changes`

### pdf-to-md.py

1. "Convert every PDF in the reference folder to Markdown." → `convert_all_pdfs`
2. "Convert just the 'procurement policy' PDF to Markdown." → `convert_pdf_to_markdown`
3. "Reconvert all PDFs to Markdown even though some already have a .md file, since the source PDFs changed." → `convert_all_pdfs` with `force=true`
4. "Convert all our compliance PDFs (including those in sub-folders) to Markdown so the knowledge-base server can search them." → `convert_all_pdfs` (with `--recursive` set at startup) feeding into `knowledge-base.py`'s `reference_search`
