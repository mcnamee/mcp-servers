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
| `jira.py` | _none_ | standard library only (read-only, Jira Data Center v2 REST API) |
| `knowledge-base.py` | _none_ | standard library only (keyword search) |
| `knowledge-base-rag.py` | `pip install chromadb` | true RAG: local ChromaDB vector index + your embeddings API (HTTP is stdlib `urllib`, no `requests`) |
| `ms-excel.py` | _none_ | standard library only (parses .xlsx as a zip of XML) |
| `ms-word.py` | `pip install python-docx` | also pulls in `lxml` (compiled) and `typing_extensions` |
| `ms-outlook.py` | `pip install pywin32` | Windows only (COM automation of classic Outlook) |
| `pdf-to-md.py` | `pip install pymupdf pymupdf4llm` | OCR of scanned PDFs additionally requires Tesseract installed on the machine (not a pip package) |

## Install everything at once

```
pip install python-docx pymupdf pymupdf4llm pywin32 chromadb
```

(Drop `pywin32` if you're not on Windows / not using `ms-outlook.py`.)

## File access policy

Every server that touches the filesystem is confined to the folder(s) named
in its configuration, and that configuration is **required** — a server will
refuse to start unconfined rather than fall back to "anywhere":

| Server | Confined to | Required setting |
|---|---|---|
| `ms-word.py` | open/save only inside the document root (and the output folder, if set); new docs written to the output folder; Markdown mirrored to the knowledge-base folder | `--document-root` / `MSWORD_DOCUMENT_ROOT` / `DOCUMENT_ROOT` constant (required); optional `--output-dir` and `--kb-dir` |
| `ms-excel.py` | reads only inside the workbook folder | `--folder` / `EXCEL_WORKBOOK_FOLDER` / `WORKBOOK_FOLDER` constant |
| `knowledge-base.py` | reads only inside the docs folder | `--docs-dir` / `REFERENCE_DOCS_DIR` |
| `knowledge-base-rag.py` | reads only inside the docs folder; writes only the vector-index folder (default `<docs-dir>\.kb-rag-index`); network only to the endpoint(s) you configure | `--docs-dir` / `KB_DOCS_DIR` + `--embed-url` / `KB_EMBED_URL` |
| `pdf-to-md.py` | reads only the input folder, writes only the output folder | `--input-dir` + `--output-dir` (or `PDF2MD_INPUT_DIR`/`PDF2MD_OUTPUT_DIR`) |
| `confluence.py` | no local file access unless `CONFLUENCE_KB_DIR` is set; then writes only inside that folder | n/a (mirroring is optional; unset = no file access) |
| `jira.py` | no local file access (HTTP GET to Jira only; reads the optional `JIRA_CA_CERT` bundle once at startup) | n/a |
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
  "C:\path\to\python.exe" -m pip install python-docx pymupdf pymupdf4llm pywin32 chromadb
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

### jira.py

Read-only access to Jira Data Center (v2 REST API). Every request is an HTTP
GET — there is no code path that creates, edits, transitions, comments on, or
deletes anything. Like `confluence.py`, **credentials are env-var only** (no
`--token`/`--user`/`--password` flags).

| Env var | CLI flag | Purpose |
|---|---|---|
| `JIRA_BASE_URL` | `--base-url` | Base URL incl. any context path, no trailing slash |
| `JIRA_TOKEN` | _(env only)_ | Personal Access Token, sent as Bearer (preferred; Jira DC 8.14+) |
| `JIRA_USER` | _(env only)_ | Username for basic auth (fallback if no token) |
| `JIRA_PASSWORD` | _(env only)_ | Password for basic auth |
| `JIRA_PROJECTS` | `--projects` | Optional comma-separated **project-key allowlist** (e.g. `"ABC,DEF"`). When set, every tool is confined to those projects: searches are scoped with an AND clause, issue keys outside the list are refused, and other projects are hidden from `jira_list_projects` |
| `JIRA_CA_CERT` | `--ca-cert` | Path to a PEM CA bundle for an internal CA |
| `JIRA_VERIFY_SSL=false` | `--insecure` | Disable TLS certificate verification |
| `JIRA_TIMEOUT` | `--timeout` | Request timeout in seconds (default 30) |
| `JIRA_MAX_BODY` | `--max-body` | Truncate issue descriptions to N chars, 0 = unlimited (default) |
| — | `--check` | Connect to Jira, print who you are authenticated as + visible project count to stderr, then exit (no server) |

```yaml
mcpServers:
  - name: jira
    command: C:\path\to\python.exe
    args:
      - C:\path\to\jira.py
    env:
      JIRA_BASE_URL: https://jira.internal.example.com
      JIRA_TOKEN: your-personal-access-token
      JIRA_PROJECTS: "ABC,DEF"        # optional allowlist
      PYTHONUTF8: "1"
```

> Targets Jira **Data Center / Server** (plain-text descriptions via the v2
> API). Jira Cloud's v3 API returns rich-text documents and is not supported.

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

> For true vector (RAG) retrieval over the same kind of folder, see
> `knowledge-base-rag.py` below. You can run both at once (different
> `name:` and folder).

### knowledge-base-rag.py

True RAG (Retrieval-Augmented Generation) over a folder of your own markdown
files, in three stages:

1. **Index** — documents are split into heading-aware chunks, embedded via
   your embeddings API endpoint, and stored in a local ChromaDB vector
   database on disk. Indexing is incremental: only new/changed files are
   re-embedded, deleted files are removed.
2. **Retrieve** — a question is embedded the same way and the most
   semantically similar chunks come back with source file, section heading
   and similarity score.
3. **Generate** — (optional) the retrieved chunks + question go to a
   chat-completions endpoint, which writes a grounded answer citing its
   sources. With no chat endpoint configured, `kb_ask` returns the retrieved
   context and the agent you're already talking to writes the answer — so
   generation works either way.

Requires `pip install chromadb` (ChromaDB's anonymised telemetry is disabled
in the file; HTTP to your endpoints is stdlib `urllib`). **API keys are
env-var only** — there are deliberately no `--*-api-key` flags, because
command-line arguments are visible to other local users in process listings.
Note that retrieved document text *is* sent to the endpoints you configure —
that is what RAG is — so point it only at material appropriate for those APIs.

| Tool | Purpose |
|---|---|
| `kb_index` | Build/update the vector index (incremental; `force=true` rebuilds — needed after changing embedding model) |
| `kb_retrieve` | Semantic vector search: top-k most similar chunks, with source file, heading and similarity score |
| `kb_ask` | Full RAG: retrieve, then generate a grounded cited answer (or return context for the agent, if no chat endpoint) |
| `kb_status` | Documents vs index freshness + configuration summary (never shows keys) |

| Env var | CLI flag | Purpose |
|---|---|---|
| `KB_DOCS_DIR` | `--docs-dir` | **Required.** Folder of `.md`/`.markdown`/`.txt` docs, searched recursively |
| `KB_INDEX_DIR` | `--index-dir` | ChromaDB folder (default `<docs-dir>\.kb-rag-index`) |
| `KB_COLLECTION` | `--collection` | ChromaDB collection name (default `kb-rag`) |
| `KB_EMBED_URL` | `--embed-url` | **Required.** Full URL of the embeddings endpoint |
| `KB_EMBED_MODEL` | `--embed-model` | Model name sent in embed requests (omit if the endpoint fixes one) |
| `KB_EMBED_API_KEY` | _(env only)_ | API key for the embeddings endpoint |
| `KB_EMBED_AUTH_HEADER` | `--embed-auth-header` | Header the key is sent in — default `Authorization` (as `Bearer <key>`); any other name (e.g. Azure's `api-key`) sends the raw key |
| `KB_EMBED_STYLE` | `--embed-style` | Request format: `openai` (default; batch `{"input": [...]}`) or `ollama` (`{"prompt": ...}` one-per-request). Response parsing also accepts bare `embedding`/`embeddings` shapes, so most bespoke internal endpoints work unchanged |
| `KB_EMBED_BATCH` | `--embed-batch` | Texts per embeddings request, openai style (default 16) |
| `KB_EMBED_QUERY_PREFIX` | `--embed-query-prefix` | Prefix for query embeds, for models that need it (e5-style `"query: "`) |
| `KB_EMBED_DOC_PREFIX` | `--embed-doc-prefix` | Prefix for document embeds (`"passage: "`) |
| `KB_EMBED_EXTRA_HEADERS` | _(env only)_ | JSON object of extra HTTP headers for the embed endpoint |
| `KB_CHAT_URL` | `--chat-url` | *Optional.* Chat-completions endpoint for the generate step (OpenAI shape; Ollama `/api/chat` and `/api/generate` response shapes also parsed) |
| `KB_CHAT_MODEL` | `--chat-model` | Generation model name |
| `KB_CHAT_API_KEY` | _(env only)_ | API key for the chat endpoint (falls back to `KB_EMBED_API_KEY`) |
| `KB_CHAT_AUTH_HEADER` | `--chat-auth-header` | As per `KB_EMBED_AUTH_HEADER` |
| `KB_CHAT_MAX_TOKENS` | `--chat-max-tokens` | `max_tokens` for generation (default 1024; 0 omits the field) |
| `KB_CHAT_EXTRA_HEADERS` | _(env only)_ | JSON object of extra HTTP headers for the chat endpoint |
| `KB_CA_CERT` | `--ca-cert` | Path to a PEM CA bundle for an internal CA |
| `KB_VERIFY_SSL=false` | `--insecure` | Disable TLS certificate verification |
| `KB_TIMEOUT` | `--timeout` | HTTP timeout in seconds (default 120) |
| `KB_CHUNK_CHARS` | `--chunk-chars` | Soft max characters per chunk (default 1500) |
| `KB_CHUNK_OVERLAP` | `--chunk-overlap` | Overlap between adjacent chunks (default 200) |
| `KB_TOP_K` | `--top-k` | Default chunks retrieved (default 5) |
| — | `--check` | Validate config, call the endpoint(s) once, report index status, then exit |
| — | `--reindex` | Build/update the vector index, then exit (add `--force` to rebuild from scratch) |
| — | `--search QUERY` | Test retrieval from the command line, then exit |
| — | `--ask QUESTION` | Test full RAG (retrieve + generate) from the command line, then exit |
| — | `--version` | Print version and exit |

First run (before wiring into the MCP client): `--check`, then `--reindex`,
then `--search "some topic"` to confirm retrieval — the docstring at the top
of the file walks through it. If you change embedding model, run
`--reindex --force` once (vector dimensions differ between models).

```yaml
mcpServers:
  - name: kb-rag
    command: C:\path\to\python.exe
    args:
      - C:\path\to\knowledge-base-rag.py
      - --docs-dir
      - C:\Users\me\knowledge-base
      - --embed-url
      - https://ai-gateway.internal.example.com/v1/embeddings
      - --embed-model
      - text-embedding-3-small
    env:
      KB_EMBED_API_KEY: your-api-key
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
| `--output-dir` | Optional folder where `msword_create` writes **new** `.docx` files, kept **separate** from the knowledge-base folder. Falls back to the `MSWORD_OUTPUT_DIR` env var, then the `OUTPUT_DIR` config value, and finally to the document root. Also treated as a permitted open/save location so created documents can be reopened and edited |
| `--kb-dir` | Optional. If set, **every document opened** with `msword_open` is *also* written out as a Markdown file into this folder for a local RAG knowledge base (like `confluence.py`'s `--kb-dir`). Falls back to the `MSWORD_KB_DIR` env var, then the `KB_DIR` config value. Files are named `Word - <name>.md` and overwritten each open; the folder is created if missing. Omit to disable mirroring |

**Building a RAG knowledge base.** Point `--kb-dir` at the same folder your
knowledge-base servers (`knowledge-base.py` / `knowledge-base-semantic.py`)
index. Each time a `.docx` is opened, a Markdown copy is dropped there —
headings become `#`/`##`, `List Bullet`/`List Number` paragraphs become `-`/`1.`
lists, and tables become GitHub-style pipe tables — so Word content lands
alongside the Confluence pages in the same RAG index.

**Creating documents.** `msword_create` makes a new blank `.docx` in the
`--output-dir` folder (falling back to the document root) and opens it as a
session; build it up with `msword_add_heading` / `msword_add_paragraph` /
`msword_add_table` and persist with `msword_save` (omit its `path` to save in
place). Any directory part in the requested filename is stripped, so new files
always land inside the output folder.

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
      - --output-dir
      - C:\Users\me\Documents\ai_generated
      - --kb-dir
      - C:\Users\me\Documents\rag_kb
    env:
      PYTHONUTF8: "1"
```

`--output-dir` and `--kb-dir` are optional — drop those four lines to keep new
documents in the document root and disable Markdown mirroring.

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

### jira.py

1. "What's assigned to me right now, highest priority first?" → `jira_my_issues`
2. "Find any tickets mentioning the login timeout bug — has anyone reported this before?" → `jira_search`
3. "Show me ABC-123 in full, including the comments and who changed its status." → `jira_get_issue` with `include_changelog=true`
4. "Everything resolved in project ABC in the last week, for the release notes." → `jira_search_jql` with `project = ABC AND resolved >= -7d`
5. "How healthy is project ABC — what's open, in progress, unassigned?" → `jira_project_status`
6. "Which projects can I see in Jira?" → `jira_list_projects`
7. "Draft a status report from my open tickets as a Word doc with tracked changes." → `jira_my_issues` + `ms-word.py`'s editing tools

### knowledge-base.py

1. "What reference documents do we have available?" → `reference_list`
2. "Find any reference material about our procurement policy." → `reference_search`
3. "Read the full expense-reporting policy document and tell me the approval limits." → `reference_get`
4. "Search our reference docs for anything about onboarding, then read whichever one covers IT equipment." → `reference_search` followed by `reference_get`

### knowledge-base-rag.py

1. "Using my knowledge base, can I extend my work trip by 2 days, pay for my own accommodation for the weekend, and fly back Monday?" → `kb_ask` (retrieves the travel policy's trip-extension chunks and generates a cited answer — or hands the agent the chunks to answer from, if no chat endpoint is configured)
2. "Find the parts of our policies about accommodation and per diem." → `kb_retrieve`
3. "I've added some new documents to the knowledge base folder — pick them up." → `kb_index` (incremental: only new/changed files are embedded)
4. "Is the knowledge base index up to date? What's actually indexed?" → `kb_status`
5. "Rebuild the whole knowledge base index from scratch (we switched embedding model)." → `kb_index` with `force=true`

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
2. "Open every .docx in my docs folder so it gets mirrored into the RAG knowledge base as Markdown." → `msword_open` with `--kb-dir` set (each open writes `Word - <name>.md` next to the Confluence pages for `knowledge-base.py` / `knowledge-base-semantic.py` to index)
3. "Create a new status report document and draft it with a title, headings and a summary table, then save it to my generated-docs folder." → `msword_create` (writes to `--output-dir`) + `msword_add_heading` + `msword_add_paragraph` + `msword_add_table` + `msword_save`
5. "Find every mention of 'Acme Corp' in the contract and replace it with 'Acme Corporation'." → `msword_search` + `msword_replace_text`
6. "Add a 'Next Steps' heading and a summary paragraph to the end of the report, then save it." → `msword_add_heading` + `msword_add_paragraph` + `msword_save`
7. "Pull out the data from every table in the document as structured rows." → `msword_get_tables`
8. "Add a 3x4 pricing table to the end of the quote document with these values, using the 'Table Grid' style." → `msword_add_table` + `msword_save`
9. "Change 'DRAFT' to 'FINAL' throughout the report as a tracked change so it shows up as a Word revision for review." → `msword_replace_text` with `track_changes=true`
10. "Rewrite the third paragraph to be more concise, showing your edits as tracked changes — only mark the words you actually changed." → `msword_set_paragraph_text` with `track_changes=true` (old vs new text is diffed word-by-word, like editing in Word with Track Changes on)
11. "Add a new paragraph after the introduction as a tracked insertion, so reviewers can reject it if they disagree." → `msword_insert_paragraph` with `track_changes=true`
12. "Delete the whole limitation-of-liability paragraph as a tracked change — struck out, so legal can accept or reject it." → `msword_delete_paragraph` with `track_changes=true`
13. "What tracked changes are currently in this document, and who made them?" → `msword_list_changes`
14. "Accept Jane's two changes in the pricing section but leave everything else pending." → `msword_list_changes` + `msword_accept_changes` with those change ids
15. "Reject just the change that deleted the warranty sentence." → `msword_list_changes` + `msword_reject_changes` with that change id
16. "Accept all the tracked changes in this document now that legal has signed off." → `msword_accept_all_changes`
17. "Reject all the tracked changes and revert this document to its original wording." → `msword_reject_all_changes`

### pdf-to-md.py

1. "Convert every PDF in the reference folder to Markdown." → `convert_all_pdfs`
2. "Convert just the 'procurement policy' PDF to Markdown." → `convert_pdf_to_markdown`
3. "Reconvert all PDFs to Markdown even though some already have a .md file, since the source PDFs changed." → `convert_all_pdfs` with `force=true`
4. "Convert all our compliance PDFs (including those in sub-folders) to Markdown so the knowledge-base server can search them." → `convert_all_pdfs` (with `--recursive` set at startup) feeding into `knowledge-base.py`'s `reference_search`
