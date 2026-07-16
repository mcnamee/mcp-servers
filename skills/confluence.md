---
name: confluence
description: Search and read Confluence pages via the confluence MCP server. Use when the user asks to find, read, summarise or pull content from Confluence (runbooks, handbooks, wiki pages, spaces), or to mirror Confluence pages into the local knowledge base.
---

# Confluence (via the `confluence` MCP server)

Requires the `confluence.py` MCP server (read-only, Confluence Data Center).
If its tools are not available, tell the user to wire it in first (see the
repo README) and to verify connectivity with `python confluence.py --check`.

## Tools

| Tool | Use for |
|---|---|
| `confluence_search` | Free-text search for pages by topic |
| `confluence_search_cql` | Advanced search with raw CQL (spaces, dates, labels) |
| `confluence_get_page` | Full content of one page by numeric ID |
| `confluence_get_page_by_title` | Full content by exact title + space key |
| `confluence_list_pages_under` | Children of a page (navigate a page tree) |

## Workflow

1. Start with `confluence_search` using 2–4 topic keywords. Prefer fewer,
   more distinctive words over full sentences.
2. If the user names a space, date range or label, use `confluence_search_cql`
   instead, e.g. `space = DOCS AND text ~ "release notes" AND lastmodified >= now("-30d")`.
3. Fetch the winning result with `confluence_get_page` (by the ID from the
   search results) and answer from the page body. Quote the page title and ID
   so the user can find it.
4. For "everything under X" requests, walk `confluence_list_pages_under`.

## Notes

- The server is read-only; it cannot create or edit pages.
- If `--kb-dir` / `CONFLUENCE_KB_DIR` is configured, every page you read is
  automatically mirrored to Markdown for the local knowledge-base server —
  reading a page IS how you import it into the RAG index.
- Long pages may be truncated in the returned text if `--max-body` is set;
  say so if an answer might sit past the truncation point.
