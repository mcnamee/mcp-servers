---
name: knowledge-base
description: Search and read the local reference-document folder via the knowledge-base MCP server (keyword search). Use when the user asks what reference docs exist, or to find/read policy or reference material by topic or name.
---

# Local knowledge base (via the `reference` MCP server)

Requires the `knowledge-base.py` MCP server (read-only keyword search over a
folder of `.md`/`.markdown`/`.txt` files). If its tools are not available,
tell the user to wire it in first (see the repo README) and to verify with
`python knowledge-base.py --docs-dir <folder> --check`.

## Tools

| Tool | Use for |
|---|---|
| `reference_list` | List every available document (+ title) |
| `reference_search` | Ranked keyword search across all documents, with snippets |
| `reference_get` | Read ONE document in full, by loose name or topic |

## Workflow

1. "What do we have?" → `reference_list`.
2. Topic questions → `reference_search` with 2–4 keywords; scan the ranked
   snippets, then `reference_get` the best match and answer from the full
   text. Cite the document name.
3. "Read the X policy" → `reference_get` directly (it resolves loose names).

## Notes

- Keyword search only — for semantic ("meaning-based") retrieval use the
  `kb-rag` server's skill instead, if it is configured.
- Content arrives in this folder from the other servers' `--kb-dir`
  mirroring (Confluence pages, Word docs, Outlook emails, converted PDFs);
  if something is missing, suggest reading/converting it once via the
  relevant server so it lands here.
