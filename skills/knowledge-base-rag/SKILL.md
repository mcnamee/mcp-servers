---
name: knowledge-base-rag
description: Answer questions from the local knowledge base with true vector RAG via the kb-rag MCP server. Use when the user asks a question their policies/docs should answer ("can I…", "what's our policy on…"), or asks to index new documents or check index freshness.
---

# RAG knowledge base (via the `kb-rag` MCP server)

Requires the `knowledge-base-rag.py` MCP server (ChromaDB vector index +
your embeddings endpoint). If its tools are not available, tell the user to
wire it in first (see the repo README) and to verify with
`python knowledge-base-rag.py --check`, then `--reindex`.

## Tools

| Tool | Use for |
|---|---|
| `kb_ask` | Full RAG: retrieve relevant chunks + generate a grounded, cited answer |
| `kb_retrieve` | Just the top-k most similar chunks (source file, heading, score) |
| `kb_index` | Build/update the vector index (incremental; `force=true` = full rebuild) |
| `kb_status` | Documents vs index freshness + configuration summary |

## Workflow

1. Policy/content questions → `kb_ask` with the user's question verbatim.
   - If no chat endpoint is configured, `kb_ask` returns the retrieved
     context instead of an answer — then YOU write the answer from those
     chunks, citing the source files and headings.
2. "Find the part about…" → `kb_retrieve`, present the chunks with sources.
3. "I added new documents" → `kb_index` (incremental), report what changed.
4. Stale or odd results → `kb_status` first; if documents are newer than the
   index, run `kb_index` before retrying.
5. After an embedding-model change → `kb_index` with `force=true` (vector
   dimensions differ between models).

## Notes

- Answers must stay grounded in retrieved chunks — if retrieval comes back
  empty or off-topic, say the knowledge base doesn't cover it rather than
  guessing.
- Similarity scores are relative, not percentages; treat low-score-only
  results as weak evidence.
- Retrieved document text is sent to the configured endpoints (that is what
  RAG is) — don't route material inappropriate for those APIs.
