#!/usr/bin/env python3
"""
knowledge-base-rag.py
=====================

A single-file MCP (Model Context Protocol) server providing true RAG
(Retrieval-Augmented Generation) over a folder of your own markdown
documents. It replaces knowledge-base-semantic.py with real vector
retrieval:

  1. INDEX    : documents are split into heading-aware chunks, each chunk is
                embedded via your embeddings API, and the vectors are stored
                in a local ChromaDB database on disk (persistent, incremental
                - only new/changed files are re-embedded).
  2. RETRIEVE : a question is embedded with the same API and the most
                semantically similar chunks are returned, with their source
                file, section heading and similarity score.
  3. GENERATE : (optional) the retrieved chunks + the question are sent to a
                chat-completions API, which writes a grounded answer citing
                its sources. If no chat endpoint is configured, kb_ask
                returns the retrieved context and the agent you're already
                talking to writes the answer instead - so generation works
                either way.

Transport: newline-delimited JSON-RPC 2.0 over stdio (the standard MCP stdio
transport, and what the VSCode Continue extension speaks).

DEPENDENCY
----------
    pip install chromadb

Everything else is standard library (HTTP to your endpoints is done with
urllib - no requests/httpx needed). chromadb pulls in several packages,
some with compiled wheels; install through your pip proxy with the SAME
interpreter your MCP client launches:

    C:\\path\\to\\python.exe -m pip install chromadb

ChromaDB's anonymised telemetry is explicitly disabled in this file, so the
server makes no network calls other than to the two endpoints you configure.

CONFIGURATION
-------------
CLI flags take priority over environment variables. API keys are env-var
ONLY (command lines are visible to other local users in process listings).

| Env var                 | CLI flag             | Purpose                                                     |
|-------------------------|----------------------|-------------------------------------------------------------|
| KB_DOCS_DIR             | --docs-dir           | REQUIRED. Folder of .md/.markdown/.txt docs (recursive)     |
| KB_INDEX_DIR            | --index-dir          | ChromaDB folder (default: <docs-dir>/.kb-rag-index)         |
| KB_COLLECTION           | --collection         | ChromaDB collection name (default: kb-rag; 3-512 chars of   |
|                         |                      | [a-zA-Z0-9._-], starting/ending alphanumeric)               |
| KB_EMBED_URL            | --embed-url          | REQUIRED. Full URL of the embeddings endpoint               |
| KB_EMBED_MODEL          | --embed-model        | Model name sent in the request (omit if endpoint has one)   |
| KB_EMBED_API_KEY        | (env only)           | API key for the embeddings endpoint                         |
| KB_EMBED_AUTH_HEADER    | --embed-auth-header  | Header the key is sent in. Default Authorization (Bearer);  |
|                         |                      | any other name (e.g. api-key for Azure) sends the raw key   |
| KB_EMBED_STYLE          | --embed-style        | Request format: openai (default) or ollama                  |
| KB_EMBED_BATCH          | --embed-batch        | Texts per embeddings request (default 16; ollama is 1-by-1) |
| KB_EMBED_QUERY_PREFIX   | --embed-query-prefix | Prefix for query embeds (e5-style models: "query: ")        |
| KB_EMBED_DOC_PREFIX     | --embed-doc-prefix   | Prefix for document embeds ("passage: ")                    |
| KB_EMBED_EXTRA_HEADERS  | (env only)           | JSON object of extra HTTP headers for the embed endpoint    |
| KB_CHAT_URL             | --chat-url           | OPTIONAL. Chat-completions endpoint for the generate step   |
| KB_CHAT_MODEL           | --chat-model         | Model name for generation                                   |
| KB_CHAT_API_KEY         | (env only)           | API key for the chat endpoint (falls back to embed key)     |
| KB_CHAT_AUTH_HEADER     | --chat-auth-header   | As per KB_EMBED_AUTH_HEADER                                 |
| KB_CHAT_MAX_TOKENS      | --chat-max-tokens    | max_tokens for generation (default 1024, 0 = omit field)    |
| KB_CHAT_EXTRA_HEADERS   | (env only)           | JSON object of extra HTTP headers for the chat endpoint     |
| KB_CA_CERT              | --ca-cert            | PEM CA bundle for an internal CA                            |
| KB_CLIENT_CERT          | --client-cert        | PEM client certificate, for gateways that require mutual    |
|                         |                      | TLS (mTLS). Presented to BOTH endpoints                     |
| KB_CLIENT_KEY           | --client-key         | PEM private key for the client certificate (omit if the     |
|                         |                      | --client-cert file contains both cert and key)              |
| KB_CLIENT_KEY_PASSWORD  | (env only)           | Passphrase, if the client private key is encrypted          |
| KB_VERIFY_SSL=false     | --insecure           | Disable TLS certificate verification                        |
| KB_TIMEOUT              | --timeout            | HTTP timeout seconds (default 120)                          |
| KB_CHUNK_CHARS          | --chunk-chars        | Soft max characters per chunk (default 1500)                |
| KB_CHUNK_OVERLAP        | --chunk-overlap      | Overlap characters between adjacent chunks (default 200)    |
| KB_TOP_K                | --top-k              | Default number of chunks retrieved (default 5)              |

Endpoint formats ("where you have unknowns"):
  --embed-style openai  : POST {"input": [texts], "model": m}
                          reads response["data"][i]["embedding"]
  --embed-style ollama  : POST {"model": m, "prompt": text} (one per request)
                          reads response["embedding"]
  Response parsing also falls back to top-level "embedding"/"embeddings",
  so most bespoke internal endpoints work with style=openai unchanged.
  Generation POSTs OpenAI chat-completions JSON and falls back to Ollama
  /api/chat ("message"."content") and /api/generate ("response") shapes.

In Continue's config.yaml:

    mcpServers:
      - name: kb-rag
        command: C:\\path\\to\\python.exe
        args:
          - C:\\path\\to\\knowledge-base-rag.py
          - --docs-dir
          - C:\\Users\\me\\knowledge-base
          - --embed-url
          - https://ai-gateway.internal.example.com/v1/embeddings
          - --embed-model
          - text-embedding-3-small
        env:
          KB_EMBED_API_KEY: your-api-key
          PYTHONUTF8: "1"

FIRST RUN / TESTING (do this before wiring into the MCP client)
---------------------------------------------------------------
    set KB_EMBED_API_KEY=...                        (PowerShell: $env:KB_EMBED_API_KEY="...")
    python knowledge-base-rag.py --docs-dir C:\\kb --embed-url https://... --check
        -> validates config, calls the embeddings endpoint once, reports index status
    python knowledge-base-rag.py --docs-dir C:\\kb --embed-url https://... --reindex
        -> builds/updates the vector index (add --force to rebuild from scratch)
    python knowledge-base-rag.py --docs-dir C:\\kb --embed-url https://... --search "trip extension"
        -> test retrieval from the command line
    python knowledge-base-rag.py --docs-dir C:\\kb --embed-url https://... --chat-url https://... --ask "Can I extend my trip?"
        -> test full RAG (retrieve + generate) from the command line

TOOLS EXPOSED
-------------
- kb_index    : build/update the vector index (incremental; force=true rebuilds)
- kb_retrieve : semantic search - top-k most similar chunks for a question
- kb_ask      : full RAG - retrieve, then generate a grounded, cited answer
                (or return the context for the agent to answer from, if no
                chat endpoint is configured)
- kb_status   : index freshness + configuration summary (never shows keys)

NOTES
-----
- ALL diagnostic output goes to stderr; stdout carries only JSON-RPC.
- Set PYTHONUTF8=1 in the launching environment so non-ASCII content does not
  crash on the default Windows cp1252 codec.
- File access is confined to --docs-dir (paths are resolved, symlinks
  included, before the containment check). The index folder and dot-folders
  are never indexed. The only network calls are to the endpoints you set.
- Retrieved document text IS sent to the configured endpoints (that's what
  RAG is) - point the server only at material appropriate for those APIs.
- If you change embedding model, the vector dimensions change: run
  --reindex --force once to rebuild the index.
- Mutual TLS: a gateway error like CERTIFICATE_NOT_PROVIDED (or a TLS
  handshake failure / connection reset during --check) means the gateway
  requires a CLIENT certificate. Configure --client-cert/--client-key;
  --insecure cannot fix this, because it only disables YOUR verification of
  the server - it does not change the certificate you present to it. The
  client certificate/key is loaded once at startup, so a bad path or wrong
  passphrase fails immediately with a clear message.
"""

import os
import re
import ssl
import sys
import json
import time
import hashlib
import argparse
import traceback
import urllib.error
import urllib.request


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
# Configuration (populated in main(); flags take priority over env vars)
# ---------------------------------------------------------------------------

DOC_EXTENSIONS = {".md", ".markdown", ".txt"}

# Hard cap on how much retrieved context is stuffed into a generation prompt.
MAX_CONTEXT_CHARS = 16000


class Config(object):
    """All runtime settings in one place. Filled in by main()."""
    docs_dir = None            # absolute, real path of the documents folder
    index_dir = None           # absolute path of the ChromaDB folder
    collection = "kb-rag"
    embed_url = None
    embed_model = ""
    embed_key = ""
    embed_auth_header = "Authorization"
    embed_style = "openai"
    embed_batch = 16
    embed_query_prefix = ""
    embed_doc_prefix = ""
    embed_extra_headers = {}
    chat_url = ""
    chat_model = ""
    chat_key = ""
    chat_auth_header = "Authorization"
    chat_max_tokens = 1024
    chat_extra_headers = {}
    ca_cert = ""
    client_cert = ""
    client_key = ""
    client_key_password = ""
    verify_ssl = True
    timeout = 120
    chunk_chars = 1500
    chunk_overlap = 200
    top_k = 5


CFG = Config()


class RagError(Exception):
    """A failure with a message meant to be shown to the caller as-is."""


# ---------------------------------------------------------------------------
# Filesystem helpers (same confinement model as the other servers)
# ---------------------------------------------------------------------------

def is_within(path, base):
    """
    True if `path` resolves to a location inside `base`. Guards against path
    traversal and symlinks pointing outside the configured folder.
    """
    try:
        real_path = os.path.realpath(path)
        real_base = os.path.realpath(base)
        return os.path.commonpath([real_path, real_base]) == real_base
    except Exception:
        return False


def to_rel(path):
    """Relative path from docs_dir, using forward slashes for stable display."""
    return os.path.relpath(path, CFG.docs_dir).replace("\\", "/")


def scan_documents():
    """
    Return {relative_path: absolute_path} for every document in the folder.
    Dot-folders and the index folder are pruned; anything resolving outside
    the docs folder (e.g. a symlink) is excluded.
    """
    found = {}
    index_real = os.path.realpath(CFG.index_dir) if CFG.index_dir else None
    for root, dirs, files in os.walk(CFG.docs_dir):
        # Prune hidden folders and the vector index itself from the walk.
        dirs[:] = [
            d for d in dirs
            if not d.startswith(".")
            and os.path.realpath(os.path.join(root, d)) != index_real
        ]
        for filename in files:
            if filename.startswith("."):
                continue
            if os.path.splitext(filename)[1].lower() not in DOC_EXTENSIONS:
                continue
            full = os.path.join(root, filename)
            if not is_within(full, CFG.docs_dir):
                log("Excluded (resolves outside the docs folder): {0}".format(full))
                continue
            found[to_rel(full)] = full
    return found


def read_text(path):
    with open(path, "r", encoding="utf-8", errors="replace") as handle:
        return handle.read()


def file_hash(path):
    """SHA-256 of a file's bytes - used to detect changed documents."""
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for block in iter(lambda: handle.read(65536), b""):
            digest.update(block)
    return digest.hexdigest()


# ---------------------------------------------------------------------------
# HTTP plumbing (stdlib only)
# ---------------------------------------------------------------------------

def build_ssl_context():
    if CFG.ca_cert and CFG.verify_ssl:
        context = ssl.create_default_context(cafile=CFG.ca_cert)
    else:
        context = ssl.create_default_context()
    if not CFG.verify_ssl:
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE
    if CFG.client_cert:
        # Mutual TLS: present our certificate to the gateway. Independent of
        # verify_ssl, which only controls how we verify the SERVER.
        try:
            context.load_cert_chain(
                certfile=CFG.client_cert,
                keyfile=CFG.client_key or None,
                password=CFG.client_key_password or None,
            )
        except (ssl.SSLError, OSError) as exc:
            raise RagError(
                "Could not load the mutual-TLS client certificate/key "
                "({0} / {1}): {2}. If the key is encrypted, set "
                "KB_CLIENT_KEY_PASSWORD.".format(
                    CFG.client_cert,
                    CFG.client_key or "key expected in the cert file",
                    exc))
    return context


def auth_headers(key, header_name):
    """Authorization -> 'Bearer <key>'; any other header carries the raw key."""
    if not key:
        return {}
    if header_name.strip().lower() == "authorization":
        return {"Authorization": "Bearer " + key}
    return {header_name.strip(): key}


def _mtls_hint(error):
    """Append advice when a TLS failure looks like a missing client certificate."""
    text = str(error)
    if ("CERTIFICATE_REQUIRED" in text or "certificate required" in text.lower()
            or "CERTIFICATE_NOT_PROVIDED" in text or "handshake failure" in text.lower()):
        return (" This looks like the endpoint requires a CLIENT certificate "
                "(mutual TLS): configure --client-cert / --client-key. "
                "--insecure cannot fix this - it only disables verification "
                "of the server, not the certificate you present.")
    return ""


def http_post_json(url, payload, headers):
    """POST JSON, return the decoded JSON response. Raises RagError on failure."""
    body = json.dumps(payload).encode("utf-8")
    request = urllib.request.Request(url, data=body, method="POST")
    request.add_header("Content-Type", "application/json")
    request.add_header("Accept", "application/json")
    for name, value in headers.items():
        request.add_header(name, value)
    try:
        with urllib.request.urlopen(request, timeout=CFG.timeout,
                                    context=build_ssl_context()) as response:
            return json.loads(response.read().decode("utf-8", errors="replace"))
    except urllib.error.HTTPError as exc:
        detail = ""
        try:
            detail = exc.read().decode("utf-8", errors="replace")[:500]
        except Exception:
            pass
        raise RagError("HTTP {0} from {1}: {2}{3}".format(
            exc.code, url, detail or exc.reason, _mtls_hint(detail or exc.reason)))
    except urllib.error.URLError as exc:
        raise RagError("Could not reach {0}: {1}{2}".format(
            url, exc.reason, _mtls_hint(exc.reason)))
    except OSError as exc:
        # TLS alerts can surface on the first read (TLS 1.3), outside
        # urllib's URLError wrapping - e.g. a gateway rejecting the handshake
        # because no client certificate was presented.
        raise RagError("Connection to {0} failed: {1}{2}".format(
            url, exc, _mtls_hint(exc)))
    except json.JSONDecodeError:
        raise RagError("Non-JSON response from {0}.".format(url))


# ---------------------------------------------------------------------------
# Embeddings client
# ---------------------------------------------------------------------------

def _parse_embedding_response(data, expected):
    """
    Pull `expected` vectors out of an embeddings response, tolerating the
    common shapes: OpenAI {"data":[{"embedding":[...]}]}, bare {"embedding":
    [...]}, and {"embeddings":[[...]]}.
    """
    vectors = None
    if isinstance(data, dict):
        if isinstance(data.get("data"), list):
            items = sorted(data["data"], key=lambda item: item.get("index", 0))
            vectors = [item.get("embedding") for item in items]
        elif isinstance(data.get("embeddings"), list):
            vectors = data["embeddings"]
        elif isinstance(data.get("embedding"), list):
            vectors = [data["embedding"]]
    if (not vectors or len(vectors) != expected
            or any(not isinstance(v, list) or not v for v in vectors)):
        raise RagError(
            "Unexpected embeddings response shape (expected {0} vector(s)). "
            "Check --embed-style / the endpoint URL. Response keys: {1}".format(
                expected, list(data.keys()) if isinstance(data, dict) else type(data).__name__))
    return vectors


def embed_texts(texts, is_query=False):
    """
    Embed a list of strings via the configured endpoint; returns one vector
    per input, in order. Batching applies to the openai style; the ollama
    style is one request per text (its classic API takes a single prompt).
    """
    prefix = CFG.embed_query_prefix if is_query else CFG.embed_doc_prefix
    inputs = [prefix + text for text in texts]
    headers = dict(CFG.embed_extra_headers)
    headers.update(auth_headers(CFG.embed_key, CFG.embed_auth_header))

    vectors = []
    if CFG.embed_style == "ollama":
        for text in inputs:
            payload = {"prompt": text}
            if CFG.embed_model:
                payload["model"] = CFG.embed_model
            data = http_post_json(CFG.embed_url, payload, headers)
            vectors.extend(_parse_embedding_response(data, 1))
    else:  # openai-compatible (the default)
        for start in range(0, len(inputs), CFG.embed_batch):
            batch = inputs[start:start + CFG.embed_batch]
            payload = {"input": batch}
            if CFG.embed_model:
                payload["model"] = CFG.embed_model
            data = http_post_json(CFG.embed_url, payload, headers)
            vectors.extend(_parse_embedding_response(data, len(batch)))
    return vectors


# ---------------------------------------------------------------------------
# Generation client (OpenAI chat-completions shape, with Ollama fallbacks)
# ---------------------------------------------------------------------------

GENERATION_SYSTEM_PROMPT = (
    "You are a careful assistant answering questions from a personal knowledge "
    "base. Answer ONLY from the provided context. Quote or paraphrase the "
    "relevant passages and cite the source file for each claim, e.g. "
    "[travel-policy.md]. If the context does not contain the answer, say so "
    "plainly - do not invent one."
)


def generate_answer(question, context):
    """Send the retrieved context + question to the chat endpoint; return the answer text."""
    headers = dict(CFG.chat_extra_headers)
    headers.update(auth_headers(CFG.chat_key, CFG.chat_auth_header))
    payload = {
        "messages": [
            {"role": "system", "content": GENERATION_SYSTEM_PROMPT},
            {"role": "user",
             "content": "Context:\n\n{0}\n\nQuestion: {1}".format(context, question)},
        ],
        "temperature": 0.1,
        "stream": False,
    }
    if CFG.chat_model:
        payload["model"] = CFG.chat_model
    if CFG.chat_max_tokens > 0:
        payload["max_tokens"] = CFG.chat_max_tokens

    data = http_post_json(CFG.chat_url, payload, headers)

    # OpenAI shape, then Ollama /api/chat, then Ollama /api/generate.
    try:
        return data["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError):
        pass
    try:
        return data["message"]["content"]
    except (KeyError, TypeError):
        pass
    if isinstance(data.get("response"), str):
        return data["response"]
    raise RagError(
        "Unexpected chat response shape. Response keys: {0}".format(
            list(data.keys()) if isinstance(data, dict) else type(data).__name__))


# ---------------------------------------------------------------------------
# Markdown chunking (heading-aware, code-fence-aware)
# ---------------------------------------------------------------------------

# ATX heading, allowing up to 3 leading spaces and trailing '#'s (per CommonMark).
HEADING_RE = re.compile(r"^ {0,3}(#{1,6})\s+(.+?)\s*#*\s*$")


def parse_headings(content):
    """
    Return [{level, text, line}, ...] for every ATX heading, in document
    order. Lines inside fenced code blocks are ignored so a '#' comment in a
    code sample is not mistaken for a heading.
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
            heads.append({"level": len(match.group(1)),
                          "text": match.group(2).strip(), "line": idx})
    return heads


def split_text(text, size, overlap):
    """
    Split `text` into pieces of at most ~`size` characters, preferring
    paragraph boundaries, with ~`overlap` characters carried between adjacent
    pieces so a sentence cut at a boundary is still retrievable. `size` is a
    soft target: a carried tail can push a piece slightly over it.
    """
    if len(text) <= size:
        return [text]

    # Break into paragraph units, hard-splitting any single oversized paragraph.
    units = []
    for para in re.split(r"\n\s*\n", text):
        para = para.strip()
        if not para:
            continue
        while len(para) > size:
            units.append(para[:size])
            para = para[size - overlap:]
        if para:
            units.append(para)

    pieces = []
    current = ""
    for unit in units:
        if current and len(current) + 2 + len(unit) > size:
            pieces.append(current)
            tail = current[-overlap:].strip() if overlap else ""
            current = (tail + "\n\n" + unit) if tail else unit
        else:
            current = (current + "\n\n" + unit) if current else unit
    if current:
        pieces.append(current)
    return pieces


def chunk_document(content):
    """
    Split a markdown document into chunks along its heading structure:
    each heading's section becomes one or more chunks, each tagged with the
    full heading path (e.g. 'Travel Policy > Expenses > Per diem') so the
    embedding and the retrieved result both carry that context.

    Returns [{"heading": path, "text": chunk_text}, ...].
    """
    heads = parse_headings(content)
    lines = content.splitlines()

    sections = []  # (heading_path, start_line, end_line)
    if heads:
        if heads[0]["line"] > 0:
            sections.append(("", 0, heads[0]["line"]))
        stack = []  # [(level, text), ...] - the open headings above this point
        for i, head in enumerate(heads):
            while stack and stack[-1][0] >= head["level"]:
                stack.pop()
            stack.append((head["level"], head["text"]))
            end = heads[i + 1]["line"] if i + 1 < len(heads) else len(lines)
            sections.append((" > ".join(t for _lvl, t in stack), head["line"], end))
    else:
        sections.append(("", 0, len(lines)))

    chunks = []
    for heading_path, start, end in sections:
        section_text = "\n".join(lines[start:end]).strip()
        if not section_text:
            continue
        for piece in split_text(section_text, CFG.chunk_chars, CFG.chunk_overlap):
            chunks.append({"heading": heading_path, "text": piece})
    return chunks


# ---------------------------------------------------------------------------
# Vector index (ChromaDB)
# ---------------------------------------------------------------------------

_CHROMA_CLIENT = None


def chroma_client():
    """Create (once) and return the persistent ChromaDB client, telemetry off."""
    global _CHROMA_CLIENT
    if _CHROMA_CLIENT is not None:
        return _CHROMA_CLIENT
    try:
        import chromadb
        from chromadb.config import Settings
    except ImportError:
        raise RagError(
            "chromadb is not installed in this interpreter ({0}). "
            "Install it with:  {0} -m pip install chromadb".format(sys.executable))
    _CHROMA_CLIENT = chromadb.PersistentClient(
        path=CFG.index_dir,
        settings=Settings(anonymized_telemetry=False),
    )
    return _CHROMA_CLIENT


def get_collection(reset=False):
    """Return the vector collection (cosine space), optionally dropping it first."""
    client = chroma_client()
    if reset:
        try:
            client.delete_collection(CFG.collection)
        except Exception:
            pass  # nothing to drop on first run
    return client.get_or_create_collection(
        name=CFG.collection, metadata={"hnsw:space": "cosine"})


def indexed_file_hashes(collection):
    """Return {source_relpath: file_hash} for everything currently indexed."""
    hashes = {}
    offset = 0
    while True:
        page = collection.get(include=["metadatas"], limit=1000, offset=offset)
        metadatas = page.get("metadatas") or []
        if not metadatas:
            break
        for meta in metadatas:
            if meta and "source" in meta:
                hashes[meta["source"]] = meta.get("file_hash", "")
        if len(metadatas) < 1000:
            break
        offset += len(metadatas)
    return hashes


def index_sync(force=False):
    """
    Bring the vector index in line with the docs folder. Only new or changed
    files are re-embedded; files deleted from the folder are removed from the
    index. force=True drops the collection and rebuilds everything (needed
    after changing embedding model, since vector dimensions change).

    Returns a stats dict.
    """
    started = time.time()
    collection = get_collection(reset=force)
    documents = scan_documents()
    indexed = {} if force else indexed_file_hashes(collection)

    removed = [rel for rel in indexed if rel not in documents]
    for rel in removed:
        collection.delete(where={"source": rel})

    unchanged = 0
    updated = []
    errors = []
    total_chunks = 0
    for rel, path in sorted(documents.items()):
        try:
            digest = file_hash(path)
        except OSError as exc:
            errors.append("{0}: {1}".format(rel, exc))
            continue
        if indexed.get(rel) == digest:
            unchanged += 1
            continue

        try:
            content = read_text(path)
        except OSError as exc:
            errors.append("{0}: {1}".format(rel, exc))
            continue
        chunks = chunk_document(content)
        if rel in indexed:
            collection.delete(where={"source": rel})
        if not chunks:
            updated.append(rel)
            continue

        # Embed with the heading path prepended for context; store the raw
        # chunk text so what the model reads back is the document itself.
        embed_inputs = [
            ("{0} — {1}\n\n{2}".format(rel, c["heading"], c["text"])
             if c["heading"] else "{0}\n\n{1}".format(rel, c["text"]))
            for c in chunks
        ]
        try:
            vectors = embed_texts(embed_inputs, is_query=False)
        except RagError as exc:
            # Surface an embedding-dimension mismatch as the fix, not a mystery.
            raise RagError(
                "Embedding failed while indexing '{0}': {1}".format(rel, exc))

        try:
            collection.add(
                ids=["{0}#{1}".format(rel, i) for i in range(len(chunks))],
                embeddings=vectors,
                documents=[c["text"] for c in chunks],
                metadatas=[{
                    "source": rel,
                    "heading": c["heading"],
                    "chunk": i,
                    "file_hash": digest,
                } for i, c in enumerate(chunks)],
            )
        except Exception as exc:
            raise RagError(
                "Storing vectors for '{0}' failed: {1}. If you changed "
                "embedding model, rebuild with kb_index(force=true) or "
                "--reindex --force (vector dimensions differ between models)."
                .format(rel, exc))
        updated.append(rel)
        total_chunks += len(chunks)
        log("Indexed {0} ({1} chunks)".format(rel, len(chunks)))

    return {
        "documents": len(documents),
        "updated": updated,
        "unchanged": unchanged,
        "removed": removed,
        "new_chunks": total_chunks,
        "index_chunks": collection.count(),
        "errors": errors,
        "seconds": round(time.time() - started, 1),
    }


def retrieve(query, top_k):
    """
    Embed `query` and return the top_k most similar chunks:
    [{"source", "heading", "text", "similarity"}, ...] best first.
    """
    collection = get_collection()
    count = collection.count()
    if count == 0:
        raise RagError(
            "The vector index is empty. Run kb_index first (or launch once "
            "with --reindex).")
    vector = embed_texts([query], is_query=True)[0]
    result = collection.query(
        query_embeddings=[vector],
        n_results=min(top_k, count),
        include=["documents", "metadatas", "distances"],
    )
    hits = []
    for text, meta, distance in zip(result["documents"][0],
                                    result["metadatas"][0],
                                    result["distances"][0]):
        hits.append({
            "source": meta.get("source", "?"),
            "heading": meta.get("heading", ""),
            "text": text,
            # cosine distance in [0,2] -> similarity in [-1,1]
            "similarity": round(1.0 - distance, 3),
        })
    return hits


def format_hits(hits):
    """Human/agent-readable rendering of retrieved chunks, best first."""
    blocks = []
    for rank, hit in enumerate(hits, 1):
        where = hit["source"] + (" · " + hit["heading"] if hit["heading"] else "")
        blocks.append("[{0}] {1}  (similarity {2})\n{3}".format(
            rank, where, hit["similarity"], hit["text"]))
    return "\n\n---\n\n".join(blocks)


def build_context(hits):
    """Concatenate hits into a generation context, capped at MAX_CONTEXT_CHARS."""
    parts = []
    total = 0
    for hit in hits:
        where = hit["source"] + (" · " + hit["heading"] if hit["heading"] else "")
        block = "[Source: {0}]\n{1}".format(where, hit["text"])
        if parts and total + len(block) > MAX_CONTEXT_CHARS:
            break
        parts.append(block)
        total += len(block)
    return "\n\n".join(parts), len(parts)


# ---------------------------------------------------------------------------
# Tool implementations (each returns a human-readable text string)
# ---------------------------------------------------------------------------

def _arg_int(args, name, default):
    try:
        return int(args.get(name, default))
    except (TypeError, ValueError):
        return default


def tool_kb_index(args):
    force = bool(args.get("force", False))
    stats = index_sync(force=force)
    lines = [
        "Index {0} in {1}s.".format("rebuilt" if force else "updated", stats["seconds"]),
        "- documents in folder : {0}".format(stats["documents"]),
        "- re-indexed          : {0}".format(len(stats["updated"])),
        "- unchanged (skipped) : {0}".format(stats["unchanged"]),
        "- removed from index  : {0}".format(len(stats["removed"])),
        "- chunks in index     : {0}".format(stats["index_chunks"]),
    ]
    if stats["updated"]:
        lines.append("Re-indexed files:\n" + "\n".join(
            "  - " + rel for rel in stats["updated"][:30]))
        if len(stats["updated"]) > 30:
            lines.append("  ... and {0} more".format(len(stats["updated"]) - 30))
    if stats["errors"]:
        lines.append("Unreadable files (skipped):\n" + "\n".join(
            "  - " + err for err in stats["errors"]))
    return "\n".join(lines)


def tool_kb_retrieve(args):
    query = (args.get("query") or "").strip()
    if not query:
        return "Error: 'query' is required."
    top_k = max(1, min(20, _arg_int(args, "top_k", CFG.top_k)))
    hits = retrieve(query, top_k)
    return (
        "Top {0} chunk(s) for '{1}', most similar first. Cite the source file "
        "when using them.\n\n{2}".format(len(hits), query, format_hits(hits)))


def tool_kb_ask(args):
    question = (args.get("question") or "").strip()
    if not question:
        return "Error: 'question' is required."
    top_k = max(1, min(20, _arg_int(args, "top_k", CFG.top_k)))
    hits = retrieve(question, top_k)
    context, used = build_context(hits)

    if not CFG.chat_url:
        # No generation endpoint: hand the agent the context to answer from.
        return (
            "No generation endpoint is configured (KB_CHAT_URL / --chat-url), "
            "so answer the question yourself STRICTLY from the retrieved "
            "context below, citing source files. If the context doesn't "
            "contain the answer, say so.\n\nQuestion: {0}\n\n{1}".format(
                question, format_hits(hits)))

    answer = generate_answer(question, context)
    sources = []
    for hit in hits[:used]:
        entry = hit["source"] + (" · " + hit["heading"] if hit["heading"] else "")
        if entry not in sources:
            sources.append(entry)
    return "{0}\n\nSources ({1} chunk(s) retrieved):\n{2}".format(
        answer, used, "\n".join("- " + s for s in sources))


def tool_kb_status(_args):
    documents = scan_documents()
    lines = [
        "Knowledge base folder : {0}".format(CFG.docs_dir),
        "Vector index folder   : {0}".format(CFG.index_dir),
        "Documents in folder   : {0}".format(len(documents)),
    ]
    try:
        collection = get_collection()
        indexed = indexed_file_hashes(collection)
        stale = [rel for rel, path in documents.items()
                 if indexed.get(rel) != file_hash(path)]
        removed = [rel for rel in indexed if rel not in documents]
        lines.append("Files indexed         : {0}".format(len(indexed)))
        lines.append("Chunks in index       : {0}".format(collection.count()))
        if stale or removed:
            lines.append(
                "STALE: {0} file(s) new/changed, {1} deleted - run kb_index "
                "to refresh.".format(len(stale), len(removed)))
        else:
            lines.append("Index is up to date with the folder.")
    except Exception as exc:
        lines.append("Index status          : unavailable ({0})".format(exc))
    lines.append("Embeddings endpoint   : {0} (style: {1}, model: {2}, key: {3})".format(
        CFG.embed_url, CFG.embed_style, CFG.embed_model or "(none)",
        "set" if CFG.embed_key else "NOT SET"))
    lines.append("Mutual TLS (client)   : {0}".format(
        "cert: {0}, key: {1}, passphrase: {2}".format(
            CFG.client_cert, CFG.client_key or "(in cert file)",
            "set" if CFG.client_key_password else "not set")
        if CFG.client_cert else "(none)"))
    lines.append("Generation endpoint   : {0}".format(
        "{0} (model: {1}, key: {2})".format(
            CFG.chat_url, CFG.chat_model or "(none)",
            "set" if CFG.chat_key else "NOT SET")
        if CFG.chat_url else "(none - kb_ask returns context for the agent to answer from)"))
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# MCP tool registry
# ---------------------------------------------------------------------------

TOOLS = [
    {
        "name": "kb_index",
        "description": (
            "Build or update the vector index of the knowledge base. Incremental: "
            "only new or changed files are re-embedded, and deleted files are "
            "removed. Run this after documents change, or when kb_status reports "
            "the index is stale. Set force=true to rebuild from scratch (required "
            "after changing embedding model)."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "force": {
                    "type": "boolean",
                    "description": "Drop the index and re-embed everything (default false).",
                },
            },
        },
    },
    {
        "name": "kb_retrieve",
        "description": (
            "Semantic (vector) search: returns the chunks of the knowledge base "
            "most similar in MEANING to the query, with source file, section "
            "heading and similarity score. Use this to pull the relevant policy/"
            "notes passages for a question, then answer from them, citing the "
            "source files."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "query": {
                    "type": "string",
                    "description": "The question or topic, in natural language.",
                },
                "top_k": {
                    "type": "integer",
                    "description": "Number of chunks to return, 1-20 (default {0}).".format(CFG.top_k),
                },
            },
            "required": ["query"],
        },
    },
    {
        "name": "kb_ask",
        "description": (
            "Full RAG in one call: retrieves the most relevant chunks for the "
            "question and generates a grounded answer that cites its source files. "
            "If no generation endpoint is configured, it returns the retrieved "
            "context and YOU write the answer from it."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "question": {
                    "type": "string",
                    "description": "The question to answer from the knowledge base.",
                },
                "top_k": {
                    "type": "integer",
                    "description": "Number of chunks to retrieve, 1-20 (default {0}).".format(CFG.top_k),
                },
            },
            "required": ["question"],
        },
    },
    {
        "name": "kb_status",
        "description": (
            "Report the state of the knowledge base: documents in the folder, "
            "files/chunks in the vector index, whether the index is stale, and "
            "which endpoints are configured (never shows keys). Use to diagnose "
            "empty or odd retrieval results."
        ),
        "inputSchema": {"type": "object", "properties": {}},
    },
]

TOOL_DISPATCH = {
    "kb_index": tool_kb_index,
    "kb_retrieve": tool_kb_retrieve,
    "kb_ask": tool_kb_ask,
    "kb_status": tool_kb_status,
}


# ---------------------------------------------------------------------------
# JSON-RPC / MCP plumbing
# ---------------------------------------------------------------------------

PROTOCOL_VERSION_DEFAULT = "2024-11-05"
SERVER_INFO = {"name": "knowledge-base-rag", "version": "1.1.0"}

SERVER_INSTRUCTIONS = (
    "This server is a RAG pipeline over a personal knowledge base of markdown "
    "documents. For a question, call kb_retrieve (or kb_ask for a generated, "
    "cited answer) and ground your reply in the returned chunks, citing source "
    "files. If retrieval reports an empty or stale index, call kb_index first "
    "(it embeds only new/changed files). kb_status shows index freshness and "
    "configuration. Document access is read-only."
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
        except RagError as exc:
            return rpc_result(req_id, text_content("Knowledge-base error: {0}".format(exc), is_error=True))
        except Exception as exc:
            log("Tool '{0}' failed:\n{1}".format(name, traceback.format_exc()))
            return rpc_result(req_id, text_content("Knowledge-base tool error: {0}".format(exc), is_error=True))

    if is_notification:
        return None
    return rpc_error(req_id, -32601, "Method not found: {0}".format(method))


def run_server():
    """Main stdio loop: read newline-delimited JSON-RPC, dispatch, respond."""
    log("knowledge-base-rag server started (stdio). Docs: {0}  Index: {1}".format(
        CFG.docs_dir, CFG.index_dir))
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
        log("knowledge-base-rag server stopped.")


# ---------------------------------------------------------------------------
# CLI modes (--check / --reindex / --search / --ask)
# ---------------------------------------------------------------------------

def run_check():
    """Validate config, test the endpoints, report index status. Exit code 0/1."""
    ok = True
    log(tool_kb_status({}))
    try:
        vector = embed_texts(["connectivity test"], is_query=True)[0]
        log("Embeddings endpoint   : OK ({0} dimensions)".format(len(vector)))
    except RagError as exc:
        log("Embeddings endpoint   : FAILED - {0}".format(exc))
        ok = False
    if CFG.chat_url:
        try:
            reply = generate_answer("Reply with the single word: ok",
                                    "(connectivity test - no context)")
            log("Generation endpoint   : OK (replied: {0})".format(reply.strip()[:60]))
        except RagError as exc:
            log("Generation endpoint   : FAILED - {0}".format(exc))
            ok = False
    log("CHECK OK" if ok else "CHECK FAILED")
    return 0 if ok else 1


def run_reindex(force):
    try:
        log(tool_kb_index({"force": force}))
    except RagError as exc:
        log("REINDEX FAILED: {0}".format(exc))
        return 1
    except Exception:
        log("REINDEX FAILED:\n{0}".format(traceback.format_exc()))
        return 1
    return 0


def run_query(mode, text):
    """CLI retrieval/RAG test: mode is 'search' or 'ask'."""
    try:
        if mode == "search":
            log(tool_kb_retrieve({"query": text}))
        else:
            log(tool_kb_ask({"question": text}))
    except RagError as exc:
        log("FAILED: {0}".format(exc))
        return 1
    except Exception:
        log("FAILED:\n{0}".format(traceback.format_exc()))
        return 1
    return 0


def env_flag_false(name):
    """True if env var `name` is an explicit 'off' value (false/0/no)."""
    return os.environ.get(name, "").strip().lower() in {"false", "0", "no"}


def parse_extra_headers(env_name):
    raw = os.environ.get(env_name, "").strip()
    if not raw:
        return {}
    try:
        headers = json.loads(raw)
        if not isinstance(headers, dict):
            raise ValueError("not a JSON object")
        return {str(k): str(v) for k, v in headers.items()}
    except ValueError as exc:
        log("FATAL: {0} is not a JSON object of headers ({1}).".format(env_name, exc))
        sys.exit(2)


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Single-file MCP server providing RAG (index / retrieve / generate) over "
            "a folder of local markdown documents, using a ChromaDB vector index and "
            "your own embeddings (and optionally chat) API endpoints. With no mode "
            "flag it runs as an stdio MCP server. See the docstring for full "
            "configuration; API keys are env-var only (KB_EMBED_API_KEY, "
            "KB_CHAT_API_KEY)."
        )
    )
    env = os.environ.get
    parser.add_argument("--docs-dir", default=env("KB_DOCS_DIR"),
                        help="Folder of markdown documents (env: KB_DOCS_DIR). Required.")
    parser.add_argument("--index-dir", default=env("KB_INDEX_DIR"),
                        help="ChromaDB folder (env: KB_INDEX_DIR; default: <docs-dir>/.kb-rag-index).")
    parser.add_argument("--collection", default=env("KB_COLLECTION", "kb-rag"),
                        help="ChromaDB collection name (env: KB_COLLECTION; default: kb-rag).")
    parser.add_argument("--embed-url", default=env("KB_EMBED_URL"),
                        help="Embeddings endpoint URL (env: KB_EMBED_URL). Required.")
    parser.add_argument("--embed-model", default=env("KB_EMBED_MODEL", ""),
                        help="Embedding model name sent in requests (env: KB_EMBED_MODEL).")
    parser.add_argument("--embed-auth-header", default=env("KB_EMBED_AUTH_HEADER", "Authorization"),
                        help="Header for the embed API key; 'Authorization' sends 'Bearer <key>', "
                             "anything else (e.g. 'api-key') sends the raw key (env: KB_EMBED_AUTH_HEADER).")
    parser.add_argument("--embed-style", choices=["openai", "ollama"],
                        default=env("KB_EMBED_STYLE", "openai"),
                        help="Embeddings request format (env: KB_EMBED_STYLE; default: openai).")
    parser.add_argument("--embed-batch", type=int, default=int(env("KB_EMBED_BATCH", "16")),
                        help="Texts per embeddings request, openai style (env: KB_EMBED_BATCH; default: 16).")
    parser.add_argument("--embed-query-prefix", default=env("KB_EMBED_QUERY_PREFIX", ""),
                        help="Prefix prepended to query text before embedding, for models "
                             "that require it, e.g. 'query: ' (env: KB_EMBED_QUERY_PREFIX).")
    parser.add_argument("--embed-doc-prefix", default=env("KB_EMBED_DOC_PREFIX", ""),
                        help="Prefix prepended to document chunks before embedding, "
                             "e.g. 'passage: ' (env: KB_EMBED_DOC_PREFIX).")
    parser.add_argument("--chat-url", default=env("KB_CHAT_URL", ""),
                        help="Optional chat-completions endpoint for the generate step "
                             "(env: KB_CHAT_URL). Unset: kb_ask returns context for the agent.")
    parser.add_argument("--chat-model", default=env("KB_CHAT_MODEL", ""),
                        help="Generation model name (env: KB_CHAT_MODEL).")
    parser.add_argument("--chat-auth-header", default=env("KB_CHAT_AUTH_HEADER", "Authorization"),
                        help="Header for the chat API key, as per --embed-auth-header "
                             "(env: KB_CHAT_AUTH_HEADER).")
    parser.add_argument("--chat-max-tokens", type=int, default=int(env("KB_CHAT_MAX_TOKENS", "1024")),
                        help="max_tokens for generation; 0 omits the field (env: KB_CHAT_MAX_TOKENS; default: 1024).")
    parser.add_argument("--ca-cert", default=env("KB_CA_CERT", ""),
                        help="Path to a PEM CA bundle for an internal CA (env: KB_CA_CERT).")
    parser.add_argument("--client-cert", default=env("KB_CLIENT_CERT", ""),
                        help="Path to a PEM client certificate, for gateways requiring "
                             "mutual TLS (env: KB_CLIENT_CERT). Presented to both endpoints.")
    parser.add_argument("--client-key", default=env("KB_CLIENT_KEY", ""),
                        help="Path to the PEM private key for --client-cert; omit if the "
                             "cert file contains both (env: KB_CLIENT_KEY). An encrypted "
                             "key's passphrase goes in KB_CLIENT_KEY_PASSWORD (env only).")
    parser.add_argument("--insecure", action="store_true",
                        default=env_flag_false("KB_VERIFY_SSL"),
                        help="Disable TLS certificate verification (env: KB_VERIFY_SSL=false).")
    parser.add_argument("--timeout", type=int, default=int(env("KB_TIMEOUT", "120")),
                        help="HTTP timeout in seconds (env: KB_TIMEOUT; default: 120).")
    parser.add_argument("--chunk-chars", type=int, default=int(env("KB_CHUNK_CHARS", "1500")),
                        help="Soft max characters per chunk (env: KB_CHUNK_CHARS; default: 1500).")
    parser.add_argument("--chunk-overlap", type=int, default=int(env("KB_CHUNK_OVERLAP", "200")),
                        help="Overlap characters between adjacent chunks (env: KB_CHUNK_OVERLAP; default: 200).")
    parser.add_argument("--top-k", type=int, default=int(env("KB_TOP_K", "5")),
                        help="Default number of chunks retrieved (env: KB_TOP_K; default: 5).")
    parser.add_argument("--check", action="store_true",
                        help="Validate config, test the endpoint(s), report index status, then exit.")
    parser.add_argument("--reindex", action="store_true",
                        help="Build/update the vector index, then exit (no server).")
    parser.add_argument("--force", action="store_true",
                        help="With --reindex: drop the index and re-embed everything.")
    parser.add_argument("--search", metavar="QUERY",
                        help="Test retrieval from the CLI: print the top chunks for QUERY, then exit.")
    parser.add_argument("--ask", metavar="QUESTION",
                        help="Test full RAG from the CLI: retrieve + generate for QUESTION, then exit.")
    parser.add_argument("--version", action="version",
                        version="knowledge-base-rag {0}".format(SERVER_INFO["version"]))
    args = parser.parse_args()

    if not args.docs_dir:
        log("FATAL: no knowledge-base folder set. Pass --docs-dir or set KB_DOCS_DIR.")
        sys.exit(2)
    if not os.path.isdir(args.docs_dir):
        log("FATAL: knowledge-base folder does not exist or is not a directory: {0}".format(args.docs_dir))
        sys.exit(2)
    if not args.embed_url:
        log("FATAL: no embeddings endpoint set. Pass --embed-url or set KB_EMBED_URL.")
        sys.exit(2)
    if not re.match(r"^[a-zA-Z0-9][a-zA-Z0-9._-]{1,510}[a-zA-Z0-9]$", args.collection):
        log("FATAL: --collection must be 3-512 characters of [a-zA-Z0-9._-], "
            "starting and ending alphanumeric: {0}".format(args.collection))
        sys.exit(2)
    if args.chunk_overlap >= args.chunk_chars:
        log("FATAL: --chunk-overlap must be smaller than --chunk-chars.")
        sys.exit(2)
    if args.ca_cert and not os.path.isfile(args.ca_cert):
        log("FATAL: --ca-cert file not found: {0}".format(args.ca_cert))
        sys.exit(2)
    if args.client_key and not args.client_cert:
        log("FATAL: --client-key was given without --client-cert.")
        sys.exit(2)
    if args.client_cert and not os.path.isfile(args.client_cert):
        log("FATAL: --client-cert file not found: {0}".format(args.client_cert))
        sys.exit(2)
    if args.client_key and not os.path.isfile(args.client_key):
        log("FATAL: --client-key file not found: {0}".format(args.client_key))
        sys.exit(2)

    CFG.docs_dir = os.path.realpath(args.docs_dir)
    CFG.index_dir = os.path.realpath(
        args.index_dir or os.path.join(CFG.docs_dir, ".kb-rag-index"))
    CFG.collection = args.collection
    CFG.embed_url = args.embed_url
    CFG.embed_model = args.embed_model
    CFG.embed_key = os.environ.get("KB_EMBED_API_KEY", "")
    CFG.embed_auth_header = args.embed_auth_header
    CFG.embed_style = args.embed_style
    CFG.embed_batch = max(1, args.embed_batch)
    CFG.embed_query_prefix = args.embed_query_prefix
    CFG.embed_doc_prefix = args.embed_doc_prefix
    CFG.embed_extra_headers = parse_extra_headers("KB_EMBED_EXTRA_HEADERS")
    CFG.chat_url = args.chat_url
    CFG.chat_model = args.chat_model
    CFG.chat_key = os.environ.get("KB_CHAT_API_KEY", "") or CFG.embed_key
    CFG.chat_auth_header = args.chat_auth_header
    CFG.chat_max_tokens = args.chat_max_tokens
    CFG.chat_extra_headers = parse_extra_headers("KB_CHAT_EXTRA_HEADERS")
    CFG.ca_cert = args.ca_cert
    CFG.client_cert = args.client_cert
    CFG.client_key = args.client_key
    CFG.client_key_password = os.environ.get("KB_CLIENT_KEY_PASSWORD", "")
    CFG.verify_ssl = not args.insecure
    CFG.timeout = max(1, args.timeout)
    CFG.chunk_chars = max(200, args.chunk_chars)
    CFG.chunk_overlap = max(0, args.chunk_overlap)
    CFG.top_k = max(1, min(20, args.top_k))

    if not CFG.verify_ssl:
        log("WARNING: TLS certificate verification is DISABLED.")

    # Fail fast on an unloadable client cert/key (bad file, wrong passphrase)
    # rather than surfacing it on the first tool call.
    if CFG.client_cert:
        try:
            build_ssl_context()
        except RagError as exc:
            log("FATAL: {0}".format(exc))
            sys.exit(2)

    if args.check:
        sys.exit(run_check())
    if args.reindex:
        sys.exit(run_reindex(args.force))
    if args.search:
        sys.exit(run_query("search", args.search))
    if args.ask:
        sys.exit(run_query("ask", args.ask))
    run_server()


if __name__ == "__main__":
    main()
