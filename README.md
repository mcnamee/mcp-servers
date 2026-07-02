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
| `knowledge-base.py` | _none_ | standard library only |
| `ms-excel.py` | _none_ | standard library only (parses .xlsx as a zip of XML) |
| `ms-word.py` | `pip install python-docx` | also pulls in `lxml` (compiled) and `typing_extensions` |
| `ms-outlook.py` | `pip install pywin32` | Windows only (COM automation of classic Outlook) |
| `pdf-to-md.py` | `pip install pymupdf pymupdf4llm` | OCR of scanned PDFs additionally requires Tesseract installed on the machine (not a pip package) |

## Install everything at once

```
pip install python-docx pymupdf pymupdf4llm pywin32
```

(Drop `pywin32` if you're not on Windows / not using `ms-outlook.py`.)

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
