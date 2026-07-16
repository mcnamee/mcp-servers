---
name: ms-word
description: Read, edit and create Word .docx files via the msword MCP server, including Word tracked changes and creating documents from a template. Use when the user asks to open/summarise/search a Word document, edit or redline one with tracked changes, accept/reject revisions, or draft a new .docx (optionally from a template).
---

# Word (via the `msword-py` MCP server)

Requires the `ms-word.py` MCP server (the only write-capable server in the
suite; sandboxed to its configured docs folder). If its tools are not
available, tell the user to wire it in first (see the repo README) and to
verify with `python ms-word.py --check`.

## Core workflow: open → edit → save

1. Find the file: `msword_list_documents` (optional `query` substring) if
   the exact name is unknown. Bare names resolve against the docs folder —
   `msword_open` with `path: "Policy 103.docx"` works without a full path.
2. `msword_open` returns a `session_id`; every later call needs it.
3. Read with `msword_get_content` / `msword_search` / `msword_get_tables`.
4. Edit with `msword_replace_text`, `msword_set_paragraph_text`,
   `msword_insert_paragraph`, `msword_delete_paragraph`,
   `msword_add_heading` / `msword_add_paragraph` / `msword_add_table`.
5. Persist with `msword_save` (omit `path` to save in place).

## Tracked changes (redlining)

- Pass `track_changes=true` on the editing tools to record edits as real
  Word revisions (word-by-word diff, author-stamped) that the user can
  Accept/Reject in Word. Prefer this whenever the user says "redline",
  "tracked changes", "for review", or edits someone else's document.
- Review flow: `msword_list_changes` → `msword_accept_changes` /
  `msword_reject_changes` (by id), or `msword_accept_all_changes` /
  `msword_reject_all_changes`.
- While changes are pending, content/search show the final "No Markup" view.

## Creating documents

`msword_create` makes a new .docx (written to the configured output folder)
and opens it as a session — build it with the add_* tools, then
`msword_save`. Directory parts in the requested filename are stripped.

### From a template

To start from a defined template (a letterhead, report layout, contract
boilerplate, etc.), pass `template` to `msword_create`:

- `msword_create` with `filename: "Q3 Report.docx"` and
  `template: "Report Template.docx"`. The template is an existing `.docx` in
  the docs folder, resolved the same forgiving way as `msword_open` (bare name,
  relative path, or a fuzzy near-miss). Its styles, headers/footers, page setup
  and boilerplate are inherited into the new file, which is written to the
  output folder. The template file itself is never modified, and the result
  includes `template` (the resolved template) so you can confirm the right one
  was used.
- Then edit and `msword_save` as usual. Omit `title` when the template already
  carries its own title.
- To discover available templates, use `msword_list_documents` (e.g.
  `query: "template"`); keep templates in the docs folder (a `templates/`
  subfolder is a tidy convention). Only `.docx` templates are supported.

## Notes

- All paths must be inside the configured docs folder (plus the output
  folder); requests outside it are refused — don't fight the sandbox.
- Opening a document with `--kb-dir` configured also mirrors it to Markdown
  for the knowledge base.
- Not supported: comments, tracked moves, formatting-only revisions,
  headers/footers editing.
