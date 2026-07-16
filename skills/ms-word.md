---
name: ms-word
description: Read, edit and create Word .docx files via the msword MCP server, including Word tracked changes, editing tables (set cells, add/delete rows), and creating documents from a template — including filling out an example template (placeholders + example tables). Use when the user asks to open/summarise/search a Word document, edit or redline one with tracked changes, accept/reject revisions, fill in or edit tables, or draft a new .docx (optionally from a template).
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

### Filling out a template (placeholders + example tables)

When the template is an *example* to fill in (e.g. an agenda with placeholder
text and an example items table), after `msword_create(template=...)`:

1. **See what's there:** `msword_get_content` with `mode: "structured"` (shows
   each table's `table_index`, `rows`, `cols` in order) and `msword_get_tables`
   (shows every cell's text). Identify the placeholders, the repeating table,
   its header row and its example rows.
2. **Text placeholders:** `msword_replace_text` to swap them (works inside table
   cells too). Templates are easiest to fill when they use explicit
   `{{TOKEN}}` markers (e.g. `{{MEETING_DATE}}`, `{{CHAIR}}`), which make the
   `find` unambiguous. If the template just has example prose, use that example
   text as the `find` string — or, when example values repeat, set the cell by
   coordinate with `msword_set_cell` instead.
3. **Grow the table:** for each real item, `msword_add_table_row` with
   `table_index`, `copy_from_row` = a styled example data row (so borders,
   shading and fonts are inherited), and `values` = the row's cell texts. Or add
   the row and fill cells individually with `msword_set_cell`
   (`table_index`, `row`, `col`, `text`). New rows are always appended last.
4. **Remove leftover example rows:** `msword_delete_table_row`. **Row indices
   shift after each delete**, so delete from the highest index down (or re-read
   `msword_get_tables` between deletes). Never delete the header row unless you
   mean to; the only remaining row can't be deleted.
5. **Save:** `msword_save`.

These table edits are always plain (untracked) — table row/cell changes can't be
recorded as Word tracked changes. `values` maps to *cells* left-to-right, so a
row with horizontally merged cells has fewer cells than grid columns; passing
more values than the row has cells is an error.

## Notes

- All paths must be inside the configured docs folder (plus the output
  folder); requests outside it are refused — don't fight the sandbox.
- Opening a document with `--kb-dir` configured also mirrors it to Markdown
  for the knowledge base.
- Not supported: comments, tracked moves, formatting-only revisions,
  headers/footers editing.
