---
name: jira
description: Query Jira issues via the jira MCP server. Use when the user asks about their tickets, sprint/project status, issue details or history, or wants a report drafted from Jira data.
---

# Jira (via the `jira` MCP server)

Requires the `jira.py` MCP server (STRICTLY read-only, Jira Data Center v2
REST API). If its tools are not available, tell the user to wire it in first
(see the repo README) and to verify with `python jira.py --check`.

## Tools

| Tool | Use for |
|---|---|
| `jira_my_issues` | "What's assigned to me?" — the user's open issues |
| `jira_search` | Free-text search (safely quoted into JQL) |
| `jira_search_jql` | Advanced search with raw JQL |
| `jira_get_issue` | One issue in full (set `include_changelog=true` for history) |
| `jira_project_status` | Health summary of one project (counts by status, unassigned, top open) |
| `jira_list_projects` | Which project keys are visible |

## Workflow

1. "My work" questions → `jira_my_issues`, sort/summarise by priority.
2. "Has anyone seen…" / topic questions → `jira_search` with distinctive
   keywords from the problem description.
3. Time-boxed or precise questions → `jira_search_jql`, e.g.
   `project = ABC AND resolved >= -7d ORDER BY resolved DESC`.
4. Deep-dive on one ticket → `jira_get_issue` (add the changelog only when
   the user asks who changed what).
5. Reports: combine `jira_my_issues`/`jira_search_jql` results, then (if the
   ms-word server is available) draft the report as a .docx with tracked
   changes.

## Notes

- Read-only by design: the server cannot create, edit, transition or comment
  on issues — never promise to update Jira.
- A `JIRA_PROJECTS` allowlist may be configured; issues outside it are
  refused and other projects hidden. If a key is refused, say why.
- Jira Cloud is not supported (v3 rich-text API); this targets Data Center.
