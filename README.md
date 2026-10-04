# ms-project-mcp — MCP server for Microsoft Project

**English** · [Español](README.es.md)

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue)
![Platform: Windows](https://img.shields.io/badge/platform-Windows-0078D6)
![MCP](https://img.shields.io/badge/MCP-server-8A2BE2)

A [Model Context Protocol](https://modelcontextprotocol.io) (MCP) server that lets **Claude** and other AI agents drive **Microsoft Project desktop** (MS Project 2016 / 2019 / 2021 / Microsoft 365) on Windows. Ask in plain language and the AI builds, edits and analyzes real `.mpp` schedules: Gantt charts, work breakdown structures (WBS), dependencies, resources, calendars, baselines, critical path, resource leveling and earned value.

It automates your installed Project through COM. Changes appear live in the Project window, and each AI edit is a single **Undo** step.

> *"Build a 3-month schedule for a small wooden house, with crews, hourly rates and Chilean public holidays."*
> → 43 tasks across 7 phases, 7 resources, a 42-hour work week, holidays and a critical path with no overallocated resources, saved as `.mpp`. See [`examples/wooden_house.py`](examples/wooden_house.py).

## Features

| Area | Tools |
|---|---|
| Files | `project_status`, `open_project`, `new_project`, `activate_project`, `save_project` (.mpp / .xml), `close_project` |
| Project | `set_project_properties` (start/finish, status date, calendar, hours per day…), `get_project_summary` (dates, work, cost, % complete, baseline, **earned value** SPI/CPI) |
| Tasks | `list_tasks` (filters: critical, late, milestones, overallocated…), `get_task`, `add_task`, **`bulk_create_tasks`** (whole WBS in one call), `update_task`, `bulk_update_tasks` (progress tracking), `delete_task`, `indent_task` |
| Dependencies | `link_tasks` (FS/SS/FF/SF + lag), `link_chain`, `unlink_tasks` |
| Resources | `list_resources`, `add_resource` (work / material / cost), `update_resource`, `delete_resource` |
| Assignments | `assign_resource`, `unassign_resource`, `list_assignments`, `level_resources` |
| Calendars | `list_calendars`, `add_calendar_exceptions` (holidays), `set_work_week` (shifts) |
| Analysis | `get_critical_path`, **`check_schedule_health`** (DCMA-style checks), `set_baseline` / `clear_baseline` (0–10), `get_baseline_variance`, `get_timephased_data` (S-curves, resource histograms) |
| Output | `list_views`, `apply_view`, `export_pdf`, `export_tasks_csv`, `rename_custom_field`, `undo`, `recalculate` |

**Compact input syntax**, which is easy for an LLM to write:

- Durations: `"5d"`, `"3h"`, `"2w"`, `"1mo"`. Dates: `"2026-11-02"` or `"2026-11-02 08:00"`.
- Dependencies: `"3"`, `"5SS+2d"`, `"#2FF-1d"` (`#n` = n-th task of the same batch). The Spanish codes FC/CC/CF are accepted too.
- Assignments: `"Ana"`, `"Carpenters[200%]"`. Missing resources are created automatically.
- Tasks can be referenced by ID, by `"uid:<UniqueID>"` or by exact name.
- Any Project field by name: `{"Text1": "Subcontractor", "Number2": 5}`.

## Requirements

- Windows with **Microsoft Project desktop** installed and activated. Tested with Project Professional 2019 (16.0), 64-bit.
- Python 3.10+.
- An MCP client: Claude Desktop, Claude Code, or any other MCP-compatible client.

Project Online / Project for the web are **not** supported, because this server uses the desktop COM API.

## Installation

```bash
git clone https://github.com/tomascarvallo/ms-project-mcp.git
cd ms-project-mcp
python -m venv .venv
.venv\Scripts\python -m pip install -e .
```

### Claude Desktop

Add the server to `%APPDATA%\Claude\claude_desktop_config.json`, using your own clone path:

```json
{
  "mcpServers": {
    "ms-project": {
      "command": "C:\\path\\to\\ms-project-mcp\\.venv\\Scripts\\python.exe",
      "args": ["-m", "ms_project_mcp"],
      "env": { "PYTHONIOENCODING": "utf-8" }
    }
  }
}
```

### Claude Code

```bash
claude mcp add ms-project --scope user -- C:\path\to\ms-project-mcp\.venv\Scripts\python.exe -m ms_project_mcp
```

Restart the client. The server starts automatically with Claude. If Project is not running, it is launched on the first request. If Project is already open, the server attaches to it and works on the **active project**.

> **Note:** Microsoft Store Python cannot see `%APPDATA%\Roaming`. Edit the Claude Desktop config with another tool, not with a script run by that Python.

## Example prompts

- *"Open `C:\Projects\Warehouse.mpp` and tell me which tasks are on the critical path."*
- *"Create a software release plan starting Monday: design 2w, development 6w, QA 2w overlapping development by 1 week."*
- *"Add Christmas and New Year as holidays and tell me how much the finish date moves."*
- *"Save a baseline, mark tasks 4–9 as 100% complete and show me the variance."*
- *"Run a schedule health check and fix the tasks without successors."*
- *"Give me the weekly work of each resource for an S-curve and export the Gantt chart to PDF."*

## How it works

- All COM calls run on a single STA thread. Calls are retried automatically while Project is busy.
- Every editing tool is wrapped in an undo transaction (`OpenUndoTransaction`), so **Ctrl+Z in Project reverts a whole AI operation**.
- Durations are converted using the project's hours per day and week. Dates without a time use the project's default start or finish time.
- Errors are returned as readable messages to the model, for example *"There are 2 tasks named 'Design'; use their Id: [3, 12]"*.
- Tool descriptions are written in Spanish. Claude and other LLMs use them in any language.

## Limitations

- Windows only, because it uses COM.
- If Project has a modal dialog open (for example Task Information), calls wait about 10 seconds and then fail with a message asking to close it.
- Opening non-native formats other than `.xml` may show Project's import wizard.

## Testing

```bash
.venv\Scripts\python tests\smoke_test.py
```

The test creates a temporary project, exercises all 43 tools and closes it without saving. The 2 reported errors are intentional negative checks.

## Keywords

Microsoft Project MCP, MS Project AI, MS Project automation Python, Claude Microsoft Project, AI project scheduling, Gantt chart AI, construction scheduling, critical path method, earned value, pywin32 COM, Model Context Protocol server.

## License

[MIT](LICENSE). This is not affiliated with or endorsed by Microsoft. *Microsoft Project* is a trademark of Microsoft Corporation.
