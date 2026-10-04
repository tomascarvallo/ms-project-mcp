"""Smoke test / prueba de humo.

Launches the server over stdio and exercises every tool on a temporary project that is
closed without saving. Requires Microsoft Project installed. Run: python tests/smoke_test.py
"""

import asyncio
import json
import os
import sys
import tempfile

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


async def main():
    params = StdioServerParameters(command=sys.executable, args=["-m", "ms_project_mcp"], cwd=ROOT)
    async with stdio_client(params) as (r, w):
        async with ClientSession(r, w) as s:
            await s.initialize()
            tools = await s.list_tools()
            print("TOOLS", len(tools.tools), [t.name for t in tools.tools])

            failures = 0

            async def call(_tool, **args):
                nonlocal failures
                res = await s.call_tool(_tool, args)
                text = res.content[0].text if res.content else ""
                flag = "ERR " if res.is_error else "ok  "
                if res.is_error:
                    failures += 1
                print(f"{flag}{_tool}: {text[:300]}")
                try:
                    return json.loads(text)
                except Exception:
                    return text

            await call("new_project", start="2026-11-02", title="Prueba MCP")
            await call("set_work_week", days=["lun", "mar", "mié", "jue", "vie"], shifts=[["08:00", "13:00"], ["14:00", "18:00"]])
            await call("add_resource", name="Ana", standard_rate="20000/h", group="Ingeniería")
            await call("add_resource", name="Hormigón", type="material", material_label="m3", standard_rate=95000)
            await call("bulk_create_tasks", tasks=[
                {"name": "Diseño", "level": 1},
                {"name": "Levantamiento", "level": 2, "duration": "3d", "resources": ["Ana"]},
                {"name": "Planos", "level": 2, "duration": "5d", "predecessors": ["#2"], "resources": ["Ana[50%]", "Luis"]},
                {"name": "Construcción", "level": 1},
                {"name": "Fundaciones", "level": 2, "duration": "1w", "predecessors": ["#3FS+2d"], "resources": ["Luis"]},
                {"name": "Obra gruesa", "level": 2, "duration": "10d", "predecessors": ["#5SS+3d"], "resources": ["Luis", "Ana"]},
                {"name": "Entrega", "level": 1, "milestone": True, "predecessors": ["#6"]},
            ])
            await call("add_task", name="Permisos", parent="Diseño", duration="4d", predecessors=["2"], fields={"Text1": "Municipalidad", "deadline": "2026-11-05"})
            await call("link_tasks", predecessor="Permisos", successor="Fundaciones", type="FS")
            await call("assign_resource", task="Fundaciones", resource="Hormigón", units_percent=1200)
            await call("add_calendar_exceptions", exceptions=[{"name": "Feriado prueba", "start": "2026-11-10"}])
            await call("list_tasks")
            await call("get_task", task="Obra gruesa")
            await call("list_resources")
            await call("list_assignments", resource="Ana")
            await call("get_critical_path")
            await call("check_schedule_health")
            await call("set_baseline")
            await call("update_task", task="Levantamiento", fields={"percent_complete": 100, "duration": "4d"})
            await call("bulk_update_tasks", updates=[{"task": "Planos", "fields": {"percent_complete": 40}}])
            await call("get_baseline_variance")
            await call("set_project_properties", status_date="2026-11-12")
            await call("get_project_summary")
            await call("get_timephased_data", target="resource:Ana", data="work", unit="weeks")
            await call("get_timephased_data", target="project", data="cost", unit="weeks")
            await call("level_resources")
            await call("list_calendars")
            views = await call("list_views")
            await call("link_chain", tasks=["Levantamiento", "Planos"], type="SS")
            await call("update_resource", resource="Ana", fields={"max_units_percent": 80, "Text1": "Planta"})
            await call("unassign_resource", task="Obra gruesa", resource="Ana")
            await call("set_baseline", number=1, tasks=["Planos", "Fundaciones"])
            await call("get_baseline_variance", number=1, only_slipped=False)
            await call("clear_baseline", number=1)
            await call("apply_view", name=views["views"][1] if isinstance(views, dict) else "Calendario")
            await call("recalculate")
            await call("list_tasks", filter="critical", extra_fields=["Text1"])
            await call("get_task", task="no existe")
            await call("rename_custom_field", field="Text1", alias="Responsable externo")
            tmp = tempfile.gettempdir()
            await call("export_tasks_csv", path=os.path.join(tmp, "pj_test.csv"), extra_fields=["Text1"])
            await call("export_pdf", path=os.path.join(tmp, "pj_test.pdf"))
            await call("save_project", path=os.path.join(tmp, "pj_test.xml"))
            await call("unlink_tasks", predecessor="Permisos", successor="Fundaciones")
            await call("indent_task", task="Entrega", outdent=True)
            await call("delete_task", task="Permisos")
            await call("undo")
            await call("project_status")
            await call("close_project", save=False)
            await call("open_project", path=os.path.join(tmp, "pj_test.xml"))
            await call("get_project_summary")
            await call("close_project", save=False)
            # 2 errores son intencionales: tarea inexistente y quitar sangría a una tarea de nivel 1
            print("FAILURES", failures, "(expected 2)")


asyncio.run(main())
