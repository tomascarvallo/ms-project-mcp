"""Example: build a small wooden-house construction schedule through the MCP server.

Ejemplo: crea la programación de una casa de madera de 60 m2 (Chile) usando el servidor MCP:
jornada de 42 h, feriados, 7 recursos con tarifas en CLP, 43 tareas con EDT y dependencias.
Run / ejecutar: python examples/wooden_house.py
"""

import asyncio, json, os, sys
from pathlib import Path
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

ROOT = Path(__file__).resolve().parent.parent
OUT = str(Path.home() / "Documents" / "Casa de madera (ejemplo).mpp")

L = lambda name, **k: {"name": name, "level": 1, **k}
T = lambda name, dur, preds, res=None, **k: {"name": name, "level": 2, "duration": dur, "predecessors": preds, "resources": res or [], **k}
M = lambda name, preds, lvl=2, **k: {"name": name, "level": lvl, "milestone": True, "predecessors": preds, **k}

TASKS = [
    M("Inicio de obra", [], lvl=1),                                                          # 1
    L("Obras preliminares"),                                                                  # 2
    T("Revisión de permiso de edificación y planos", "3d", ["#1"], ["Maestro carpintero"]),   # 3
    T("Instalación de faena (bodega, empalme provisorio)", "2d", ["#3"], ["Jornales[200%]"]), # 4
    T("Limpieza y nivelación del terreno", "2d", ["#4"], ["Jornales[200%]"]),                 # 5
    T("Trazado y niveles", "1d", ["#5"], ["Maestro carpintero", "Jornales"]),                 # 6
    L("Fundaciones"),                                                                         # 7
    T("Excavación de poyos", "2d", ["#6"], ["Jornales[200%]"]),                               # 8
    T("Moldajes y enfierradura", "2d", ["#8"], ["Carpinteros", "Jornales"]),                  # 9
    T("Hormigonado de poyos", "1d", ["#9"], ["Maestro carpintero", "Jornales[200%]"]),        # 10
    M("Fundaciones listas (tras curado)", ["#10FS+5d"], notes="Incluye 5 días hábiles de curado del hormigón antes de cargar la estructura."),  # 11
    L("Estructura"),                                                                          # 12
    T("Envigado de piso", "3d", ["#11"], ["Maestro carpintero", "Carpinteros[200%]"]),        # 13
    T("Plataforma de piso (OSB)", "2d", ["#13"], ["Carpinteros[200%]", "Jornales"]),          # 14
    T("Tabiquería perimetral", "5d", ["#14"], ["Maestro carpintero", "Carpinteros"]),         # 15
    T("Tabiquería interior", "3d", ["#15SS+3d"], ["Carpinteros", "Jornales"]),                # 16
    T("Estructura de techumbre (cerchas)", "4d", ["#15", "#16"], ["Maestro carpintero", "Carpinteros[200%]"]),  # 17
    M("Estructura terminada", ["#17"]),                                                       # 18
    L("Envolvente"),                                                                          # 19
    T("Cubierta: OSB y fieltro", "2d", ["#18"], ["Carpinteros[200%]"]),                       # 20
    T("Cubierta: planchas de zinc", "3d", ["#20"], ["Carpinteros[200%]", "Jornales"]),        # 21
    T("Hojalatería y canaletas", "1d", ["#21"], ["Maestro carpintero"]),                      # 22
    T("Barrera de humedad y revestimiento exterior", "6d", ["#21"], ["Carpinteros[200%]", "Jornales"]),  # 23
    T("Ventanas y puerta exterior", "2d", ["#23SS+3d"], ["Maestro carpintero", "Jornales"]),  # 24
    M("Casa cerrada", ["#22", "#23", "#24"]),                                                 # 25
    L("Instalaciones"),                                                                       # 26
    T("Instalación eléctrica (canalización y cableado)", "4d", ["#18"], ["Electricista"]),    # 27
    T("Agua potable y alcantarillado", "4d", ["#18"], ["Gasfíter"]),                          # 28
    T("Pruebas e inspección de instalaciones", "1d", ["#27", "#28", "#25"], ["Electricista", "Gasfíter"]),  # 29
    L("Terminaciones"),                                                                       # 30
    T("Aislación térmica (lana de vidrio)", "2d", ["#29"], ["Jornales[200%]"]),               # 31
    T("Revestimiento interior (yeso-cartón)", "6d", ["#31"], ["Carpinteros[200%]"]),          # 32
    T("Huinchas, pasta y lijado", "4d", ["#32"], ["Pintor"]),                                 # 33
    T("Cerámica de baño y cocina", "3d", ["#33"], ["Ceramista"]),                             # 34
    T("Piso flotante", "3d", ["#34", "#36"], ["Carpinteros[200%]"]),                          # 35
    T("Pintura interior", "4d", ["#33"], ["Pintor"]),                                         # 36
    T("Puertas interiores y molduras", "3d", ["#35"], ["Carpinteros"]),                       # 37
    T("Muebles de cocina y artefactos sanitarios", "3d", ["#35"], ["Gasfíter", "Carpinteros"]),  # 38
    T("Artefactos eléctricos y luminarias", "2d", ["#36"], ["Electricista"]),                 # 39
    T("Protección y pintura exterior", "4d", ["#25"], ["Pintor"]),                            # 40
    L("Entrega"),                                                                             # 41
    T("Aseo final", "1d", ["#37", "#38", "#39", "#40"], ["Jornales[200%]"]),                  # 42
    M("Recepción final y entrega de llaves", ["#42"]),                                        # 43
]

RESOURCES = [
    dict(name="Maestro carpintero", standard_rate="10000/h", group="Carpintería", initials="MC"),
    dict(name="Carpinteros", standard_rate="8000/h", group="Carpintería", initials="CA", max_units_percent=200),
    dict(name="Jornales", standard_rate="5500/h", group="Obra", initials="JO", max_units_percent=200),
    dict(name="Electricista", standard_rate="10000/h", group="Especialidades", initials="EL"),
    dict(name="Gasfíter", standard_rate="10000/h", group="Especialidades", initials="GA"),
    dict(name="Pintor", standard_rate="7000/h", group="Terminaciones", initials="PI"),
    dict(name="Ceramista", standard_rate="8000/h", group="Terminaciones", initials="CE"),
]


async def main():
    params = StdioServerParameters(command=sys.executable, args=["-m", "ms_project_mcp"], cwd=str(ROOT),
                                    env={**os.environ, "PYTHONIOENCODING": "utf-8"})
    async with stdio_client(params) as (r, w):
        async with ClientSession(r, w) as s:
            await s.initialize()

            async def call(_tool, show=True, **args):
                res = await s.call_tool(_tool, args)
                text = res.content[0].text if res.content else ""
                if res.is_error:
                    print(f"ERR {_tool}: {text}")
                    sys.exit(1)
                data = json.loads(text)
                if show:
                    print(f"ok  {_tool}")
                return data

            await call("new_project", start="2026-11-02 08:00", title="Construcción casa de madera 60 m²")
            await call("set_project_properties", hours_per_day=8.4, hours_per_week=42, show_summary_task=True)
            await call("set_work_week", days=["lun", "mar", "mié", "jue"], shifts=[["08:00", "13:00"], ["14:00", "17:30"]])
            await call("set_work_week", days=["vie"], shifts=[["08:00", "13:00"], ["14:00", "17:00"]])
            await call("add_calendar_exceptions", exceptions=[
                {"name": "Inmaculada Concepción", "start": "2026-12-08"},
                {"name": "Navidad", "start": "2026-12-25"},
                {"name": "Año Nuevo", "start": "2027-01-01"},
            ])
            for r in RESOURCES:
                await call("add_resource", show=False, **r)
            print("ok  recursos", len(RESOURCES))
            await call("bulk_create_tasks", tasks=TASKS)

            health = await call("check_schedule_health")
            print("sobreasignados:", health["overallocated_resources"])
            if health["overallocated_resources"]:
                lv = await call("level_resources")
                print("tras redistribuir:", lv["overallocated_after"])
            await call("save_project", path=OUT)

            summ = await call("get_project_summary", show=False)
            crit = await call("get_critical_path", show=False)
            health = await call("check_schedule_health", show=False)
            tasks = await call("list_tasks", show=False)
            res = await call("list_resources", show=False)
            print(json.dumps({
                "summary": {k: summ[k] for k in ("start", "finish", "duration_days", "work_hours", "cost")},
                "counts": summ["counts"],
                "critical": [t["name"] for t in crit["tasks"]],
                "health": {k: v["count"] for k, v in health["summary"].items() if v["count"]},
                "phases": [(t["wbs"], t["name"], t["start"][:10], t["finish"][:10], t["duration_days"]) for t in tasks["tasks"] if t["level"] == 1],
                "resources": [(x["name"], x["work_hours"], x["cost"], x["overallocated"]) for x in res["resources"]],
            }, ensure_ascii=False, indent=1))


asyncio.run(main())
