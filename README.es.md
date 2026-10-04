# ms-project-mcp — Servidor MCP para Microsoft Project

[English](README.md) · **Español**

[![Licencia: MIT](https://img.shields.io/badge/Licencia-MIT-blue.svg)](LICENSE)
![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue)
![Plataforma: Windows](https://img.shields.io/badge/plataforma-Windows-0078D6)
![MCP](https://img.shields.io/badge/MCP-servidor-8A2BE2)

Un servidor [Model Context Protocol](https://modelcontextprotocol.io) (MCP) que permite que **Claude** y otros agentes de inteligencia artificial (IA) manejen **Microsoft Project de escritorio** (MS Project 2016 / 2019 / 2021 / Microsoft 365) en Windows. Pides en lenguaje natural y la IA crea, edita y analiza cronogramas `.mpp` reales: carta Gantt, estructura de desglose del trabajo (EDT), dependencias, recursos, calendarios, líneas base, ruta crítica, redistribución de recursos y valor ganado.

Automatiza tu Project instalado mediante COM. Los cambios se ven en vivo en la ventana de Project, y cada edición de la IA es un solo paso de **Deshacer**.

> *"Crea la programación de la construcción de una casa pequeña de madera."*
> → 43 tareas en 7 fases, 7 recursos con tarifas, jornada de 42 horas, feriados de Chile y ruta crítica sin sobreasignaciones, guardado como `.mpp`. Revisa [`examples/wooden_house.py`](examples/wooden_house.py).

## Funcionalidades

| Área | Herramientas |
|---|---|
| Archivos | `project_status`, `open_project`, `new_project`, `activate_project`, `save_project` (.mpp / .xml), `close_project` |
| Proyecto | `set_project_properties` (inicio/fin, fecha de estado, calendario, horas por día…), `get_project_summary` (fechas, trabajo, costo, % completado, línea base, **valor ganado** SPI/CPI) |
| Tareas | `list_tasks` (filtros: críticas, atrasadas, hitos, sobreasignadas…), `get_task`, `add_task`, **`bulk_create_tasks`** (una EDT completa en una llamada), `update_task`, `bulk_update_tasks` (registro de avance), `delete_task`, `indent_task` |
| Dependencias | `link_tasks` (FC/CC/FF/CF + posposición), `link_chain`, `unlink_tasks` |
| Recursos | `list_resources`, `add_resource` (trabajo / material / costo), `update_resource`, `delete_resource` |
| Asignaciones | `assign_resource`, `unassign_resource`, `list_assignments`, `level_resources` |
| Calendarios | `list_calendars`, `add_calendar_exceptions` (feriados), `set_work_week` (jornada y turnos) |
| Análisis | `get_critical_path`, **`check_schedule_health`** (revisión de calidad tipo DCMA), `set_baseline` / `clear_baseline` (0–10), `get_baseline_variance`, `get_timephased_data` (curvas S, histogramas de recursos) |
| Salida | `list_views`, `apply_view`, `export_pdf`, `export_tasks_csv`, `rename_custom_field`, `undo`, `recalculate` |

**Sintaxis compacta**, fácil de escribir para un LLM:

- Duraciones: `"5d"`, `"3h"`, `"2w"`, `"1mo"`. Fechas: `"2026-11-02"` o `"2026-11-02 08:00"`.
- Dependencias: `"3"`, `"5CC+2d"`, `"#2FF-1d"` (`#n` = n-ésima tarea del mismo lote). También acepta FS/SS/SF en inglés.
- Asignaciones: `"Ana"`, `"Carpinteros[200%]"`. Los recursos que no existen se crean solos.
- Las tareas se identifican por Id, por `"uid:<UniqueID>"` o por su nombre exacto.
- Cualquier campo de Project por su nombre: `{"Text1": "Subcontratista", "Number2": 5}`.

## Requisitos

- Windows con **Microsoft Project de escritorio** instalado y activado. Probado con Project Professional 2019 (16.0) de 64 bits.
- Python 3.10 o superior.
- Un cliente MCP: Claude Desktop, Claude Code u otro compatible.

**No** es compatible con Project Online ni con Project para la web, porque usa la API COM de escritorio.

## Instalación

```bash
git clone https://github.com/tomascarvallo/ms-project-mcp.git
cd ms-project-mcp
python -m venv .venv
.venv\Scripts\python -m pip install -e .
```

### Claude Desktop

Agrega el servidor en `%APPDATA%\Claude\claude_desktop_config.json`, usando la ruta donde clonaste el repositorio:

```json
{
  "mcpServers": {
    "ms-project": {
      "command": "C:\\ruta\\a\\ms-project-mcp\\.venv\\Scripts\\python.exe",
      "args": ["-m", "ms_project_mcp"],
      "env": { "PYTHONIOENCODING": "utf-8" }
    }
  }
}
```

### Claude Code

```bash
claude mcp add ms-project --scope user -- C:\ruta\a\ms-project-mcp\.venv\Scripts\python.exe -m ms_project_mcp
```

Reinicia el cliente. El servidor se inicia solo con Claude. Si Project está cerrado, se abre en la primera petición. Si ya está abierto, el servidor se conecta a esa ventana y trabaja sobre el **proyecto activo**.

> **Nota:** si usas Python de Microsoft Store, ese Python no ve la carpeta `%APPDATA%\Roaming`. Edita la configuración de Claude Desktop con otra herramienta, no con un script de ese Python.

## Ejemplos de uso

- *"Abre `C:\Proyectos\Bodega.mpp` y dime qué tareas están en la ruta crítica."*
- *"Arma un cronograma de remodelación de 3 meses con sus cuadrillas y los feriados de Chile."*
- *"Agrega Navidad y Año Nuevo como feriados y dime cuánto se mueve la fecha de término."*
- *"Guarda la línea base, marca las tareas 4 a 9 al 100 % y muéstrame la variación."*
- *"Revisa la calidad del cronograma y corrige las tareas sin sucesoras."*
- *"Dame el trabajo semanal de cada recurso para una curva S y exporta la carta Gantt a PDF."*

## Cómo funciona

- Todas las llamadas COM van por un único hilo STA. Si Project está ocupado, se reintentan automáticamente.
- Cada herramienta de edición se agrupa en una transacción de deshacer, así que **Ctrl+Z en Project revierte una operación completa de la IA**.
- Las duraciones se convierten con las horas por día y por semana del proyecto. Las fechas sin hora usan la hora de comienzo o fin predeterminada del proyecto.
- Los errores llegan al modelo como mensajes legibles, por ejemplo *"Hay 2 tareas llamadas 'Diseño'; use su Id: [3, 12]"*.

## Limitaciones

- Solo funciona en Windows, porque usa COM.
- Si Project tiene un cuadro de diálogo abierto (por ejemplo, Información de la tarea), las llamadas esperan unos 10 segundos y luego fallan con un aviso para cerrarlo.
- Abrir formatos no nativos distintos de `.xml` puede mostrar el asistente de importación de Project.

## Pruebas

```bash
.venv\Scripts\python tests\smoke_test.py
```

La prueba crea un proyecto temporal, ejecuta las 43 herramientas y lo cierra sin guardar. Los 2 errores que reporta son verificaciones negativas intencionales.

## Palabras clave

MCP para Microsoft Project, MS Project con inteligencia artificial, automatizar MS Project con Python, Claude y Microsoft Project, programación de obras con IA, carta Gantt con IA, cronograma de construcción, método de la ruta crítica, valor ganado, gestión de proyectos, servidor Model Context Protocol.

## Licencia

[MIT](LICENSE). Este proyecto no está afiliado ni respaldado por Microsoft. *Microsoft Project* es una marca registrada de Microsoft Corporation.
