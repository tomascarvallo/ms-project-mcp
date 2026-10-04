"""Servidor MCP para Microsoft Project de escritorio (vía COM).

Controla la instancia de Project instalada en este equipo: abre/crea/guarda
archivos, lee y edita tareas, recursos, asignaciones, dependencias,
calendarios, líneas base, nivelación, datos en fase temporal y exporta.

Convenciones:
- Las tareas se identifican por su Id (número de fila), por "uid:<UniqueID>"
  o por su nombre exacto.
- Duraciones como texto: "5d", "3h", "2w", "1mo", "30m" (número solo = días).
- Fechas ISO: "2026-10-15" o "2026-10-15 08:00".
- Dependencias abreviadas: "3", "3FS", "5SS+2d", "#2FF-1d" (# = fila dentro
  del mismo lote en bulk_create_tasks). También acepta FC/CC/CF en español.
- Cada operación que modifica el proyecto es un único paso de Deshacer.
"""

from __future__ import annotations

import csv
import functools
import re
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from typing import Any

import pythoncom
import pywintypes
import win32com.client as w32
from win32com.client import gencache

from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

mcp = MCPServer(
    "ms-project",
    instructions=__doc__,
)

# ---------------------------------------------------------------------------
# Hilo COM único: Project es un servidor STA, todas las llamadas van por aquí.
# ---------------------------------------------------------------------------

_executor = ThreadPoolExecutor(max_workers=1, initializer=pythoncom.CoInitialize)
_app = None
_C = None  # constantes del type library

RPC_E_CALL_REJECTED = -2147418111
RPC_E_SERVERCALL_RETRYLATER = -2147417846


class ProjectError(ToolError):
    """Error esperado: el mensaje llega al modelo."""


def _run(fn):
    def call():
        for attempt in range(20):
            try:
                return fn()
            except pywintypes.com_error as e:
                if e.hresult in (RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER) and attempt < 19:
                    time.sleep(0.5)  # Project ocupado (diálogo abierto, recalculando)
                    continue
                raise ProjectError(_com_msg(e)) from None
            except ProjectError:
                raise
            except Exception as e:  # cualquier otra falla también se informa al modelo
                raise ProjectError(f"{type(e).__name__}: {e}") from e
    return _executor.submit(call).result()


def _com_msg(e: pywintypes.com_error) -> str:
    detail = ""
    if e.excepinfo and e.excepinfo[2]:
        detail = e.excepinfo[2]
    elif e.strerror:
        detail = e.strerror
    hint = ""
    if e.hresult in (RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER):
        hint = " (Project está ocupado: cierre cualquier cuadro de diálogo abierto en Project)"
    return f"Error de Project: {detail}{hint}"


def app():
    global _app, _C
    if _app is not None:
        try:
            _app.Version
            return _app
        except Exception:
            _app = None
    try:
        obj = w32.GetActiveObject("MSProject.Application")
        _app = gencache.EnsureDispatch(obj)
    except pywintypes.com_error:
        _app = gencache.EnsureDispatch("MSProject.Application")
    _app.Visible = True
    _C = w32.constants
    return _app


def C(name: str) -> int:
    app()
    return getattr(_C, name)


def proj():
    a = app()
    if a.Projects.Count == 0:
        raise ProjectError("No hay ningún proyecto abierto. Use open_project o new_project.")
    return a.ActiveProject


def pj_tool(fn=None, *, undo: str | None = None):
    """Registra una tool cuyo cuerpo corre en el hilo COM.

    Si `undo` se indica, la operación se agrupa en una transacción de Deshacer.
    """
    def deco(f):
        @functools.wraps(f)
        def wrapper(*args, **kwargs):
            def body():
                if undo:
                    a = app()
                    opened = False
                    if a.Projects.Count:
                        try:
                            a.OpenUndoTransaction(f"Claude: {undo}")
                            opened = True
                        except Exception:
                            pass
                    try:
                        return f(*args, **kwargs)
                    finally:
                        if opened:
                            try:
                                a.CloseUndoTransaction()
                            except Exception:
                                pass
                return f(*args, **kwargs)
            return _run(body)
        return mcp.tool()(wrapper)
    return deco(fn) if fn else deco


# ---------------------------------------------------------------------------
# Conversión de valores
# ---------------------------------------------------------------------------

def _date_out(v) -> str | None:
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d %H:%M")
    return None  # "NA"/"ND" u otro marcador de vacío


def _default_time(kind: str) -> tuple[int, int]:
    """Hora por defecto del proyecto (Opciones > Programación) para fechas sin hora."""
    try:
        p = app().ActiveProject
        v = p.DefaultStartTime if kind == "start" else p.DefaultFinishTime
        if isinstance(v, datetime):
            return v.hour, v.minute
        hh, mm = str(v).split(":")[:2]
        return int(hh), int(mm[:2])
    except Exception:
        return (8, 0) if kind == "start" else (17, 0)


def _date_in(s: str | None, kind: str = "start"):
    """Texto ISO -> fecha COM. Sin hora, usa la hora de comienzo/fin predeterminada del proyecto.

    pywin32 convierte las fechas como UTC; se marca la hora local como UTC para que llegue intacta.
    """
    if s is None or s == "":
        return None
    s = s.strip().replace("T", " ")
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            dt = datetime.strptime(s, fmt)
        except ValueError:
            continue
        if fmt == "%Y-%m-%d":
            hh, mm = _default_time(kind)
            dt = dt.replace(hour=hh, minute=mm)
        return pywintypes.Time(dt.replace(tzinfo=timezone.utc))
    raise ProjectError(f"Fecha no válida: {s!r}. Use AAAA-MM-DD o AAAA-MM-DD HH:MM.")


def _minutes_per(p) -> dict[str, float]:
    hpd = float(p.HoursPerDay)
    return {
        "m": 1, "min": 1,
        "h": 60, "hr": 60, "hrs": 60, "hora": 60, "horas": 60,
        "d": hpd * 60, "dia": hpd * 60, "dias": hpd * 60, "día": hpd * 60, "días": hpd * 60, "day": hpd * 60, "days": hpd * 60,
        "w": float(p.HoursPerWeek) * 60, "s": float(p.HoursPerWeek) * 60, "sem": float(p.HoursPerWeek) * 60,
        "semana": float(p.HoursPerWeek) * 60, "semanas": float(p.HoursPerWeek) * 60, "wk": float(p.HoursPerWeek) * 60,
        "mo": float(p.DaysPerMonth) * hpd * 60, "mes": float(p.DaysPerMonth) * hpd * 60, "meses": float(p.DaysPerMonth) * hpd * 60,
    }


_DUR_RE = re.compile(r"^\s*([+-]?\d+(?:[.,]\d+)?)\s*([a-zA-Záí]*)\s*$")


def _dur_in(p, s) -> int:
    if isinstance(s, (int, float)):
        return round(float(s) * float(p.HoursPerDay) * 60)
    m = _DUR_RE.match(str(s))
    if not m:
        raise ProjectError(f"Duración no válida: {s!r}. Ejemplos: 5d, 3h, 2w, 1mo.")
    val = float(m.group(1).replace(",", "."))
    unit = (m.group(2) or "d").lower()
    factors = _minutes_per(p)
    if unit not in factors:
        raise ProjectError(f"Unidad de duración desconocida: {unit!r}")
    return round(val * factors[unit])


def _days(p, minutes) -> float | None:
    try:
        return round(float(minutes) / (float(p.HoursPerDay) * 60), 2)
    except (TypeError, ValueError):
        return None


def _hours(minutes) -> float | None:
    try:
        return round(float(minutes) / 60, 2)
    except (TypeError, ValueError):
        return None


def _num(v):
    try:
        return round(float(v), 2)
    except (TypeError, ValueError):
        return None


_DEP_TYPES = {
    "FS": "pjFinishToStart", "FC": "pjFinishToStart",
    "SS": "pjStartToStart", "CC": "pjStartToStart",
    "FF": "pjFinishToFinish",
    "SF": "pjStartToFinish", "CF": "pjStartToFinish",
}
_DEP_NAMES = {}  # valor -> "FS"...


def _dep_code(value: int) -> str:
    if not _DEP_NAMES:
        for code in ("FS", "SS", "FF", "SF"):
            _DEP_NAMES[C(_DEP_TYPES[code])] = code
    return _DEP_NAMES.get(value, str(value))


_CONSTRAINTS = {
    "ASAP": "pjASAP", "ALAP": "pjALAP", "MSO": "pjMSO", "MFO": "pjMFO",
    "SNET": "pjSNET", "SNLT": "pjSNLT", "FNET": "pjFNET", "FNLT": "pjFNLT",
}
_TASK_TYPES = {"fixed_units": "pjFixedUnits", "fixed_duration": "pjFixedDuration", "fixed_work": "pjFixedWork"}
_RES_TYPES = {"work": "pjResourceTypeWork", "material": "pjResourceTypeMaterial", "cost": "pjResourceTypeCost"}


def _const_name(mapping: dict[str, str], value) -> str:
    for k, v in mapping.items():
        if C(v) == value:
            return k
    return str(value)


# ---------------------------------------------------------------------------
# Búsqueda de tareas / recursos
# ---------------------------------------------------------------------------

def _iter_tasks(p):
    for t in p.Tasks:
        if t is not None:
            yield t


def _task(p, ref):
    if isinstance(ref, str):
        r = ref.strip()
        if r.lower().startswith("uid:"):
            try:
                return p.Tasks.UniqueID(int(r[4:]))
            except pywintypes.com_error:
                raise ProjectError(f"No existe la tarea con UniqueID {r[4:]}") from None
        if r.isdigit():
            ref = int(r)
        else:
            matches = [t for t in _iter_tasks(p) if t.Name == r]
            if not matches:
                matches = [t for t in _iter_tasks(p) if t.Name.lower() == r.lower()]
            if len(matches) == 1:
                return matches[0]
            if not matches:
                raise ProjectError(f"No existe una tarea llamada {r!r}")
            raise ProjectError(f"Hay {len(matches)} tareas llamadas {r!r}; use su Id: {[t.ID for t in matches]}")
    try:
        t = p.Tasks(int(ref))
    except pywintypes.com_error:
        t = None
    if t is None:
        raise ProjectError(f"No existe la tarea con Id {ref}")
    return t


def _iter_resources(p):
    for r in p.Resources:
        if r is not None:
            yield r


def _resource(p, ref, create: bool = False):
    if isinstance(ref, int) or (isinstance(ref, str) and ref.strip().isdigit()):
        try:
            r = p.Resources(int(ref))
            if r is not None:
                return r
        except pywintypes.com_error:
            pass
        raise ProjectError(f"No existe el recurso con Id {ref}")
    name = str(ref).strip()
    for r in _iter_resources(p):
        if r.Name.lower() == name.lower():
            return r
    if create:
        r = p.Resources.Add(name)
        return r
    raise ProjectError(f"No existe el recurso {name!r}")


# ---------------------------------------------------------------------------
# Serialización
# ---------------------------------------------------------------------------

def _task_brief(p, t) -> dict:
    return {
        "id": t.ID,
        "uid": t.UniqueID,
        "wbs": t.WBS,
        "name": t.Name,
        "level": t.OutlineLevel,
        "summary": bool(t.Summary),
        "milestone": bool(t.Milestone),
        "duration_days": _days(p, t.Duration),
        "start": _date_out(t.Start),
        "finish": _date_out(t.Finish),
        "percent_complete": t.PercentComplete,
        "predecessors": t.Predecessors,
        "resources": t.ResourceNames,
        "critical": bool(t.Critical),
    }


def _task_full(p, t) -> dict:
    d = _task_brief(p, t)
    preds, succs = [], []
    for dep in t.TaskDependencies:
        item = {
            "type": _dep_code(dep.Type),
            "lag_days": _days(p, dep.Lag),
        }
        if dep.To.UniqueID == t.UniqueID:
            preds.append({"id": dep.From.ID, "name": dep.From.Name, **item})
        else:
            succs.append({"id": dep.To.ID, "name": dep.To.Name, **item})
    assigns = []
    for a in t.Assignments:
        assigns.append({
            "resource": a.ResourceName,
            "units_percent": _num(float(a.Units) * 100),
            "work_hours": _hours(a.Work),
            "cost": _num(a.Cost),
        })
    d.update({
        "parent_id": t.OutlineParent.ID if t.OutlineLevel > 1 else None,
        "manual": bool(getattr(t, "Manual", False)),
        "active": bool(getattr(t, "Active", True)),
        "task_type": _const_name(_TASK_TYPES, t.Type),
        "effort_driven": bool(t.EffortDriven),
        "work_hours": _hours(t.Work),
        "cost": _num(t.Cost),
        "fixed_cost": _num(t.FixedCost),
        "priority": t.Priority,
        "constraint_type": _const_name(_CONSTRAINTS, t.ConstraintType),
        "constraint_date": _date_out(t.ConstraintDate),
        "deadline": _date_out(t.Deadline),
        "total_slack_days": _days(p, t.TotalSlack),
        "free_slack_days": _days(p, t.FreeSlack),
        "actual_start": _date_out(t.ActualStart),
        "actual_finish": _date_out(t.ActualFinish),
        "remaining_duration_days": _days(p, t.RemainingDuration),
        "baseline_start": _date_out(t.BaselineStart),
        "baseline_finish": _date_out(t.BaselineFinish),
        "start_variance_days": _days(p, t.StartVariance),
        "finish_variance_days": _days(p, t.FinishVariance),
        "calendar": t.Calendar,
        "notes": t.Notes,
        "predecessor_detail": preds,
        "successor_detail": succs,
        "assignments": assigns,
    })
    return d


def _resource_dict(p, r) -> dict:
    return {
        "id": r.ID,
        "uid": r.UniqueID,
        "name": r.Name,
        "type": _const_name(_RES_TYPES, r.Type),
        "initials": r.Initials,
        "group": r.Group,
        "max_units_percent": _num(float(r.MaxUnits) * 100) if r.Type == C("pjResourceTypeWork") else None,
        "standard_rate": r.StandardRate,
        "overtime_rate": r.OvertimeRate,
        "cost_per_use": _num(r.CostPerUse),
        "material_label": r.MaterialLabel if r.Type == C("pjResourceTypeMaterial") else None,
        "email": r.EMailAddress,
        "work_hours": _hours(r.Work),
        "cost": _num(r.Cost),
        "overallocated": bool(r.Overallocated),
        "assignment_count": r.Assignments.Count,
    }


def _extra_fields(t, names: list[str] | None, kind: str = "pjTask") -> dict:
    out = {}
    if not names:
        return out
    a = app()
    for n in names:
        try:
            out[n] = t.GetField(a.FieldNameToFieldConstant(n, C(kind)))
        except Exception:
            out[n] = None
    return out


# ---------------------------------------------------------------------------
# Archivos y proyectos
# ---------------------------------------------------------------------------

@pj_tool
def project_status() -> dict:
    """Estado de Project: versión, proyectos abiertos y cuál está activo."""
    a = app()
    projects = [{"name": pr.Name, "path": pr.FullName, "saved": bool(pr.Saved)} for pr in a.Projects]
    active = a.ActiveProject.Name if a.Projects.Count else None
    return {"version": a.Version, "build": a.Build, "open_projects": projects, "active_project": active}


@pj_tool
def open_project(path: str, read_only: bool = False) -> dict:
    """Abre un archivo .mpp, .xml u otro formato soportado por Project y lo deja activo."""
    a = app()
    kw: dict[str, Any] = {"Name": path, "ReadOnly": read_only}
    if path.lower().endswith(".xml"):
        kw["FormatID"] = "MSProject.XML"
    alerts = a.DisplayAlerts
    a.DisplayAlerts = False  # evita el Asistente para importación, que bloquearía el servidor
    try:
        a.FileOpenEx(**kw)
    finally:
        a.DisplayAlerts = alerts
    p = a.ActiveProject
    return {"opened": p.FullName, "tasks": p.Tasks.Count, "resources": p.Resources.Count}


@pj_tool
def new_project(start: str | None = None, title: str | None = None, template: str | None = None) -> dict:
    """Crea un proyecto nuevo (opcionalmente desde una plantilla .mpt) y lo activa.

    start: fecha de inicio del proyecto. title: nombre de la tarea resumen del proyecto.
    """
    a = app()
    if template:
        a.FileOpenEx(Name=template)
    else:
        a.FileNew(SummaryInfo=False)
    p = a.ActiveProject
    if start:
        p.ProjectStart = _date_in(start)
    if title:
        p.ProjectSummaryTask.Name = title
        p.DisplayProjectSummaryTask = True
    return {"name": p.Name, "start": _date_out(p.ProjectStart)}


@pj_tool
def activate_project(name: str) -> dict:
    """Activa uno de los proyectos abiertos (por nombre de archivo o ruta)."""
    a = app()
    for pr in a.Projects:
        if name.lower() in (pr.Name.lower(), pr.FullName.lower()):
            pr.Activate()
            return {"active_project": pr.Name}
    raise ProjectError(f"No hay un proyecto abierto llamado {name!r}")


@pj_tool
def save_project(path: str | None = None) -> dict:
    """Guarda el proyecto activo. Con `path` hace Guardar como; .xml guarda en formato XML de Project."""
    a = app()
    p = proj()
    if not path:
        a.FileSave()
        return {"saved": p.FullName}
    alerts = a.DisplayAlerts
    a.DisplayAlerts = False
    try:
        if path.lower().endswith(".xml"):
            a.FileSaveAs(Name=path, FormatID="MSProject.XML")
        else:
            a.FileSaveAs(Name=path)
    finally:
        a.DisplayAlerts = alerts
    return {"saved": a.ActiveProject.FullName}


@pj_tool
def close_project(save: bool = False) -> dict:
    """Cierra el proyecto activo. save=False descarta los cambios no guardados."""
    a = app()
    name = proj().Name
    a.FileCloseEx(Save=C("pjSave") if save else C("pjDoNotSave"))
    return {"closed": name, "saved": save}


@pj_tool(undo="propiedades del proyecto")
def set_project_properties(
    start: str | None = None,
    finish: str | None = None,
    status_date: str | None = None,
    title: str | None = None,
    calendar: str | None = None,
    hours_per_day: float | None = None,
    hours_per_week: float | None = None,
    days_per_month: float | None = None,
    manager: str | None = None,
    company: str | None = None,
    show_summary_task: bool | None = None,
) -> dict:
    """Cambia propiedades del proyecto activo.

    start/finish: programar desde inicio o desde fin (si se da finish, el proyecto se programa hacia atrás).
    status_date: fecha de estado (para seguimiento y valor ganado). calendar: calendario base del proyecto.
    """
    a = app()
    p = proj()
    if start:
        p.ScheduleFromStart = True
        p.ProjectStart = _date_in(start)
    if finish:
        p.ScheduleFromStart = False
        p.ProjectFinish = _date_in(finish, "finish")
    if status_date:
        p.StatusDate = _date_in(status_date)
    if title:
        p.ProjectSummaryTask.Name = title
    if calendar:
        p.Calendar = p.BaseCalendars(calendar)
    if hours_per_day is not None:
        p.HoursPerDay = hours_per_day
    if hours_per_week is not None:
        p.HoursPerWeek = hours_per_week
    if days_per_month is not None:
        p.DaysPerMonth = days_per_month
    if manager is not None:
        p.Manager = manager
    if company is not None:
        p.Company = company
    if show_summary_task is not None:
        p.DisplayProjectSummaryTask = show_summary_task
    return _summary(p)


def _summary(p) -> dict:
    s = p.ProjectSummaryTask
    out = {
        "name": p.Name,
        "path": p.FullName,
        "title": s.Name,
        "schedule_from_start": bool(p.ScheduleFromStart),
        "start": _date_out(s.Start),
        "finish": _date_out(s.Finish),
        "duration_days": _days(p, s.Duration),
        "work_hours": _hours(s.Work),
        "cost": _num(s.Cost),
        "percent_complete": s.PercentComplete,
        "status_date": _date_out(p.StatusDate),
        "calendar": p.Calendar.Name,
        "hours_per_day": _num(p.HoursPerDay),
        "hours_per_week": _num(p.HoursPerWeek),
        "baseline_finish": _date_out(s.BaselineFinish),
        "finish_variance_days": _days(p, s.FinishVariance),
        "baseline_cost": _num(s.BaselineCost),
        "cost_variance": _num(s.CostVariance),
    }
    if _date_out(s.BaselineFinish):
        for k, attr in (("BCWS", "BCWS"), ("BCWP", "BCWP"), ("ACWP", "ACWP"), ("SV", "SV"), ("CV", "CV"),
                        ("SPI", "SPI"), ("CPI", "CPI"), ("EAC", "EAC"), ("VAC", "VAC")):
            try:
                out.setdefault("earned_value", {})[k] = _num(getattr(s, attr))
            except Exception:
                pass
    return out


@pj_tool
def get_project_summary() -> dict:
    """Resumen del proyecto activo: fechas, duración, trabajo, costo, avance, línea base, valor ganado y conteos."""
    p = proj()
    out = _summary(p)
    tasks = [t for t in _iter_tasks(p)]
    leaf = [t for t in tasks if not t.Summary]
    status = p.StatusDate if isinstance(p.StatusDate, datetime) else datetime.now()
    out["counts"] = {
        "tasks": len(leaf),
        "summary_tasks": len(tasks) - len(leaf),
        "milestones": sum(1 for t in leaf if t.Milestone),
        "completed": sum(1 for t in leaf if t.PercentComplete == 100),
        "in_progress": sum(1 for t in leaf if 0 < t.PercentComplete < 100),
        "not_started": sum(1 for t in leaf if t.PercentComplete == 0),
        "critical": sum(1 for t in leaf if t.Critical),
        "late": sum(1 for t in leaf if t.PercentComplete < 100 and isinstance(t.Finish, datetime)
                    and t.Finish.replace(tzinfo=None) < status.replace(tzinfo=None)),
        "resources": sum(1 for _ in _iter_resources(p)),
        "overallocated_resources": sum(1 for r in _iter_resources(p) if r.Overallocated),
    }
    return out


# ---------------------------------------------------------------------------
# Tareas
# ---------------------------------------------------------------------------

@pj_tool
def list_tasks(
    filter: str | None = None,
    name_contains: str | None = None,
    max_outline_level: int | None = None,
    include_summary: bool = True,
    extra_fields: list[str] | None = None,
    limit: int = 500,
) -> dict:
    """Lista tareas del proyecto activo.

    filter: "critical", "milestones", "late", "in_progress", "not_started", "completed",
    "overallocated" (con recursos sobreasignados), "manual", "no_predecessors", "no_resources".
    extra_fields: nombres de campos de Project adicionales a leer (ej. "Text1", "Number1", "Cost1", "Flag1").
    """
    p = proj()
    status = p.StatusDate if isinstance(p.StatusDate, datetime) else datetime.now()
    status = status.replace(tzinfo=None)
    out = []
    for t in _iter_tasks(p):
        if not include_summary and t.Summary:
            continue
        if max_outline_level and t.OutlineLevel > max_outline_level:
            continue
        if name_contains and name_contains.lower() not in t.Name.lower():
            continue
        f = (filter or "").lower()
        if f == "critical" and not t.Critical:
            continue
        if f == "milestones" and not t.Milestone:
            continue
        if f == "late" and not (t.PercentComplete < 100 and isinstance(t.Finish, datetime)
                                and t.Finish.replace(tzinfo=None) < status):
            continue
        if f == "in_progress" and not (0 < t.PercentComplete < 100):
            continue
        if f == "not_started" and t.PercentComplete != 0:
            continue
        if f == "completed" and t.PercentComplete != 100:
            continue
        if f == "overallocated" and not t.Overallocated:
            continue
        if f == "manual" and not getattr(t, "Manual", False):
            continue
        if f == "no_predecessors" and (t.Summary or t.Predecessors):
            continue
        if f == "no_resources" and (t.Summary or t.Milestone or t.ResourceNames):
            continue
        d = _task_brief(p, t)
        d.update(_extra_fields(t, extra_fields))
        out.append(d)
        if len(out) >= limit:
            break
    return {"count": len(out), "tasks": out}


@pj_tool
def get_task(task: str) -> dict:
    """Detalle completo de una tarea (Id, "uid:N" o nombre): dependencias, asignaciones, holguras, restricciones, línea base, notas."""
    p = proj()
    return _task_full(p, _task(p, task))


def _parse_dep(p, spec, batch_ids: list[int] | None = None) -> tuple[Any, int, int]:
    """'5SS+2d' -> (tarea, tipo, lag_minutos)."""
    if isinstance(spec, dict):
        ref = spec.get("task")
        code = (spec.get("type") or "FS").upper()
        lag = spec.get("lag") or 0
    else:
        m = re.match(r"^\s*(#?\w[\w:]*?)\s*(FS|SS|FF|SF|FC|CC|CF)?\s*([+-].+)?\s*$", str(spec), re.I)
        if not m:
            raise ProjectError(f"Dependencia no válida: {spec!r}")
        ref, code, lag = m.group(1), (m.group(2) or "FS").upper(), (m.group(3) or "0").replace(" ", "")
    if isinstance(ref, str) and ref.startswith("#"):
        if not batch_ids:
            raise ProjectError("'#n' solo se puede usar dentro de bulk_create_tasks")
        ref = batch_ids[int(ref[1:]) - 1]
    if code not in _DEP_TYPES:
        raise ProjectError(f"Tipo de dependencia desconocido: {code}")
    lag_min = _dur_in(p, lag) if lag not in (0, "0", "+0") else 0
    return _task(p, ref), C(_DEP_TYPES[code]), lag_min


def _parse_assignment(spec) -> tuple[str, float | None]:
    if isinstance(spec, dict):
        return spec["resource"], spec.get("units")
    m = re.match(r"^\s*(.+?)\s*(?:\[\s*([\d.,]+)\s*%?\s*\])?\s*$", str(spec))
    units = float(m.group(2).replace(",", ".")) if m.group(2) else None
    return m.group(1), units


def _assign(p, t, spec, create_resources: bool):
    name, units = _parse_assignment(spec)
    r = _resource(p, name, create=create_resources)
    if units is None:
        return t.Assignments.Add(ResourceID=r.ID)
    return t.Assignments.Add(ResourceID=r.ID, Units=units / 100.0)


def _last_descendant_id(p, t) -> int:
    last = t.ID
    lvl = t.OutlineLevel
    i = t.ID + 1
    n = p.Tasks.Count
    while i <= n:
        x = p.Tasks(i)
        if x is not None:
            if x.OutlineLevel <= lvl:
                break
            last = x.ID
        i += 1
    return last


def _create_task(p, name: str, parent=None, before=None, outline_level=None):
    if parent is not None:
        par = _task(p, parent)
        pos = _last_descendant_id(p, par) + 1
        t = p.Tasks.Add(Name=name, Before=pos) if pos <= p.Tasks.Count else p.Tasks.Add(Name=name)
        t.OutlineLevel = par.OutlineLevel + 1
    elif before is not None:
        t = p.Tasks.Add(Name=name, Before=_task(p, before).ID)
        if outline_level:
            t.OutlineLevel = outline_level
    else:
        t = p.Tasks.Add(Name=name)
        if outline_level:
            t.OutlineLevel = outline_level
    try:
        t.Manual = False  # programación automática
    except Exception:
        pass
    return t


_SIMPLE_FIELDS = {
    # clave -> (propiedad COM, conversor)
    "name": ("Name", None),
    "notes": ("Notes", None),
    "priority": ("Priority", int),
    "percent_complete": ("PercentComplete", int),
    "percent_work_complete": ("PercentWorkComplete", int),
    "milestone": ("Milestone", bool),
    "manual": ("Manual", bool),
    "active": ("Active", bool),
    "effort_driven": ("EffortDriven", bool),
    "fixed_cost": ("FixedCost", float),
    "outline_level": ("OutlineLevel", int),
    "calendar": ("Calendar", str),
    "wbs": ("WBS", str),
    "hyperlink": ("HyperlinkAddress", str),
}
_DATE_FIELDS = {
    "start": "Start", "finish": "Finish", "deadline": "Deadline", "constraint_date": "ConstraintDate",
    "actual_start": "ActualStart", "actual_finish": "ActualFinish",
}
_DUR_FIELDS = {
    "duration": "Duration", "actual_duration": "ActualDuration", "remaining_duration": "RemainingDuration",
    "work": "Work", "actual_work": "ActualWork", "remaining_work": "RemainingWork",
}


def _apply_task_fields(p, t, fields: dict) -> list[str]:
    a = app()
    applied = []
    # el tipo de restricción antes de la fecha de restricción; duración antes de % completado
    order = ["task_type", "constraint_type", "duration", "milestone"]
    keys = sorted(fields, key=lambda k: order.index(k) if k in order else len(order))
    for k in keys:
        v = fields[k]
        kl = k.lower()
        if kl == "constraint_type":
            t.ConstraintType = C(_CONSTRAINTS[str(v).upper()])
        elif kl == "task_type":
            t.Type = C(_TASK_TYPES[str(v).lower()])
        elif kl in _SIMPLE_FIELDS:
            attr, conv = _SIMPLE_FIELDS[kl]
            setattr(t, attr, conv(v) if conv else v)
        elif kl in _DATE_FIELDS:
            kind = "finish" if kl in ("finish", "deadline", "actual_finish") else "start"
            setattr(t, _DATE_FIELDS[kl], _date_in(v, kind) if v else "NA")
        elif kl in _DUR_FIELDS:
            # trabajo se expresa en horas por defecto, duración en días
            if kl.endswith("work") and isinstance(v, (int, float)):
                v = f"{v}h"
            setattr(t, _DUR_FIELDS[kl], _dur_in(p, v))
        else:
            # cualquier otro campo de Project por su nombre (Text1, Number3, Flag2, Cost1, Date1...)
            fid = a.FieldNameToFieldConstant(k, C("pjTask"))
            t.SetField(fid, str(v))
        applied.append(k)
    return applied


@pj_tool(undo="agregar tarea")
def add_task(
    name: str,
    duration: str | None = None,
    parent: str | None = None,
    before: str | None = None,
    outline_level: int | None = None,
    start: str | None = None,
    milestone: bool = False,
    predecessors: list[str] | None = None,
    resources: list[str] | None = None,
    notes: str | None = None,
    fields: dict[str, Any] | None = None,
    create_missing_resources: bool = True,
) -> dict:
    """Agrega una tarea con programación automática.

    parent: tarea resumen bajo la cual insertarla (al final de sus hijas). before: insertar antes de esa tarea.
    predecessors: ["3", "5SS+2d"]. resources: ["Ana", "Luis[50%]"] (los recursos inexistentes se crean).
    start: fija una restricción "No comenzar antes de". fields: otros campos (ver update_task).
    """
    p = proj()
    t = _create_task(p, name, parent, before, outline_level)
    if milestone:
        t.Duration = 0
        t.Milestone = True
    elif duration is not None:
        t.Duration = _dur_in(p, duration)
    if start:
        t.Start = _date_in(start)
    if notes:
        t.Notes = notes
    for spec in predecessors or []:
        pred, typ, lag = _parse_dep(p, spec)
        t.TaskDependencies.Add(From=pred, Type=typ, Lag=lag)
    for spec in resources or []:
        _assign(p, t, spec, create_missing_resources)
    if fields:
        _apply_task_fields(p, t, fields)
    return _task_brief(p, t)


@pj_tool(undo="crear tareas en lote")
def bulk_create_tasks(tasks: list[dict[str, Any]], create_missing_resources: bool = True) -> dict:
    """Crea muchas tareas de una vez (ideal para armar una EDT/cronograma completo).

    Cada elemento: {"name", "level" (1 = nivel superior, define la jerarquía), "duration", "milestone",
    "predecessors": ["#1", "#2SS+1d", "15"], "resources": ["Ana[50%]"], "start", "notes", "fields": {...}}.
    "#n" referencia al n-ésimo elemento de esta misma lista; un número sin # es un Id existente.
    Las tareas se agregan al final del proyecto. Las que tienen hijas se vuelven resumen automáticamente.
    """
    p = proj()
    base_level = 0
    ids: list[int] = []
    created = []
    # 1) crear todas las filas y la jerarquía
    for spec in tasks:
        t = _create_task(p, spec["name"])
        lvl = int(spec.get("level", 1)) + base_level
        if t.OutlineLevel != lvl:
            t.OutlineLevel = lvl
        ids.append(t.ID)
        created.append(t)
    # 2) duraciones, campos, dependencias y recursos (después de la jerarquía)
    for spec, t in zip(tasks, created):
        is_summary = bool(t.Summary)
        if spec.get("milestone"):
            t.Duration = 0
            t.Milestone = True
        elif spec.get("duration") is not None and not is_summary:
            t.Duration = _dur_in(p, spec["duration"])
        if spec.get("start"):
            t.Start = _date_in(spec["start"])
        if spec.get("notes"):
            t.Notes = spec["notes"]
        for dep in spec.get("predecessors") or []:
            pred, typ, lag = _parse_dep(p, dep, ids)
            t.TaskDependencies.Add(From=pred, Type=typ, Lag=lag)
        if not is_summary:
            for r in spec.get("resources") or []:
                _assign(p, t, r, create_missing_resources)
        if spec.get("fields"):
            _apply_task_fields(p, t, spec["fields"])
    return {"created": len(created), "ids": ids, "project": _summary(p)}


@pj_tool(undo="actualizar tarea")
def update_task(task: str, fields: dict[str, Any]) -> dict:
    """Modifica campos de una tarea.

    Claves: name, duration ("5d"), start, finish, deadline, percent_complete, percent_work_complete,
    actual_start, actual_finish, actual_duration, remaining_duration, work ("40h"), actual_work,
    milestone, manual, active, priority (0-1000), notes, fixed_cost, calendar, outline_level, wbs,
    constraint_type (ASAP, ALAP, MSO, MFO, SNET, SNLT, FNET, FNLT), constraint_date,
    task_type (fixed_units, fixed_duration, fixed_work), effort_driven, hyperlink.
    Cualquier otra clave se trata como nombre de campo de Project (Text1, Number1, Flag1, Cost1, Date1, ...).
    """
    p = proj()
    t = _task(p, task)
    applied = _apply_task_fields(p, t, fields)
    return {"applied": applied, "task": _task_brief(p, t)}


@pj_tool(undo="actualizar tareas en lote")
def bulk_update_tasks(updates: list[dict[str, Any]]) -> dict:
    """Actualiza varias tareas: [{"task": "12", "fields": {"percent_complete": 50}}, ...]. Útil para registrar avance."""
    p = proj()
    results = []
    for u in updates:
        t = _task(p, u["task"])
        results.append({"id": t.ID, "applied": _apply_task_fields(p, t, u["fields"])})
    return {"updated": len(results), "results": results, "project": _summary(p)}


@pj_tool(undo="eliminar tarea")
def delete_task(task: str) -> dict:
    """Elimina una tarea (si es resumen, también sus subtareas)."""
    p = proj()
    t = _task(p, task)
    info = {"id": t.ID, "name": t.Name, "summary": bool(t.Summary)}
    t.Delete()
    return {"deleted": info}


@pj_tool(undo="sangría")
def indent_task(task: str, outdent: bool = False) -> dict:
    """Aplica sangría (convierte en subtarea de la anterior) o quita sangría a una tarea."""
    p = proj()
    t = _task(p, task)
    if outdent:
        t.OutlineOutdent()
    else:
        t.OutlineIndent()
    return _task_brief(p, t)


# ---------------------------------------------------------------------------
# Dependencias
# ---------------------------------------------------------------------------

@pj_tool(undo="vincular tareas")
def link_tasks(predecessor: str, successor: str, type: str = "FS", lag: str = "0d") -> dict:
    """Crea una dependencia. type: FS/SS/FF/SF (o FC/CC/CF). lag: "2d", "-1d", "4h"."""
    p = proj()
    pred = _task(p, predecessor)
    succ = _task(p, successor)
    succ.TaskDependencies.Add(From=pred, Type=C(_DEP_TYPES[type.upper()]), Lag=_dur_in(p, lag))
    return _task_brief(p, succ)


@pj_tool(undo="vincular cadena")
def link_chain(tasks: list[str], type: str = "FS") -> dict:
    """Vincula tareas en secuencia (1→2→3...) con el mismo tipo de dependencia."""
    p = proj()
    ts = [_task(p, x) for x in tasks]
    for a_, b_ in zip(ts, ts[1:]):
        b_.TaskDependencies.Add(From=a_, Type=C(_DEP_TYPES[type.upper()]), Lag=0)
    return {"linked": [t.ID for t in ts]}


@pj_tool(undo="desvincular tareas")
def unlink_tasks(predecessor: str, successor: str) -> dict:
    """Elimina la dependencia entre dos tareas."""
    p = proj()
    pred = _task(p, predecessor)
    succ = _task(p, successor)
    for dep in succ.TaskDependencies:
        if dep.From.UniqueID == pred.UniqueID and dep.To.UniqueID == succ.UniqueID:
            dep.Delete()
            return {"removed": f"{pred.ID} -> {succ.ID}"}
    raise ProjectError(f"No existe dependencia {pred.ID} -> {succ.ID}")


# ---------------------------------------------------------------------------
# Recursos y asignaciones
# ---------------------------------------------------------------------------

@pj_tool
def list_resources(extra_fields: list[str] | None = None) -> dict:
    """Lista los recursos con tipo, tasas, capacidad, trabajo, costo y si están sobreasignados."""
    p = proj()
    out = []
    for r in _iter_resources(p):
        d = _resource_dict(p, r)
        d.update(_extra_fields(r, extra_fields, "pjResource"))
        out.append(d)
    return {"count": len(out), "resources": out}


def _apply_resource_fields(p, r, f: dict):
    a = app()
    if "type" in f:
        r.Type = C(_RES_TYPES[f["type"].lower()])
    for k, v in f.items():
        kl = k.lower()
        if kl == "type":
            continue
        if kl == "name":
            r.Name = v
        elif kl == "max_units_percent":
            r.MaxUnits = float(v) / 100.0
        elif kl == "standard_rate":
            r.StandardRate = v
        elif kl == "overtime_rate":
            r.OvertimeRate = v
        elif kl == "cost_per_use":
            r.CostPerUse = float(v)
        elif kl == "initials":
            r.Initials = v
        elif kl == "group":
            r.Group = v
        elif kl == "email":
            r.EMailAddress = v
        elif kl == "material_label":
            r.MaterialLabel = v
        elif kl == "notes":
            r.Notes = v
        elif kl == "calendar":
            r.BaseCalendar = v
        else:
            r.SetField(a.FieldNameToFieldConstant(k, C("pjResource")), str(v))


@pj_tool(undo="agregar recurso")
def add_resource(
    name: str,
    type: str = "work",
    max_units_percent: float | None = None,
    standard_rate: str | float | None = None,
    overtime_rate: str | float | None = None,
    cost_per_use: float | None = None,
    group: str | None = None,
    initials: str | None = None,
    email: str | None = None,
    material_label: str | None = None,
) -> dict:
    """Agrega un recurso. type: work (personas/equipos), material, cost.

    standard_rate: número (por hora) o texto con unidad ("400/d", "50/h"). max_units_percent: 100 = jornada completa.
    """
    p = proj()
    r = p.Resources.Add(name)
    f = {k: v for k, v in dict(type=type, max_units_percent=max_units_percent, standard_rate=standard_rate,
                                overtime_rate=overtime_rate, cost_per_use=cost_per_use, group=group,
                                initials=initials, email=email, material_label=material_label).items()
         if v is not None}
    _apply_resource_fields(p, r, f)
    return _resource_dict(p, r)


@pj_tool(undo="actualizar recurso")
def update_resource(resource: str, fields: dict[str, Any]) -> dict:
    """Modifica un recurso. Claves: name, type, max_units_percent, standard_rate, overtime_rate, cost_per_use,
    initials, group, email, material_label, notes, calendar, u otro campo de Project por nombre (Text1...)."""
    p = proj()
    r = _resource(p, resource)
    _apply_resource_fields(p, r, fields)
    return _resource_dict(p, r)


@pj_tool(undo="eliminar recurso")
def delete_resource(resource: str) -> dict:
    """Elimina un recurso y todas sus asignaciones."""
    p = proj()
    r = _resource(p, resource)
    name = r.Name
    r.Delete()
    return {"deleted": name}


@pj_tool(undo="asignar recurso")
def assign_resource(task: str, resource: str, units_percent: float | None = None, work: str | None = None,
                    create_missing_resource: bool = True) -> dict:
    """Asigna un recurso a una tarea. units_percent: 50 = media jornada. work: trabajo total ("40h")."""
    p = proj()
    t = _task(p, task)
    spec = {"resource": resource, "units": units_percent}
    a = _assign(p, t, spec, create_missing_resource)
    if work:
        a.Work = _dur_in(p, work if not isinstance(work, (int, float)) else f"{work}h")
    return _task_full(p, t)


@pj_tool(undo="quitar asignación")
def unassign_resource(task: str, resource: str) -> dict:
    """Quita la asignación de un recurso en una tarea."""
    p = proj()
    t = _task(p, task)
    for a in t.Assignments:
        if a.ResourceName.lower() == resource.lower():
            a.Delete()
            return {"removed": f"{resource} de {t.Name}"}
    raise ProjectError(f"{resource!r} no está asignado a {t.Name!r}")


@pj_tool
def list_assignments(resource: str | None = None, task: str | None = None) -> dict:
    """Lista asignaciones (de un recurso, de una tarea o todas) con unidades, trabajo, costo y fechas."""
    p = proj()
    if task:
        source = _task(p, task).Assignments
    elif resource:
        source = _resource(p, resource).Assignments
    else:
        source = [a for t in _iter_tasks(p) for a in t.Assignments]
    out = []
    for a in source:
        out.append({
            "task_id": a.TaskID,
            "task": a.TaskName,
            "resource": a.ResourceName,
            "units_percent": _num(float(a.Units) * 100),
            "work_hours": _hours(a.Work),
            "actual_work_hours": _hours(a.ActualWork),
            "cost": _num(a.Cost),
            "start": _date_out(a.Start),
            "finish": _date_out(a.Finish),
            "overallocated": bool(a.Overallocated),
        })
    return {"count": len(out), "assignments": out}


@pj_tool(undo="redistribuir recursos")
def level_resources(clear: bool = False) -> dict:
    """Redistribuye (nivela) los recursos sobreasignados de todo el proyecto. clear=True elimina la redistribución."""
    a = app()
    p = proj()
    if clear:
        a.LevelingClear(All=True)
    else:
        a.LevelNow(All=True)
    return {
        "overallocated_after": [r.Name for r in _iter_resources(p) if r.Overallocated],
        "project": _summary(p),
    }


# ---------------------------------------------------------------------------
# Calendarios
# ---------------------------------------------------------------------------

@pj_tool
def list_calendars() -> dict:
    """Lista los calendarios base, el calendario del proyecto y sus excepciones (feriados)."""
    p = proj()
    out = []
    for cal in p.BaseCalendars:
        exc = []
        try:
            for e in cal.Exceptions:
                exc.append({"name": e.Name, "start": _date_out(e.Start), "finish": _date_out(e.Finish)})
        except Exception:
            pass
        days = {}
        names = ["dom", "lun", "mar", "mié", "jue", "vie", "sáb"]
        for i in range(1, 8):
            wd = cal.WeekDays(i)
            shifts = []
            if wd.Working:
                for s in range(1, 6):
                    sh = getattr(wd, f"Shift{s}")
                    if sh.Start not in (None, "", "NA") and _time_str(sh.Start):
                        shifts.append(f"{_time_str(sh.Start)}-{_time_str(sh.Finish)}")
            days[names[i - 1]] = shifts
        out.append({"name": cal.Name, "is_project_calendar": cal.Name == p.Calendar.Name,
                    "work_week": days, "exceptions": exc})
    return {"calendars": out}


def _time_str(v) -> str | None:
    if isinstance(v, datetime):
        return v.strftime("%H:%M")
    if isinstance(v, str) and v and v not in ("NA", "ND"):
        return v
    return None


@pj_tool(undo="agregar feriados")
def add_calendar_exceptions(exceptions: list[dict[str, str]], calendar: str | None = None) -> dict:
    """Agrega días no laborables (feriados, vacaciones) a un calendario (por defecto el del proyecto).

    exceptions: [{"name": "Fiestas Patrias", "start": "2026-09-18", "finish": "2026-09-19"}, ...]
    """
    p = proj()
    cal = p.BaseCalendars(calendar) if calendar else p.Calendar
    added = []
    for e in exceptions:
        start = _date_in(e["start"])
        finish = _date_in(e.get("finish") or e["start"])
        cal.Exceptions.Add(Type=C("pjDaily"), Start=start, Finish=finish, Name=e.get("name", "Excepción"))
        added.append(e.get("name", e["start"]))
    return {"calendar": cal.Name, "added": added, "project_finish": _summary(p)["finish"]}


@pj_tool(undo="semana laboral")
def set_work_week(
    days: list[str],
    shifts: list[list[str]] | None = None,
    calendar: str | None = None,
) -> dict:
    """Define la jornada para días de la semana en un calendario base.

    days: ["lun","mar","mié","jue","vie"] (o mon..sun). shifts: [["08:00","13:00"],["14:00","18:00"]];
    lista vacía o null = día no laborable.
    """
    a = app()
    p = proj()
    cal_name = calendar or p.Calendar.Name
    idx = {"dom": 1, "sun": 1, "lun": 2, "mon": 2, "mar": 3, "tue": 3, "mié": 4, "mie": 4, "wed": 4,
           "jue": 5, "thu": 5, "vie": 6, "fri": 6, "sáb": 7, "sab": 7, "sat": 7}
    shifts = shifts or []
    if len(shifts) > 5:
        raise ProjectError("Máximo 5 turnos por día")
    for d in days:
        kw: dict[str, Any] = {"Name": cal_name, "WeekDay": idx[d.lower()[:3]], "Working": bool(shifts)}
        if shifts:
            for i in range(5):
                f, t = (shifts[i] if i < len(shifts) else ("", ""))
                kw[f"From{i + 1}"] = f
                kw[f"To{i + 1}"] = t
        a.BaseCalendarEditDays(**kw)
    return {"calendar": cal_name, "days": days, "shifts": shifts}


# ---------------------------------------------------------------------------
# Análisis
# ---------------------------------------------------------------------------

@pj_tool
def get_critical_path() -> dict:
    """Tareas de la ruta crítica (sin resúmenes) ordenadas por inicio, con su holgura."""
    p = proj()
    crit = [t for t in _iter_tasks(p) if t.Critical and not t.Summary]
    crit.sort(key=lambda t: (t.Start if isinstance(t.Start, datetime) else datetime.max).replace(tzinfo=None))
    return {
        "project_finish": _date_out(p.ProjectSummaryTask.Finish),
        "count": len(crit),
        "tasks": [{**_task_brief(p, t), "total_slack_days": _days(p, t.TotalSlack)} for t in crit],
    }


@pj_tool
def check_schedule_health() -> dict:
    """Diagnóstico de calidad del cronograma (estilo DCMA): tareas sin predecesoras/sucesoras, restricciones
    rígidas, holgura negativa, fechas límite incumplidas, tareas manuales, duraciones largas, resúmenes con
    recursos, tareas sin recursos, sobreasignaciones y tareas atrasadas."""
    p = proj()
    status = p.StatusDate if isinstance(p.StatusDate, datetime) else datetime.now()
    status = status.replace(tzinfo=None)
    hard = {C(x) for x in ("pjMSO", "pjMFO", "pjSNLT", "pjFNLT")}
    soft = {C(x) for x in ("pjSNET", "pjFNET")}
    issues: dict[str, list] = {k: [] for k in (
        "no_predecessors", "no_successors", "hard_constraints", "soft_constraints", "negative_slack",
        "missed_deadlines", "manual_tasks", "long_durations", "summary_with_resources",
        "tasks_without_resources", "overallocated_tasks", "late_tasks", "lags_negative",
    )}
    leaf = [t for t in _iter_tasks(p) if not t.Summary]
    for t in _iter_tasks(p):
        ref = {"id": t.ID, "name": t.Name}
        if t.Summary:
            if t.Assignments.Count:
                issues["summary_with_resources"].append(ref)
            continue
        preds = [d for d in t.TaskDependencies if d.To.UniqueID == t.UniqueID]
        succs = [d for d in t.TaskDependencies if d.From.UniqueID == t.UniqueID]
        if not preds:
            issues["no_predecessors"].append(ref)
        if not succs:
            issues["no_successors"].append(ref)
        if any(float(d.Lag) < 0 for d in preds):
            issues["lags_negative"].append(ref)
        if t.ConstraintType in hard:
            issues["hard_constraints"].append({**ref, "type": _const_name(_CONSTRAINTS, t.ConstraintType)})
        elif t.ConstraintType in soft:
            issues["soft_constraints"].append({**ref, "type": _const_name(_CONSTRAINTS, t.ConstraintType)})
        try:
            if float(t.TotalSlack) < 0:
                issues["negative_slack"].append({**ref, "slack_days": _days(p, t.TotalSlack)})
        except (TypeError, ValueError):
            pass
        if isinstance(t.Deadline, datetime) and isinstance(t.Finish, datetime) and t.Finish > t.Deadline:
            issues["missed_deadlines"].append({**ref, "deadline": _date_out(t.Deadline), "finish": _date_out(t.Finish)})
        if getattr(t, "Manual", False):
            issues["manual_tasks"].append(ref)
        if (_days(p, t.Duration) or 0) > 44:
            issues["long_durations"].append({**ref, "duration_days": _days(p, t.Duration)})
        if not t.Milestone and t.Assignments.Count == 0:
            issues["tasks_without_resources"].append(ref)
        if t.Overallocated:
            issues["overallocated_tasks"].append(ref)
        if t.PercentComplete < 100 and isinstance(t.Finish, datetime) and t.Finish.replace(tzinfo=None) < status:
            issues["late_tasks"].append({**ref, "finish": _date_out(t.Finish), "percent_complete": t.PercentComplete})
    # la primera y la última tarea es normal que no tengan predecesora / sucesora
    n = len(leaf) or 1
    return {
        "leaf_tasks": len(leaf),
        "status_date": status.strftime("%Y-%m-%d"),
        "overallocated_resources": [r.Name for r in _iter_resources(p) if r.Overallocated],
        "summary": {k: {"count": len(v), "pct": round(100 * len(v) / n, 1)} for k, v in issues.items()},
        "details": {k: v[:50] for k, v in issues.items() if v},
    }


@pj_tool(undo="línea base")
def set_baseline(number: int = 0, tasks: list[str] | None = None) -> dict:
    """Guarda la línea base (0 = Línea base, 1-10 = Línea base N) para todo el proyecto o para tareas específicas."""
    a = app()
    p = proj()
    into = C("pjIntoBaseline") if number == 0 else C(f"pjIntoBaseline{number}")
    if tasks:
        for x in tasks:
            a.EditGoTo(ID=_task(p, x).ID)  # selecciona la tarea por Id, independiente de filtros/orden
            a.BaselineSave(All=False, Copy=C("pjCopyCurrent"), Into=into)
    else:
        a.BaselineSave(All=True, Copy=C("pjCopyCurrent"), Into=into)
    return {"baseline": number, "scope": tasks or "proyecto completo", "project": _summary(p)}


@pj_tool(undo="borrar línea base")
def clear_baseline(number: int = 0) -> dict:
    """Borra la línea base indicada (0 = Línea base, 1-10) de todo el proyecto."""
    a = app()
    p = proj()
    frm = C("pjIntoBaseline") if number == 0 else C(f"pjIntoBaseline{number}")
    a.BaselineClear(All=True, From=frm)
    return {"cleared": number}


@pj_tool
def get_baseline_variance(number: int = 0, only_slipped: bool = True, threshold_days: float = 0) -> dict:
    """Compara el cronograma actual con la línea base N: variación de inicio, fin, duración, trabajo y costo."""
    p = proj()
    pre = "Baseline" if number == 0 else f"Baseline{number}"
    out = []
    for t in _iter_tasks(p):
        bs, bf = getattr(t, f"{pre}Start"), getattr(t, f"{pre}Finish")
        if not isinstance(bf, datetime):
            continue
        fv = (t.Finish.replace(tzinfo=None) - bf.replace(tzinfo=None)) if isinstance(t.Finish, datetime) else None
        fv_days = round(fv.total_seconds() / 86400, 2) if fv is not None else None
        if only_slipped and (fv_days is None or fv_days <= threshold_days):
            continue
        out.append({
            "id": t.ID, "name": t.Name, "summary": bool(t.Summary),
            "baseline_start": _date_out(bs), "start": _date_out(t.Start),
            "baseline_finish": _date_out(bf), "finish": _date_out(t.Finish),
            "finish_slip_calendar_days": fv_days,
            "baseline_duration_days": _days(p, getattr(t, f"{pre}Duration")),
            "duration_days": _days(p, t.Duration),
            "baseline_work_hours": _hours(getattr(t, f"{pre}Work")),
            "work_hours": _hours(t.Work),
            "baseline_cost": _num(getattr(t, f"{pre}Cost")),
            "cost": _num(t.Cost),
        })
    if not out and not any(isinstance(getattr(t, f"{pre}Finish"), datetime) for t in _iter_tasks(p)):
        raise ProjectError(f"El proyecto no tiene guardada la línea base {number}. Use set_baseline.")
    return {"baseline": number, "count": len(out), "tasks": out}


@pj_tool
def get_timephased_data(
    target: str = "project",
    data: str = "work",
    unit: str = "weeks",
    start: str | None = None,
    finish: str | None = None,
) -> dict:
    """Datos en fase temporal (distribución en el tiempo), útiles para curvas S e histogramas de recursos.

    target: "project", "task:<ref>" o "resource:<nombre>".
    data: work, cost, actual_work, baseline_work (tareas/proyecto) | work, cost, overallocation, peak_units (recursos).
    unit: days, weeks, months, quarters, years. Trabajo en horas.
    """
    p = proj()
    unit_c = C({"days": "pjTimescaleDays", "weeks": "pjTimescaleWeeks", "months": "pjTimescaleMonths",
                "quarters": "pjTimescaleQuarters", "years": "pjTimescaleYears"}[unit])
    if target.startswith("resource:"):
        obj = _resource(p, target.split(":", 1)[1])
        typ = C({"work": "pjResourceTimescaledWork", "cost": "pjResourceTimescaledCost",
                 "overallocation": "pjResourceTimescaledOverallocation",
                 "peak_units": "pjResourceTimescaledPeakUnits"}[data])
    else:
        obj = p.ProjectSummaryTask if target == "project" else _task(p, target.split(":", 1)[1])
        typ = C({"work": "pjTaskTimescaledWork", "cost": "pjTaskTimescaledCost",
                 "actual_work": "pjTaskTimescaledActualWork", "baseline_work": "pjTaskTimescaledBaselineWork"}[data])
    s = _date_in(start) if start else p.ProjectSummaryTask.Start
    f = _date_in(finish, "finish") if finish else p.ProjectSummaryTask.Finish
    vals = obj.TimeScaleData(StartDate=s, EndDate=f, Type=typ, TimeScaleUnit=unit_c, Count=1)
    series, cumulative = [], 0.0
    for v in vals:
        raw = v.Value
        try:
            num = float(raw)
        except (TypeError, ValueError):
            num = 0.0
        if data in ("work", "actual_work", "baseline_work", "overallocation"):
            num = num / 60
        elif data == "peak_units":
            num = num * 100
        cumulative += num
        series.append({"start": _date_out(v.StartDate), "value": round(num, 2), "cumulative": round(cumulative, 2)})
    return {"target": target, "data": data, "unit": unit, "series": series}


# ---------------------------------------------------------------------------
# Vistas, exportación y utilidades
# ---------------------------------------------------------------------------

@pj_tool
def list_views() -> dict:
    """Lista las vistas disponibles y la vista actual."""
    a = app()
    p = proj()
    views = [v.Name for v in p.Views]
    current = p.CurrentView
    return {"current": current, "views": views}


@pj_tool
def apply_view(name: str) -> dict:
    """Cambia la vista activa (ej. "Diagrama de Gantt", "Gantt de seguimiento", "Uso de recursos", "Calendario")."""
    a = app()
    proj()
    a.ViewApplyEx(Name=name)
    return {"view": a.ActiveProject.CurrentView}


@pj_tool
def export_pdf(path: str, start: str | None = None, finish: str | None = None) -> dict:
    """Exporta la vista actual a PDF (opcionalmente solo un rango de fechas)."""
    a = app()
    proj()
    kw: dict[str, Any] = {"Filename": path, "FileType": C("pjPDF")}
    if start:
        kw["FromDate"] = _date_in(start)
    if finish:
        kw["ToDate"] = _date_in(finish, "finish")
    a.DocumentExport(**kw)
    return {"exported": path}


@pj_tool
def export_tasks_csv(path: str, extra_fields: list[str] | None = None, delimiter: str = ";") -> dict:
    """Exporta la tabla de tareas a CSV (UTF-8, abre directo en Excel). delimiter ';' para Excel en español."""
    p = proj()
    rows = []
    for t in _iter_tasks(p):
        d = _task_brief(p, t)
        d["work_hours"] = _hours(t.Work)
        d["cost"] = _num(t.Cost)
        d["baseline_finish"] = _date_out(t.BaselineFinish)
        d.update(_extra_fields(t, extra_fields))
        rows.append(d)
    if not rows:
        raise ProjectError("El proyecto no tiene tareas")
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        wr = csv.DictWriter(fh, fieldnames=list(rows[0].keys()), delimiter=delimiter)
        wr.writeheader()
        wr.writerows(rows)
    return {"exported": path, "rows": len(rows)}


@pj_tool(undo="renombrar campo")
def rename_custom_field(field: str, alias: str, entity: str = "task") -> dict:
    """Pone un nombre visible (alias) a un campo personalizado, ej. Text1 -> "Contratista"."""
    a = app()
    proj()
    fid = a.FieldNameToFieldConstant(field, C("pjTask" if entity == "task" else "pjResource"))
    a.CustomFieldRename(FieldID=fid, NewName=alias)
    return {"field": field, "alias": alias}


@pj_tool
def undo(levels: int = 1) -> dict:
    """Deshace las últimas operaciones en Project (cada tool de edición cuenta como una)."""
    a = app()
    proj()
    for _ in range(levels):
        a.Undo()
    return {"undone": levels}


@pj_tool
def recalculate() -> dict:
    """Fuerza el recálculo del cronograma (F9)."""
    a = app()
    p = proj()
    a.CalculateProject()
    return _summary(p)


def main():
    mcp.run("stdio")


if __name__ == "__main__":
    main()
