import json
import os
from typing import List, Dict, Optional

CUENTAS_FILE = "cuentas.json"

Cuenta = Dict[str, object]

def cargar_cuentas() -> List[Cuenta]:
    if os.path.exists(CUENTAS_FILE):
        try:
            with open(CUENTAS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                return data
        except json.JSONDecodeError:
            pass
    return []

def guardar_cuentas(cuentas: List[Cuenta]) -> None:
    with open(CUENTAS_FILE, "w", encoding="utf-8") as f:
        json.dump(cuentas, f, indent=4, ensure_ascii=False)

def _display(c: Cuenta) -> str:
    codigo = str(c.get("codigo", "") or "").strip()
    nombre = str(c.get("nombre", "") or "").strip()
    if codigo and nombre:
        return f"{codigo} - {nombre}"
    return nombre or codigo

def cuentas_activas(cuentas: List[Cuenta]) -> List[Cuenta]:
    out = []
    for c in cuentas:
        if bool(c.get("activa", True)):
            out.append(c)
    return out

def opciones_select(cuentas: List[Cuenta]) -> List[str]:
    # devuelve strings display para selectbox
    return [_display(c) for c in cuentas_activas(cuentas) if _display(c)]

def nombre_desde_display(display: str) -> str:
    # si viene "XXXX - Nombre", devuelve "Nombre"
    if " - " in display:
        return display.split(" - ", 1)[1].strip()
    return display.strip()

def asegurar_cuentas_por_defecto(nombres: List[str]) -> None:
    """Si no existe cuentas.json, lo crea con una lista mínima.
    nombres: nombres de cuentas detectadas en el sistema."""
    if os.path.exists(CUENTAS_FILE):
        return
    cuentas = []
    seen = set()
    # IDs simples C0001...
    n = 1
    for name in nombres:
        nm = str(name or "").strip()
        if not nm:
            continue
        key = nm.lower()
        if key in seen:
            continue
        cuentas.append({
            "codigo": f"C{str(n).zfill(4)}",
            "nombre": nm,
            "tipo": "",          # activo/pasivo/patrimonio/ingreso/gasto (opcional por ahora)
            "naturaleza": "",    # deudora/acreedora (opcional por ahora)
            "activa": True
        })
        seen.add(key)
        n += 1
    guardar_cuentas(cuentas)


def opciones_nombres(cuentas: List[Cuenta]) -> List[str]:
    return [str(c.get('nombre','') or '').strip() for c in cuentas_activas(cuentas) if str(c.get('nombre','') or '').strip()]
