import streamlit as st
import pandas as pd
import json
import os
import io
import calendar
import re
import sqlite3
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
import shutil
import tempfile
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent


def resolver_ruta(ruta):
    ruta = str(ruta)
    if os.path.isabs(ruta):
        return ruta
    return str(BASE_DIR / ruta)


def _json_default(obj):
    if isinstance(obj, pd.Timestamp):
        return str(obj)

    if hasattr(obj, "item"):
        try:
            return obj.item()
        except Exception:
            pass

    raise TypeError(f"Tipo no serializable a JSON: {type(obj).__name__}")


def guardar_json_seguro(ruta, data, crear_backup=True):
    ruta_real = resolver_ruta(ruta)
    carpeta = os.path.dirname(ruta_real) or "."
    os.makedirs(carpeta, exist_ok=True)

    fd, ruta_tmp = tempfile.mkstemp(
        prefix=".tmp_",
        suffix=".json",
        dir=carpeta
    )

    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(
                data,
                f,
                ensure_ascii=False,
                indent=4,
                default=_json_default
            )
            f.flush()
            os.fsync(f.fileno())

        if crear_backup and os.path.exists(ruta_real):
            shutil.copy2(ruta_real, ruta_real + ".bak")

        os.replace(ruta_tmp, ruta_real)

    except Exception:
        if os.path.exists(ruta_tmp):
            os.remove(ruta_tmp)
        raise


def cargar_json_lista(ruta):
    ruta_real = resolver_ruta(ruta)

    if not os.path.exists(ruta_real):
        return []

    try:
        with open(ruta_real, "r", encoding="utf-8") as f:
            data = json.load(f)
    except json.JSONDecodeError:
        st.error(
            f"El archivo {os.path.basename(ruta_real)} está dañado o incompleto. "
            f"No se cargó para evitar sobrescribir datos. "
            f"Revisa el respaldo .bak si existe."
        )
        st.stop()

    if not isinstance(data, list):
        st.error(
            f"El archivo {os.path.basename(ruta_real)} no tiene formato de lista válido."
        )
        st.stop()

    return data


def cargar_json_dict(ruta):
    ruta_real = resolver_ruta(ruta)

    if not os.path.exists(ruta_real):
        return {}

    try:
        with open(ruta_real, "r", encoding="utf-8") as f:
            data = json.load(f)
    except json.JSONDecodeError:
        st.error(
            f"El archivo {os.path.basename(ruta_real)} está dañado o incompleto. "
            f"No se cargó para evitar sobrescribir datos. "
            f"Revisa el respaldo .bak si existe."
        )
        st.stop()

    if not isinstance(data, dict):
        st.error(
            f"El archivo {os.path.basename(ruta_real)} no tiene formato de diccionario válido."
        )
        st.stop()

    return data


def guardar_json_dict(ruta, data):
    guardar_json_seguro(ruta, data)

st.set_page_config(
    page_title="DIARIO",
    page_icon="🧾",
    layout="wide"
)

VENTAS_FILE = "partidas.json"
COMPRAS_FILE = "partidascompras.json"
EMPRESAS_FILE = "empresas.json"
CONFIG_UI_FILE = "config_ui_partidas.json"
CUENTAS_FILE = "cuentas.json"
SALARIOS_MINIMOS_FILE = "salarios_minimos.json"
PARTIDAS_GENERADAS_FILE = "partidas_generadas.json"
ISR_TRIMESTRAL_FILE = "isr_trimestral_por_pagar.json"
SQLITE_BACKUP_FILE = "contabilidad.db"
CONFIG_UI_DB_FILE = "contabilidad.db"
PARTIDAS_SQLITE_PENDIENTES_FILE = "partidas_generadas_sqlite_pendientes.json"
BONIFICACION_INCENTIVO = 250.00
PORC_IGSS_LABORAL = 0.0483
PORC_IGSS_PATRONAL = 0.1067
PORC_IRTRA = 0.01
PORC_INTECAP = 0.01
PORC_PRESTACIONES_DEFAULT = 0.2638

MESES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}
MESES_INV = {v.lower(): k for k, v in MESES.items()}
CUENTAS_CAJA = {"caja", "bancos", "banco", "banco 1", "banco 2", "banco 3"}
CUENTAS_CXC = {"clientes", "cuentas por cobrar", "cxc"}


def aplicar_estilos():
    st.markdown("""
    <style>
    .block-container {
        padding-top: 2.6rem;
        padding-bottom: 1rem;
        max-width: 100%;
    }

    .panel {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 16px;
        padding: 18px 18px 16px 18px;
        margin-bottom: 14px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.04);
    }

    .panel-title {
        font-size: 1.05rem;
        font-weight: 700;
        margin-bottom: 12px;
        color: #111827;
    }

    .kpi-box {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 14px;
        padding: 14px 16px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        margin-bottom: 10px;
        min-height: 84px;
    }

    .kpi-title {
        font-size: 12px;
        color: #6b7280;
        margin-bottom: 6px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: .3px;
    }

    .kpi-value {
        font-size: 18px;
        font-weight: 700;
        color: #111827;
        line-height: 1.25;
    }

    .small-muted {
        color: #6b7280;
        font-size: 13px;
    }

    div[data-testid="stExpander"] {
        border: 1px solid #e5e7eb !important;
        border-radius: 12px !important;
        overflow: hidden;
    }

    div[data-testid="stExpander"] summary {
        font-weight: 600;
    }

    .resumen-vacio {
        padding: 18px;
        border: 1px dashed #d1d5db;
        border-radius: 12px;
        color: #6b7280;
        background: #fafafa;
    }

    .header-diario {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 0.4rem;
    }

    .header-diario-titulo {
        font-size: 1.6rem;
        font-weight: 700;
        color: #111827;
        line-height: 1.2;
    }

    .header-diario-regimen {
        text-align: right;
        font-size: 0.95rem;
        font-weight: 700;
        color: #2563eb;
        background: #eff6ff;
        border: 1px solid #bfdbfe;
        padding: 6px 12px;
        border-radius: 999px;
        display: inline-block;
    }
    
    </style>
    """, unsafe_allow_html=True)

    st.markdown("""
<style>
/* ===== Separadores compactos tipo ERP ===== */
.seccion-mini {
    margin: 8px 0 10px 0;
}

.seccion-mini-titulo {
    font-size: 12px;
    font-weight: 800;
    text-transform: uppercase;
    letter-spacing: .5px;
    margin-bottom: 6px;
    display: flex;
    align-items: center;
    gap: 8px;
}

.seccion-mini-linea {
    height: 2px;
    border-radius: 999px;
    width: 100%;
    opacity: 0.95;
}

.seccion-general .seccion-mini-titulo {
    color: #475569;
}
.seccion-general .seccion-mini-linea {
    background: linear-gradient(90deg, #94a3b8, #e2e8f0);
}

.seccion-ingresos .seccion-mini-titulo {
    color: #15803d;
}
.seccion-ingresos .seccion-mini-linea {
    background: linear-gradient(90deg, #16a34a, #bbf7d0);
}

.seccion-egresos .seccion-mini-titulo {
    color: #c2410c;
}
.seccion-egresos .seccion-mini-linea {
    background: linear-gradient(90deg, #ea580c, #fed7aa);
}

/* ===== Expander con menos peso visual ===== */
div[data-testid="stExpander"] {
    border: 1px solid #e5e7eb !important;
    border-radius: 12px !important;
    overflow: hidden;
    margin-bottom: 8px !important;
}

div[data-testid="stExpander"] summary {
    padding-top: 0.35rem !important;
    padding-bottom: 0.35rem !important;
    font-weight: 600;
}

/* Compactar algunos controles */
div[data-testid="stNumberInput"],
div[data-testid="stSelectbox"],
div[data-testid="stTextInput"] {
    margin-bottom: 0.2rem;
}
</style>
""", unsafe_allow_html=True)


def mostrar_resumen_superior(empresa, regimen, anio, mes, activar_distribucion, cantidad_cuentas):
    c1, c2, c3, c4, c5 = st.columns(5)

def cargar_isr_trimestral_guardado():
    data = cargar_json_lista(ISR_TRIMESTRAL_FILE)
    salida = []

    for item in data:
        if not isinstance(item, dict):
            continue

        salida.append({
            "empresa": str(item.get("empresa", "") or "").strip(),
            "empresa_nit": normalizar_nit(item.get("empresa_nit")),
            "anio": int(item.get("anio", 0) or 0),
            "trimestre": int(item.get("trimestre", 0) or 0),
            "mes_corte": int(item.get("mes_corte", 0) or 0),
            "isr_trimestral_por_pagar": round(float(item.get("isr_trimestral_por_pagar", 0) or 0), 2),
            "isr_acumulado_hasta_corte": round(float(item.get("isr_acumulado_hasta_corte", 0) or 0), 2),
            "isr_trimestres_anteriores": round(float(item.get("isr_trimestres_anteriores", 0) or 0), 2),
            "fecha_guardado": str(item.get("fecha_guardado", "") or "").strip(),
            "detalle_acreditado": item.get("detalle_acreditado", []) or [],
        })

    return salida


def trimestre_desde_mes(mes):
    mes = int(mes or 0)
    if mes <= 3:
        return 1
    if mes <= 6:
        return 2
    if mes <= 9:
        return 3
    return 4


def buscar_isr_trimestral_guardado(data, empresa_id, anio, trimestre):
    empresa_id = normalizar_nit(empresa_id)
    trimestre = int(trimestre or 0)

    for item in data or []:
        if (
            normalizar_nit(item.get("empresa_nit")) == empresa_id
            and int(item.get("anio", 0) or 0) == int(anio)
            and int(item.get("trimestre", 0) or 0) == trimestre
        ):
            return item

    return None


def obtener_isr_trimestral_desde_cierres(empresa_id, anio, mes):
    trimestre = trimestre_desde_mes(mes)
    data = cargar_isr_trimestral_guardado()
    return buscar_isr_trimestral_guardado(data, empresa_id, anio, trimestre)


def log_sqlite_respaldo(msg):
    try:
        print(f"[sqlite-respaldo] {msg}")
    except Exception:
        pass


def sqlite_resolver_ruta_db(ruta=SQLITE_BACKUP_FILE):
    return resolver_ruta(ruta)


def sqlite_resolver_ruta_pendientes(ruta=PARTIDAS_SQLITE_PENDIENTES_FILE):
    return resolver_ruta(ruta)


def clave_periodo_partida_generada(empresa_nit, anio, mes):
    return f"{normalizar_nit(empresa_nit)}__{int(anio)}__{int(mes)}"


def sqlite_respaldo_partidas_conn():
    ruta_db = sqlite_resolver_ruta_db(SQLITE_BACKUP_FILE)
    carpeta = os.path.dirname(ruta_db) or "."
    os.makedirs(carpeta, exist_ok=True)

    conn = sqlite3.connect(ruta_db)
    conn.row_factory = sqlite3.Row
    return conn


def cargar_pendientes_partidas_sqlite():
    ruta = sqlite_resolver_ruta_pendientes(PARTIDAS_SQLITE_PENDIENTES_FILE)
    if not os.path.exists(ruta):
        return {}

    try:
        with open(ruta, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception as e:
        log_sqlite_respaldo(f"No se pudieron cargar los pendientes de SQLite: {e}")
        return {}


def guardar_pendientes_partidas_sqlite(data):
    try:
        guardar_json_seguro(PARTIDAS_SQLITE_PENDIENTES_FILE, data or {}, crear_backup=False)
    except Exception as e:
        log_sqlite_respaldo(f"No se pudieron guardar los pendientes de SQLite: {e}")


def marcar_periodo_pendiente_sqlite(registro, motivo=""):
    try:
        empresa_nit = normalizar_nit(registro.get("empresa_nit"))
        anio = int(registro.get("anio", 0) or 0)
        mes = int(registro.get("mes", 0) or 0)
        if not empresa_nit or anio <= 0 or mes <= 0:
            return False

        key = clave_periodo_partida_generada(empresa_nit, anio, mes)
        pendientes = cargar_pendientes_partidas_sqlite()
        pendientes[key] = {
            "empresa": str(registro.get("empresa", "") or "").strip(),
            "empresa_nit": empresa_nit,
            "anio": anio,
            "mes": mes,
            "motivo": str(motivo or "").strip(),
            "actualizado_en": str(pd.Timestamp.now()),
        }
        guardar_pendientes_partidas_sqlite(pendientes)
        return True
    except Exception as e:
        log_sqlite_respaldo(f"No se pudo marcar como pendiente el período en SQLite: {e}")
        return False


def quitar_periodo_pendiente_sqlite(empresa_nit, anio, mes):
    try:
        key = clave_periodo_partida_generada(empresa_nit, anio, mes)
        pendientes = cargar_pendientes_partidas_sqlite()
        if key in pendientes:
            pendientes.pop(key, None)
            guardar_pendientes_partidas_sqlite(pendientes)
        return True
    except Exception as e:
        log_sqlite_respaldo(f"No se pudo limpiar un período pendiente de SQLite: {e}")
        return False


def contar_pendientes_partidas_sqlite():
    return len(cargar_pendientes_partidas_sqlite())


def init_sqlite_respaldo_partidas_generadas():
    try:
        with sqlite_respaldo_partidas_conn() as conn:
            conn.execute("PRAGMA journal_mode=WAL;")
            conn.execute("""
                CREATE TABLE IF NOT EXISTS partidas_generadas_respaldo (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    empresa_nit TEXT NOT NULL,
                    empresa TEXT,
                    anio INTEGER NOT NULL,
                    mes INTEGER NOT NULL,
                    fecha_guardado TEXT,
                    origen TEXT,
                    partidas_json TEXT NOT NULL,
                    actualizado_en TEXT NOT NULL,
                    UNIQUE (empresa_nit, anio, mes)
                )
            """)
            conn.execute("""
                CREATE INDEX IF NOT EXISTS idx_partidas_generadas_respaldo_empresa_periodo
                ON partidas_generadas_respaldo (empresa_nit, anio, mes)
            """)
            conn.commit()
    except Exception as e:
        log_sqlite_respaldo(f"No se pudo inicializar SQLite de respaldo: {e}")


def upsert_partida_generada_en_sqlite(registro):
    if not isinstance(registro, dict):
        return False

    try:
        empresa_nit = normalizar_nit(registro.get("empresa_nit"))
        anio = int(registro.get("anio", 0) or 0)
        mes = int(registro.get("mes", 0) or 0)
        partidas = registro.get("partidas", []) or []

        if not empresa_nit or anio <= 0 or mes <= 0:
            return False

        partidas_json = json.dumps(
            partidas,
            ensure_ascii=False,
            default=_json_default
        )
        actualizado_en = str(pd.Timestamp.now())

        with sqlite_respaldo_partidas_conn() as conn:
            conn.execute("""
                INSERT INTO partidas_generadas_respaldo (
                    empresa_nit, empresa, anio, mes, fecha_guardado, origen, partidas_json, actualizado_en
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(empresa_nit, anio, mes)
                DO UPDATE SET
                    empresa = excluded.empresa,
                    fecha_guardado = excluded.fecha_guardado,
                    origen = excluded.origen,
                    partidas_json = excluded.partidas_json,
                    actualizado_en = excluded.actualizado_en
            """, (
                empresa_nit,
                str(registro.get("empresa", "") or "").strip(),
                anio,
                mes,
                str(registro.get("fecha_guardado", "") or "").strip(),
                str(registro.get("origen", "generado") or "generado").strip() or "generado",
                partidas_json,
                actualizado_en,
            ))
            conn.commit()

        quitar_periodo_pendiente_sqlite(empresa_nit, anio, mes)
        return True
    except Exception as e:
        log_sqlite_respaldo(f"No se pudo respaldar en SQLite el período {registro.get('empresa_nit')} {registro.get('anio')}-{registro.get('mes')}: {e}")
        return False


def respaldar_periodo_partida_generada_sqlite(registro, motivo_error="guardado"):
    ok = upsert_partida_generada_en_sqlite(registro)
    if not ok:
        marcar_periodo_pendiente_sqlite(registro, motivo=motivo_error)
    return ok


def contar_partidas_generadas_respaldo_sqlite():
    try:
        init_sqlite_respaldo_partidas_generadas()
        with sqlite_respaldo_partidas_conn() as conn:
            row = conn.execute("SELECT COUNT(*) AS total FROM partidas_generadas_respaldo").fetchone()
            return int(row["total"] or 0) if row else 0
    except Exception as e:
        log_sqlite_respaldo(f"No se pudo contar el respaldo SQLite de partidas generadas: {e}")
        return 0


def obtener_claves_partidas_generadas_respaldo_sqlite():
    try:
        init_sqlite_respaldo_partidas_generadas()
        with sqlite_respaldo_partidas_conn() as conn:
            rows = conn.execute("SELECT empresa_nit, anio, mes FROM partidas_generadas_respaldo").fetchall()
        return {
            clave_periodo_partida_generada(row["empresa_nit"], row["anio"], row["mes"])
            for row in (rows or [])
        }
    except Exception as e:
        log_sqlite_respaldo(f"No se pudieron obtener las claves del respaldo SQLite: {e}")
        return set()


def normalizar_registro_partida_generada(item):
    if not isinstance(item, dict):
        return None
    if not isinstance(item.get("partidas"), list):
        return None

    empresa_nit = normalizar_nit(item.get("empresa_nit"))
    anio = int(item.get("anio", 0) or 0)
    mes = int(item.get("mes", 0) or 0)
    if not empresa_nit or anio <= 0 or mes <= 0:
        return None

    return {
        "empresa": str(item.get("empresa", "") or "").strip(),
        "empresa_nit": empresa_nit,
        "anio": anio,
        "mes": mes,
        "partidas": item.get("partidas", []) or [],
        "fecha_guardado": str(item.get("fecha_guardado", "") or "").strip(),
        "origen": str(item.get("origen", "generado") or "generado").strip() or "generado",
    }


def sincronizar_partidas_generadas_faltantes_json_a_sqlite(data=None):
    try:
        init_sqlite_respaldo_partidas_generadas()

        if data is None:
            data = cargar_json_lista(PARTIDAS_GENERADAS_FILE)

        registros = []
        for item in data or []:
            registro = normalizar_registro_partida_generada(item)
            if registro:
                registros.append(registro)

        if not registros:
            return 0

        total_sqlite = contar_partidas_generadas_respaldo_sqlite()
        claves_sqlite = set() if total_sqlite == 0 else obtener_claves_partidas_generadas_respaldo_sqlite()

        total = 0
        for registro in registros:
            key = clave_periodo_partida_generada(
                registro.get("empresa_nit"),
                registro.get("anio"),
                registro.get("mes")
            )
            if total_sqlite > 0 and key in claves_sqlite:
                continue

            if respaldar_periodo_partida_generada_sqlite(registro, motivo_error="sincronizacion_faltante"):
                total += 1

        return total
    except Exception as e:
        log_sqlite_respaldo(f"No se pudo sincronizar faltantes de partidas_generadas.json hacia SQLite: {e}")
        return 0


def reintentar_periodos_pendientes_partidas_sqlite(data=None):
    try:
        pendientes = cargar_pendientes_partidas_sqlite()
        if not pendientes:
            return 0

        if data is None:
            data = cargar_json_lista(PARTIDAS_GENERADAS_FILE)

        registros_por_key = {}
        for item in data or []:
            registro = normalizar_registro_partida_generada(item)
            if not registro:
                continue
            key = clave_periodo_partida_generada(registro.get("empresa_nit"), registro.get("anio"), registro.get("mes"))
            registros_por_key[key] = registro

        total = 0
        for key, meta in list(pendientes.items()):
            registro = registros_por_key.get(key)
            if not registro:
                continue

            if respaldar_periodo_partida_generada_sqlite(registro, motivo_error="reintento_pendiente"):
                total += 1

        return total
    except Exception as e:
        log_sqlite_respaldo(f"No se pudieron reintentar los períodos pendientes hacia SQLite: {e}")
        return 0


def sincronizar_respaldo_partidas_generadas_sqlite(data=None):
    try:
        init_sqlite_respaldo_partidas_generadas()
        faltantes = sincronizar_partidas_generadas_faltantes_json_a_sqlite(data)
        pendientes = reintentar_periodos_pendientes_partidas_sqlite(data)
        return {
            "faltantes": int(faltantes or 0),
            "pendientes_reintentados": int(pendientes or 0),
            "pendientes_restantes": int(contar_pendientes_partidas_sqlite() or 0),
        }
    except Exception as e:
        log_sqlite_respaldo(f"No se pudo sincronizar el respaldo de partidas generadas en SQLite: {e}")
        return {
            "faltantes": 0,
            "pendientes_reintentados": 0,
            "pendientes_restantes": int(contar_pendientes_partidas_sqlite() or 0),
        }


def cargar_partidas_generadas():
    data = cargar_json_lista(PARTIDAS_GENERADAS_FILE)
    salida = []

    for item in data:
        if not isinstance(item, dict):
            continue
        if not isinstance(item.get("partidas"), list):
            continue

        salida.append({
            "empresa": str(item.get("empresa", "")).strip(),
            "empresa_nit": normalizar_nit(item.get("empresa_nit")),
            "anio": int(item.get("anio", 0) or 0),
            "mes": int(item.get("mes", 0) or 0),
            "partidas": item.get("partidas", []),
            "fecha_guardado": str(item.get("fecha_guardado", "")).strip(),
            "origen": str(item.get("origen", "generado")).strip() or "generado"
        })

    return salida


def guardar_partidas_generadas(data):
    guardar_json_seguro(PARTIDAS_GENERADAS_FILE, data)


def buscar_partidas_generadas(data, empresa_id, anio, mes):
    empresa_id = normalizar_nit(empresa_id)

    for item in data:
        if (
            normalizar_nit(item.get("empresa_nit")) == empresa_id
            and int(item.get("anio", 0) or 0) == int(anio)
            and int(item.get("mes", 0) or 0) == int(mes)
        ):
            return item
    return None


def guardar_o_actualizar_partidas_generadas(data, empresa, empresa_id, anio, mes, partidas):
    empresa_id = normalizar_nit(empresa_id)
    anio = int(anio)
    mes = int(mes)

    registro = {
        "empresa": str(empresa).strip(),
        "empresa_nit": empresa_id,
        "anio": anio,
        "mes": mes,
        "partidas": partidas,
        "fecha_guardado": str(pd.Timestamp.now()),
        "origen": "generado"
    }

    actualizado = False
    for i, item in enumerate(data):
        if (
            normalizar_nit(item.get("empresa_nit")) == empresa_id
            and int(item.get("anio", 0) or 0) == anio
            and int(item.get("mes", 0) or 0) == mes
        ):
            data[i] = registro
            actualizado = True
            break

    if not actualizado:
        data.append(registro)

    data.sort(key=lambda x: (
        str(x.get("empresa", "")).lower(),
        int(x.get("anio", 0) or 0),
        int(x.get("mes", 0) or 0)
    ))

    guardar_partidas_generadas(data)
    respaldar_periodo_partida_generada_sqlite(registro, motivo_error="guardado_periodo")
    return registro


def cargar_salarios_minimos():
    return cargar_json_dict(SALARIOS_MINIMOS_FILE)

def obtener_partidas_mes_guardado(partidas_generadas_data, empresa_id, anio, mes):
    registro = buscar_partidas_generadas(
        partidas_generadas_data or [],
        empresa_id,
        int(anio),
        int(mes)
    )
    if not registro:
        return []
    return registro.get("partidas", [])

def obtener_partidas_acumuladas_hasta_mes(partidas_generadas_data, empresa_id, anio, mes):
    empresa_id = normalizar_nit(empresa_id)
    anio = int(anio)
    mes = int(mes)

    registros = []
    for item in partidas_generadas_data or []:
        try:
            item_anio = int(item.get("anio", 0) or 0)
            item_mes = int(item.get("mes", 0) or 0)
        except (TypeError, ValueError):
            continue

        if normalizar_nit(item.get("empresa_nit")) != empresa_id:
            continue
        if item_anio != anio:
            continue
        if item_mes < 1 or item_mes > mes:
            continue

        registros.append(item)

    registros.sort(key=lambda x: (
        int(x.get("anio", 0) or 0),
        int(x.get("mes", 0) or 0),
        str(x.get("fecha_guardado", ""))
    ))

    partidas = []
    for registro in registros:
        for partida in registro.get("partidas", []) or []:
            if es_partida_apertura_guardada(partida):
                continue
            partidas.append(partida)

    return partidas


def es_partida_apertura_guardada(partida):
    glosa = str(partida.get("glosa", "")).strip().lower()
    codigo = str(partida.get("codigo", "")).strip().upper()
    tipo_partida = str(partida.get("tipo_partida", "")).strip().lower()

    if tipo_partida == "apertura":
        return True

    if glosa == "partida de apertura":
        return True

    if codigo in {"PDA 0", "PDA 1"}:
        cuentas = partida.get("cuentas", [])
        if any(str(c.get("cuenta", "")).startswith("A: ") for c in cuentas):
            return True

    return False


def obtener_apertura_y_movimientos_para_mayor_local(
    ventas_data,
    compras_data,
    partidas_generadas_data,
    empresa_id,
    anio,
    mes
):
    apertura = buscar_partida_apertura(ventas_data, empresa_id, anio)
    if not apertura:
        apertura = buscar_partida_apertura(compras_data, empresa_id, anio)

    apertura = normalizar_apertura_contrapartidas(apertura)

    # Para reconstruir saldos reales del mayor, se toma la apertura del año
    # más todas las partidas generadas acumuladas hasta el mes indicado.
    partidas_mes = obtener_partidas_acumuladas_hasta_mes(
        partidas_generadas_data,
        empresa_id,
        int(anio),
        int(mes)
    )

    return apertura, partidas_mes


def obtener_saldos_por_pagar_desde_mayor_mes(
    ventas_data,
    compras_data,
    partidas_generadas_data,
    empresa_id,
    anio,
    mes,
    excluir_cuentas=None
):
    excluir_cuentas = {
        str(x).replace("A: ", "").strip().lower()
        for x in (excluir_cuentas or [])
        if str(x).strip()
    }

    apertura, partidas_mes = obtener_apertura_y_movimientos_para_mayor_local(
        ventas_data=ventas_data,
        compras_data=compras_data,
        partidas_generadas_data=partidas_generadas_data,
        empresa_id=empresa_id,
        anio=int(anio),
        mes=int(mes)
    )

    partidas_fuente = []
    if apertura:
        partidas_fuente.append(apertura)
    partidas_fuente.extend(partidas_mes or [])

    saldos = calcular_saldos_por_pagar(
        partidas_fuente,
        excluir_cuentas=excluir_cuentas
    )

    return saldos

def limpiar_nombre_cuenta(cuenta):
    return str(cuenta or "").replace("A: ", "").strip()
def obtener_partidas_para_saldo_hasta_mes_anterior(
    ventas_data,
    compras_data,
    partidas_generadas_data,
    empresa_id,
    anio,
    mes
):
    apertura = buscar_partida_apertura(ventas_data, empresa_id, anio)
    if not apertura:
        apertura = buscar_partida_apertura(compras_data, empresa_id, anio)

    apertura = normalizar_apertura_contrapartidas(apertura)

    partidas_fuente = []
    if apertura:
        partidas_fuente.append(apertura)

    mes = int(mes or 0)
    if mes > 1:
        partidas_previas = obtener_partidas_acumuladas_hasta_mes(
            partidas_generadas_data,
            empresa_id,
            int(anio),
            mes - 1
        )
        partidas_fuente.extend(partidas_previas or [])

    return partidas_fuente
def saldo_neto_cuenta_en_partidas(partidas_fuente, cuenta_objetivo):
    cuenta_objetivo = limpiar_nombre_cuenta(cuenta_objetivo).lower()
    total_debe = 0.0
    total_haber = 0.0

    for partida in partidas_fuente or []:
        for mov in partida.get("cuentas", []) or []:
            cuenta_mov = limpiar_nombre_cuenta(mov.get("cuenta") or mov.get("nombre") or "").lower()
            if cuenta_mov != cuenta_objetivo:
                continue
            total_debe += round(float(mov.get("debe", 0) or 0), 2)
            total_haber += round(float(mov.get("haber", 0) or 0), 2)

    return round(total_debe - total_haber, 2)
def obtener_saldo_neto_cuenta_hasta_mes_anterior(
    ventas_data,
    compras_data,
    partidas_generadas_data,
    empresa_id,
    anio,
    mes,
    cuenta
):
    partidas_fuente = obtener_partidas_para_saldo_hasta_mes_anterior(
        ventas_data=ventas_data,
        compras_data=compras_data,
        partidas_generadas_data=partidas_generadas_data,
        empresa_id=empresa_id,
        anio=anio,
        mes=mes
    )
    return saldo_neto_cuenta_en_partidas(partidas_fuente, cuenta)
def inferir_naturaleza_cuenta_desde_catalogo(cuenta):
    cuenta_base = limpiar_nombre_cuenta(cuenta).lower()
    for item in cargar_cuentas() or []:
        nombre = str(item.get("nombre", "")).strip().lower()
        if nombre != cuenta_base:
            continue

        valores = []
        for k in ["tipo", "naturaleza", "clasificacion", "clasificación", "grupo", "categoria", "categoría", "rubro", "subtipo"]:
            v = item.get(k)
            if v is not None:
                valores.append(str(v).strip().lower())
        texto = " ".join(valores)

        if any(x in texto for x in ["activo", "deudora", "deudor", "gasto", "costo"]):
            return "deudora"
        if any(x in texto for x in ["pasivo", "acreedora", "acreedor", "patrimonio", "capital", "ingreso", "venta"]):
            return "acreedora"
    return None
def inferir_naturaleza_cuenta_por_nombre(cuenta):
    nombre = limpiar_nombre_cuenta(cuenta).lower()

    patrones_deudores = [
        "cliente", "cuentas por cobrar", "cxc", "caja", "banco", "inventario",
        "anticipo s/compras", "anticipo sobre compras", "iva credito", "iva crédito",
        "iva por cobrar"
    ]
    patrones_acreedores = [
        "proveedor", "cuentas por pagar", "cpp", "anticipo s/ventas", "anticipo sobre ventas",
        "por pagar", "retencion", "retención", "venta", "ingreso", "capital",
        "patrimonio", "iva debito", "iva débito"
    ]

    if any(p in nombre for p in patrones_deudores):
        return "deudora"
    if any(p in nombre for p in patrones_acreedores):
        return "acreedora"
    return None
def obtener_naturaleza_cuenta_para_cuadre(
    ventas_data,
    compras_data,
    partidas_generadas_data,
    empresa_id,
    anio,
    mes,
    cuenta
):
    saldo_neto = obtener_saldo_neto_cuenta_hasta_mes_anterior(
        ventas_data=ventas_data,
        compras_data=compras_data,
        partidas_generadas_data=partidas_generadas_data,
        empresa_id=empresa_id,
        anio=anio,
        mes=mes,
        cuenta=cuenta
    )
    if saldo_neto > 0.009:
        return "deudora", saldo_neto
    if saldo_neto < -0.009:
        return "acreedora", saldo_neto

    naturaleza = inferir_naturaleza_cuenta_desde_catalogo(cuenta)
    if naturaleza:
        return naturaleza, saldo_neto

    naturaleza = inferir_naturaleza_cuenta_por_nombre(cuenta)
    if naturaleza:
        return naturaleza, saldo_neto

    return "deudora", saldo_neto
def crear_movimiento_para_cuadre(cuenta, lado, monto):
    monto = round(float(monto or 0), 2)
    if monto <= 0:
        return None

    cuenta_txt = limpiar_nombre_cuenta(cuenta)
    if lado == "haber":
        cuenta_txt = f"A: {cuenta_txt}" if not cuenta_txt.startswith("A: ") else cuenta_txt
        return {"cuenta": cuenta_txt, "debe": 0.0, "haber": monto}

    return {"cuenta": cuenta_txt, "debe": monto, "haber": 0.0}
def cuadrar_con_saldo_real(
    cuentas,
    cuenta_cuadre,
    empresa_id,
    anio,
    mes,
    ventas_data,
    compras_data,
    partidas_generadas_data,
    cuenta_respaldo_debe,
    cuenta_respaldo_haber,
):
    diferencia = descuadre_partida(cuentas)
    info = {
        "diferencia_original": round(diferencia, 2),
        "cuenta_principal": limpiar_nombre_cuenta(cuenta_cuadre),
        "cuenta_respaldo_debe": limpiar_nombre_cuenta(cuenta_respaldo_debe),
        "cuenta_respaldo_haber": limpiar_nombre_cuenta(cuenta_respaldo_haber),
        "lado_necesario": None,
        "naturaleza_cuenta_principal": None,
        "saldo_neto_anterior": 0.0,
        "saldo_disponible": 0.0,
        "monto_usado_principal": 0.0,
        "monto_usado_respaldo": 0.0,
        "redujo_cuenta_principal": False,
    }
    if abs(diferencia) <= 0.009:
        return cuentas, info

    lado_necesario = "haber" if diferencia > 0 else "debe"
    monto_necesario = round(abs(diferencia), 2)

    naturaleza, saldo_neto = obtener_naturaleza_cuenta_para_cuadre(
        ventas_data=ventas_data,
        compras_data=compras_data,
        partidas_generadas_data=partidas_generadas_data,
        empresa_id=empresa_id,
        anio=anio,
        mes=mes,
        cuenta=cuenta_cuadre
    )

    info["lado_necesario"] = lado_necesario
    info["naturaleza_cuenta_principal"] = naturaleza
    info["saldo_neto_anterior"] = round(saldo_neto, 2)

    reduce_principal = (
        (naturaleza == "deudora" and lado_necesario == "haber")
        or (naturaleza == "acreedora" and lado_necesario == "debe")
    )
    info["redujo_cuenta_principal"] = reduce_principal

    if naturaleza == "deudora":
        saldo_disponible = round(max(saldo_neto, 0.0), 2)
    else:
        saldo_disponible = round(max(-saldo_neto, 0.0), 2)
    info["saldo_disponible"] = saldo_disponible

    monto_principal = monto_necesario
    monto_respaldo = 0.0

    if reduce_principal:
        monto_principal = round(min(saldo_disponible, monto_necesario), 2)
        monto_respaldo = round(monto_necesario - monto_principal, 2)

    if monto_principal > 0:
        mov = crear_movimiento_para_cuadre(cuenta_cuadre, lado_necesario, monto_principal)
        if mov:
            cuentas.append(mov)
        info["monto_usado_principal"] = monto_principal

    if monto_respaldo > 0:
        cuenta_respaldo = cuenta_respaldo_haber if lado_necesario == "haber" else cuenta_respaldo_debe
        mov = crear_movimiento_para_cuadre(cuenta_respaldo, lado_necesario, monto_respaldo)
        if mov:
            cuentas.append(mov)
        info["monto_usado_respaldo"] = monto_respaldo

    return cuentas, info


def guardar_salarios_minimos(data):
    guardar_json_dict(SALARIOS_MINIMOS_FILE, data)


def obtener_salario_minimo(anio, circunscripcion, tipo_suscripcion, salarios_data):
    anio = str(anio).strip()
    circunscripcion = str(circunscripcion).strip().upper()
    tipo_suscripcion = str(tipo_suscripcion).strip().lower()

    return float(
        salarios_data
        .get(anio, {})
        .get(circunscripcion, {})
        .get(tipo_suscripcion, 0.0) or 0.0
    )

def calcular_planilla_laboral(
    cantidad_trabajadores,
    salario_mensual,
    incluir_igss=True,
    incluir_irtra=True,
    incluir_intecap=True,
    bonificacion_incentivo=250.00
):
    cantidad_trabajadores = int(cantidad_trabajadores or 0)
    salario_mensual = float(salario_mensual or 0.0)

    salario_base_total = round(cantidad_trabajadores * salario_mensual, 2)
    bonificacion_total = round(cantidad_trabajadores * float(bonificacion_incentivo or 0.0), 2)

    cuota_laboral = round(salario_base_total * PORC_IGSS_LABORAL, 2) if incluir_igss else 0.0
    cuota_patronal = round(salario_base_total * PORC_IGSS_PATRONAL, 2) if incluir_igss else 0.0
    irtra = round(salario_base_total * PORC_IRTRA, 2) if incluir_irtra else 0.0
    intecap = round(salario_base_total * PORC_INTECAP, 2) if incluir_intecap else 0.0

    sueldo_neto_total = round(salario_base_total - cuota_laboral, 2)

    return {
        "cantidad_trabajadores": cantidad_trabajadores,
        "salario_mensual": round(salario_mensual, 2),
        "salario_base_total": salario_base_total,
        "bonificacion_total": bonificacion_total,
        "cuota_laboral": cuota_laboral,
        "cuota_patronal": cuota_patronal,
        "irtra": irtra,
        "intecap": intecap,
        "sueldo_neto_total": sueldo_neto_total
    }

def calcular_planilla_laboral_detallada(
    trabajadores,
    incluir_igss=True,
    incluir_irtra=True,
    incluir_intecap=True,
    bonificacion_incentivo=250.00
):
    salario_base_total = 0.0
    bonificacion_total = 0.0

    for t in trabajadores or []:
        salario = round(float(t.get("salario", 0) or 0), 2)
        recibe_bonificacion = bool(t.get("bonificacion", True))

        salario_base_total += salario
        if recibe_bonificacion:
            bonificacion_total += round(float(bonificacion_incentivo or 0.0), 2)

    salario_base_total = round(salario_base_total, 2)
    bonificacion_total = round(bonificacion_total, 2)

    cuota_laboral = round(salario_base_total * PORC_IGSS_LABORAL, 2) if incluir_igss else 0.0
    cuota_patronal = round(salario_base_total * PORC_IGSS_PATRONAL, 2) if incluir_igss else 0.0
    irtra = round(salario_base_total * PORC_IRTRA, 2) if incluir_irtra else 0.0
    intecap = round(salario_base_total * PORC_INTECAP, 2) if incluir_intecap else 0.0

    sueldo_neto_total = round(salario_base_total - cuota_laboral, 2)

    return {
        "cantidad_trabajadores": len(trabajadores or []),
        "salario_mensual": 0.0,
        "salario_base_total": salario_base_total,
        "bonificacion_total": bonificacion_total,
        "cuota_laboral": cuota_laboral,
        "cuota_patronal": cuota_patronal,
        "irtra": irtra,
        "intecap": intecap,
        "sueldo_neto_total": sueldo_neto_total,
        "detalle_trabajadores": trabajadores or []
    }

def normalizar_apertura_contrapartidas(apertura):
    if not apertura:
        return None

    nueva_apertura = apertura.copy()
    cuentas_nuevas = []

    for mov in apertura.get("cuentas", []):
        cuenta = str(mov.get("cuenta", "")).strip()
        debe = float(mov.get("debe", 0) or 0)
        haber = float(mov.get("haber", 0) or 0)

        # 🔥 Si es HABER → agregar A:
        if haber > 0:
            if not cuenta.startswith("A: "):
                cuenta = f"A: {cuenta}"

        cuentas_nuevas.append({
            "cuenta": cuenta,
            "debe": debe,
            "haber": haber
        })

    nueva_apertura["cuentas"] = cuentas_nuevas
    return nueva_apertura

def movimientos_planilla_laboral(
    datos_planilla,
    modo_pago="mensual",
    sueldo_por_pagar_manual=0.0,
    bonificacion_por_pagar_manual=0.0,
    monto_retencion_isr_planilla=0.0,
):
    movimientos = []

    salario_base_total = round(float(datos_planilla.get("salario_base_total", 0) or 0), 2)
    bonificacion_total = round(float(datos_planilla.get("bonificacion_total", 0) or 0), 2)
    cuota_laboral = round(float(datos_planilla.get("cuota_laboral", 0) or 0), 2)
    cuota_patronal = round(float(datos_planilla.get("cuota_patronal", 0) or 0), 2)
    irtra = round(float(datos_planilla.get("irtra", 0) or 0), 2)
    intecap = round(float(datos_planilla.get("intecap", 0) or 0), 2)
    sueldo_neto_total = round(float(datos_planilla.get("sueldo_neto_total", 0) or 0), 2)

    # ===== DEBE =====
    if salario_base_total > 0:
        movimientos.append({"cuenta": "Sueldos y Salarios", "debe": salario_base_total, "haber": 0.0})

    if bonificacion_total > 0:
        movimientos.append({"cuenta": "Bonificación", "debe": bonificacion_total, "haber": 0.0})

    if cuota_patronal > 0:
        movimientos.append({"cuenta": "Cuota Patronal", "debe": cuota_patronal, "haber": 0.0})

    if irtra > 0:
        movimientos.append({"cuenta": "IRTRA", "debe": irtra, "haber": 0.0})

    if intecap > 0:
        movimientos.append({"cuenta": "INTECAP", "debe": intecap, "haber": 0.0})

    # ===== HABER que siempre queda por pagar =====
    if cuota_laboral > 0:
        movimientos.append({"cuenta": "A: Cuota Laboral por pagar", "debe": 0.0, "haber": cuota_laboral})

    if cuota_patronal > 0:
        movimientos.append({"cuenta": "A: Cuota Patronal por pagar", "debe": 0.0, "haber": cuota_patronal})

    if irtra > 0:
        movimientos.append({"cuenta": "A: IRTRA por pagar", "debe": 0.0, "haber": irtra})

    if intecap > 0:
        movimientos.append({"cuenta": "A: INTECAP por pagar", "debe": 0.0, "haber": intecap})

    # ===== Sueldos y bonificación por pagar =====
    # La retención ISR planilla rebaja primero el sueldo neto total del período.
    # Luego, sobre ese saldo neto ya rebajado, se calcula mensual/quincenal/manual.
    monto_retencion_isr_planilla = round(
        min(float(monto_retencion_isr_planilla or 0), max(sueldo_neto_total, 0.0)),
        2
    )
    sueldo_neto_despues_isr = round(max(sueldo_neto_total - monto_retencion_isr_planilla, 0.0), 2)

    modo_pago = str(modo_pago or "mensual").strip().lower()

    if modo_pago == "quincenal":
        sueldo_por_pagar = round(sueldo_neto_despues_isr / 2, 2)
        bonificacion_por_pagar = round(bonificacion_total / 2, 2)

    elif modo_pago == "manual":
        sueldo_por_pagar = round(min(float(sueldo_por_pagar_manual or 0), sueldo_neto_despues_isr), 2)
        bonificacion_por_pagar = round(min(float(bonificacion_por_pagar_manual or 0), bonificacion_total), 2)

    else:  # mensual
        sueldo_por_pagar = sueldo_neto_despues_isr
        bonificacion_por_pagar = bonificacion_total

    if sueldo_por_pagar > 0:
        movimientos.append({"cuenta": "A: Sueldos y Salarios por pagar", "debe": 0.0, "haber": sueldo_por_pagar})

    if monto_retencion_isr_planilla > 0:
        movimientos.append({"cuenta": "A: Retención ISR por pagar", "debe": 0.0, "haber": monto_retencion_isr_planilla})

    if bonificacion_por_pagar > 0:
        movimientos.append({"cuenta": "A: Bonificación por pagar", "debe": 0.0, "haber": bonificacion_por_pagar})

    return consolidar_movimientos(movimientos)

def calcular_provision_prestaciones(salario_base_total, porcentaje=0.2638):
    salario_base_total = round(float(salario_base_total or 0), 2)
    porcentaje = float(porcentaje or 0)
    return round(salario_base_total * porcentaje, 2)

def cargar_empresas():
    return cargar_json_lista(EMPRESAS_FILE)

def cargar_cuentas():
    return cargar_json_lista(CUENTAS_FILE)


def opciones_cuentas_catalogo():
    cuentas = cargar_cuentas()
    nombres = []
    for c in cuentas:
        nombre = str(c.get("nombre", "")).strip()
        if nombre:
            nombres.append(nombre)
    return sorted(set(nombres))

def opciones_cuentas_cobro():
    cuentas = cargar_cuentas()
    nombres = []

    palabras_validas = {
        "caja", "bancos", "banco", "banco 1", "banco 2", "banco 3"
    }

    for c in cuentas:
        nombre = str(c.get("nombre", "")).strip()
        activa = bool(c.get("activa", True))

        if not nombre or not activa:
            continue

        nombre_lower = nombre.lower()

        if nombre_lower in palabras_validas or nombre_lower.startswith("banco"):
            nombres.append(nombre)

    return sorted(set(nombres))

def normalizar_nit(nit):
    return str(nit or "").strip().replace("-", "").replace(" ", "").upper()


def buscar_empresa_por_nit(nit, empresas):
    nit_norm = normalizar_nit(nit)
    for e in empresas:
        if normalizar_nit(e.get("nit")) == nit_norm:
            return e
    return None


def buscar_empresa_por_nombre(nombre, empresas):
    nombre_norm = str(nombre or "").strip()
    for e in empresas:
        if str(e.get("nombre", "")).strip() == nombre_norm:
            return e
    return None


def obtener_id_empresa_desde_partida(partida, empresas):
    nit_partida = normalizar_nit(partida.get("empresa_nit"))
    nombre_partida = str(partida.get("empresa", "")).strip()

    if nit_partida:
        emp = buscar_empresa_por_nit(nit_partida, empresas)
        if emp:
            return normalizar_nit(emp.get("nit"))

    if nombre_partida:
        emp = buscar_empresa_por_nombre(nombre_partida, empresas)
        if emp:
            return normalizar_nit(emp.get("nit"))

    return ""


def ruta_config_ui_db():
    return resolver_ruta(CONFIG_UI_DB_FILE)


def get_config_ui_conn():
    ruta_db = ruta_config_ui_db()
    carpeta = os.path.dirname(ruta_db) or "."
    os.makedirs(carpeta, exist_ok=True)

    conn = sqlite3.connect(ruta_db)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def init_config_ui_db():
    conn = get_config_ui_conn()
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS config_ui (
                key_config TEXT PRIMARY KEY,
                empresa_id TEXT,
                anio INTEGER,
                mes INTEGER,
                datos_json TEXT NOT NULL,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_config_ui_empresa_periodo ON config_ui(empresa_id, anio, mes)"
        )
        conn.commit()
    finally:
        conn.close()



def parsear_key_config(key):
    key = str(key or "").strip()
    partes = key.split("__")
    if len(partes) == 3:
        empresa_id = str(partes[0]).strip()
        try:
            anio = int(partes[1])
            mes = int(partes[2])
            return empresa_id, anio, mes
        except (TypeError, ValueError):
            return key, None, None
    return key, None, None



def contar_registros_config_ui_sqlite():
    init_config_ui_db()
    conn = get_config_ui_conn()
    try:
        row = conn.execute("SELECT COUNT(*) AS total FROM config_ui").fetchone()
        return int(row["total"] or 0)
    finally:
        conn.close()



def guardar_config_ui_sqlite_por_key(key, datos):
    init_config_ui_db()
    empresa_id, anio, mes = parsear_key_config(key)
    payload = json.dumps(datos or {}, ensure_ascii=False, default=_json_default)
    conn = get_config_ui_conn()
    try:
        conn.execute(
            """
            INSERT INTO config_ui (key_config, empresa_id, anio, mes, datos_json, updated_at)
            VALUES (?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(key_config) DO UPDATE SET
                empresa_id=excluded.empresa_id,
                anio=excluded.anio,
                mes=excluded.mes,
                datos_json=excluded.datos_json,
                updated_at=CURRENT_TIMESTAMP
            """,
            (str(key).strip(), empresa_id, anio, mes, payload)
        )
        conn.commit()
    finally:
        conn.close()



def cargar_config_ui_desde_sqlite():
    init_config_ui_db()
    conn = get_config_ui_conn()
    try:
        rows = conn.execute(
            "SELECT key_config, datos_json FROM config_ui ORDER BY key_config"
        ).fetchall()
    finally:
        conn.close()

    salida = {}
    for row in rows:
        key = str(row["key_config"] or "").strip()
        if not key:
            continue
        try:
            datos = json.loads(row["datos_json"] or "{}")
        except json.JSONDecodeError:
            continue
        if isinstance(datos, dict):
            salida[key] = datos

    return salida



def migrar_config_ui_json_a_sqlite_si_hace_falta():
    init_config_ui_db()
    if contar_registros_config_ui_sqlite() > 0:
        return

    config_json = cargar_json_dict(CONFIG_UI_FILE)
    if not isinstance(config_json, dict) or not config_json:
        return

    for key, datos in config_json.items():
        if isinstance(datos, dict):
            guardar_config_ui_sqlite_por_key(key, datos)



def cargar_config_ui():
    migrar_config_ui_json_a_sqlite_si_hace_falta()
    config_sqlite = cargar_config_ui_desde_sqlite()
    if config_sqlite:
        return config_sqlite
    return cargar_json_dict(CONFIG_UI_FILE)



def guardar_config_ui(config):
    if not isinstance(config, dict):
        config = {}

    init_config_ui_db()
    conn = get_config_ui_conn()
    try:
        conn.execute("DELETE FROM config_ui")
        conn.commit()
    finally:
        conn.close()

    for key, datos in config.items():
        if isinstance(datos, dict):
            guardar_config_ui_sqlite_por_key(key, datos)

    guardar_json_seguro(CONFIG_UI_FILE, config)



def clave_config_periodo(empresa_id, anio, mes):
    return f"{str(empresa_id).strip()}__{int(anio)}__{int(mes)}"



def existe_config_periodo(config_ui, empresa_id, anio, mes):
    return clave_config_periodo(empresa_id, anio, mes) in config_ui



def obtener_config_empresa(config_ui, empresa_id, anio=None, mes=None):
    empresa_id = str(empresa_id).strip()

    if anio is not None and mes is not None:
        key_periodo = clave_config_periodo(empresa_id, anio, mes)
        if key_periodo in config_ui and isinstance(config_ui.get(key_periodo), dict):
            return dict(config_ui.get(key_periodo, {}))

        objetivo = (int(anio), int(mes))
        mejor_periodo = None
        mejor_config = None
        prefijo = f"{empresa_id}__"

        for key, value in config_ui.items():
            if not (isinstance(key, str) and key.startswith(prefijo) and isinstance(value, dict)):
                continue

            partes = key.split("__")
            if len(partes) != 3:
                continue

            try:
                anio_cfg = int(partes[1])
                mes_cfg = int(partes[2])
            except (TypeError, ValueError):
                continue

            periodo_cfg = (anio_cfg, mes_cfg)
            if periodo_cfg >= objetivo:
                continue

            if mejor_periodo is None or periodo_cfg > mejor_periodo:
                mejor_periodo = periodo_cfg
                mejor_config = value

        if isinstance(mejor_config, dict):
            return dict(mejor_config)

    # fallback: config general por empresa si existiera
    config_general = config_ui.get(empresa_id, {})
    return dict(config_general) if isinstance(config_general, dict) else {}



def guardar_config_empresa(config_ui, empresa_id, anio, mes, datos):
    empresa_id = str(empresa_id).strip()
    key_periodo = clave_config_periodo(empresa_id, anio, mes)
    datos_limpios = dict(datos or {}) if isinstance(datos, dict) else {}
    config_ui[key_periodo] = datos_limpios
    guardar_config_ui_sqlite_por_key(key_periodo, datos_limpios)
    guardar_json_seguro(CONFIG_UI_FILE, config_ui)


def periodo_anterior(anio, mes):
    anio = int(anio)
    mes = int(mes)
    if mes <= 1:
        return anio - 1, 12
    return anio, mes - 1


def extraer_numero_pda(codigo):
    txt = str(codigo or "").strip().upper()
    if not txt:
        return None

    m = re.search(r"PDA\s*(\d+)", txt)
    if m:
        try:
            return int(m.group(1))
        except (TypeError, ValueError):
            return None

    solo_num = re.search(r"(\d+)", txt)
    if solo_num:
        try:
            return int(solo_num.group(1))
        except (TypeError, ValueError):
            return None

    return None


def obtener_siguiente_pda_default(partidas_generadas_data, empresa_id, anio, mes, fallback=58):
    empresa_id = normalizar_nit(empresa_id)
    objetivo = (int(anio), int(mes))
    mejor_registro = None
    mejor_periodo = None

    for item in partidas_generadas_data or []:
        if normalizar_nit(item.get("empresa_nit")) != empresa_id:
            continue

        try:
            periodo_item = (int(item.get("anio", 0) or 0), int(item.get("mes", 0) or 0))
        except (TypeError, ValueError):
            continue

        if periodo_item >= objetivo:
            continue

        if mejor_periodo is None or periodo_item > mejor_periodo:
            mejor_periodo = periodo_item
            mejor_registro = item

    if mejor_registro:
        ultimo_pda = 0
        for partida in mejor_registro.get("partidas", []) or []:
            numero = extraer_numero_pda(partida.get("codigo", ""))
            if numero is not None and numero > ultimo_pda:
                ultimo_pda = numero

        if ultimo_pda > 0:
            return ultimo_pda + 1

    try:
        return int(fallback)
    except (TypeError, ValueError):
        return 58


def obtener_regimen_empresa(empresa_id, empresas):
    emp = buscar_empresa_por_nit(empresa_id, empresas)
    if emp:
        return str(emp.get("regimen", "Mensual")).strip()
    return "Mensual"


def mes_a_numero(valor):
    if isinstance(valor, int):
        return valor
    if isinstance(valor, float):
        return int(valor)
    txt = str(valor).strip().lower()
    if txt.isdigit():
        return int(txt)
    return MESES_INV.get(txt, 0)


def ultimo_dia_mes(anio, mes):
    return calendar.monthrange(int(anio), int(mes))[1]


def fmt_fecha_fin_mes(anio, mes):
    dia = ultimo_dia_mes(anio, mes)
    return f"{dia:02d}/{int(mes):02d}/{int(anio)}"


def normalizar_texto_clave(valor):
    return re.sub(r"\s+", " ", str(valor or "").strip().casefold())


def clave_estable_partida(partida):
    """Identifica la partida sin depender del numero PDA visible."""
    tipo_partida = normalizar_texto_clave(partida.get("tipo_partida"))
    glosa = normalizar_texto_clave(partida.get("glosa"))
    fecha = str(partida.get("fecha", "") or "").strip()
    empresa_nit = normalizar_nit(partida.get("empresa_nit", ""))
    anio = str(partida.get("anio", "") or "").strip()
    mes = str(partida.get("mes", "") or "").strip()

    if not (tipo_partida or glosa):
        return ""

    descriptor = glosa or f"tipo:{tipo_partida}"
    return "|".join(
        [
            empresa_nit,
            anio,
            mes,
            fecha,
            descriptor,
        ]
    )


def normalizar_mov(mov):
    cuenta = str(mov.get("nombre") or mov.get("cuenta") or "").strip()
    debe = float(mov.get("debe", 0) or 0)
    haber = float(mov.get("haber", 0) or 0)
    anotacion = str(mov.get("anotacion", "") or "").strip()
    return {"cuenta": cuenta, "debe": debe, "haber": haber, "anotacion": anotacion}


def consolidar_movimientos(movimientos):
    acum = {}
    orden = []
    for mov in movimientos:
        m = normalizar_mov(mov)
        if not m["cuenta"]:
            continue
        if abs(m["debe"]) < 1e-9 and abs(m["haber"]) < 1e-9:
            continue
        key = m["cuenta"]
        if key not in acum:
            acum[key] = {"cuenta": key, "debe": 0.0, "haber": 0.0, "anotacion": ""}
            orden.append(key)
        acum[key]["debe"] += m["debe"]
        acum[key]["haber"] += m["haber"]
        if m["anotacion"] and not acum[key]["anotacion"]:
            acum[key]["anotacion"] = m["anotacion"]

    lineas = [acum[k] for k in orden if round(acum[k]["debe"], 2) != 0 or round(acum[k]["haber"], 2) != 0]
    lineas_debe = [x for x in lineas if round(x["debe"], 2) > 0]
    lineas_haber = [x for x in lineas if round(x["haber"], 2) > 0]

    return [
        {
            "cuenta": x["cuenta"],
            "debe": round(x["debe"], 2),
            "haber": round(x["haber"], 2),
            **({"anotacion": x["anotacion"]} if x["anotacion"] else {}),
        }
        for x in (lineas_debe + lineas_haber)
    ]


def preservar_anotaciones_partidas(partidas_nuevas, partidas_anteriores):
    """Copia notas por identidad de partida, cuenta y aparicion dentro de la partida."""
    anteriores_por_clave = {}
    for p in partidas_anteriores or []:
        clave = clave_estable_partida(p)
        if clave:
            anteriores_por_clave.setdefault(clave, []).append(p)

    anteriores_por_codigo = {
        str(p.get("codigo", "")).strip(): p
        for p in partidas_anteriores or []
    }

    usos_por_clave = {}
    for partida in partidas_nuevas or []:
        anterior = None
        clave = clave_estable_partida(partida)
        if clave:
            posicion_clave = usos_por_clave.get(clave, 0)
            usos_por_clave[clave] = posicion_clave + 1
            anteriores_misma_clave = anteriores_por_clave.get(clave, [])
            if posicion_clave < len(anteriores_misma_clave):
                anterior = anteriores_misma_clave[posicion_clave]

        if not anterior:
            anterior = anteriores_por_codigo.get(str(partida.get("codigo", "")).strip())

        if not anterior:
            continue

        notas_por_cuenta = {}
        for mov in anterior.get("cuentas", []) or []:
            cuenta = str(mov.get("cuenta", "")).strip().casefold()
            notas_por_cuenta.setdefault(cuenta, []).append(
                str(mov.get("anotacion", "") or "").strip()
            )

        apariciones = {}
        for mov in partida.get("cuentas", []) or []:
            cuenta = str(mov.get("cuenta", "")).strip().casefold()
            posicion = apariciones.get(cuenta, 0)
            apariciones[cuenta] = posicion + 1
            notas = notas_por_cuenta.get(cuenta, [])
            if posicion < len(notas):
                if notas[posicion]:
                    mov["anotacion"] = notas[posicion]
                else:
                    mov.pop("anotacion", None)

    return partidas_nuevas


def total_debe_haber(cuentas):
    debe = round(sum(float(x.get("debe", 0) or 0) for x in cuentas), 2)
    haber = round(sum(float(x.get("haber", 0) or 0) for x in cuentas), 2)
    return debe, haber


def descuadre_partida(cuentas):
    debe, haber = total_debe_haber(consolidar_movimientos(cuentas))
    return round(debe - haber, 2)


def cuadrar_con_clientes(cuentas, nombre_clientes="Clientes"):
    diferencia = descuadre_partida(cuentas)
    if abs(diferencia) <= 0.009:
        return cuentas

    if diferencia > 0:
        cuentas.append({"cuenta": f"A: {nombre_clientes}", "debe": 0.0, "haber": round(diferencia, 2)})
    else:
        cuentas.append({"cuenta": nombre_clientes, "debe": round(abs(diferencia), 2), "haber": 0.0})
    return cuentas

def cuadrar_con_cuenta(cuentas, cuenta_cuadre="Proveedores"):
    diferencia = descuadre_partida(cuentas)
    if abs(diferencia) <= 0.009:
        return cuentas

    if diferencia > 0:
        cuentas.append({
            "cuenta": f"A: {cuenta_cuadre}",
            "debe": 0.0,
            "haber": round(diferencia, 2)
        })
    else:
        cuentas.append({
            "cuenta": cuenta_cuadre,
            "debe": round(abs(diferencia), 2),
            "haber": 0.0
        })
    return cuentas

def cuenta_caja_preferida(venta=None, compra=None):
    candidatos = []
    for p in [venta, compra]:
        if not p:
            continue
        for mov in p.get("cuentas", []):
            c = str(mov.get("cuenta") or mov.get("nombre") or "").strip()
            base = c.replace("A: ", "").strip().lower()
            if base in CUENTAS_CAJA:
                candidatos.append(c.replace("A: ", "").strip())
    return candidatos[0] if candidatos else "Caja"


def buscar_partida(data, empresa_id, anio, mes):
    empresa_id = normalizar_nit(empresa_id)

    for p in data:
        nit_partida = normalizar_nit(p.get("empresa_nit"))
        nombre_partida = str(p.get("empresa", "")).strip()

        coincide_empresa = False

        if nit_partida:
            coincide_empresa = (nit_partida == empresa_id)
        else:
            emp = buscar_empresa_por_nit(empresa_id, cargar_empresas())
            if emp:
                coincide_empresa = (nombre_partida == str(emp.get("nombre", "")).strip())

        if (
            coincide_empresa
            and int(p.get("anio", 0) or 0) == int(anio)
            and mes_a_numero(p.get("mes")) == int(mes)
        ):
            return p
    return None


def ventas_netas(venta):
    total = 0.0
    for mov in venta.get("cuentas", []):
        m = normalizar_mov(mov)
        nombre = m["cuenta"].lower()
        if m["haber"] <= 0:
            continue
        if (
            "iva" in nombre
            or "retencion" in nombre
            or nombre.replace("a: ", "").strip() in {"caja", "bancos", "banco", "clientes"}
        ):
            continue
        total += m["haber"]
    return round(total, 2)


def total_cobrado_venta(venta):
    total = 0.0
    for mov in venta.get("cuentas", []):
        m = normalizar_mov(mov)
        nombre = m["cuenta"].replace("A: ", "").strip().lower()
        if nombre in CUENTAS_CAJA and m["debe"] > 0:
            total += m["debe"]
    if total > 0:
        return round(total, 2)
    return round(ventas_netas(venta) + monto_iva_debito(venta), 2)

def buscar_partida_apertura(data, empresa_id, anio):
    empresa_id = normalizar_nit(empresa_id)

    for p in data:
        nit_p = normalizar_nit(p.get("empresa_nit"))
        anio_p = int(p.get("anio", 0) or 0)
        mes_p = mes_a_numero(p.get("mes"))
        tipo_partida = str(p.get("tipo_partida", "")).strip().lower()
        codigo = str(p.get("codigo", "")).strip().upper()

        if (
            nit_p == empresa_id
            and anio_p == int(anio)
            and (
                mes_p == 0
                or tipo_partida == "apertura"
                or codigo == "PDA 0"
            )
        ):
            return p

    return None

def saldo_cuenta_en_apertura(apertura, nombres_posibles):
    if not apertura:
        return 0.0

    nombres = {str(x).strip().lower() for x in nombres_posibles}
    saldo = 0.0

    for mov in apertura.get("cuentas", []):
        cuenta = str(mov.get("cuenta", "")).replace("A: ", "").strip().lower()
        debe = float(mov.get("debe", 0) or 0)
        haber = float(mov.get("haber", 0) or 0)

        if cuenta in nombres:
            saldo += round(debe - haber, 2)

    return round(saldo, 2)


def saldo_cuenta_en_partidas(partidas_fuente, nombres_posibles):
    nombres = {str(x).replace("A: ", "").strip().lower() for x in (nombres_posibles or []) if str(x).strip()}
    if not nombres or not partidas_fuente:
        return 0.0

    if isinstance(partidas_fuente, dict):
        partidas_fuente = [partidas_fuente]

    saldo = 0.0
    for partida in partidas_fuente:
        for mov in partida.get("cuentas", []):
            cuenta = str(mov.get("cuenta", "")).replace("A: ", "").strip().lower()
            if cuenta not in nombres:
                continue
            debe = float(mov.get("debe", 0) or 0)
            haber = float(mov.get("haber", 0) or 0)
            saldo += round(debe - haber, 2)

    return round(saldo, 2)


def obtener_saldo_arrastrado_desde_mayor_mes(
    ventas_data,
    compras_data,
    partidas_generadas_data,
    empresa_id,
    anio,
    mes,
    nombres_posibles,
):
    apertura, partidas_mes = obtener_apertura_y_movimientos_para_mayor_local(
        ventas_data=ventas_data,
        compras_data=compras_data,
        partidas_generadas_data=partidas_generadas_data,
        empresa_id=empresa_id,
        anio=int(anio),
        mes=int(mes)
    )

    partidas_fuente = []
    if apertura:
        partidas_fuente.append(apertura)
    partidas_fuente.extend(partidas_mes or [])

    saldo = saldo_cuenta_en_partidas(partidas_fuente, nombres_posibles)
    return round(max(saldo, 0.0), 2)


def obtener_arrastres_iva_para_regularizacion(
    ventas_data,
    compras_data,
    partidas_generadas_data,
    empresa_id,
    anio,
    mes,
):
    if int(mes) == 1:
        apertura = buscar_partida_apertura(ventas_data, empresa_id, anio)
        if not apertura:
            apertura = buscar_partida_apertura(compras_data, empresa_id, anio)

        apertura = normalizar_apertura_contrapartidas(apertura)

        saldo_iva_por_cobrar = round(
            max(saldo_cuenta_en_apertura(apertura, ["iva por cobrar"]), 0.0),
            2
        )
        saldo_credito_retencion_iva = round(
            max(saldo_cuenta_en_apertura(apertura, ["credito de retencion iva", "crédito de retención iva"]), 0.0),
            2
        )

        return saldo_iva_por_cobrar, saldo_credito_retencion_iva

    prev_mes = 12 if int(mes) == 1 else int(mes) - 1
    prev_anio = int(anio) - 1 if int(mes) == 1 else int(anio)

    saldo_iva_por_cobrar = obtener_saldo_arrastrado_desde_mayor_mes(
        ventas_data=ventas_data,
        compras_data=compras_data,
        partidas_generadas_data=partidas_generadas_data,
        empresa_id=empresa_id,
        anio=prev_anio,
        mes=prev_mes,
        nombres_posibles=["iva por cobrar"],
    )
    saldo_credito_retencion_iva = obtener_saldo_arrastrado_desde_mayor_mes(
        ventas_data=ventas_data,
        compras_data=compras_data,
        partidas_generadas_data=partidas_generadas_data,
        empresa_id=empresa_id,
        anio=prev_anio,
        mes=prev_mes,
        nombres_posibles=["credito de retencion iva", "crédito de retención iva"],
    )

    return saldo_iva_por_cobrar, saldo_credito_retencion_iva

def obtener_saldo_iso_por_acreditar_hasta_fecha(
    ventas_data,
    compras_data,
    partidas_generadas_data,
    empresa_id,
    anio,
    mes,
    incluir_mes_actual=False,
):
    nombres_iso = ["iso por acreditar", "i.s.o. por acreditar", "impuesto de solidaridad por acreditar"]

    if incluir_mes_actual:
        saldo = obtener_saldo_arrastrado_desde_mayor_mes(
            ventas_data=ventas_data,
            compras_data=compras_data,
            partidas_generadas_data=partidas_generadas_data,
            empresa_id=empresa_id,
            anio=int(anio),
            mes=int(mes),
            nombres_posibles=nombres_iso,
        )
    else:
        partidas_fuente = obtener_partidas_para_saldo_hasta_mes_anterior(
            ventas_data=ventas_data,
            compras_data=compras_data,
            partidas_generadas_data=partidas_generadas_data,
            empresa_id=empresa_id,
            anio=int(anio),
            mes=int(mes),
        )
        saldo = saldo_cuenta_en_partidas(partidas_fuente, nombres_iso)

    return round(max(float(saldo or 0), 0.0), 2)

def monto_iva_debito(venta):
    total = 0.0
    for mov in venta.get("cuentas", []):
        m = normalizar_mov(mov)
        if "iva débito" in m["cuenta"].lower() or "iva debito" in m["cuenta"].lower():
            total += m["haber"] if m["haber"] > 0 else m["debe"]
    return round(total, 2)


def monto_iva_credito(compra):
    total = 0.0
    for mov in compra.get("cuentas", []):
        m = normalizar_mov(mov)
        if "iva crédito" in m["cuenta"].lower() or "iva credito" in m["cuenta"].lower():
            total += m["debe"] if m["debe"] > 0 else m["haber"]
    return round(total, 2)


def obtener_comision_desde_compras(compra):
    comision = 0.0
    if not compra:
        return 0.0
    for mov in compra.get("cuentas", []):
        m = normalizar_mov(mov)
        if m["cuenta"].strip().lower() == "comisiones pagadas":
            comision += round(m["debe"], 2)
    return round(comision, 2)


def calcular_iva_comision(comision):
    return round(float(comision or 0) * 0.12, 2)


def limpiar_movimientos_ingreso_config(items, permitir_solo_cuenta=False):
    salida = []

    for item in items or []:
        cuenta = str(item.get("cuenta", "")).replace("A: ", "").strip()
        monto = round(float(item.get("monto", 0) or 0), 2)
        iva = round(float(item.get("iva", 0) or 0), 2)

        if not cuenta:
            continue

        if not permitir_solo_cuenta and monto <= 0 and iva <= 0:
            continue

        salida.append({
            "cuenta": cuenta,
            "monto": monto,
            "iva": iva,
        })

    return salida


def limpiar_haberes_manuales(items):
    salida = []

    for item in items or []:
        cuenta = str(item.get("cuenta", "")).replace("A: ", "").strip()
        monto = round(float(item.get("monto", 0) or 0), 2)

        if not cuenta or monto <= 0:
            continue

        salida.append({
            "cuenta": cuenta,
            "monto": monto,
        })

    return salida


def limpiar_haberes_manuales_ingreso(items):
    return limpiar_haberes_manuales(items)


def resumir_salarios_sujetos_isr_planilla(datos_planilla, umbral=4000.0):
    datos_planilla = datos_planilla or {}
    umbral = round(float(umbral or 0), 2)
    detalle = datos_planilla.get("detalle_trabajadores", []) or []

    cantidad = 0
    base = 0.0

    if detalle:
        for item in detalle:
            salario = round(float(item.get("salario", 0) or 0), 2)
            if salario > umbral:
                cantidad += 1
                base = round(base + salario, 2)
    else:
        salario_ref = round(float(datos_planilla.get("salario_mensual", 0) or 0), 2)
        cantidad_trabajadores = int(datos_planilla.get("cantidad_trabajadores", 0) or 0)
        if salario_ref > umbral and cantidad_trabajadores > 0:
            cantidad = cantidad_trabajadores
            base = round(salario_ref * cantidad_trabajadores, 2)

    return {
        "cantidad": int(cantidad),
        "base_referencia": round(base, 2),
        "umbral": umbral,
    }


def obtener_cuentas_debe_compra_para_traslado(compra):
    opciones = []
    vistos = set()

    if not compra:
        return opciones

    for mov in compra.get("cuentas", []):
        m = normalizar_mov(mov)
        cuenta = str(m.get("cuenta", "")).replace("A: ", "").strip()
        cuenta_lower = cuenta.lower()

        if not cuenta or m["debe"] <= 0:
            continue

        if cuenta_lower in {"iva crédito", "iva credito"}:
            continue

        if cuenta_lower in vistos:
            continue

        vistos.add(cuenta_lower)
        opciones.append(cuenta)

    return opciones


def obtener_mapa_cuentas_debe_compra(compra):
    mapa = {}

    if not compra:
        return mapa

    for mov in compra.get("cuentas", []):
        m = normalizar_mov(mov)
        cuenta = str(m.get("cuenta", "")).replace("A: ", "").strip()
        cuenta_lower = cuenta.lower()

        if not cuenta or m["debe"] <= 0:
            continue

        if cuenta_lower in {"iva crédito", "iva credito"}:
            continue

        mapa[cuenta] = round(mapa.get(cuenta, 0.0) + float(m["debe"] or 0), 2)

    return mapa


def resolver_traslados_egreso_ingreso(items, compra):
    items_limpios = limpiar_movimientos_ingreso_config(items, permitir_solo_cuenta=True)
    mapa_compra = obtener_mapa_cuentas_debe_compra(compra)
    mapa_compra_lower = {str(k).strip().lower(): round(float(v or 0), 2) for k, v in mapa_compra.items()}

    salida = []
    vistos = set()

    for item in items_limpios:
        cuenta = str(item.get("cuenta", "")).replace("A: ", "").strip()
        cuenta_lower = cuenta.lower()
        if not cuenta or cuenta_lower in vistos:
            continue

        monto = round(float(mapa_compra_lower.get(cuenta_lower, item.get("monto", 0) or 0)), 2)
        if monto <= 0:
            continue

        salida.append({
            "cuenta": cuenta,
            "monto": monto,
            "iva": round(monto * 0.12, 2),
        })
        vistos.add(cuenta_lower)

    return salida


def calcular_isr_mensual(ventas_netas_mes, regimen):
    if str(regimen).strip().lower() != "mensual":
        return 0.0
    ventas_netas_mes = float(ventas_netas_mes or 0)
    if ventas_netas_mes > 30000:
        return round(((ventas_netas_mes - 30000) * 0.07) + 1500, 2)
    return round(ventas_netas_mes * 0.05, 2)


def es_fin_trimestre(mes):
    return int(mes or 0) in {3, 6, 9, 12}


def obtener_rango_trimestre(mes):
    mes = int(mes or 0)
    if mes <= 0:
        return 1, 3
    inicio = ((mes - 1) // 3) * 3 + 1
    fin = inicio + 2
    return inicio, fin


def obtener_tipo_cuenta_resultado(cuenta):
    cuenta_base = limpiar_nombre_cuenta(cuenta).lower()
    if not cuenta_base:
        return ""

    for item in cargar_cuentas() or []:
        nombre = str(item.get("nombre", "") or "").strip().lower()
        if nombre != cuenta_base:
            continue

        tipo = str(item.get("tipo", "") or "").strip().lower()
        if tipo in {"ingreso", "gasto"}:
            return tipo

        texto = " ".join(
            str(item.get(k, "") or "").strip().lower()
            for k in ["tipo", "naturaleza", "clasificacion", "clasificación", "grupo", "categoria", "categoría", "rubro", "subtipo"]
        )
        if any(x in texto for x in ["ingreso", "venta", "producto", "servicio"]):
            return "ingreso"
        if any(x in texto for x in ["gasto", "costo"]):
            return "gasto"

    if any(x in cuenta_base for x in ["venta", "ventas", "ingreso", "ingresos", "servicio", "servicios", "producto", "productos"]):
        return "ingreso"
    if any(x in cuenta_base for x in ["gasto", "gastos", "costo", "costos", "compra", "compras", "isr", "sueldo", "sueldos", "salario", "salarios", "bonificacion", "bonificación", "igss", "irtra", "intecap", "prestacion", "prestación", "comision", "comisión"]):
        return "gasto"
    return ""


def obtener_ventas_netas_trimestre(ventas_data, empresa_id, anio, mes):
    inicio, fin = obtener_rango_trimestre(mes)
    total = 0.0
    meses_con_datos = []

    for mes_tmp in range(inicio, fin + 1):
        venta = buscar_partida(ventas_data, empresa_id, int(anio), mes_tmp)
        if not venta:
            continue
        netas_mes = round(float(ventas_netas(venta) or 0), 2)
        if netas_mes > 0:
            meses_con_datos.append(mes_tmp)
        total += netas_mes

    return round(total, 2), meses_con_datos, inicio, fin


def obtener_partidas_trimestre_para_cierre(partidas_generadas_data, empresa_id, anio, mes, partidas_actuales=None):
    inicio, fin = obtener_rango_trimestre(mes)
    acumuladas = []
    meses_faltantes = []

    for mes_tmp in range(inicio, fin + 1):
        if mes_tmp == int(mes):
            if partidas_actuales is None:
                continue
            if partidas_actuales:
                acumuladas.extend([clonar_partida(p) for p in partidas_actuales])
            else:
                meses_faltantes.append(mes_tmp)
            continue

        registro = buscar_partidas_generadas(partidas_generadas_data or [], empresa_id, int(anio), mes_tmp)
        if registro and registro.get("partidas"):
            acumuladas.extend([clonar_partida(p) for p in (registro.get("partidas") or [])])
        else:
            meses_faltantes.append(mes_tmp)

    return acumuladas, meses_faltantes, inicio, fin


def construir_partida_cierre_trimestral(
    empresa,
    empresa_id,
    anio,
    mes,
    codigo_num,
    partidas_trimestre,
    trimestre_inicio,
    trimestre_fin,
):
    acumulado = {}

    for partida in partidas_trimestre or []:
        for mov in partida.get("cuentas", []) or []:
            cuenta = limpiar_nombre_cuenta(mov.get("cuenta") or mov.get("nombre") or "")
            if not cuenta:
                continue

            cuenta_base = cuenta.lower()
            if cuenta_base in {"ganancia del ejercicio", "ganancia acumulada"}:
                continue

            tipo = obtener_tipo_cuenta_resultado(cuenta)
            if tipo not in {"ingreso", "gasto"}:
                continue

            key = cuenta.lower()
            if key not in acumulado:
                acumulado[key] = {
                    "cuenta": cuenta,
                    "tipo": tipo,
                    "debe": 0.0,
                    "haber": 0.0,
                }

            acumulado[key]["debe"] += round(float(mov.get("debe", 0) or 0), 2)
            acumulado[key]["haber"] += round(float(mov.get("haber", 0) or 0), 2)

    movimientos_cierre = []

    for item in sorted(acumulado.values(), key=lambda x: x["cuenta"].lower()):
        cuenta = item["cuenta"]
        debe = round(float(item.get("debe", 0) or 0), 2)
        haber = round(float(item.get("haber", 0) or 0), 2)
        tipo = item["tipo"]

        if tipo == "ingreso":
            saldo = round(haber - debe, 2)
            if abs(saldo) <= 0.009:
                continue
            if saldo > 0:
                movimientos_cierre.append({"cuenta": cuenta, "debe": saldo, "haber": 0.0})
            else:
                movimientos_cierre.append({"cuenta": f"A: {cuenta}", "debe": 0.0, "haber": abs(saldo)})
        else:
            saldo = round(debe - haber, 2)
            if abs(saldo) <= 0.009:
                continue
            if saldo > 0:
                movimientos_cierre.append({"cuenta": f"A: {cuenta}", "debe": 0.0, "haber": saldo})
            else:
                movimientos_cierre.append({"cuenta": cuenta, "debe": abs(saldo), "haber": 0.0})

    movimientos_cierre = consolidar_movimientos(movimientos_cierre)
    total_debe, total_haber = total_debe_haber(movimientos_cierre)
    diferencia = round(total_debe - total_haber, 2)
    monto_cierre = round(abs(diferencia), 2)

    if monto_cierre <= 0.009 and not movimientos_cierre:
        return None, 0.0, "sin_movimiento"

    if diferencia > 0:
        movimientos_cierre.append({"cuenta": "A: Ganancia del Ejercicio", "debe": 0.0, "haber": diferencia})
        naturaleza_cierre = "ganancia"
    elif diferencia < 0:
        movimientos_cierre.append({"cuenta": "Ganancia del Ejercicio", "debe": abs(diferencia), "haber": 0.0})
        naturaleza_cierre = "perdida"
    else:
        naturaleza_cierre = "cerrado"

    movimientos_cierre = consolidar_movimientos(movimientos_cierre)

    partida = base_partida(
        empresa,
        empresa_id,
        anio,
        mes,
        codigo_num,
        f"Cierre trimestral de resultados de {MESES[int(trimestre_inicio)]} a {MESES[int(trimestre_fin)]}",
        movimientos_cierre,
        extras={
            "tipo_partida": "cierre_trimestral",
            "trimestre_inicio": int(trimestre_inicio),
            "trimestre_fin": int(trimestre_fin),
            "monto_cierre_trimestral": monto_cierre,
            "naturaleza_cierre_trimestral": naturaleza_cierre,
        }
    )
    return partida, monto_cierre, naturaleza_cierre


def base_partida(empresa, empresa_id, anio, mes, codigo_num, glosa, cuentas, extras=None, auto_cuadrar_clientes=False, nombre_clientes="Clientes"):
    cuentas_limpias = consolidar_movimientos(cuentas)
    if auto_cuadrar_clientes:
        cuentas_limpias = consolidar_movimientos(
            cuadrar_con_clientes(cuentas_limpias, nombre_clientes=nombre_clientes)
        )
    debe, haber = total_debe_haber(cuentas_limpias)
    data = {
        "empresa": empresa,
        "empresa_nit": normalizar_nit(empresa_id),
        "anio": int(anio),
        "mes": int(mes),
        "fecha": fmt_fecha_fin_mes(anio, mes),
        "codigo": f"PDA {codigo_num}",
        "glosa": glosa,
        "cuentas": cuentas_limpias,
        "total_debe": debe,
        "total_haber": haber,
    }
    if extras:
        data.update(extras)
    return data

def clonar_partida(partida):
    return {
        "empresa": partida.get("empresa", ""),
        "empresa_nit": partida.get("empresa_nit", ""),
        "anio": int(partida.get("anio", 0) or 0),
        "mes": int(partida.get("mes", 0) or 0),
        "fecha": partida.get("fecha", ""),
        "codigo": str(partida.get("codigo", "")).strip(),
        "glosa": partida.get("glosa", ""),
        "cuentas": [
            {
                "cuenta": str(m.get("cuenta", "")),
                "debe": float(m.get("debe", 0) or 0),
                "haber": float(m.get("haber", 0) or 0),
                **(
                    {"anotacion": str(m.get("anotacion", "") or "").strip()}
                    if str(m.get("anotacion", "") or "").strip()
                    else {}
                ),
            }
            for m in partida.get("cuentas", [])
        ],
        "total_debe": float(partida.get("total_debe", 0) or 0),
        "total_haber": float(partida.get("total_haber", 0) or 0),
    }


def renumerar_partidas_visual(partidas):
    nuevas = []

    for i, p in enumerate(partidas, start=1):
        cp = dict(p)
        cp["codigo"] = f"PDA {i}"
        nuevas.append(cp)

    return nuevas


def es_cuenta_por_pagar(nombre_cuenta):
    base = str(nombre_cuenta or "").replace("A: ", "").strip().lower()

    if not base:
        return False

    return (
        "por pagar" in base
        or base.startswith("cuentas por pagar")
        or base.startswith("documentos por pagar")
    )


def calcular_saldos_por_pagar(partidas_fuente, excluir_cuentas=None):
    """
    Recibe una partida o una lista de partidas y devuelve saldos netos
    pendientes por pagar, calculados como HABER - DEBE.
    """
    excluir_cuentas = {
        str(x).replace("A: ", "").strip().lower()
        for x in (excluir_cuentas or [])
        if str(x).strip()
    }

    if not partidas_fuente:
        return []

    if isinstance(partidas_fuente, dict):
        partidas_fuente = [partidas_fuente]

    acumulado = {}
    orden = []

    for partida in partidas_fuente:
        for mov in partida.get("cuentas", []):
            cuenta_original = str(mov.get("cuenta", "")).strip()
            cuenta_base = cuenta_original.replace("A: ", "").strip()
            cuenta_base_lower = cuenta_base.lower()

            if not es_cuenta_por_pagar(cuenta_original):
                continue

            if cuenta_base_lower in excluir_cuentas:
                continue

            debe = round(float(mov.get("debe", 0) or 0), 2)
            haber = round(float(mov.get("haber", 0) or 0), 2)

            saldo = round(haber - debe, 2)  # saldo pendiente por pagar

            if cuenta_base not in acumulado:
                acumulado[cuenta_base] = 0.0
                orden.append(cuenta_base)

            acumulado[cuenta_base] += saldo

    salida = []
    for cuenta in orden:
        monto = round(acumulado[cuenta], 2)
        if monto > 0:
            salida.append({
                "cuenta": cuenta,
                "monto": monto
            })

    return salida


def obtener_saldos_por_pagar_para_pago(
    ventas_data,
    compras_data,
    partidas_generadas_data,
    empresa_id,
    anio,
    mes,
    excluir_cuentas=None
):
    """
    Enero:
        toma saldos por pagar desde la partida de apertura del año.
    Febrero en adelante:
        toma saldos por pagar desde el saldo final del mayor del mes anterior,
        calculado como apertura + movimientos guardados de ese mes anterior.
    """
    excluir = {
        str(x).replace("A: ", "").strip().lower()
        for x in (excluir_cuentas or [])
        if str(x).strip()
    }

    # ENERO -> apertura
    if int(mes) == 1:
        apertura = buscar_partida_apertura(ventas_data, empresa_id, anio)
        if not apertura:
            apertura = buscar_partida_apertura(compras_data, empresa_id, anio)

        apertura = normalizar_apertura_contrapartidas(apertura)

        return calcular_saldos_por_pagar(apertura, excluir_cuentas=excluir)

    # FEBRERO EN ADELANTE -> saldo final del mayor del mes anterior
    prev_mes = 12 if int(mes) == 1 else int(mes) - 1
    prev_anio = int(anio) - 1 if int(mes) == 1 else int(anio)

    return obtener_saldos_por_pagar_desde_mayor_mes(
        ventas_data=ventas_data,
        compras_data=compras_data,
        partidas_generadas_data=partidas_generadas_data,
        empresa_id=empresa_id,
        anio=prev_anio,
        mes=prev_mes,
        excluir_cuentas=excluir
    )


def insertar_apertura_en_enero(partidas, ventas_data, compras_data, empresa_id, anio, mes):
    partidas = partidas or []

    if int(mes) != 1:
        return partidas

    apertura = buscar_partida_apertura(ventas_data, empresa_id, anio)
    if not apertura:
        apertura = buscar_partida_apertura(compras_data, empresa_id, anio)

    # 🔥 aplicar normalización
    apertura = normalizar_apertura_contrapartidas(apertura)

    if not apertura:
        return partidas

    apertura_copia = clonar_partida(apertura)

    if not apertura_copia.get("fecha"):
        apertura_copia["fecha"] = "01/01/" + str(int(anio))

    if not apertura_copia.get("glosa"):
        apertura_copia["glosa"] = "Partida de apertura"

    partidas_con_apertura = [apertura_copia] + [clonar_partida(p) for p in partidas]

    return renumerar_partidas_visual(partidas_con_apertura)


def limpiar_distribucion(distribucion_cobro):
    salida = []
    for item in distribucion_cobro or []:
        cuenta = str(item.get("cuenta", "")).strip()
        monto = round(float(item.get("monto", 0) or 0), 2)
        if not cuenta or monto <= 0:
            continue
        salida.append({"cuenta": cuenta, "monto": monto})
    return salida

def total_distribucion(distribucion):
    distribucion = limpiar_distribucion(distribucion)
    return round(sum(float(x.get("monto", 0) or 0) for x in distribucion), 2)

def repartir_monto_en_cuentas(monto_total, distribucion):
    monto_total = round(float(monto_total or 0), 2)
    distribucion = limpiar_distribucion(distribucion)

    if monto_total <= 0 or not distribucion:
        return [], monto_total, distribucion

    restante = monto_total
    movimientos = []
    distribucion_restante = []

    for item in distribucion:
        cuenta = str(item.get("cuenta", "")).strip()
        disponible = round(float(item.get("monto", 0) or 0), 2)

        if not cuenta or disponible <= 0:
            continue

        usar = 0.0
        if restante > 0:
            usar = round(min(disponible, restante), 2)

        saldo = round(disponible - usar, 2)

        if usar > 0:
            movimientos.append({
                "cuenta": f"A: {cuenta}" if not cuenta.startswith("A: ") else cuenta,
                "debe": 0.0,
                "haber": usar
            })
            restante = round(restante - usar, 2)

        if saldo > 0:
            distribucion_restante.append({
                "cuenta": cuenta,
                "monto": saldo
            })

    return movimientos, restante, distribucion_restante


def construir_partidas_5(
    ventas_data, compras_data, empresa, empresa_id, anio, mes,
    regimen="Mensual", pda_inicial=1,
    incluir_ret_iva=False, monto_ret_iva=0.0,
    incluir_exencion_iva=False, monto_exencion_iva=0.0,
    incluir_ret_isr=False, monto_ret_isr=0.0,
    distribucion_cobro=None,
    modo_pago_planilla="mensual",
    sueldo_por_pagar_manual=0.0,
    bonificacion_por_pagar_manual=0.0,
    monto_retencion_isr_planilla=0.0,
    cuenta_egreso="Caja",
    activar_distribucion_egreso=False,
    distribucion_egreso=None,
    cuenta_cuadre_egreso="Proveedores",
    activar_distribucion_impuestos=False,
    distribucion_impuestos=None,
    gastos_extraordinarios=None,
    traslados_egreso_ingreso=None,
    ingresos_extra_manual=None,
    haberes_manuales_ingreso=None,
    haberes_manuales_egreso=None,
    usar_clientes_para_cuadre=True,
    cuenta_cuadre_ingreso="Clientes",
    cuenta_cobro_auto="Caja",
    activar_planilla=False,
    datos_planilla=None,
    activar_provision_prestaciones=True,
    porcentaje_prestaciones=26.38,
    partidas_generadas_data=None,
    incluir_saldos_por_pagar_anteriores=True,
    cuentas_excluir_pago_impuestos=None,
    activar_cierre_trimestral=False,
    usar_isr_trimestral_renta_bruta=False,
    usar_isr_trimestral_manual=False,
    monto_isr_trimestral_manual=0.0,
    generar_isr_trimestral=False,
    modo_isr_trimestral="ninguno",
    usar_isr_desde_cierres_parciales=False,
    activar_iso_trimestral=False,
    es_primer_anio_iso=False,
    monto_iso_trimestral_manual=0.0,
    acreditar_isr_a_iso=False,
):
    venta = buscar_partida(ventas_data, empresa_id, anio, mes)
    compra = buscar_partida(compras_data, empresa_id, anio, mes)

    apertura = buscar_partida_apertura(ventas_data, empresa_id, anio)
    if not apertura:
        apertura = buscar_partida_apertura(compras_data, empresa_id, anio)
        

    iva_apertura = max(
        abs(min(saldo_cuenta_en_apertura(apertura, ["iva por pagar"]), 0.0)),
        0.0
    )

    isr_apertura = max(
        abs(min(saldo_cuenta_en_apertura(apertura, ["isr por pagar"]), 0.0)),
        0.0
    )
    

    prev_mes = 12 if int(mes) == 1 else int(mes) - 1
    prev_anio = int(anio) - 1 if int(mes) == 1 else int(anio)
    venta_prev = buscar_partida(ventas_data, empresa_id, prev_anio, prev_mes)
    compra_prev = buscar_partida(compras_data, empresa_id, prev_anio, prev_mes)

    caja = cuenta_cobro_auto
    iva_debito_prev = monto_iva_debito(venta_prev) if venta_prev else 0.0
    iva_credito_prev = monto_iva_credito(compra_prev) if compra_prev else 0.0
    iva_por_pagar_prev_calculado = round(max(iva_debito_prev - iva_credito_prev, 0), 2)

    ventas_netas_prev = ventas_netas(venta_prev) if venta_prev else 0.0
    isr_prev_calculado = calcular_isr_mensual(ventas_netas_prev, regimen)

    # Enero: priorizar apertura. Febrero en adelante: usar el mayor del mes anterior
    # cuando está activada la opción de traer saldos por pagar anteriores.
    if int(mes) == 1:
        if apertura:
            iva_por_pagar_prev = round(iva_apertura, 2)
            isr_prev = round(isr_apertura, 2)
        else:
            iva_por_pagar_prev = round(iva_por_pagar_prev_calculado, 2)
            isr_prev = round(isr_prev_calculado, 2)
    else:
        if incluir_saldos_por_pagar_anteriores:
            iva_por_pagar_prev = 0.0
            isr_prev = 0.0
        else:
            iva_por_pagar_prev = round(iva_por_pagar_prev_calculado, 2)
            isr_prev = round(isr_prev_calculado, 2)

    saldo_iva_por_cobrar_prev, saldo_credito_retencion_iva_prev = obtener_arrastres_iva_para_regularizacion(
        ventas_data=ventas_data,
        compras_data=compras_data,
        partidas_generadas_data=partidas_generadas_data,
        empresa_id=empresa_id,
        anio=anio,
        mes=mes,
    )

    partidas = []
    n = int(pda_inicial)

    cuentas_pago = []

    saldos_por_pagar_anteriores = []
    if incluir_saldos_por_pagar_anteriores:
        saldos_por_pagar_anteriores = obtener_saldos_por_pagar_para_pago(
            ventas_data=ventas_data,
            compras_data=compras_data,
            partidas_generadas_data=partidas_generadas_data,
            empresa_id=empresa_id,
            anio=anio,
            mes=mes,
            excluir_cuentas=cuentas_excluir_pago_impuestos or []
        )
    if (
        not venta
        and not compra
        and iva_por_pagar_prev <= 0
        and isr_prev <= 0
        and not saldos_por_pagar_anteriores
    ):
        return []

    cuentas_excluidas_pago = {
        str(x).replace("A: ", "").strip().lower()
        for x in (cuentas_excluir_pago_impuestos or [])
        if str(x).strip()
    }

    cuentas_ya_agregadas_pago = {
        str(item.get("cuenta", "")).replace("A: ", "").strip().lower()
        for item in saldos_por_pagar_anteriores
    }

    if (
        isr_prev > 0
        and "isr por pagar" not in cuentas_ya_agregadas_pago
        and "isr por pagar" not in cuentas_excluidas_pago
    ):
        cuentas_pago.append({
            "cuenta": "ISR por pagar",
            "debe": isr_prev,
            "haber": 0.0
        })

    if (
        iva_por_pagar_prev > 0
        and "iva por pagar" not in cuentas_ya_agregadas_pago
        and "iva por pagar" not in cuentas_excluidas_pago
    ):
        cuentas_pago.append({
            "cuenta": "IVA por pagar",
            "debe": iva_por_pagar_prev,
            "haber": 0.0
        })

    for item in saldos_por_pagar_anteriores:
        cuentas_pago.append({
            "cuenta": item["cuenta"],
            "debe": round(float(item.get("monto", 0) or 0), 2),
            "haber": 0.0
        })

    total_saldos_por_pagar_anteriores = round(
        sum(float(x.get("monto", 0) or 0) for x in saldos_por_pagar_anteriores),
        2
    )

    total_pago = round(
        sum(float(x.get("debe", 0) or 0) for x in cuentas_pago),
        2
    )
    monto_pago_impuestos_usado = 0.0

    distribucion_egreso_restante = limpiar_distribucion(distribucion_egreso)
    distribucion_impuestos_limpia = limpiar_distribucion(distribucion_impuestos)

    if total_pago > 0:
        if activar_distribucion_impuestos:
            movimientos_pago, restante_pago, _ = repartir_monto_en_cuentas(
                total_pago,
                distribucion_impuestos_limpia
            )
            cuentas_pago.extend(movimientos_pago)

            if restante_pago > 0:
                cuentas_pago.append({
                    "cuenta": f"A: {cuenta_egreso}",
                    "debe": 0.0,
                    "haber": restante_pago
                })

            monto_pago_impuestos_usado = round(total_pago - restante_pago, 2)

        elif activar_distribucion_egreso:
            movimientos_pago, restante_pago, distribucion_egreso_restante = repartir_monto_en_cuentas(
                total_pago,
                distribucion_egreso
            )
            cuentas_pago.extend(movimientos_pago)

            if restante_pago > 0:
                cuentas_pago.append({
                    "cuenta": f"A: {cuenta_egreso}",
                    "debe": 0.0,
                    "haber": restante_pago
                })

            monto_pago_impuestos_usado = round(total_pago - restante_pago, 2)

        else:
            cuentas_pago.append({
                "cuenta": f"A: {cuenta_egreso}",
                "debe": 0.0,
                "haber": total_pago
            })
            monto_pago_impuestos_usado = total_pago

        partidas.append(
            base_partida(
                empresa,
                empresa_id,
                anio,
                mes,
                n,
                "Registro del pago de impuestos del mes anterior",
                cuentas_pago,
                extras={
                    "iva_apertura_detectado": iva_apertura,
                    "isr_apertura_detectado": isr_apertura,
                    "iva_por_pagar_prev": iva_por_pagar_prev,
                    "isr_prev": isr_prev,
                    "saldos_por_pagar_anteriores": saldos_por_pagar_anteriores,
                    "total_saldos_por_pagar_anteriores": total_saldos_por_pagar_anteriores,
                    "cuentas_excluir_pago_impuestos": cuentas_excluir_pago_impuestos or [],
                }
            )
        )
    n += 1

    comision = obtener_comision_desde_compras(compra)
    iva_comision = calcular_iva_comision(comision)

    traslados_egreso_ingreso = resolver_traslados_egreso_ingreso(traslados_egreso_ingreso, compra)
    ingresos_extra_manual = limpiar_movimientos_ingreso_config(ingresos_extra_manual)
    haberes_manuales_ingreso = limpiar_haberes_manuales(haberes_manuales_ingreso)
    haberes_manuales_egreso = limpiar_haberes_manuales(haberes_manuales_egreso)

    cuentas_traslado_lower = {
        str(item.get("cuenta", "")).replace("A: ", "").strip().lower()
        for item in traslados_egreso_ingreso
        if str(item.get("cuenta", "")).strip()
    }

    traslados_aplicados = list(traslados_egreso_ingreso)
    if comision > 0 and "comisiones pagadas" not in cuentas_traslado_lower:
        traslados_aplicados.append({
            "cuenta": "Comisiones Pagadas",
            "monto": comision,
            "iva": iva_comision,
        })

    total_traslados = round(sum(float(x.get("monto", 0) or 0) for x in traslados_aplicados), 2)
    total_iva_traslados = round(sum(float(x.get("iva", 0) or 0) for x in traslados_aplicados), 2)
    total_ingresos_extra_manual = round(sum(float(x.get("monto", 0) or 0) for x in ingresos_extra_manual), 2)
    total_iva_ingresos_extra_manual = round(sum(float(x.get("iva", 0) or 0) for x in ingresos_extra_manual), 2)
    total_haberes_manuales_ingreso = round(sum(float(x.get("monto", 0) or 0) for x in haberes_manuales_ingreso), 2)
    total_haberes_manuales_egreso = round(sum(float(x.get("monto", 0) or 0) for x in haberes_manuales_egreso), 2)

    traslados_pendientes_por_cuenta = {}
    for item in traslados_aplicados:
        cuenta_base = str(item.get("cuenta", "")).replace("A: ", "").strip().lower()
        if not cuenta_base:
            continue
        traslados_pendientes_por_cuenta[cuenta_base] = round(
            traslados_pendientes_por_cuenta.get(cuenta_base, 0.0) + float(item.get("monto", 0) or 0),
            2
        )

    if venta:
        total_cobrado = total_cobrado_venta(venta)
        monto_ret_iva = round(float(monto_ret_iva or 0), 2) if incluir_ret_iva else 0.0
        monto_exencion_iva = round(float(monto_exencion_iva or 0), 2) if incluir_exencion_iva else 0.0
        monto_ret_isr = round(float(monto_ret_isr or 0), 2) if incluir_ret_isr else 0.0

        movimientos = []
        distribucion = limpiar_distribucion(distribucion_cobro)

        if distribucion:
            for item in distribucion:
                movimientos.append({"cuenta": item["cuenta"], "debe": item["monto"], "haber": 0.0})
        else:
            caja_neta = round(
                total_cobrado
                - monto_ret_iva
                - monto_exencion_iva
                - monto_ret_isr
                - total_traslados
                - total_iva_traslados
                - total_ingresos_extra_manual
                - total_iva_ingresos_extra_manual,
                2
            )
            if caja_neta > 0:
                movimientos.append({"cuenta": caja, "debe": caja_neta, "haber": 0.0})

        for item in traslados_aplicados:
            monto_item = round(float(item.get("monto", 0) or 0), 2)
            iva_item = round(float(item.get("iva", 0) or 0), 2)
            cuenta_item = str(item.get("cuenta", "")).replace("A: ", "").strip()

            if monto_item > 0 and cuenta_item:
                movimientos.append({"cuenta": cuenta_item, "debe": monto_item, "haber": 0.0})
            if iva_item > 0:
                movimientos.append({"cuenta": "IVA Crédito", "debe": iva_item, "haber": 0.0})

        for item in ingresos_extra_manual:
            monto_item = round(float(item.get("monto", 0) or 0), 2)
            iva_item = round(float(item.get("iva", 0) or 0), 2)
            cuenta_item = str(item.get("cuenta", "")).replace("A: ", "").strip()

            if monto_item > 0 and cuenta_item:
                movimientos.append({"cuenta": cuenta_item, "debe": monto_item, "haber": 0.0})
            if iva_item > 0:
                movimientos.append({"cuenta": "IVA Crédito", "debe": iva_item, "haber": 0.0})

        if monto_ret_iva > 0:
            movimientos.append({"cuenta": "Retención IVA", "debe": monto_ret_iva, "haber": 0.0})
        if monto_exencion_iva > 0:
            movimientos.append({"cuenta": "Exencion IVA", "debe": monto_exencion_iva, "haber": 0.0})
        if monto_ret_isr > 0:
            movimientos.append({"cuenta": "Retención ISR", "debe": monto_ret_isr, "haber": 0.0})

        for m in [normalizar_mov(x) for x in venta.get("cuentas", [])]:
            if abs(m["debe"]) < 1e-9 and abs(m["haber"]) < 1e-9:
                continue
            nombre_base = m["cuenta"].replace("A: ", "").strip().lower()
            if nombre_base in CUENTAS_CAJA or nombre_base in CUENTAS_CXC:
                continue
            c = m["cuenta"]
            if m["haber"] > 0 and not c.startswith("A: "):
                c = f"A: {c}"
            movimientos.append({"cuenta": c, "debe": m["debe"], "haber": m["haber"]})

        for item in haberes_manuales_ingreso:
            cuenta_haber = str(item.get("cuenta", "")).replace("A: ", "").strip()
            monto_haber = round(float(item.get("monto", 0) or 0), 2)

            if not cuenta_haber or monto_haber <= 0:
                continue

            movimientos.append({
                "cuenta": f"A: {cuenta_haber}",
                "debe": 0.0,
                "haber": monto_haber
            })

        descuadre_antes_ingreso = descuadre_partida(movimientos)
        cuadre_info_ingreso = {
            "diferencia_original": round(descuadre_antes_ingreso, 2),
            "monto_usado_principal": 0.0,
            "monto_usado_respaldo": 0.0,
            "cuenta_respaldo_haber": "Anticipo S/Ventas",
            "cuenta_respaldo_debe": "Clientes",
            "redujo_cuenta_principal": False,
            "saldo_disponible": 0.0,
        }
        if usar_clientes_para_cuadre:
            movimientos, cuadre_info_ingreso = cuadrar_con_saldo_real(
                cuentas=movimientos,
                cuenta_cuadre=cuenta_cuadre_ingreso,
                empresa_id=empresa_id,
                anio=anio,
                mes=mes,
                ventas_data=ventas_data,
                compras_data=compras_data,
                partidas_generadas_data=partidas_generadas_data,
                cuenta_respaldo_debe="Clientes",
                cuenta_respaldo_haber="Anticipo S/Ventas",
            )

        partidas.append(
            base_partida(
                empresa, empresa_id, anio, mes, n,
                "Registro de los ingresos obtenidos en el presente mes",
                movimientos,
                extras={
                    "retencion_iva_manual": monto_ret_iva,
                    "exencion_iva_manual": monto_exencion_iva,
                    "retencion_isr_manual": monto_ret_isr,
                    "comision_detectada": comision,
                    "iva_comision_calculado": iva_comision,
                    "traslados_egreso_ingreso": traslados_aplicados,
                    "total_traslados_egreso_ingreso": total_traslados,
                    "total_iva_traslados_egreso_ingreso": total_iva_traslados,
                    "ingresos_extra_manual": ingresos_extra_manual,
                    "total_ingresos_extra_manual": total_ingresos_extra_manual,
                    "total_iva_ingresos_extra_manual": total_iva_ingresos_extra_manual,
                    "haberes_manuales_ingreso": haberes_manuales_ingreso,
                    "total_haberes_manuales_ingreso": total_haberes_manuales_ingreso,
                    "regimen": regimen,
                    "total_cobrado_referencia": total_cobrado,
                    "distribucion_cobro_manual": distribucion,
                    "descuadre_antes_clientes": descuadre_antes_ingreso,
                    "cuadre_inteligente_ingreso": cuadre_info_ingreso,
                },
                auto_cuadrar_clientes=False,
                nombre_clientes=cuenta_cuadre_ingreso,
            )
        )
    n += 1

    if compra:
        movimientos = []
        iva_credito_total = monto_iva_credito(compra)
        iva_credito_egresos = round(max(iva_credito_total - total_iva_traslados, 0), 2)
        iva_credito_reemplazado = False

        for m in [normalizar_mov(x) for x in compra.get("cuentas", [])]:
            nombre = m["cuenta"].strip().lower()
            nombre_base = m["cuenta"].replace("A: ", "").strip().lower()

            if abs(m["debe"]) < 1e-9 and abs(m["haber"]) < 1e-9:
                continue

            if m["debe"] > 0 and nombre_base in traslados_pendientes_por_cuenta:
                monto_pendiente = round(float(traslados_pendientes_por_cuenta.get(nombre_base, 0) or 0), 2)
                if monto_pendiente > 0:
                    aplicar = round(min(m["debe"], monto_pendiente), 2)
                    m["debe"] = round(max(m["debe"] - aplicar, 0), 2)
                    traslados_pendientes_por_cuenta[nombre_base] = round(monto_pendiente - aplicar, 2)

            if abs(m["debe"]) < 1e-9 and abs(m["haber"]) < 1e-9:
                continue

            if nombre in {"iva crédito", "iva credito"}:
                if iva_credito_reemplazado:
                    continue
                iva_credito_reemplazado = True
                if iva_credito_egresos > 0:
                    movimientos.append({
                        "cuenta": m["cuenta"],
                        "debe": iva_credito_egresos,
                        "haber": 0.0
                    })
                continue

            c = m["cuenta"]
            base_c = c.replace("A: ", "").strip().lower()

            # ya no tomar las cuentas de salida originales
            if m["haber"] > 0 and base_c in CUENTAS_CAJA:
                continue

            if m["haber"] > 0 and not c.startswith("A: "):
                c = f"A: {c}"

            movimientos.append({
                "cuenta": c,
                "debe": m["debe"],
                "haber": m["haber"]
            })

        gastos_extraordinarios = gastos_extraordinarios or []
        
        gastos_extra_limpios = []

        for item in gastos_extraordinarios:
            cuenta = str(item.get("cuenta", "")).strip()
            monto = round(float(item.get("monto", 0) or 0), 2)
            if cuenta and monto > 0:
                gastos_extra_limpios.append({
                    "cuenta": cuenta,
                    "monto": monto
                })

        total_gastos_extra = round(sum(x["monto"] for x in gastos_extra_limpios), 2)

        for gx in gastos_extra_limpios:
            movimientos.append({
                "cuenta": gx["cuenta"],
                "debe": gx["monto"],
                "haber": 0.0
            })

        if activar_planilla and datos_planilla:
            movimientos.extend(
                movimientos_planilla_laboral(
                    datos_planilla,
                    modo_pago=modo_pago_planilla,
                    sueldo_por_pagar_manual=sueldo_por_pagar_manual,
                    bonificacion_por_pagar_manual=bonificacion_por_pagar_manual,
                    monto_retencion_isr_planilla=monto_retencion_isr_planilla,
                )
            )
        total_egreso_configurado = 0.0
        total_egreso_bruto_antes_haberes_manuales = 0.0
        total_haberes_manuales_egreso_aplicado = 0.0

        if activar_distribucion_egreso:
            total_egreso_bruto_antes_haberes_manuales = total_distribucion(distribucion_egreso_restante)
            total_haberes_manuales_egreso_aplicado = 0.0
            total_egreso_configurado = round(total_egreso_bruto_antes_haberes_manuales, 2)

            if total_egreso_configurado > 0:
                movimientos_egreso, _, _ = repartir_monto_en_cuentas(
                    total_egreso_configurado,
                    distribucion_egreso_restante
                )
                movimientos.extend(movimientos_egreso)

        else:
            total_egreso_bruto_antes_haberes_manuales = 0.0
            for m in [normalizar_mov(x) for x in compra.get("cuentas", [])]:
                c = m["cuenta"]
                base_c = c.replace("A: ", "").strip().lower()

                if m["haber"] > 0 and base_c in CUENTAS_CAJA:
                    total_egreso_bruto_antes_haberes_manuales += m["haber"]

            total_egreso_bruto_antes_haberes_manuales = round(
                max(total_egreso_bruto_antes_haberes_manuales - monto_pago_impuestos_usado, 0) + total_gastos_extra,
                2
            )
            total_haberes_manuales_egreso_aplicado = 0.0
            total_egreso_configurado = round(total_egreso_bruto_antes_haberes_manuales, 2)

            if total_egreso_configurado > 0:
                movimientos.append({
                    "cuenta": f"A: {cuenta_egreso}",
                    "debe": 0.0,
                    "haber": total_egreso_configurado
                })

        for item in haberes_manuales_egreso:
            cuenta_haber = str(item.get("cuenta", "")).replace("A: ", "").strip()
            monto_haber = round(float(item.get("monto", 0) or 0), 2)

            if not cuenta_haber or monto_haber <= 0:
                continue

            movimientos.append({
                "cuenta": f"A: {cuenta_haber}",
                "debe": 0.0,
                "haber": monto_haber
            })

        descuadre_antes_egreso = descuadre_partida(movimientos)
        movimientos, cuadre_info_egreso = cuadrar_con_saldo_real(
            cuentas=movimientos,
            cuenta_cuadre=cuenta_cuadre_egreso,
            empresa_id=empresa_id,
            anio=anio,
            mes=mes,
            ventas_data=ventas_data,
            compras_data=compras_data,
            partidas_generadas_data=partidas_generadas_data,
            cuenta_respaldo_debe="Anticipo S/Compras",
            cuenta_respaldo_haber="Anticipo S/Compras",
        )
        movimientos = consolidar_movimientos(movimientos)

        partidas.append(
            base_partida(
                empresa, empresa_id, anio, mes, n,
                "Registro de los egresos efectuados en el presente mes",
                movimientos,
                extras={
                    "iva_credito_total_compras": iva_credito_total,
                    "iva_credito_egresos": iva_credito_egresos,
                    "monto_pago_impuestos_usado": monto_pago_impuestos_usado,
                    "total_egreso_bruto_antes_haberes_manuales": total_egreso_bruto_antes_haberes_manuales,
                    "total_haberes_manuales_egreso_aplicado": total_haberes_manuales_egreso_aplicado,
                    "total_haberes_manuales_egreso_excedente": round(
                        max(total_haberes_manuales_egreso - total_haberes_manuales_egreso_aplicado, 0),
                        2
                    ),
                    "total_egreso_configurado": total_egreso_configurado,
                    "cuenta_cuadre_egreso": cuenta_cuadre_egreso,
                    "descuadre_antes_egreso": descuadre_antes_egreso,
                    "cuadre_inteligente_egreso": cuadre_info_egreso,
                    "gastos_extraordinarios": gastos_extra_limpios,
                    "haberes_manuales_egreso": haberes_manuales_egreso,
                    "total_haberes_manuales_egreso": total_haberes_manuales_egreso,
                    "traslados_egreso_ingreso_aplicados": traslados_aplicados,
                    "traslados_egreso_ingreso_pendientes": [
                        {
                            "cuenta": cuenta,
                            "monto": round(monto, 2)
                        }
                        for cuenta, monto in traslados_pendientes_por_cuenta.items()
                        if round(float(monto or 0), 2) > 0
                    ],
                }
            )
        )
    n += 1

    if activar_planilla and datos_planilla and activar_provision_prestaciones:
        salario_base_total_planilla = round(float(datos_planilla.get("salario_base_total", 0) or 0), 2)
        porcentaje_prestaciones_decimal = round(float(porcentaje_prestaciones or 0) / 100, 6)

        monto_provision_prestaciones = calcular_provision_prestaciones(
            salario_base_total_planilla,
            porcentaje=porcentaje_prestaciones_decimal
        )

        if monto_provision_prestaciones > 0:
            cuentas_prestaciones = [
                {
                    "cuenta": "Prestaciones Laborales",
                    "debe": monto_provision_prestaciones,
                    "haber": 0.0
                },
                {
                    "cuenta": "A: Provisión para Prestaciones Laborales",
                    "debe": 0.0,
                    "haber": monto_provision_prestaciones
                }
            ]

            partidas.append(
                base_partida(
                    empresa,
                    empresa_id,
                    anio,
                    mes,
                    n,
                    "Registro de la provisión mensual de prestaciones laborales",
                    cuentas_prestaciones,
                    extras={
                        "salario_base_planilla": salario_base_total_planilla,
                        "porcentaje_prestaciones": round(float(porcentaje_prestaciones or 0), 2),
                        "monto_provision_prestaciones": monto_provision_prestaciones
                    }
                )
            )
            n += 1

    if venta and str(regimen).strip().lower() == "mensual":
        v_netas = ventas_netas(venta)
        isr_mes = calcular_isr_mensual(v_netas, regimen)
        monto_ret_isr = round(float(monto_ret_isr or 0), 2) if incluir_ret_isr else 0.0
        if isr_mes > 0:
            cuentas_isr = [{"cuenta": "ISR Mensual", "debe": isr_mes, "haber": 0.0}]
            aplicado = min(isr_mes, monto_ret_isr) if monto_ret_isr > 0 else 0.0
            if aplicado > 0:
                cuentas_isr.append({"cuenta": "A: Retención ISR", "debe": 0.0, "haber": aplicado})
            neto_isr = round(isr_mes - aplicado, 2)
            if neto_isr > 0:
                cuentas_isr.append({"cuenta": "A: ISR por pagar", "debe": 0.0, "haber": neto_isr})

            partidas.append(
                base_partida(
                    empresa, empresa_id, anio, mes, n,
                    "Registro del ISR por pagar del presente mes",
                    cuentas_isr,
                    extras={"isr_calculado": isr_mes, "regimen": regimen}
                )
            )
    n += 1

    if venta or compra or saldo_iva_por_cobrar_prev > 0 or saldo_credito_retencion_iva_prev > 0:
        iva_debito = monto_iva_debito(venta) if venta else 0.0
        iva_credito_total = monto_iva_credito(compra) if compra else 0.0
        monto_ret_iva = round(float(monto_ret_iva or 0), 2) if incluir_ret_iva else 0.0
        monto_exencion_iva = round(float(monto_exencion_iva or 0), 2) if incluir_exencion_iva else 0.0

        if (
            iva_debito > 0
            or iva_credito_total > 0
            or monto_ret_iva > 0
            or monto_exencion_iva > 0
            or saldo_iva_por_cobrar_prev > 0
            or saldo_credito_retencion_iva_prev > 0
        ):
            cuentas_iva = []
            if iva_debito > 0:
                cuentas_iva.append({"cuenta": "IVA Débito", "debe": iva_debito, "haber": 0.0})

            restante_por_regularizar = round(iva_debito, 2)

            aplicado_iva_credito = round(min(iva_credito_total, restante_por_regularizar), 2)
            restante_por_regularizar = round(restante_por_regularizar - aplicado_iva_credito, 2)

            aplicado_iva_por_cobrar_prev = round(min(saldo_iva_por_cobrar_prev, restante_por_regularizar), 2)
            restante_por_regularizar = round(restante_por_regularizar - aplicado_iva_por_cobrar_prev, 2)

            aplicado_credito_retencion_prev = round(min(saldo_credito_retencion_iva_prev, restante_por_regularizar), 2)
            restante_por_regularizar = round(restante_por_regularizar - aplicado_credito_retencion_prev, 2)

            aplicado_ret_iva_mes = round(min(monto_ret_iva, restante_por_regularizar), 2)
            restante_por_regularizar = round(restante_por_regularizar - aplicado_ret_iva_mes, 2)

            aplicado_exencion_iva_mes = round(min(monto_exencion_iva, restante_por_regularizar), 2)
            restante_por_regularizar = round(restante_por_regularizar - aplicado_exencion_iva_mes, 2)

            remanente_iva_credito = round(iva_credito_total - aplicado_iva_credito, 2)
            remanente_ret_iva_mes = round(monto_ret_iva - aplicado_ret_iva_mes, 2)
            remanente_exencion_iva_mes = round(monto_exencion_iva - aplicado_exencion_iva_mes, 2)
            remanente_iva_por_cobrar = round(remanente_iva_credito + remanente_exencion_iva_mes, 2)

            # Neto del movimiento de la cuenta Crédito de Retención IVA:
            # - si se consume más saldo previo del que ingresa este mes, queda en HABER (A: Crédito...)
            # - si ingresa más de lo que se consumió del saldo previo, queda en DEBE (Crédito...)
            neto_credito_retencion_iva = round(
                aplicado_credito_retencion_prev - remanente_ret_iva_mes,
                2
            )
            if remanente_iva_por_cobrar > 0:
                cuentas_iva.append({"cuenta": "IVA por cobrar", "debe": remanente_iva_por_cobrar, "haber": 0.0})
            if neto_credito_retencion_iva < 0:
                cuentas_iva.append({
                    "cuenta": "Crédito de Retención IVA",
                    "debe": abs(neto_credito_retencion_iva),
                    "haber": 0.0
                })

            if iva_credito_total > 0:
                cuentas_iva.append({"cuenta": "A: IVA Crédito", "debe": 0.0, "haber": iva_credito_total})
            if aplicado_iva_por_cobrar_prev > 0:
                cuentas_iva.append({"cuenta": "A: IVA por cobrar", "debe": 0.0, "haber": aplicado_iva_por_cobrar_prev})
            if monto_ret_iva > 0:
                cuentas_iva.append({"cuenta": "A: Retención IVA", "debe": 0.0, "haber": monto_ret_iva})
            if monto_exencion_iva > 0:
                cuentas_iva.append({"cuenta": "A: Exencion IVA", "debe": 0.0, "haber": monto_exencion_iva})
            if neto_credito_retencion_iva > 0:
                cuentas_iva.append({
                    "cuenta": "A: Crédito de Retención IVA",
                    "debe": 0.0,
                    "haber": neto_credito_retencion_iva
                })
            if restante_por_regularizar > 0:
                cuentas_iva.append({"cuenta": "A: IVA por pagar", "debe": 0.0, "haber": restante_por_regularizar})

            partidas.append(
                base_partida(
                    empresa, empresa_id, anio, mes, n,
                    "Regularización del IVA del presente mes",
                    cuentas_iva,
                    extras={
                        "saldo_iva_por_cobrar_prev": saldo_iva_por_cobrar_prev,
                        "saldo_credito_retencion_iva_prev": saldo_credito_retencion_iva_prev,
                        "aplicado_iva_credito": aplicado_iva_credito,
                        "aplicado_iva_por_cobrar_prev": aplicado_iva_por_cobrar_prev,
                        "aplicado_credito_retencion_iva_prev": aplicado_credito_retencion_prev,
                        "aplicado_retencion_iva_mes": aplicado_ret_iva_mes,
                        "aplicado_exencion_iva_mes": aplicado_exencion_iva_mes,
                        "remanente_iva_credito": remanente_iva_credito,
                        "remanente_retencion_iva_mes": remanente_ret_iva_mes,
                        "remanente_exencion_iva_mes": remanente_exencion_iva_mes,
                        "remanente_iva_por_cobrar": remanente_iva_por_cobrar,
                        "neto_credito_retencion_iva": neto_credito_retencion_iva,
                    }
                )
            )

    if str(regimen).strip().lower() == "trimestral" and es_fin_trimestre(mes):
        ventas_netas_trim, meses_con_ventas_trim, trimestre_inicio, trimestre_fin = obtener_ventas_netas_trimestre(
            ventas_data,
            empresa_id,
            anio,
            mes,
        )

        modo_isr_trimestral_final = str(modo_isr_trimestral or "").strip().lower()
        if modo_isr_trimestral_final not in {"cierres_parciales", "renta_bruta", "manual", "ninguno"}:
            if usar_isr_trimestral_manual:
                modo_isr_trimestral_final = "manual"
            elif usar_isr_trimestral_renta_bruta:
                modo_isr_trimestral_final = "renta_bruta"
            elif usar_isr_desde_cierres_parciales:
                modo_isr_trimestral_final = "cierres_parciales"
            else:
                modo_isr_trimestral_final = "ninguno"

        isr_trimestral_base = 0.0
        formula_isr_trimestral = ""
        registro_isr_cierre = None

        if generar_isr_trimestral and modo_isr_trimestral_final == "manual":
            isr_trimestral_base = round(float(monto_isr_trimestral_manual or 0), 2)
            formula_isr_trimestral = "manual"
        elif generar_isr_trimestral and modo_isr_trimestral_final == "renta_bruta":
            isr_trimestral_base = round(float(ventas_netas_trim or 0) * 0.08 * 0.25, 2)
            formula_isr_trimestral = "ventas_netas_trimestre * 8% * 25%"
        elif generar_isr_trimestral and modo_isr_trimestral_final == "cierres_parciales":
            registro_isr_cierre = obtener_isr_trimestral_desde_cierres(empresa_id, anio, mes)
            if registro_isr_cierre:
                isr_trimestral_base = round(float(registro_isr_cierre.get("isr_trimestral_por_pagar", 0) or 0), 2)
                formula_isr_trimestral = "isr_trimestral_por_pagar guardado en cierres parciales"

        saldo_iso_por_acreditar = obtener_saldo_iso_por_acreditar_hasta_fecha(
            ventas_data=ventas_data,
            compras_data=compras_data,
            partidas_generadas_data=partidas_generadas_data,
            empresa_id=empresa_id,
            anio=anio,
            mes=mes,
            incluir_mes_actual=False,
        )
        aplicado_iso_por_acreditar_isr = round(min(saldo_iso_por_acreditar, isr_trimestral_base), 2)
        isr_trimestral_por_pagar = round(max(isr_trimestral_base - aplicado_iso_por_acreditar_isr, 0.0), 2)

        if generar_isr_trimestral and isr_trimestral_base > 0:
            cuentas_isr_trimestral = [
                {"cuenta": "ISR Trimestral", "debe": isr_trimestral_base, "haber": 0.0},
            ]
            if aplicado_iso_por_acreditar_isr > 0:
                cuentas_isr_trimestral.append({
                    "cuenta": "A: ISO por acreditar",
                    "debe": 0.0,
                    "haber": aplicado_iso_por_acreditar_isr
                })
            if isr_trimestral_por_pagar > 0:
                cuentas_isr_trimestral.append({
                    "cuenta": "A: ISR Trimestral por pagar",
                    "debe": 0.0,
                    "haber": isr_trimestral_por_pagar
                })

            partidas.append(
                base_partida(
                    empresa,
                    empresa_id,
                    anio,
                    mes,
                    n,
                    f"Registro del ISR trimestral de {MESES[int(trimestre_inicio)]} a {MESES[int(trimestre_fin)]}",
                    cuentas_isr_trimestral,
                    extras={
                        "tipo_partida": "isr_trimestral",
                        "regimen": regimen,
                        "trimestre_inicio": int(trimestre_inicio),
                        "trimestre_fin": int(trimestre_fin),
                        "ventas_netas_trimestre": round(float(ventas_netas_trim or 0), 2),
                        "meses_con_ventas_trimestre": meses_con_ventas_trim,
                        "modo_isr_trimestral": modo_isr_trimestral_final,
                        "isr_trimestral_base": isr_trimestral_base,
                        "formula_isr_trimestral": formula_isr_trimestral,
                        "saldo_iso_por_acreditar": saldo_iso_por_acreditar,
                        "aplicado_iso_por_acreditar_isr": aplicado_iso_por_acreditar_isr,
                        "isr_trimestral_por_pagar": isr_trimestral_por_pagar,
                        "isr_fuente_cierre_parcial": registro_isr_cierre or {},
                    }
                )
            )
            n += 1

        iso_trimestral_base = round(float(monto_iso_trimestral_manual or 0), 2)
        aplicado_isr_a_iso = 0.0
        iso_trimestral_por_pagar = 0.0

        if activar_iso_trimestral and not es_primer_anio_iso and iso_trimestral_base > 0:
            if acreditar_isr_a_iso:
                aplicado_isr_a_iso = round(min(iso_trimestral_base, isr_trimestral_por_pagar), 2)
            iso_trimestral_por_pagar = round(max(iso_trimestral_base - aplicado_isr_a_iso, 0.0), 2)

            if iso_trimestral_por_pagar > 0:
                cuentas_iso_trimestral = [
                    {"cuenta": "ISO", "debe": iso_trimestral_por_pagar, "haber": 0.0},
                    {"cuenta": "A: ISO Trimestral por pagar", "debe": 0.0, "haber": iso_trimestral_por_pagar},
                ]

                partidas.append(
                    base_partida(
                        empresa,
                        empresa_id,
                        anio,
                        mes,
                        n,
                        f"Registro del ISO trimestral de {MESES[int(trimestre_inicio)]} a {MESES[int(trimestre_fin)]}",
                        cuentas_iso_trimestral,
                        extras={
                            "tipo_partida": "iso_trimestral",
                            "regimen": regimen,
                            "trimestre_inicio": int(trimestre_inicio),
                            "trimestre_fin": int(trimestre_fin),
                            "iso_trimestral_base": iso_trimestral_base,
                            "acreditar_isr_a_iso": bool(acreditar_isr_a_iso),
                            "aplicado_isr_a_iso": aplicado_isr_a_iso,
                            "iso_trimestral_por_pagar": iso_trimestral_por_pagar,
                            "isr_trimestral_por_pagar_periodo": isr_trimestral_por_pagar,
                        }
                    )
                )
                n += 1

    return partidas


def a_dataframe_visual(partidas):
    filas = []
    for p in partidas:
        filas.append({"Partida": p["codigo"], "Fecha": p["fecha"], "Cuenta": "", "Debe": "", "Haber": ""})
        for mov in p["cuentas"]:
            filas.append({
                "Partida": "",
                "Fecha": "",
                "Cuenta": mov["cuenta"],
                "Debe": mov["debe"] if mov["debe"] else "",
                "Haber": mov["haber"] if mov["haber"] else ""
            })
        filas.append({
            "Partida": "",
            "Fecha": "",
            "Cuenta": p["glosa"],
            "Debe": p["total_debe"],
            "Haber": p["total_haber"]
        })
    return pd.DataFrame(filas)


def exportar_excel(partidas, mes=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Partidas Mensuales"

    normal_font = Font(bold=False)
    bold_font = Font(bold=True)
    center_align = Alignment(horizontal="center", vertical="center")

    doble_abajo = Side(style="double", color="000000")
    borde_total = Border(bottom=doble_abajo)

    # Color por mes
    colores_por_mes = {
        1: "EAF2FF",   # Enero
        2: "FFF4E5",   # Febrero
        3: "EAFBF0",   # Marzo
        4: "F3E8FF",   # Abril
        5: "FFF1F2",   # Mayo
        6: "ECFEFF",   # Junio
        7: "FEF9C3",   # Julio
        8: "FCE7F3",   # Agosto
        9: "E0F2FE",   # Septiembre
        10: "F5F3FF",  # Octubre
        11: "ECFCCB",  # Noviembre
        12: "FFE4E6",  # Diciembre
    }

    color_mes = colores_por_mes.get(int(mes or 0), "F3F4F6")
    relleno_total = PatternFill(fill_type="solid", fgColor=color_mes)

    fila = 1

    for p in partidas:
        fila_inicio_movimientos = fila + 1

        # Encabezado de partida
        celda_codigo = ws.cell(row=fila, column=1, value=p["codigo"])
        celda_codigo.font = bold_font
        celda_codigo.alignment = center_align

        celda_fecha = ws.cell(row=fila, column=2, value=p["fecha"])
        celda_fecha.font = bold_font
        celda_fecha.alignment = center_align
        fila += 1

        # Movimientos
        for mov in p["cuentas"]:
            ws.cell(row=fila, column=2, value=mov["cuenta"]).font = normal_font

            if mov["debe"]:
                ws.cell(row=fila, column=3, value=float(mov["debe"]))
                ws.cell(row=fila, column=3).number_format = '"Q"#,##0.00'
                ws.cell(row=fila, column=3).font = normal_font

            if mov["haber"]:
                ws.cell(row=fila, column=4, value=float(mov["haber"]))
                ws.cell(row=fila, column=4).number_format = '"Q"#,##0.00'
                ws.cell(row=fila, column=4).font = normal_font

            fila += 1

        fila_total = fila
        fila_fin_movimientos = fila - 1

        # Glosa SIN color especial y SIN borde doble
        ws.cell(row=fila_total, column=2, value=p["glosa"]).font = normal_font

        if fila_fin_movimientos >= fila_inicio_movimientos:
            ws.cell(row=fila_total, column=3, value=f"=SUM(C{fila_inicio_movimientos}:C{fila_fin_movimientos})")
            ws.cell(row=fila_total, column=4, value=f"=SUM(D{fila_inicio_movimientos}:D{fila_fin_movimientos})")
        else:
            ws.cell(row=fila_total, column=3, value=0)
            ws.cell(row=fila_total, column=4, value=0)

        ws.cell(row=fila_total, column=3).number_format = '"Q"#,##0.00'
        ws.cell(row=fila_total, column=4).number_format = '"Q"#,##0.00'
        ws.cell(row=fila_total, column=3).font = normal_font
        ws.cell(row=fila_total, column=4).font = normal_font

        # SOLO a los totales
        ws.cell(row=fila_total, column=3).fill = relleno_total
        ws.cell(row=fila_total, column=4).fill = relleno_total
        ws.cell(row=fila_total, column=3).border = borde_total
        ws.cell(row=fila_total, column=4).border = borde_total

        fila += 1

    for col in ["A", "B", "C", "D"]:
        ws.column_dimensions[col].width = 28 if col == "B" else 16

    ws_anotaciones = wb.create_sheet("Anotaciones")
    encabezados = ["No. Partida", "Fecha", "Cuenta", "Anotación"]
    for columna, encabezado in enumerate(encabezados, start=1):
        celda = ws_anotaciones.cell(row=1, column=columna, value=encabezado)
        celda.font = bold_font
        celda.alignment = center_align
        celda.fill = relleno_total

    fila_anotacion = 2
    for p in partidas:
        for mov in p.get("cuentas", []) or []:
            anotacion = str(mov.get("anotacion", "") or "").strip()
            if not anotacion:
                continue

            ws_anotaciones.cell(row=fila_anotacion, column=1, value=p.get("codigo", ""))
            ws_anotaciones.cell(row=fila_anotacion, column=2, value=p.get("fecha", ""))
            ws_anotaciones.cell(row=fila_anotacion, column=3, value=mov.get("cuenta", ""))
            celda_anotacion = ws_anotaciones.cell(row=fila_anotacion, column=4, value=anotacion)
            celda_anotacion.alignment = Alignment(vertical="top", wrap_text=True)
            fila_anotacion += 1

    ws_anotaciones.freeze_panes = "A2"
    ws_anotaciones.auto_filter.ref = f"A1:D{max(fila_anotacion - 1, 1)}"
    ws_anotaciones.column_dimensions["A"].width = 16
    ws_anotaciones.column_dimensions["B"].width = 16
    ws_anotaciones.column_dimensions["C"].width = 32
    ws_anotaciones.column_dimensions["D"].width = 70

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


def _fila_distribucion(idx, key_periodo, valor_inicial=None):
    valor_inicial = valor_inicial or {}

    cuenta_inicial = str(valor_inicial.get("cuenta", "")).strip()
    monto_inicial = float(valor_inicial.get("monto", 0) or 0)

    col1, col2 = st.columns([2.5, 1.5])

    # 🔥 Obtener catálogo de cuentas
    opciones_catalogo = opciones_cuentas_cobro()

    # Fallback si no hay cuentas
    if not opciones_catalogo:
        opciones_catalogo = ["Caja", "Bancos"]

    # Si la cuenta guardada no existe, la agregamos
    if cuenta_inicial and cuenta_inicial not in opciones_catalogo:
        opciones_catalogo = [cuenta_inicial] + opciones_catalogo

    cuenta = col1.selectbox(
        f"Cuenta #{idx + 1}",
        opciones_catalogo,
        index=opciones_catalogo.index(cuenta_inicial) if cuenta_inicial in opciones_catalogo else 0,
        key=f"cuenta_cobro_{key_periodo}_{idx}",
    )

    monto = col2.number_input(
        "Monto",
        min_value=0.0,
        value=float(monto_inicial),
        step=0.01,
        key=f"monto_cobro_{key_periodo}_{idx}"
    )

    return {
        "cuenta": str(cuenta).strip(),
        "monto": float(monto or 0)
    }


def titulo_seccion_compacto(texto, tipo="general"):
    st.markdown(f"""
    <div class="seccion-mini seccion-{tipo}">
        <div class="seccion-mini-titulo">{texto}</div>
        <div class="seccion-mini-linea"></div>
    </div>
    """, unsafe_allow_html=True)


def valor_ui_actual(key, default=None):
    return st.session_state.get(key, default)


def titulo_expander_estado(texto, activo=False):
    return f"✅ {texto}" if activo else texto


def main():
    aplicar_estilos()

    ventas_data = cargar_json_lista(VENTAS_FILE)
    compras_data = cargar_json_lista(COMPRAS_FILE)
    partidas_guardadas_data = cargar_partidas_generadas()
    sincronizar_respaldo_partidas_generadas_sqlite(partidas_guardadas_data)
    config_ui = cargar_config_ui()
    empresas_data = cargar_empresas()
    

    if not ventas_data and not compras_data:
        st.warning("No se encontraron partidas.json ni partidascompras.json válidos.")
        return

    empresas_map = {}

    for x in (ventas_data + compras_data):
        empresa_id = obtener_id_empresa_desde_partida(x, empresas_data)
        nombre_visible = str(x.get("empresa", "")).strip()

        if empresa_id and empresa_id not in empresas_map:
            emp = buscar_empresa_por_nit(empresa_id, empresas_data)
            empresas_map[empresa_id] = str(emp.get("nombre", nombre_visible)).strip() if emp else nombre_visible

        elif not empresa_id and nombre_visible:
            emp = buscar_empresa_por_nombre(nombre_visible, empresas_data)
            if emp:
                nit_emp = normalizar_nit(emp.get("nit"))
                if nit_emp and nit_emp not in empresas_map:
                    empresas_map[nit_emp] = str(emp.get("nombre", nombre_visible)).strip()

    empresas = sorted(
        [{"id": k, "nombre": v} for k, v in empresas_map.items()],
        key=lambda x: x["nombre"].lower()
    )

    if not empresas:
        st.warning("No se encontraron empresas válidas en los archivos.")
        return

    anios = sorted(set(
        int(x.get("anio", 0) or 0)
        for x in (ventas_data + compras_data)
        if int(x.get("anio", 0) or 0)
    ))

    col_izq, col_der = st.columns([1, 1.35], gap="large")

    with col_izq:
    
        empresa_sel = st.selectbox(
            "Empresa",
            empresas,
            format_func=lambda x: x["nombre"]
        )

        empresa = empresa_sel["nombre"]
        empresa_id = empresa_sel["id"]

        regimen_empresa = obtener_regimen_empresa(empresa_id, empresas_data)
        

        

        col1, col2, col3 = st.columns(3)

        with col1:
            anio = st.selectbox(
                "Año",
                anios,
                index=len(anios) - 1 if anios else 0,
                key=f"anio_{empresa_id}"
            )

        with col2:
            mes = st.selectbox(
                "Mes",
                list(MESES.keys()),
                format_func=lambda x: MESES[x],
                key=f"mes_{empresa_id}"
            )
        
        config_empresa = obtener_config_empresa(config_ui, empresa_id, int(anio), int(mes))
        tiene_config_periodo_actual = existe_config_periodo(config_ui, empresa_id, int(anio), int(mes))
        key_periodo = f"{empresa_id}_{int(anio)}_{int(mes)}"

        pda_base_fallback = int(config_empresa.get("pda_inicial", 58) or 58)
        if tiene_config_periodo_actual:
            pda_inicial_default = pda_base_fallback
        else:
            pda_inicial_default = obtener_siguiente_pda_default(
                partidas_guardadas_data,
                empresa_id,
                int(anio),
                int(mes),
                fallback=pda_base_fallback,
            )

        pda_key = f"pda_inicial_{key_periodo}"
        contexto_diario_actual = f"{empresa_id}_{int(anio)}_{int(mes)}"
        contexto_diario_anterior = st.session_state.get("diario_contexto_pda_actual")
        if contexto_diario_anterior != contexto_diario_actual:
            st.session_state["diario_contexto_pda_actual"] = contexto_diario_actual
            st.session_state[pda_key] = int(pda_inicial_default)
        elif pda_key not in st.session_state:
            st.session_state[pda_key] = int(pda_inicial_default)

        incluir_ret_iva_default = bool(config_empresa.get("incluir_ret_iva", False))
        monto_ret_iva_default = float(config_empresa.get("monto_ret_iva", 0.0) or 0.0)
        incluir_exencion_iva_default = bool(config_empresa.get("incluir_exencion_iva", False))
        monto_exencion_iva_default = float(config_empresa.get("monto_exencion_iva", 0.0) or 0.0)
        incluir_ret_isr_default = bool(config_empresa.get("incluir_ret_isr", False))
        monto_ret_isr_default = float(config_empresa.get("monto_ret_isr", 0.0) or 0.0)
        activar_distribucion_default = bool(config_empresa.get("activar_distribucion", False))
        cantidad_cuentas_default = int(config_empresa.get("cantidad_cuentas", 1))
        usar_clientes_default = bool(config_empresa.get("usar_clientes_para_cuadre", True))
        distribucion_guardada = config_empresa.get("distribucion_cobro", [])
        cuenta_egreso_default = str(config_empresa.get("cuenta_egreso", "Caja"))
        activar_distribucion_egreso_default = bool(config_empresa.get("activar_distribucion_egreso", False))
        cantidad_cuentas_egreso_default = int(config_empresa.get("cantidad_cuentas_egreso", 1))
        distribucion_egreso_guardada = config_empresa.get("distribucion_egreso", [])
        activar_distribucion_impuestos_default = bool(config_empresa.get("activar_distribucion_impuestos", False))
        cantidad_cuentas_impuestos_default = int(config_empresa.get("cantidad_cuentas_impuestos", 1))
        distribucion_impuestos_guardada = config_empresa.get("distribucion_impuestos", [])
        incluir_saldos_por_pagar_anteriores_default = bool(
            config_empresa.get("incluir_saldos_por_pagar_anteriores", True)
        )
        activar_cierre_trimestral_default = bool(config_empresa.get("activar_cierre_trimestral", False))
        usar_isr_trimestral_renta_bruta_default = bool(config_empresa.get("usar_isr_trimestral_renta_bruta", False))
        usar_isr_trimestral_manual_default = bool(config_empresa.get("usar_isr_trimestral_manual", False))
        monto_isr_trimestral_manual_default = float(config_empresa.get("monto_isr_trimestral_manual", 0.0) or 0.0)
        generar_isr_trimestral_default = bool(config_empresa.get("generar_isr_trimestral", False))
        modo_isr_trimestral_default = str(config_empresa.get("modo_isr_trimestral", "") or "").strip().lower()
        if modo_isr_trimestral_default not in {"cierres_parciales", "renta_bruta", "manual", "ninguno"}:
            if usar_isr_trimestral_manual_default:
                modo_isr_trimestral_default = "manual"
            elif usar_isr_trimestral_renta_bruta_default:
                modo_isr_trimestral_default = "renta_bruta"
            else:
                modo_isr_trimestral_default = "ninguno"
        activar_iso_trimestral_default = bool(config_empresa.get("activar_iso_trimestral", False))
        es_primer_anio_iso_default = bool(config_empresa.get("es_primer_anio_iso", False))
        monto_iso_trimestral_manual_default = float(config_empresa.get("monto_iso_trimestral_manual", 0.0) or 0.0)
        acreditar_isr_a_iso_default = bool(config_empresa.get("acreditar_isr_a_iso", False))

        cuentas_excluir_pago_impuestos_guardadas = config_empresa.get(
            "cuentas_excluir_pago_impuestos", []
        )
        traslados_guardados = config_empresa.get("traslados_egreso_ingreso", [])
        cantidad_traslados_default = int(
            config_empresa.get("cantidad_traslados_egreso_ingreso", len(traslados_guardados))
        )
        ingresos_extra_guardados = config_empresa.get("ingresos_extra_manual", [])
        cantidad_ingresos_extra_default = int(
            config_empresa.get("cantidad_ingresos_extra_manual", len(ingresos_extra_guardados))
        )
        haberes_manuales_ingreso_guardados = config_empresa.get("haberes_manuales_ingreso", [])
        cantidad_haberes_manuales_ingreso_default = int(
            config_empresa.get("cantidad_haberes_manuales_ingreso", len(haberes_manuales_ingreso_guardados))
        )
        haberes_manuales_egreso_guardados = config_empresa.get("haberes_manuales_egreso", [])
        cantidad_haberes_manuales_egreso_default = int(
            config_empresa.get("cantidad_haberes_manuales_egreso", len(haberes_manuales_egreso_guardados))
        )

        with col3:
            pda_inicial = st.number_input(
                "PDA",
                min_value=1,
                key=pda_key
            )

        titulo_seccion_compacto("🟢 Ingresos", "ingresos")

        retenciones_activas = (
            bool(valor_ui_actual(f"incluir_ret_iva_{key_periodo}", incluir_ret_iva_default))
            or bool(valor_ui_actual(f"incluir_exencion_iva_{key_periodo}", incluir_exencion_iva_default))
            or bool(valor_ui_actual(f"incluir_ret_isr_{key_periodo}", incluir_ret_isr_default))
            or float(valor_ui_actual(f"monto_ret_iva_{key_periodo}", monto_ret_iva_default) or 0) > 0
            or float(valor_ui_actual(f"monto_exencion_iva_{key_periodo}", monto_exencion_iva_default) or 0) > 0
            or float(valor_ui_actual(f"monto_ret_isr_{key_periodo}", monto_ret_isr_default) or 0) > 0
        )
        with st.expander(titulo_expander_estado("Retenciones / Exenciones", retenciones_activas), expanded=False):
            col_ret_iva, col_exencion_iva, col_ret_isr = st.columns(3)

            with col_ret_iva:
                incluir_ret_iva = st.checkbox(
                    "Hay Retención IVA",
                    value=incluir_ret_iva_default,
                    key=f"incluir_ret_iva_{key_periodo}"
                )
                monto_ret_iva = st.number_input(
                    "Retención IVA",
                    min_value=0.0,
                    value=monto_ret_iva_default,
                    step=0.01,
                    disabled=not incluir_ret_iva,
                    key=f"monto_ret_iva_{key_periodo}",
                    label_visibility="collapsed"
                )

            with col_exencion_iva:
                incluir_exencion_iva = st.checkbox(
                    "Hay Exencion IVA",
                    value=incluir_exencion_iva_default,
                    key=f"incluir_exencion_iva_{key_periodo}"
                )
                monto_exencion_iva = st.number_input(
                    "Exencion IVA",
                    min_value=0.0,
                    value=monto_exencion_iva_default,
                    step=0.01,
                    disabled=not incluir_exencion_iva,
                    key=f"monto_exencion_iva_{key_periodo}",
                    label_visibility="collapsed"
                )

            with col_ret_isr:
                incluir_ret_isr = st.checkbox(
                    "Hay Retención ISR",
                    value=incluir_ret_isr_default,
                    key=f"incluir_ret_isr_{key_periodo}"
                )
                monto_ret_isr = st.number_input(
                    "Retención ISR",
                    min_value=0.0,
                    value=monto_ret_isr_default,
                    step=0.01,
                    disabled=not incluir_ret_isr,
                    key=f"monto_ret_isr_{key_periodo}",
                    label_visibility="collapsed"
                )


        cuenta_cobro_auto_default = str(config_empresa.get("cuenta_cobro_auto", "Caja"))
        activar_distribucion_ui = bool(valor_ui_actual(f"activar_distribucion_{key_periodo}", activar_distribucion_default))
        cuenta_cobro_auto_ui = str(valor_ui_actual(f"cuenta_cobro_auto_{key_periodo}", cuenta_cobro_auto_default) or "Caja").strip()
        distribucion_cobro_activa = (
            activar_distribucion_ui
            or bool(distribucion_guardada)
            or cuenta_cobro_auto_ui.lower() != "caja"
        )
        with st.expander(titulo_expander_estado("Distribución del cobro", distribucion_cobro_activa), expanded=False):
                opciones_cobro = opciones_cuentas_cobro()
                if not opciones_cobro:
                    opciones_cobro = ["Caja", "Bancos"]

                if cuenta_cobro_auto_default not in opciones_cobro:
                    opciones_cobro = [cuenta_cobro_auto_default] + opciones_cobro

                cuenta_cobro_auto = st.selectbox(
                    "Cuenta de cobro (modo automático)",
                    opciones_cobro,
                    index=opciones_cobro.index(cuenta_cobro_auto_default),
                    key=f"cuenta_cobro_auto_{key_periodo}"
                )

                col_dist_1, col_dist_2 = st.columns([2.2, 1])

                with col_dist_1:
                    activar_distribucion = st.checkbox(
                        "Manual",
                        value=activar_distribucion_default,
                        key=f"activar_distribucion_{key_periodo}"
                    )

                with col_dist_2:
                    cantidad_cuentas = st.number_input(
                        "Cantidad",
                        min_value=1,
                        max_value=10,
                        value=cantidad_cuentas_default,
                        step=1,
                        disabled=not activar_distribucion,
                        label_visibility="collapsed",
                        key=f"cantidad_cuentas_{key_periodo}"
                    )

                distribucion_cobro = []

                if activar_distribucion:
                    for i in range(int(cantidad_cuentas)):
                        valor_inicial = distribucion_guardada[i] if i < len(distribucion_guardada) else {}
                        distribucion_cobro.append(_fila_distribucion(i, key_periodo, valor_inicial=valor_inicial))

                    total_manual = round(sum(float(x.get("monto", 0) or 0) for x in distribucion_cobro), 2)
                    st.info(f"Total manual distribuido: Q {total_manual:,.2f}")
                else:
                    cantidad_cuentas = 1

        usar_clientes_ui = bool(valor_ui_actual(f"usar_clientes_para_cuadre_{key_periodo}", usar_clientes_default))
        cuenta_cuadre_ingreso_ui = str(
            valor_ui_actual(
                f"cuenta_cuadre_ingreso_{key_periodo}",
                str(config_empresa.get("cuenta_cuadre_ingreso", "Clientes"))
            ) or "Clientes"
        ).strip()
        cuadre_ingreso_activo = (not usar_clientes_ui) or (cuenta_cuadre_ingreso_ui.lower() != "clientes")
        with st.expander(titulo_expander_estado("Cuadre", cuadre_ingreso_activo), expanded=False):
            usar_clientes_para_cuadre = st.checkbox(
                "Cuadrar automáticamente la partida de ingresos si no cuadra",
                value=usar_clientes_default,
                key=f"usar_clientes_para_cuadre_{key_periodo}"
            )

            opciones_catalogo_ingreso = opciones_cuentas_catalogo()
            if not opciones_catalogo_ingreso:
                opciones_catalogo_ingreso = ["Clientes"]

            cuenta_cuadre_ingreso_default = str(config_empresa.get("cuenta_cuadre_ingreso", "Clientes"))

            if cuenta_cuadre_ingreso_default not in opciones_catalogo_ingreso:
                opciones_catalogo_ingreso = [cuenta_cuadre_ingreso_default] + opciones_catalogo_ingreso

            cuenta_cuadre_ingreso = st.selectbox(
                "Cuenta para cuadrar ingresos si no cuadra",
                opciones_catalogo_ingreso,
                index=opciones_catalogo_ingreso.index(cuenta_cuadre_ingreso_default),
                key=f"cuenta_cuadre_ingreso_{key_periodo}"
            )

        haberes_manuales_activos = bool(
            limpiar_haberes_manuales(haberes_manuales_ingreso_guardados)
        )
        with st.expander(
            titulo_expander_estado("Cuentas manuales en el haber del ingreso", haberes_manuales_activos),
            expanded=False
        ):
            cantidad_haberes_manuales_ingreso = st.number_input(
                "Cantidad de cuentas manuales en el haber",
                min_value=0,
                max_value=20,
                value=cantidad_haberes_manuales_ingreso_default,
                step=1,
                key=f"cantidad_haberes_manuales_ingreso_{key_periodo}"
            )

            haberes_manuales_ingreso = []
            opciones_haber_ingreso = opciones_cuentas_catalogo()
            if not opciones_haber_ingreso:
                opciones_haber_ingreso = ["Reintegros", "Otros Ingresos", "Anticipo S/Ventas"]

            for i in range(int(cantidad_haberes_manuales_ingreso)):
                valor_inicial = (
                    haberes_manuales_ingreso_guardados[i]
                    if i < len(haberes_manuales_ingreso_guardados)
                    else {}
                )
                cuenta_inicial = str(valor_inicial.get("cuenta", "")).replace("A: ", "").strip()
                monto_inicial = float(valor_inicial.get("monto", 0) or 0)

                opciones_haber_fila = opciones_haber_ingreso.copy()
                if cuenta_inicial and cuenta_inicial not in opciones_haber_fila:
                    opciones_haber_fila = [cuenta_inicial] + opciones_haber_fila

                col_hi1, col_hi2 = st.columns([2.4, 1])

                cuenta_haber_ingreso = col_hi1.selectbox(
                    f"Cuenta haber ingreso #{i + 1}",
                    opciones_haber_fila,
                    index=opciones_haber_fila.index(cuenta_inicial) if cuenta_inicial in opciones_haber_fila else 0,
                    key=f"cuenta_haber_manual_ingreso_{key_periodo}_{i}"
                )

                monto_haber_ingreso = col_hi2.number_input(
                    "Monto",
                    min_value=0.0,
                    value=float(monto_inicial),
                    step=0.01,
                    key=f"monto_haber_manual_ingreso_{key_periodo}_{i}"
                )

                haberes_manuales_ingreso.append({
                    "cuenta": str(cuenta_haber_ingreso).replace("A: ", "").strip(),
                    "monto": float(monto_haber_ingreso or 0),
                })

            total_haberes_manuales_ingreso_ui = round(
                sum(float(x.get("monto", 0) or 0) for x in haberes_manuales_ingreso),
                2
            )
            if cantidad_haberes_manuales_ingreso > 0:
                st.info(
                    f"Total manual en el haber del ingreso: Q {total_haberes_manuales_ingreso_ui:,.2f}"
                )

        compra_actual_periodo = buscar_partida(compras_data, empresa_id, int(anio), int(mes))
        mapa_cuentas_compra = obtener_mapa_cuentas_debe_compra(compra_actual_periodo)
        cuentas_traslado_detectadas = obtener_cuentas_debe_compra_para_traslado(compra_actual_periodo)

        traslados_egreso_ingreso_activos = bool(
            limpiar_movimientos_ingreso_config(traslados_guardados, permitir_solo_cuenta=True)
        )
        with st.expander(
            titulo_expander_estado("Traslados del egreso al ingreso", traslados_egreso_ingreso_activos),
            expanded=False
        ):
            cantidad_traslados_egreso_ingreso = st.number_input(
                "Cantidad de cuentas a trasladar",
                min_value=0,
                max_value=20,
                value=cantidad_traslados_default,
                step=1,
                key=f"cantidad_traslados_egreso_ingreso_{key_periodo}"
            )

            traslados_egreso_ingreso = []
            cuentas_traslado_ya_usadas = set()

            if not cuentas_traslado_detectadas:
                st.caption("No se detectaron cuentas en el debe de compras para este período.")
            else:
                st.caption("Solo eliges la cuenta y el sistema toma automáticamente el monto del debe de compras y su IVA al 12%.")

            for i in range(int(cantidad_traslados_egreso_ingreso)):
                valor_inicial = traslados_guardados[i] if i < len(traslados_guardados) else {}
                cuenta_inicial = str(valor_inicial.get("cuenta", "")).replace("A: ", "").strip()

                opciones_traslado = cuentas_traslado_detectadas.copy() if cuentas_traslado_detectadas else opciones_cuentas_catalogo()
                if not opciones_traslado:
                    opciones_traslado = ["Comisiones Pagadas"]
                if cuenta_inicial and cuenta_inicial not in opciones_traslado:
                    opciones_traslado = [cuenta_inicial] + opciones_traslado

                col_t1, col_t2, col_t3 = st.columns([2.2, 1, 1])

                cuenta_traslado = col_t1.selectbox(
                    f"Cuenta a trasladar #{i + 1}",
                    opciones_traslado,
                    index=opciones_traslado.index(cuenta_inicial) if cuenta_inicial in opciones_traslado else 0,
                    key=f"cuenta_traslado_ingreso_{key_periodo}_{i}"
                )

                disponible = round(float(mapa_cuentas_compra.get(cuenta_traslado, 0) or 0), 2)
                iva_traslado = round(disponible * 0.12, 2)

                col_t2.text_input(
                    "Monto automático",
                    value=f"Q {disponible:,.2f}",
                    key=f"monto_traslado_ingreso_auto_{key_periodo}_{i}",
                    disabled=True,
                )
                col_t3.text_input(
                    "IVA automático",
                    value=f"Q {iva_traslado:,.2f}",
                    key=f"iva_traslado_ingreso_auto_{key_periodo}_{i}",
                    disabled=True,
                )

                cuenta_lower = str(cuenta_traslado).strip().lower()
                if cuenta_lower in cuentas_traslado_ya_usadas:
                    st.warning(f"La cuenta '{cuenta_traslado}' ya fue seleccionada arriba. Solo se tomará una vez.")
                    continue
                cuentas_traslado_ya_usadas.add(cuenta_lower)

                if disponible <= 0:
                    st.warning(f"La cuenta '{cuenta_traslado}' no tiene monto disponible en el debe de compras de este período.")
                    continue

                traslados_egreso_ingreso.append({
                    "cuenta": str(cuenta_traslado).strip(),
                    "monto": float(disponible or 0),
                    "iva": float(iva_traslado or 0),
                })

            total_traslados_ui = round(sum(float(x.get("monto", 0) or 0) for x in traslados_egreso_ingreso), 2)
            total_iva_traslados_ui = round(sum(float(x.get("iva", 0) or 0) for x in traslados_egreso_ingreso), 2)
            if cantidad_traslados_egreso_ingreso > 0:
                st.info(
                    f"Total trasladado al ingreso: Q {total_traslados_ui:,.2f} | IVA asociado: Q {total_iva_traslados_ui:,.2f}"
                )

        ingresos_extra_activos = bool(
            limpiar_movimientos_ingreso_config(ingresos_extra_guardados)
        )
        with st.expander(
            titulo_expander_estado("Cuentas manuales adicionales en el ingreso", ingresos_extra_activos),
            expanded=False
        ):
            cantidad_ingresos_extra_manual = st.number_input(
                "Cantidad de cuentas manuales del ingreso",
                min_value=0,
                max_value=20,
                value=cantidad_ingresos_extra_default,
                step=1,
                key=f"cantidad_ingresos_extra_manual_{key_periodo}"
            )

            ingresos_extra_manual = []
            opciones_extra_ingreso = opciones_cuentas_catalogo()
            if not opciones_extra_ingreso:
                opciones_extra_ingreso = ["Otros Productos"]

            for i in range(int(cantidad_ingresos_extra_manual)):
                valor_inicial = ingresos_extra_guardados[i] if i < len(ingresos_extra_guardados) else {}
                cuenta_inicial = str(valor_inicial.get("cuenta", "")).replace("A: ", "").strip()
                monto_inicial = float(valor_inicial.get("monto", 0) or 0)
                iva_inicial = float(valor_inicial.get("iva", 0) or 0)

                opciones_extra_fila = opciones_extra_ingreso.copy()
                if cuenta_inicial and cuenta_inicial not in opciones_extra_fila:
                    opciones_extra_fila = [cuenta_inicial] + opciones_extra_fila

                col_ie1, col_ie2, col_ie3 = st.columns([2.2, 1, 1])

                cuenta_extra_ingreso = col_ie1.selectbox(
                    f"Cuenta manual ingreso #{i + 1}",
                    opciones_extra_fila,
                    index=opciones_extra_fila.index(cuenta_inicial) if cuenta_inicial in opciones_extra_fila else 0,
                    key=f"cuenta_extra_ingreso_{key_periodo}_{i}"
                )

                monto_extra_ingreso = col_ie2.number_input(
                    "Monto",
                    min_value=0.0,
                    value=float(monto_inicial),
                    step=0.01,
                    key=f"monto_extra_ingreso_{key_periodo}_{i}"
                )

                iva_extra_ingreso = col_ie3.number_input(
                    "IVA",
                    min_value=0.0,
                    value=float(iva_inicial),
                    step=0.01,
                    key=f"iva_extra_ingreso_{key_periodo}_{i}"
                )

                ingresos_extra_manual.append({
                    "cuenta": str(cuenta_extra_ingreso).strip(),
                    "monto": float(monto_extra_ingreso or 0),
                    "iva": float(iva_extra_ingreso or 0),
                })

            total_ingresos_extra_ui = round(sum(float(x.get("monto", 0) or 0) for x in ingresos_extra_manual), 2)
            total_iva_ingresos_extra_ui = round(sum(float(x.get("iva", 0) or 0) for x in ingresos_extra_manual), 2)
            if cantidad_ingresos_extra_manual > 0:
                st.info(
                    f"Total adicional en ingreso: Q {total_ingresos_extra_ui:,.2f} | IVA asociado: Q {total_iva_ingresos_extra_ui:,.2f}"
                )

        titulo_seccion_compacto("🟠 Egresos", "egresos")    

        activar_distribucion_impuestos_ui = bool(
            valor_ui_actual(f"activar_distribucion_impuestos_{key_periodo}", activar_distribucion_impuestos_default)
        )
        incluir_saldos_por_pagar_ui = bool(
            valor_ui_actual(
                f"incluir_saldos_por_pagar_anteriores_{key_periodo}",
                incluir_saldos_por_pagar_anteriores_default
            )
        )
        pago_impuestos_activo = (
            activar_distribucion_impuestos_ui
            or bool(distribucion_impuestos_guardada)
            or bool(cuentas_excluir_pago_impuestos_guardadas)
            or (not incluir_saldos_por_pagar_ui)
        )
        with st.expander(
            titulo_expander_estado("Distribución de Pago de impuestos", pago_impuestos_activo),
            expanded=False
        ):
            col_i1, col_i2 = st.columns([2.2, 1])

            with col_i1:
                activar_distribucion_impuestos = st.checkbox(
                    "Usar distribución manual de impuestos",
                    value=activar_distribucion_impuestos_default,
                    key=f"activar_distribucion_impuestos_{key_periodo}"
                )

            with col_i2:
                cantidad_cuentas_impuestos = st.number_input(
                    "Cantidad impuestos",
                    min_value=1,
                    max_value=10,
                    value=cantidad_cuentas_impuestos_default,
                    step=1,
                    disabled=not activar_distribucion_impuestos,
                    label_visibility="collapsed",
                    key=f"cantidad_cuentas_impuestos_{key_periodo}"
                )

            distribucion_impuestos = []

            incluir_saldos_por_pagar_anteriores = st.checkbox(
                "Incluir saldos 'por pagar' del período anterior en la primera partida de pago",
                value=incluir_saldos_por_pagar_anteriores_default,
                key=f"incluir_saldos_por_pagar_anteriores_{key_periodo}"
            )

            candidatos_por_pagar = obtener_saldos_por_pagar_para_pago(
                ventas_data=ventas_data,
                compras_data=compras_data,
                partidas_generadas_data=partidas_guardadas_data,
                empresa_id=empresa_id,
                anio=int(anio),
                mes=int(mes),
                excluir_cuentas=[]
            )

            mapa_candidatos = {
                x["cuenta"]: round(float(x.get("monto", 0) or 0), 2)
                for x in candidatos_por_pagar
            }

            opciones_excluir = list(mapa_candidatos.keys())
            defaults_excluir = [
                x for x in cuentas_excluir_pago_impuestos_guardadas
                if x in opciones_excluir
            ]

            cuentas_excluir_pago_impuestos = []

            if incluir_saldos_por_pagar_anteriores:
                if opciones_excluir:
                    cuentas_excluir_pago_impuestos = st.multiselect(
                        "Quitar de esa primera partida de pago",
                        options=opciones_excluir,
                        default=defaults_excluir,
                        format_func=lambda c: f"{c} — Q {mapa_candidatos.get(c, 0):,.2f}",
                        key=f"cuentas_excluir_pago_impuestos_{key_periodo}"
                    )

                    total_candidatos = round(sum(mapa_candidatos.values()), 2)
                    total_excluido = round(
                        sum(mapa_candidatos.get(c, 0) for c in cuentas_excluir_pago_impuestos),
                        2
                    )
                    total_incluido = round(total_candidatos - total_excluido, 2)

                    st.info(
                        f"Saldos por pagar detectados: Q {total_candidatos:,.2f} | "
                        f"Excluidos: Q {total_excluido:,.2f} | "
                        f"Incluidos: Q {total_incluido:,.2f}"
                    )
                else:
                    st.caption(
                        "No se detectaron otras cuentas por pagar para arrastrar desde apertura o desde el mes anterior guardado."
                    )

            if activar_distribucion_impuestos:
                for i in range(int(cantidad_cuentas_impuestos)):
                    valor_inicial = distribucion_impuestos_guardada[i] if i < len(distribucion_impuestos_guardada) else {}
                    distribucion_impuestos.append(_fila_distribucion(200 + i, key_periodo, valor_inicial=valor_inicial))

                total_manual_impuestos = round(sum(float(x.get("monto", 0) or 0) for x in distribucion_impuestos), 2)
                st.info(f"Total manual de impuestos disponible: Q {total_manual_impuestos:,.2f}")
            else:
                cantidad_cuentas_impuestos = 1

        activar_cierre_trimestral = False
        usar_isr_trimestral_renta_bruta = False
        usar_isr_trimestral_manual = False
        monto_isr_trimestral_manual = 0.0
        generar_isr_trimestral = False
        modo_isr_trimestral = "ninguno"
        usar_isr_desde_cierres_parciales = False
        activar_iso_trimestral = False
        es_primer_anio_iso = False
        monto_iso_trimestral_manual = 0.0
        acreditar_isr_a_iso = False

        if str(regimen_empresa).strip().lower() == "trimestral":
            ventas_netas_trim_ui, meses_con_ventas_trim_ui, trimestre_inicio_ui, trimestre_fin_ui = obtener_ventas_netas_trimestre(
                ventas_data,
                empresa_id,
                int(anio),
                int(mes),
            )

            generar_isr_trimestral_ui = bool(
                valor_ui_actual(f"generar_isr_trimestral_{key_periodo}", generar_isr_trimestral_default)
            )
            activar_iso_trimestral_ui = bool(
                valor_ui_actual(f"activar_iso_trimestral_{key_periodo}", activar_iso_trimestral_default)
            )
            es_primer_anio_iso_ui = bool(
                valor_ui_actual(f"es_primer_anio_iso_{key_periodo}", es_primer_anio_iso_default)
            )
            acreditar_isr_a_iso_ui = bool(
                valor_ui_actual(f"acreditar_isr_a_iso_{key_periodo}", acreditar_isr_a_iso_default)
            )
            modo_isr_trimestral_ui_guardado = str(
                valor_ui_actual(f"modo_isr_trimestral_{key_periodo}", modo_isr_trimestral_default) or ""
            ).strip().lower()
            monto_isr_trimestral_manual_ui = float(
                valor_ui_actual(f"monto_isr_trimestral_manual_{key_periodo}", monto_isr_trimestral_manual_default) or 0
            )
            monto_iso_trimestral_manual_ui = float(
                valor_ui_actual(f"monto_iso_trimestral_manual_{key_periodo}", monto_iso_trimestral_manual_default) or 0
            )
            impuestos_trimestrales_activos = (
                generar_isr_trimestral_ui
                or activar_iso_trimestral_ui
                or es_primer_anio_iso_ui
                or acreditar_isr_a_iso_ui
                or modo_isr_trimestral_ui_guardado not in {"", "ninguno"}
                or monto_isr_trimestral_manual_ui > 0
                or monto_iso_trimestral_manual_ui > 0
            )
            with st.expander(
                titulo_expander_estado("ISR trimestral e ISO trimestral", impuestos_trimestrales_activos),
                expanded=False
            ):
                if es_fin_trimestre(int(mes)):
                    st.caption(
                        f"Trimestre actual: {MESES[int(trimestre_inicio_ui)]} a {MESES[int(trimestre_fin_ui)]}. "
                        f"Ventas netas acumuladas: Q {ventas_netas_trim_ui:,.2f}"
                    )

                    registro_isr_cierre_ui = obtener_isr_trimestral_desde_cierres(empresa_id, int(anio), int(mes))
                    saldo_iso_por_acreditar_ui = obtener_saldo_iso_por_acreditar_hasta_fecha(
                        ventas_data=ventas_data,
                        compras_data=compras_data,
                        partidas_generadas_data=partidas_guardadas_data,
                        empresa_id=empresa_id,
                        anio=int(anio),
                        mes=int(mes),
                        incluir_mes_actual=False,
                    )

                    generar_isr_trimestral = st.checkbox(
                        "Generar partida de ISR trimestral",
                        value=generar_isr_trimestral_default,
                        key=f"generar_isr_trimestral_{key_periodo}"
                    )

                    opciones_isr_trimestral = ["Cierres parciales", "Renta bruta", "Manual"]
                    mapa_isr_trimestral = {
                        "Cierres parciales": "cierres_parciales",
                        "Renta bruta": "renta_bruta",
                        "Manual": "manual",
                    }
                    valor_ui_modo = "Cierres parciales"
                    for etiqueta, valor in mapa_isr_trimestral.items():
                        if valor == modo_isr_trimestral_default:
                            valor_ui_modo = etiqueta
                            break

                    modo_isr_trimestral_ui = st.radio(
                        "Fuente del ISR trimestral",
                        options=opciones_isr_trimestral,
                        index=opciones_isr_trimestral.index(valor_ui_modo),
                        horizontal=True,
                        disabled=not generar_isr_trimestral,
                        key=f"modo_isr_trimestral_{key_periodo}"
                    )
                    modo_isr_trimestral = mapa_isr_trimestral.get(modo_isr_trimestral_ui, "ninguno") if generar_isr_trimestral else "ninguno"
                    usar_isr_desde_cierres_parciales = modo_isr_trimestral == "cierres_parciales"
                    usar_isr_trimestral_renta_bruta = modo_isr_trimestral == "renta_bruta"
                    usar_isr_trimestral_manual = modo_isr_trimestral == "manual"

                    monto_isr_renta_bruta_ui = round(float(ventas_netas_trim_ui or 0) * 0.08 * 0.25, 2)

                    if modo_isr_trimestral == "cierres_parciales":
                        if registro_isr_cierre_ui:
                            st.info(
                                f"ISR trimestral guardado en Cierres Parciales: Q {float(registro_isr_cierre_ui.get('isr_trimestral_por_pagar', 0) or 0):,.2f}"
                            )
                        else:
                            st.warning(
                                "No se encontró ISR trimestral guardado en Cierres Parciales para este trimestre."
                            )
                    elif modo_isr_trimestral == "renta_bruta":
                        st.number_input(
                            "ISR renta bruta",
                            min_value=0.0,
                            value=float(monto_isr_renta_bruta_ui),
                            step=0.01,
                            disabled=True,
                            key=f"isr_trimestral_renta_bruta_info_{key_periodo}"
                        )
                    elif modo_isr_trimestral == "manual":
                        monto_isr_trimestral_manual = st.number_input(
                            "ISR trimestral manual",
                            min_value=0.0,
                            value=float(monto_isr_trimestral_manual_default),
                            step=0.01,
                            disabled=not generar_isr_trimestral,
                            key=f"monto_isr_trimestral_manual_{key_periodo}"
                        )

                    st.caption(
                        f"Saldo disponible detectado para acreditar al ISR trimestral del período: Q {saldo_iso_por_acreditar_ui:,.2f}"
                    )

                    activar_iso_trimestral = st.checkbox(
                        "Generar partida de ISO trimestral",
                        value=activar_iso_trimestral_default,
                        key=f"activar_iso_trimestral_{key_periodo}"
                    )
                    es_primer_anio_iso = st.checkbox(
                        "Primer año de la empresa (no generar ISO trimestral)",
                        value=es_primer_anio_iso_default,
                        disabled=not activar_iso_trimestral,
                        key=f"es_primer_anio_iso_{key_periodo}"
                    )

                    col_iso_1, col_iso_2 = st.columns([1.2, 1])
                    with col_iso_1:
                        monto_iso_trimestral_manual = st.number_input(
                            "ISO trimestral manual",
                            min_value=0.0,
                            value=float(monto_iso_trimestral_manual_default),
                            step=0.01,
                            disabled=(not activar_iso_trimestral) or es_primer_anio_iso,
                            key=f"monto_iso_trimestral_manual_{key_periodo}"
                        )
                    with col_iso_2:
                        acreditar_isr_a_iso = st.checkbox(
                            "Acreditar ISR trimestral del período al ISO",
                            value=acreditar_isr_a_iso_default,
                            disabled=(not activar_iso_trimestral) or es_primer_anio_iso or (not generar_isr_trimestral),
                            key=f"acreditar_isr_a_iso_{key_periodo}"
                        )

                    if es_primer_anio_iso and activar_iso_trimestral:
                        st.info("Al estar marcado como primer año, no se formará la partida de ISO trimestral.")
                else:
                    st.caption("Este bloque solo se activa en marzo, junio, septiembre y diciembre para empresas de régimen trimestral.")

        opciones_catalogo = opciones_cuentas_catalogo()
        if not opciones_catalogo:
            opciones_catalogo = ["Proveedores", "Cuentas por Pagar"]
        activar_distribucion_egreso_ui = bool(
            valor_ui_actual(f"activar_distribucion_egreso_{key_periodo}", activar_distribucion_egreso_default)
        )
        cuenta_egreso_ui = str(valor_ui_actual(f"cuenta_egreso_{key_periodo}", cuenta_egreso_default) or "Caja").strip()
        cuenta_cuadre_egreso_ui = str(
            valor_ui_actual(
                f"cuenta_cuadre_egreso_{key_periodo}",
                str(config_empresa.get("cuenta_cuadre_egreso", "Proveedores"))
            ) or "Proveedores"
        ).strip()
        egreso_caja_activo = (
            activar_distribucion_egreso_ui
            or bool(distribucion_egreso_guardada)
            or cuenta_egreso_ui.lower() != "caja"
            or cuenta_cuadre_egreso_ui.lower() != "proveedores"
        )
        with st.expander(titulo_expander_estado("Egreso de caja / bancos", egreso_caja_activo), expanded=False):
                opciones_egreso_auto = opciones_cuentas_cobro()
                if not opciones_egreso_auto:
                    opciones_egreso_auto = ["Caja", "Bancos"]

                if cuenta_egreso_default not in opciones_egreso_auto:
                    opciones_egreso_auto = [cuenta_egreso_default] + opciones_egreso_auto

                cuenta_egreso = st.selectbox(
                    "Cuenta principal de egreso",
                    opciones_egreso_auto,
                    index=opciones_egreso_auto.index(cuenta_egreso_default),
                    key=f"cuenta_egreso_{key_periodo}"
                )

                col_e1, col_e2 = st.columns([2.2, 1])

                with col_e1:
                    activar_distribucion_egreso = st.checkbox(
                        "Usar distribución manual de egreso",
                        value=activar_distribucion_egreso_default,
                        key=f"activar_distribucion_egreso_{key_periodo}"
                    )

                with col_e2:
                    cantidad_cuentas_egreso = st.number_input(
                        "Cantidad egreso",
                        min_value=1,
                        max_value=10,
                        value=cantidad_cuentas_egreso_default,
                        step=1,
                        disabled=not activar_distribucion_egreso,
                        label_visibility="collapsed",
                        key=f"cantidad_cuentas_egreso_{key_periodo}"
                    )

                distribucion_egreso = []

                if activar_distribucion_egreso:
                    for i in range(int(cantidad_cuentas_egreso)):
                        valor_inicial = distribucion_egreso_guardada[i] if i < len(distribucion_egreso_guardada) else {}
                        distribucion_egreso.append(_fila_distribucion(100 + i, key_periodo, valor_inicial=valor_inicial))

                    total_manual_egreso = round(sum(float(x.get("monto", 0) or 0) for x in distribucion_egreso), 2)
                    st.info(f"Total manual de egreso disponible: Q {total_manual_egreso:,.2f}")
                else:
                    cantidad_cuentas_egreso = 1

                cuenta_cuadre_egreso_default = str(config_empresa.get("cuenta_cuadre_egreso", "Proveedores"))

                if cuenta_cuadre_egreso_default not in opciones_catalogo:
                    opciones_catalogo = [cuenta_cuadre_egreso_default] + opciones_catalogo

                cuenta_cuadre_egreso = st.selectbox(
                    "Cuenta para cuadrar egresos si no cuadra",
                    opciones_catalogo,
                    index=opciones_catalogo.index(cuenta_cuadre_egreso_default),
                    key=f"cuenta_cuadre_egreso_{key_periodo}"
                )

        haberes_manuales_egreso_activos = bool(
            limpiar_haberes_manuales(haberes_manuales_egreso_guardados)
        )
        with st.expander(
            titulo_expander_estado("Cuentas manuales en el haber del egreso", haberes_manuales_egreso_activos),
            expanded=False
        ):
            cantidad_haberes_manuales_egreso = st.number_input(
                "Cantidad de cuentas manuales en el haber del egreso",
                min_value=0,
                max_value=20,
                value=cantidad_haberes_manuales_egreso_default,
                step=1,
                key=f"cantidad_haberes_manuales_egreso_{key_periodo}"
            )

            haberes_manuales_egreso = []
            opciones_haber_egreso = opciones_cuentas_catalogo()
            if not opciones_haber_egreso:
                opciones_haber_egreso = ["Proveedores", "Acreedores Varios", "Anticipo S/Compras"]

            for i in range(int(cantidad_haberes_manuales_egreso)):
                valor_inicial = (
                    haberes_manuales_egreso_guardados[i]
                    if i < len(haberes_manuales_egreso_guardados)
                    else {}
                )
                cuenta_inicial = str(valor_inicial.get("cuenta", "")).replace("A: ", "").strip()
                monto_inicial = float(valor_inicial.get("monto", 0) or 0)

                opciones_haber_fila = opciones_haber_egreso.copy()
                if cuenta_inicial and cuenta_inicial not in opciones_haber_fila:
                    opciones_haber_fila = [cuenta_inicial] + opciones_haber_fila

                col_he1, col_he2 = st.columns([2.4, 1])

                cuenta_haber_egreso = col_he1.selectbox(
                    f"Cuenta haber egreso #{i + 1}",
                    opciones_haber_fila,
                    index=opciones_haber_fila.index(cuenta_inicial) if cuenta_inicial in opciones_haber_fila else 0,
                    key=f"cuenta_haber_manual_egreso_{key_periodo}_{i}"
                )

                monto_haber_egreso = col_he2.number_input(
                    "Monto",
                    min_value=0.0,
                    value=float(monto_inicial),
                    step=0.01,
                    key=f"monto_haber_manual_egreso_{key_periodo}_{i}"
                )

                haberes_manuales_egreso.append({
                    "cuenta": str(cuenta_haber_egreso).replace("A: ", "").strip(),
                    "monto": float(monto_haber_egreso or 0),
                })

            total_haberes_manuales_egreso_ui = round(
                sum(float(x.get("monto", 0) or 0) for x in haberes_manuales_egreso),
                2
            )
            if cantidad_haberes_manuales_egreso > 0:
                st.info(
                    f"Total manual en el haber del egreso: Q {total_haberes_manuales_egreso_ui:,.2f}"
                )

        extraordinarios_guardados = config_empresa.get("gastos_extraordinarios", [])
        cantidad_extraordinarios_default = int(config_empresa.get("cantidad_gastos_extraordinarios", 0))
        salarios_data = cargar_salarios_minimos()

        gastos_extraordinarios_activos = any(
            float(x.get("monto", 0) or 0) > 0 for x in (extraordinarios_guardados or [])
        )
        with st.expander(
            titulo_expander_estado("Gastos extraordinarios incluidos en el egreso", gastos_extraordinarios_activos),
            expanded=False
        ):
            cantidad_gastos_extraordinarios = st.number_input(
                "Cantidad de gastos extraordinarios",
                min_value=0,
                max_value=20,
                value=cantidad_extraordinarios_default,
                step=1,
                key=f"cantidad_gastos_extraordinarios_{key_periodo}"
            )

            gastos_extraordinarios = []

            for i in range(int(cantidad_gastos_extraordinarios)):
                valor_inicial = extraordinarios_guardados[i] if i < len(extraordinarios_guardados) else {}

                cuenta_inicial = str(valor_inicial.get("cuenta", "")).strip()
                monto_inicial = float(valor_inicial.get("monto", 0) or 0)

                opciones_extra = opciones_catalogo.copy()
                if cuenta_inicial and cuenta_inicial not in opciones_extra:
                    opciones_extra = [cuenta_inicial] + opciones_extra

                colx1, colx2 = st.columns([2.3, 1])

                cuenta_extra = colx1.selectbox(
                    f"Cuenta extraordinaria #{i + 1}",
                    opciones_extra,
                    index=opciones_extra.index(cuenta_inicial) if cuenta_inicial in opciones_extra else 0,
                    key=f"cuenta_extra_{key_periodo}_{i}"
                )

                monto_extra = colx2.number_input(
                    "Monto",
                    min_value=0.0,
                    value=float(monto_inicial),
                    step=0.01,
                    key=f"monto_extra_{key_periodo}_{i}"
                )

                gastos_extraordinarios.append({
                    "cuenta": str(cuenta_extra).strip(),
                    "monto": float(monto_extra or 0)
                })

            total_extraordinarios = round(sum(float(x.get("monto", 0) or 0) for x in gastos_extraordinarios), 2)
            if cantidad_gastos_extraordinarios > 0:
                st.info(f"Total gastos extraordinarios: Q {total_extraordinarios:,.2f}")        

                
        activar_planilla = False
        datos_planilla = None
        cantidad_trabajadores = 0
        anio_salario = int(anio)
        circunscripcion = "CE2"
        tipo_suscripcion = "no_agricola"
        usar_salario_manual = False
        salario_mensual = 0.0
        incluir_igss_planilla = True
        incluir_irtra_planilla = True
        incluir_intecap_planilla = True     
        modo_pago_planilla = "mensual"
        sueldo_por_pagar_manual = 0.0
        bonificacion_por_pagar_manual = 0.0
        monto_retencion_isr_planilla = 0.0
        modo_salario_planilla = "general"
        trabajadores_detalle = []
        activar_provision_prestaciones = bool(config_empresa.get("activar_provision_prestaciones", True))
        porcentaje_prestaciones = float(config_empresa.get("porcentaje_prestaciones", 26.38) or 26.38)   

        activar_planilla_ui = bool(
            valor_ui_actual(f"activar_planilla_{key_periodo}", bool(config_empresa.get("activar_planilla", False)))
        )
        with st.expander(titulo_expander_estado("Trabajadores / Planilla", activar_planilla_ui), expanded=False):
            activar_planilla = st.checkbox(
                "Incluir planilla de trabajadores en la partida de egreso",
                value=bool(config_empresa.get("activar_planilla", False)),
                key=f"activar_planilla_{key_periodo}"
            )

            datos_planilla = None

            if activar_planilla:
                colp1, colp2 = st.columns(2)

                cantidad_trabajadores_default = max(
                    1,
                    int(config_empresa.get("cantidad_trabajadores", 1) or 1)
                )

                cantidad_trabajadores = colp1.number_input(
                    "Cantidad de trabajadores",
                    min_value=1,
                    step=1,
                    value=cantidad_trabajadores_default,
                    key=f"cantidad_trabajadores_{key_periodo}"
                )

                anio_salario = colp2.number_input(
                    "Año salario mínimo",
                    min_value=2020,
                    step=1,
                    value=int(config_empresa.get("anio_salario", int(anio))),
                    key=f"anio_salario_{key_periodo}"
                )

                colp3, colp4 = st.columns(2)

                circunscripcion = colp3.selectbox(
                    "Circunscripción",
                    ["CE1", "CE2"],
                    index=["CE1", "CE2"].index(config_empresa.get("circunscripcion", "CE2")),
                    key=f"circunscripcion_{key_periodo}"
                )

                tipo_suscripcion = colp4.selectbox(
                    "Tipo",
                    ["agricola", "no_agricola", "maquila"],
                    index=["agricola", "no_agricola", "maquila"].index(config_empresa.get("tipo_suscripcion", "no_agricola")),
                    key=f"tipo_suscripcion_{key_periodo}"
                )

                salario_detectado = obtener_salario_minimo(
                    anio_salario,
                    circunscripcion,
                    tipo_suscripcion,
                    salarios_data
                )

                usar_salario_manual = st.checkbox(
                    "Sobrescribir salario mínimo manualmente",
                    value=bool(config_empresa.get("usar_salario_manual", False)),
                    key=f"usar_salario_manual_{key_periodo}"
                )

                modo_salario_planilla = st.selectbox(
                    "Modo de salarios",
                    ["general", "individual"],
                    index=["general", "individual"].index(
                        config_empresa.get("modo_salario_planilla", "general")
                    ),
                    key=f"modo_salario_planilla_{key_periodo}"
                )
                trabajadores_detalle_guardado = config_empresa.get("trabajadores_detalle", [])

                if modo_salario_planilla == "general":
                    salario_guardado = float(config_empresa.get("salario_mensual", 0.0) or 0.0)
                    salario_base_input = salario_guardado if salario_guardado > 0 else float(salario_detectado or 0.0)

                    salario_mensual = st.number_input(
                        "Salario mensual por trabajador",
                        min_value=0.0,
                        value=salario_base_input,
                        step=0.01,
                        key=f"salario_mensual_{key_periodo}"
                    ) if usar_salario_manual or salario_detectado > 0 else 0.0

                    trabajadores_detalle = []

                else:
                    salario_mensual = 0.0
                    trabajadores_detalle = []

                    st.markdown("### Detalle de trabajadores")

                    for i in range(int(cantidad_trabajadores)):
                        inicial = trabajadores_detalle_guardado[i] if i < len(trabajadores_detalle_guardado) else {}

                        nombre_inicial = str(inicial.get("nombre", f"Trabajador {i+1}")).strip()
                        salario_inicial = float(inicial.get("salario", salario_detectado or 0.0) or 0.0)
                        bonificacion_inicial = bool(inicial.get("bonificacion", True))

                        cta, ctb, ctc = st.columns([2.2, 1.4, 1.2])

                        nombre_trab = cta.text_input(
                            f"Nombre #{i+1}",
                            value=nombre_inicial,
                            key=f"trabajador_nombre_{key_periodo}_{i}"
                        )

                        salario_trab = ctb.number_input(
                            f"Salario #{i+1}",
                            min_value=0.0,
                            value=float(salario_inicial),
                            step=0.01,
                            key=f"trabajador_salario_{key_periodo}_{i}"
                        )

                        bonificacion_trab = ctc.checkbox(
                            f"Bonificación #{i+1}",
                            value=bonificacion_inicial,
                            key=f"trabajador_bonificacion_{key_periodo}_{i}"
                        )

                        trabajadores_detalle.append({
                            "nombre": str(nombre_trab).strip() or f"Trabajador {i+1}",
                            "salario": float(salario_trab or 0.0),
                            "bonificacion": bool(bonificacion_trab)
                        })

                colp5, colp6, colp7 = st.columns(3)

                incluir_igss_planilla = colp5.checkbox(
                    "IGSS",
                    value=bool(config_empresa.get("incluir_igss_planilla", True)),
                    key=f"incluir_igss_planilla_{key_periodo}"
                )

                incluir_irtra_planilla = colp6.checkbox(
                    "IRTRA",
                    value=bool(config_empresa.get("incluir_irtra_planilla", True)),
                    key=f"incluir_irtra_planilla_{key_periodo}"
                )

                incluir_intecap_planilla = colp7.checkbox(
                    "INTECAP",
                    value=bool(config_empresa.get("incluir_intecap_planilla", True)),
                    key=f"incluir_intecap_planilla_{key_periodo}"
                )

                # PRIMERO calcular datos_planilla
                if modo_salario_planilla == "general":
                    if salario_mensual > 0:
                        datos_planilla = calcular_planilla_laboral(
                            cantidad_trabajadores=cantidad_trabajadores,
                            salario_mensual=salario_mensual,
                            incluir_igss=incluir_igss_planilla,
                            incluir_irtra=incluir_irtra_planilla,
                            incluir_intecap=incluir_intecap_planilla,
                            bonificacion_incentivo=BONIFICACION_INCENTIVO
                        )
                else:
                    if trabajadores_detalle:
                        datos_planilla = calcular_planilla_laboral_detallada(
                            trabajadores=trabajadores_detalle,
                            incluir_igss=incluir_igss_planilla,
                            incluir_irtra=incluir_irtra_planilla,
                            incluir_intecap=incluir_intecap_planilla,
                            bonificacion_incentivo=BONIFICACION_INCENTIVO
                        )

                if datos_planilla:
                    st.info(
                        f"Base salarios: Q {datos_planilla['salario_base_total']:,.2f} | "
                        f"Bonificación: Q {datos_planilla['bonificacion_total']:,.2f} | "
                        f"IGSS laboral: Q {datos_planilla['cuota_laboral']:,.2f} | "
                        f"IGSS patronal: Q {datos_planilla['cuota_patronal']:,.2f}"
                    )

                    resumen_isr_planilla = resumir_salarios_sujetos_isr_planilla(datos_planilla)
                    if resumen_isr_planilla["cantidad"] > 0:
                        st.info(
                            f"Se detectaron {resumen_isr_planilla['cantidad']} salario(s) mayores a "
                            f"Q {resumen_isr_planilla['umbral']:,.2f}. "
                            f"Si corresponde, ingresa manualmente la retención para formar "
                            f"'Retención ISR por pagar'. Base de referencia: Q {resumen_isr_planilla['base_referencia']:,.2f}"
                        )

                    monto_retencion_isr_planilla = st.number_input(
                        "Retención ISR planilla (manual)",
                        min_value=0.0,
                        value=float(config_empresa.get("monto_retencion_isr_planilla", 0.0) or 0.0),
                        step=0.01,
                        key=f"monto_retencion_isr_planilla_{key_periodo}"
                    )

                # DESPUÉS mostrar el modo de pago
                modo_pago_planilla = st.selectbox(
                    "Modo de salarios y bonificación por pagar",
                    ["mensual", "quincenal", "manual"],
                    index=["mensual", "quincenal", "manual"].index(
                        config_empresa.get("modo_pago_planilla", "mensual")
                    ),
                    key=f"modo_pago_planilla_{key_periodo}"
                )

                st.markdown("### Provisión de prestaciones")

                activar_provision_prestaciones = st.checkbox(
                    "Generar partida de provisión para prestaciones laborales",
                    value=bool(config_empresa.get("activar_provision_prestaciones", True)),
                    key=f"activar_provision_prestaciones_{key_periodo}"
                )

                porcentaje_prestaciones = st.number_input(
                    "Porcentaje de prestaciones laborales",
                    min_value=0.0,
                    value=float(config_empresa.get("porcentaje_prestaciones", 26.38) or 26.38),
                    step=0.01,
                    key=f"porcentaje_prestaciones_{key_periodo}"
                )

                if datos_planilla and activar_provision_prestaciones:
                    salario_base_preview = round(float(datos_planilla.get("salario_base_total", 0) or 0), 2)
                    monto_preview_prestaciones = round(
                        salario_base_preview * (float(porcentaje_prestaciones or 0) / 100),
                        2
                    )
                    st.info(
                        f"Provisión estimada de prestaciones: Q {monto_preview_prestaciones:,.2f} "
                        f"({float(porcentaje_prestaciones or 0):,.2f}% sobre Q {salario_base_preview:,.2f})"
                    )

                sueldo_por_pagar_manual = 0.0
                bonificacion_por_pagar_manual = 0.0

                if datos_planilla:
                    sueldo_neto_total_preview = round(float(datos_planilla.get("sueldo_neto_total", 0) or 0), 2)
                    bonificacion_total_preview = round(float(datos_planilla.get("bonificacion_total", 0) or 0), 2)

                    if modo_pago_planilla == "quincenal":
                        st.info(
                            f"Quedará por pagar -> "
                            f"Sueldos: Q {sueldo_neto_total_preview / 2:,.2f} | "
                            f"Bonificación: Q {bonificacion_total_preview / 2:,.2f}"
                        )

                    elif modo_pago_planilla == "mensual":
                        st.info(
                            f"Quedará por pagar -> "
                            f"Sueldos: Q {sueldo_neto_total_preview:,.2f} | "
                            f"Bonificación: Q {bonificacion_total_preview:,.2f}"
                        )

                    elif modo_pago_planilla == "manual":
                        colm1, colm2 = st.columns(2)

                        sueldo_por_pagar_manual = colm1.number_input(
                            "Sueldos x pagar (manual)",
                            min_value=0.0,
                            value=float(config_empresa.get("sueldo_por_pagar_manual", 0.0) or 0.0),
                            step=0.01,
                            key=f"sueldo_por_pagar_manual_{key_periodo}"
                        )

                        bonificacion_por_pagar_manual = colm2.number_input(
                            "Bonificación x pagar (manual)",
                            min_value=0.0,
                            value=float(config_empresa.get("bonificacion_por_pagar_manual", 0.0) or 0.0),
                            step=0.01,
                            key=f"bonificacion_por_pagar_manual_{key_periodo}"
                        )
        
        
                # --- Confirmación para sobreescribir partidas existentes ---
        confirm_key = f"confirmar_sobrescritura_{empresa_id}_{int(anio)}_{int(mes)}"

        if confirm_key not in st.session_state:
            st.session_state[confirm_key] = False

        registro_existente_actual = buscar_partidas_generadas(
            partidas_guardadas_data,
            empresa_id,
            int(anio),
            int(mes)
        )

        datos_config_mes = {
            "pda_inicial": int(pda_inicial),
            "incluir_ret_iva": bool(incluir_ret_iva),
            "monto_ret_iva": float(monto_ret_iva or 0),
            "incluir_exencion_iva": bool(incluir_exencion_iva),
            "monto_exencion_iva": float(monto_exencion_iva or 0),
            "incluir_ret_isr": bool(incluir_ret_isr),
            "monto_ret_isr": float(monto_ret_isr or 0),

            "distribucion_cobro": limpiar_distribucion(distribucion_cobro),
            "modo_pago_planilla": str(modo_pago_planilla),
            "sueldo_por_pagar_manual": float(sueldo_por_pagar_manual or 0),
            "bonificacion_por_pagar_manual": float(bonificacion_por_pagar_manual or 0),

            "cuenta_egreso": str(cuenta_egreso),
            "activar_distribucion_egreso": bool(activar_distribucion_egreso),
            "distribucion_egreso": limpiar_distribucion(distribucion_egreso),
            "cuenta_cuadre_egreso": str(cuenta_cuadre_egreso),

            "activar_distribucion_impuestos": bool(activar_distribucion_impuestos),
            "distribucion_impuestos": limpiar_distribucion(distribucion_impuestos),

            "gastos_extraordinarios": gastos_extraordinarios or [],
            "traslados_egreso_ingreso": limpiar_movimientos_ingreso_config(traslados_egreso_ingreso, permitir_solo_cuenta=True),
            "ingresos_extra_manual": limpiar_movimientos_ingreso_config(ingresos_extra_manual),
            "haberes_manuales_ingreso": limpiar_haberes_manuales(haberes_manuales_ingreso),
            "cantidad_haberes_manuales_ingreso": int(cantidad_haberes_manuales_ingreso),
            "haberes_manuales_egreso": limpiar_haberes_manuales(haberes_manuales_egreso),
            "cantidad_haberes_manuales_egreso": int(cantidad_haberes_manuales_egreso),
            "usar_clientes_para_cuadre": bool(usar_clientes_para_cuadre),
            "cuenta_cuadre_ingreso": str(cuenta_cuadre_ingreso),
            "cuenta_cobro_auto": str(cuenta_cobro_auto),

            "activar_planilla": bool(activar_planilla),
            "datos_planilla": datos_planilla or {},
            "monto_retencion_isr_planilla": float(monto_retencion_isr_planilla or 0.0) if activar_planilla else 0.0,
            "activar_provision_prestaciones": bool(activar_provision_prestaciones),
            "porcentaje_prestaciones": float(porcentaje_prestaciones or 0),

            "incluir_saldos_por_pagar_anteriores": bool(incluir_saldos_por_pagar_anteriores),
            "cuentas_excluir_pago_impuestos": cuentas_excluir_pago_impuestos or [],
            "activar_cierre_trimestral": bool(activar_cierre_trimestral),
            "usar_isr_trimestral_renta_bruta": bool(usar_isr_trimestral_renta_bruta),
            "usar_isr_trimestral_manual": bool(usar_isr_trimestral_manual),
            "monto_isr_trimestral_manual": float(monto_isr_trimestral_manual or 0.0),
            "generar_isr_trimestral": bool(generar_isr_trimestral),
            "modo_isr_trimestral": str(modo_isr_trimestral).strip(),
            "usar_isr_desde_cierres_parciales": bool(usar_isr_desde_cierres_parciales),
            "activar_iso_trimestral": bool(activar_iso_trimestral),
            "es_primer_anio_iso": bool(es_primer_anio_iso),
            "monto_iso_trimestral_manual": float(monto_iso_trimestral_manual or 0.0),
            "acreditar_isr_a_iso": bool(acreditar_isr_a_iso),
        }

        generar = st.button("Guardar Cambios", use_container_width=True, type="primary")

        confirmado_generar = False

        if generar:
            if registro_existente_actual and not st.session_state[confirm_key]:
                st.session_state[confirm_key] = True
            else:
                confirmado_generar = True
                st.session_state[confirm_key] = False

        if st.session_state[confirm_key] and registro_existente_actual:
            st.warning(
                f"Ya existen partidas guardadas para {empresa} - {MESES[int(mes)]} {int(anio)}. "
                "Si continúas, se reemplazarán las partidas guardadas de ese período."
            )

            col_conf_1, col_conf_2 = st.columns(2)

            if col_conf_1.button("Sí, reemplazar partidas", use_container_width=True, type="primary"):
                confirmado_generar = True
                st.session_state[confirm_key] = False

            if col_conf_2.button("Cancelar", use_container_width=True):
                st.session_state[confirm_key] = False
                st.rerun()

        st.markdown("</div>", unsafe_allow_html=True)

    partidas_guardadas = []
    partidas_preview = []
    origen_partidas = "preview"

    registro_partidas_guardadas = buscar_partidas_generadas(
        partidas_guardadas_data,
        empresa_id,
        int(anio),
        int(mes)
    )

    partidas_preview = construir_partidas_5(
        ventas_data,
        compras_data,
        empresa,
        empresa_id,
        int(anio),
        int(mes),
        regimen=regimen_empresa,
        pda_inicial=int(pda_inicial),
        incluir_ret_iva=incluir_ret_iva,
        monto_ret_iva=float(monto_ret_iva),
        incluir_exencion_iva=incluir_exencion_iva,
        monto_exencion_iva=float(monto_exencion_iva),
        incluir_ret_isr=incluir_ret_isr,
        monto_ret_isr=float(monto_ret_isr),
        distribucion_cobro=distribucion_cobro if activar_distribucion else None,
        cuenta_egreso=cuenta_egreso,
        activar_distribucion_egreso=activar_distribucion_egreso,
        distribucion_egreso=distribucion_egreso if activar_distribucion_egreso else None,
        usar_clientes_para_cuadre=usar_clientes_para_cuadre,
        activar_distribucion_impuestos=activar_distribucion_impuestos,
        distribucion_impuestos=distribucion_impuestos if activar_distribucion_impuestos else None,
        cuenta_cuadre_ingreso=cuenta_cuadre_ingreso,
        cuenta_cuadre_egreso=cuenta_cuadre_egreso,
        gastos_extraordinarios=gastos_extraordinarios,
        traslados_egreso_ingreso=traslados_egreso_ingreso,
        ingresos_extra_manual=ingresos_extra_manual,
        haberes_manuales_ingreso=haberes_manuales_ingreso,
        haberes_manuales_egreso=haberes_manuales_egreso,
        cuenta_cobro_auto=cuenta_cobro_auto,
        activar_planilla=activar_planilla,
        datos_planilla=datos_planilla,
        modo_pago_planilla=modo_pago_planilla,
        sueldo_por_pagar_manual=sueldo_por_pagar_manual,
        bonificacion_por_pagar_manual=bonificacion_por_pagar_manual,
        monto_retencion_isr_planilla=monto_retencion_isr_planilla,
        activar_provision_prestaciones=activar_provision_prestaciones,
        porcentaje_prestaciones=porcentaje_prestaciones,
        partidas_generadas_data=partidas_guardadas_data,
        incluir_saldos_por_pagar_anteriores=incluir_saldos_por_pagar_anteriores,
        cuentas_excluir_pago_impuestos=cuentas_excluir_pago_impuestos,
        activar_cierre_trimestral=activar_cierre_trimestral,
        usar_isr_trimestral_renta_bruta=usar_isr_trimestral_renta_bruta,
        usar_isr_trimestral_manual=usar_isr_trimestral_manual,
        monto_isr_trimestral_manual=monto_isr_trimestral_manual,
        generar_isr_trimestral=generar_isr_trimestral,
        modo_isr_trimestral=modo_isr_trimestral,
        usar_isr_desde_cierres_parciales=usar_isr_desde_cierres_parciales,
        activar_iso_trimestral=activar_iso_trimestral,
        es_primer_anio_iso=es_primer_anio_iso,
        monto_iso_trimestral_manual=monto_iso_trimestral_manual,
        acreditar_isr_a_iso=acreditar_isr_a_iso,
    )
    partidas_preview = insertar_apertura_en_enero(
        partidas_preview,
        ventas_data,
        compras_data,
        empresa_id,
        int(anio),
        int(mes)
    )

    if registro_partidas_guardadas:
        partidas_preview = preservar_anotaciones_partidas(
            partidas_preview,
            registro_partidas_guardadas.get("partidas", [])
        )

    if registro_partidas_guardadas:
        partidas_guardadas = registro_partidas_guardadas.get("partidas", [])
    if confirmado_generar:
        config_empresa_actual = {
            "cuenta_cobro_auto": str(cuenta_cobro_auto).strip(),
            "pda_inicial": int(pda_inicial),
            "incluir_ret_iva": bool(incluir_ret_iva),
            "monto_ret_iva": float(monto_ret_iva or 0),
            "incluir_exencion_iva": bool(incluir_exencion_iva),
            "monto_exencion_iva": float(monto_exencion_iva or 0),
            "incluir_ret_isr": bool(incluir_ret_isr),
            "monto_ret_isr": float(monto_ret_isr or 0),
            "activar_distribucion": bool(activar_distribucion),
            "cantidad_cuentas": int(cantidad_cuentas) if activar_distribucion else 1,
            "usar_clientes_para_cuadre": bool(usar_clientes_para_cuadre),
            "modo_pago_planilla": str(modo_pago_planilla).strip() if activar_planilla else "mensual",
            "sueldo_por_pagar_manual": float(sueldo_por_pagar_manual or 0.0) if activar_planilla else 0.0,
            "bonificacion_por_pagar_manual": float(bonificacion_por_pagar_manual or 0.0) if activar_planilla else 0.0,
            "monto_retencion_isr_planilla": float(monto_retencion_isr_planilla or 0.0) if activar_planilla else 0.0,
            "cuenta_cuadre_ingreso": str(cuenta_cuadre_ingreso).strip(),
            "distribucion_cobro": distribucion_cobro if activar_distribucion else [],
            "cuenta_egreso": str(cuenta_egreso).strip(),
            "activar_distribucion_egreso": bool(activar_distribucion_egreso),
            "cantidad_cuentas_egreso": int(cantidad_cuentas_egreso) if activar_distribucion_egreso else 1,
            "distribucion_egreso": distribucion_egreso if activar_distribucion_egreso else [],
            "cuenta_cuadre_egreso": str(cuenta_cuadre_egreso).strip(),
            "activar_distribucion_impuestos": bool(activar_distribucion_impuestos),
            "cantidad_cuentas_impuestos": int(cantidad_cuentas_impuestos) if activar_distribucion_impuestos else 1,
            "distribucion_impuestos": distribucion_impuestos if activar_distribucion_impuestos else [],
            "cantidad_gastos_extraordinarios": int(cantidad_gastos_extraordinarios),
            "gastos_extraordinarios": gastos_extraordinarios,
            "cantidad_traslados_egreso_ingreso": int(cantidad_traslados_egreso_ingreso),
            "traslados_egreso_ingreso": limpiar_movimientos_ingreso_config(traslados_egreso_ingreso, permitir_solo_cuenta=True),
            "cantidad_ingresos_extra_manual": int(cantidad_ingresos_extra_manual),
            "ingresos_extra_manual": limpiar_movimientos_ingreso_config(ingresos_extra_manual),
            "cantidad_haberes_manuales_ingreso": int(cantidad_haberes_manuales_ingreso),
            "haberes_manuales_ingreso": limpiar_haberes_manuales(haberes_manuales_ingreso),
            "cantidad_haberes_manuales_egreso": int(cantidad_haberes_manuales_egreso),
            "haberes_manuales_egreso": limpiar_haberes_manuales(haberes_manuales_egreso),
            "activar_planilla": bool(activar_planilla),
            "cantidad_trabajadores": int(cantidad_trabajadores) if activar_planilla else 0,
            "anio_salario": int(anio_salario) if activar_planilla else int(anio),
            "circunscripcion": str(circunscripcion).strip() if activar_planilla else "CE2",
            "tipo_suscripcion": str(tipo_suscripcion).strip() if activar_planilla else "no_agricola",
            "usar_salario_manual": bool(usar_salario_manual) if activar_planilla else False,
            "salario_mensual": float(salario_mensual or 0) if activar_planilla else 0.0,
            "incluir_igss_planilla": bool(incluir_igss_planilla) if activar_planilla else True,
            "incluir_irtra_planilla": bool(incluir_irtra_planilla) if activar_planilla else True,
            "incluir_intecap_planilla": bool(incluir_intecap_planilla) if activar_planilla else True,
            "modo_salario_planilla": str(modo_salario_planilla).strip() if activar_planilla else "general",
            "trabajadores_detalle": trabajadores_detalle if activar_planilla and modo_salario_planilla == "individual" else [],
            "activar_provision_prestaciones": bool(activar_provision_prestaciones) if activar_planilla else True,
            "porcentaje_prestaciones": float(porcentaje_prestaciones or 26.38) if activar_planilla else 26.38,
            "incluir_saldos_por_pagar_anteriores": bool(incluir_saldos_por_pagar_anteriores),
            "cuentas_excluir_pago_impuestos": cuentas_excluir_pago_impuestos if incluir_saldos_por_pagar_anteriores else [],
            "activar_cierre_trimestral": bool(activar_cierre_trimestral),
            "usar_isr_trimestral_renta_bruta": bool(usar_isr_trimestral_renta_bruta),
            "usar_isr_trimestral_manual": bool(usar_isr_trimestral_manual),
            "monto_isr_trimestral_manual": float(monto_isr_trimestral_manual or 0.0),
            "generar_isr_trimestral": bool(generar_isr_trimestral),
            "modo_isr_trimestral": str(modo_isr_trimestral).strip(),
            "usar_isr_desde_cierres_parciales": bool(usar_isr_desde_cierres_parciales),
            "activar_iso_trimestral": bool(activar_iso_trimestral),
            "es_primer_anio_iso": bool(es_primer_anio_iso),
            "monto_iso_trimestral_manual": float(monto_iso_trimestral_manual or 0.0),
            "acreditar_isr_a_iso": bool(acreditar_isr_a_iso),
        }

        guardar_config_empresa(
            config_ui,
            empresa_id,
            int(anio),
            int(mes),
            config_empresa_actual
        )

        if partidas_preview:
            registro_partidas_guardadas = guardar_o_actualizar_partidas_generadas(
                partidas_guardadas_data,
                empresa,
                empresa_id,
                int(anio),
                int(mes),
                partidas_preview
            )
            partidas_guardadas = partidas_preview
            origen_partidas = "guardadas"

    with col_der:

        st.markdown(f"""
        <div class="header-diario">
            <div class="header-diario-titulo">🧾 Diario</div>
            <div class="header-diario-regimen">{empresa}</div>
            <div class="header-diario-regimen">Régimen: {regimen_empresa}</div>
            <div class="header-diario-regimen">{MESES[int(mes)]} {int(anio)}</div>
        </div>
        """, unsafe_allow_html=True)


        mostrar_resumen_superior(
            empresa=empresa,
            regimen=regimen_empresa,
            anio=anio,
            mes=mes,
            activar_distribucion=activar_distribucion if 'activar_distribucion' in locals() else False,
            cantidad_cuentas=cantidad_cuentas if 'cantidad_cuentas' in locals() else 1
        )

        tab1, tab2, tab3, tab4 = st.tabs(["Resumen", "Partidas", "Vista diario", "Descargas"])

        if not partidas_preview:
            with tab1:
                st.markdown("""
                <div class="resumen-vacio">
                    No se pudo formar una vista previa para esta empresa y período. Revisa los datos del panel izquierdo o verifica que existan compras/ventas para ese mes.
                </div>
                """, unsafe_allow_html=True)
            return

        with tab1:
            if confirmado_generar:
                st.success(f"Se guardaron {len(partidas_preview)} partidas.")
            elif partidas_guardadas:
                st.info(f"Vista previa actual para este período. Hay {len(partidas_guardadas)} partidas guardadas previamente.")
            else:
                st.info(f"Vista previa actual de {len(partidas_preview)} partidas para este período.")

            st.markdown('<div class="panel">', unsafe_allow_html=True)
            st.markdown('<div class="panel-title">Resumen</div>', unsafe_allow_html=True)
            if confirmado_generar:
                st.write("La configuración y las partidas fueron guardadas correctamente.")
            elif partidas_guardadas:
                st.write("Estás viendo una vista previa en vivo. Los cambios todavía no se guardan hasta presionar Guardar Cambios.")
            else:
                st.write("Estás viendo una vista previa en vivo para este período. Presiona Guardar Cambios si deseas conservarla.")

            if registro_partidas_guardadas and registro_partidas_guardadas.get("fecha_guardado"):
                st.caption(f"Último guardado: {registro_partidas_guardadas.get('fecha_guardado')}")

            if partidas_guardadas:
                st.write("La pantalla muestra una vista previa en vivo con tus cambios actuales. Lo guardado no se reemplaza hasta que presiones Guardar Cambios.")
            else:
                st.write("La pantalla muestra una vista previa en vivo. Si deseas conservar este período, presiona Guardar Cambios.")

            partidas_cierre_trim = [p for p in partidas_preview if str(p.get("tipo_partida", "")).strip().lower() == "cierre_trimestral"]
            partidas_isr_trim = [p for p in partidas_preview if str(p.get("tipo_partida", "")).strip().lower() == "isr_trimestral"]
            partidas_iso_trim = [p for p in partidas_preview if str(p.get("tipo_partida", "")).strip().lower() == "iso_trimestral"]

            if partidas_isr_trim:
                p_isr_trim = partidas_isr_trim[-1]
                st.info(
                    f"ISR trimestral detectado: Q {float(p_isr_trim.get('isr_trimestral_calculado', 0) or 0):,.2f} "
                    f"| Modo: {str(p_isr_trim.get('modo_isr_trimestral', '') or '').replace('_', ' ')}"
                )

            if partidas_isr_trim:
                p_isr_trim = partidas_isr_trim[-1]
                st.info(
                    f"ISR trimestral detectado: Base Q {float(p_isr_trim.get('isr_trimestral_base', 0) or 0):,.2f} | "
                    f"ISO por acreditar aplicado: Q {float(p_isr_trim.get('aplicado_iso_por_acreditar_isr', 0) or 0):,.2f} | "
                    f"ISR trimestral por pagar: Q {float(p_isr_trim.get('isr_trimestral_por_pagar', 0) or 0):,.2f}"
                )
            if partidas_iso_trim:
                p_iso_trim = partidas_iso_trim[-1]
                st.info(
                    f"ISO trimestral detectado: Base Q {float(p_iso_trim.get('iso_trimestral_base', 0) or 0):,.2f} | "
                    f"ISR acreditado al ISO: Q {float(p_iso_trim.get('aplicado_isr_a_iso', 0) or 0):,.2f} | "
                    f"ISO trimestral por pagar: Q {float(p_iso_trim.get('iso_trimestral_por_pagar', 0) or 0):,.2f}"
                )

            st.markdown("</div>", unsafe_allow_html=True)

        with tab2:
            st.markdown('<div class="panel">', unsafe_allow_html=True)
            st.markdown('<div class="panel-title">Vista previa de partidas</div>', unsafe_allow_html=True)

            for p in partidas_preview:
                with st.expander(f"{p['codigo']} — {p['fecha']}", expanded=False):
                    st.dataframe(pd.DataFrame(p["cuentas"]), use_container_width=True)
                    st.caption(f"{p['glosa']} — Debe: Q {p['total_debe']:,.2f} | Haber: Q {p['total_haber']:,.2f}")
                    if p.get("descuadre_antes_clientes") not in (None, 0, 0.0):
                        st.caption(f"Descuadre previo corregido con {cuenta_cuadre_ingreso}: Q {p['descuadre_antes_clientes']:,.2f}")

            st.markdown("### Anotaciones por cuenta")
            if not partidas_guardadas:
                st.info("Guarda primero las partidas del mes para poder agregar anotaciones.")
            else:
                st.caption(
                    "Las anotaciones pertenecen solo a esta empresa, mes, número de partida y cuenta. "
                    "No modifican los valores contables."
                )
                partidas_anotadas = []

                for indice_partida, p in enumerate(partidas_guardadas):
                    with st.expander(f"Anotar {p['codigo']} - {p['fecha']}", expanded=False):
                        filas_anotacion = [
                            {
                                "Cuenta": mov.get("cuenta", ""),
                                "Debe": float(mov.get("debe", 0) or 0),
                                "Haber": float(mov.get("haber", 0) or 0),
                                "Anotacion": str(mov.get("anotacion", "") or ""),
                            }
                            for mov in p.get("cuentas", []) or []
                        ]
                        df_anotaciones = st.data_editor(
                            pd.DataFrame(filas_anotacion),
                            hide_index=True,
                            use_container_width=True,
                            num_rows="fixed",
                            disabled=["Cuenta", "Debe", "Haber"],
                            column_config={
                                "Cuenta": st.column_config.TextColumn("Cuenta"),
                                "Debe": st.column_config.NumberColumn("Debe", format="Q %.2f"),
                                "Haber": st.column_config.NumberColumn("Haber", format="Q %.2f"),
                                "Anotacion": st.column_config.TextColumn(
                                    "Anotación",
                                    help="Nota exclusiva para esta cuenta en esta partida y mes.",
                                    width="large",
                                ),
                            },
                            key=f"anotaciones_{empresa_id}_{int(anio)}_{int(mes)}_{indice_partida}",
                        )

                    partida_anotada = dict(p)
                    cuentas_anotadas = []
                    for indice_mov, mov in enumerate(p.get("cuentas", []) or []):
                        mov_anotado = dict(mov)
                        valor = df_anotaciones.iloc[indice_mov]["Anotacion"]
                        anotacion = "" if pd.isna(valor) else str(valor).strip()
                        if anotacion:
                            mov_anotado["anotacion"] = anotacion
                        else:
                            mov_anotado.pop("anotacion", None)
                        cuentas_anotadas.append(mov_anotado)
                    partida_anotada["cuentas"] = cuentas_anotadas
                    partidas_anotadas.append(partida_anotada)

                if st.button(
                    "Guardar anotaciones",
                    use_container_width=True,
                    key=f"guardar_anotaciones_{empresa_id}_{int(anio)}_{int(mes)}",
                ):
                    registro_partidas_guardadas = guardar_o_actualizar_partidas_generadas(
                        partidas_guardadas_data,
                        empresa,
                        empresa_id,
                        int(anio),
                        int(mes),
                        partidas_anotadas,
                    )
                    partidas_guardadas = partidas_anotadas
                    partidas_preview = preservar_anotaciones_partidas(
                        partidas_preview,
                        partidas_anotadas,
                    )
                    st.success("Anotaciones guardadas correctamente.")

            st.markdown("</div>", unsafe_allow_html=True)

        with tab3:
            st.markdown('<div class="panel">', unsafe_allow_html=True)
            st.markdown('<div class="panel-title">Vista tipo diario</div>', unsafe_allow_html=True)
            st.dataframe(a_dataframe_visual(partidas_preview), use_container_width=True, height=520)
            st.markdown("</div>", unsafe_allow_html=True)

        with tab4:
            st.markdown('<div class="panel">', unsafe_allow_html=True)
            st.markdown('<div class="panel-title">Descargas</div>', unsafe_allow_html=True)

            st.download_button(
                "📥 Descargar partidas_mensuales.json",
                data=json.dumps(partidas_preview, ensure_ascii=False, indent=4),
                file_name=f"partidas_mensuales_{empresa}_{anio}_{mes:02d}.json",
                mime="application/json",
                use_container_width=True
            )

            st.download_button(
                "📥 Descargar partidas_mensuales.xlsx",
                data=exportar_excel(partidas_preview, mes=mes),
                file_name=f"partidas_mensuales_{empresa}_{anio}_{mes:02d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            st.markdown("</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
