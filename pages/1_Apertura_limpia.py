import streamlit as st
import json
import os
from datetime import datetime
import pandas as pd

PARTIDAS_FILE = "partidas.json"
EMPRESAS_FILE = "empresas.json"
CUENTAS_FILE = "cuentas.json"


st.set_page_config(
    page_title="Partida de Apertura",
    page_icon="📂",
    layout="wide"
)


def cargar_json_lista(ruta):
    if os.path.exists(ruta):
        with open(ruta, "r", encoding="utf-8") as f:
            try:
                data = json.load(f)
                return data if isinstance(data, list) else []
            except json.JSONDecodeError:
                return []
    return []


def guardar_json_lista(ruta, data):
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)


def normalizar_nit(nit):
    return str(nit or "").strip().replace("-", "").replace(" ", "").upper()


def normalizar_mes(valor, default=-1):
    if valor is None or valor == "":
        return default

    if isinstance(valor, (int, float)):
        try:
            return int(valor)
        except (TypeError, ValueError):
            return default

    texto = str(valor).strip()
    if not texto:
        return default

    try:
        return int(float(texto))
    except (TypeError, ValueError):
        pass

    texto_norm = (
        texto.lower()
        .replace("á", "a")
        .replace("é", "e")
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ú", "u")
    )

    meses = {
        "apertura": 0,
        "enero": 1,
        "febrero": 2,
        "marzo": 3,
        "abril": 4,
        "mayo": 5,
        "junio": 6,
        "julio": 7,
        "agosto": 8,
        "septiembre": 9,
        "setiembre": 9,
        "octubre": 10,
        "noviembre": 11,
        "diciembre": 12,
    }

    return meses.get(texto_norm, default)


def limpiar_formulario_apertura():
    claves_borrar = []
    for k in list(st.session_state.keys()):
        if (
            str(k).startswith("apertura_cuenta_")
            or str(k).startswith("apertura_debe_")
            or str(k).startswith("apertura_haber_")
            or str(k).startswith("apertura_lineas_")
        ):
            claves_borrar.append(k)

    for k in claves_borrar:
        del st.session_state[k]


def cargar_empresas():
    return cargar_json_lista(EMPRESAS_FILE)


def cargar_cuentas():
    return cargar_json_lista(CUENTAS_FILE)


def buscar_empresa_por_nit(nit, empresas):
    nit_norm = normalizar_nit(nit)
    for e in empresas:
        if normalizar_nit(e.get("nit")) == nit_norm:
            return e
    return None


def opciones_cuentas_catalogo():
    cuentas = cargar_cuentas()
    nombres = []
    for c in cuentas:
        nombre = str(c.get("nombre", "")).strip()
        if nombre:
            nombres.append(nombre)
    return sorted(set(nombres))


def obtener_partida_apertura(partidas, empresa_id, anio):
    empresa_id = normalizar_nit(empresa_id)

    for p in partidas:
        if (
            normalizar_nit(p.get("empresa_nit")) == empresa_id
            and int(p.get("anio", 0) or 0) == int(anio)
            and normalizar_mes(p.get("mes"), default=-1) == 0
        ):
            return p
    return None


def consolidar_movimientos(movimientos):
    acumulado = {}

    for mov in movimientos:
        cuenta = str(mov.get("cuenta", "")).strip()
        debe = round(float(mov.get("debe", 0) or 0), 2)
        haber = round(float(mov.get("haber", 0) or 0), 2)

        if not cuenta:
            continue
        if abs(debe) <= 0.0001 and abs(haber) <= 0.0001:
            continue

        if cuenta not in acumulado:
            acumulado[cuenta] = {"cuenta": cuenta, "debe": 0.0, "haber": 0.0}

        acumulado[cuenta]["debe"] += debe
        acumulado[cuenta]["haber"] += haber

    return [
        {
            "cuenta": k,
            "debe": round(v["debe"], 2),
            "haber": round(v["haber"], 2),
        }
        for k, v in acumulado.items()
        if round(v["debe"], 2) != 0 or round(v["haber"], 2) != 0
    ]


def total_debe_haber(cuentas):
    debe = round(sum(float(x.get("debe", 0) or 0) for x in cuentas), 2)
    haber = round(sum(float(x.get("haber", 0) or 0) for x in cuentas), 2)
    return debe, haber


def crear_partida_apertura(empresa, empresa_id, anio, cuentas):
    cuentas = consolidar_movimientos(cuentas)
    total_debe, total_haber = total_debe_haber(cuentas)

    return {
        "empresa": empresa,
        "empresa_nit": normalizar_nit(empresa_id),
        "anio": int(anio),
        "mes": 0,
        "fecha": f"01/01/{int(anio)}",
        "codigo": "PDA 0",
        "glosa": "Partida de apertura del ejercicio",
        "cuentas": cuentas,
        "total_debe": total_debe,
        "total_haber": total_haber,
        "tipo_partida": "apertura"
    }


def aplicar_estilos():
    st.markdown("""
    <style>
    .block-container {
        padding-top: 1rem;
        padding-bottom: 1rem;
        max-width: 100%;
    }
    .panel {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 16px;
        padding: 18px;
        margin-bottom: 14px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.04);
    }
    .panel-title {
        font-size: 1.05rem;
        font-weight: 700;
        margin-bottom: 10px;
        color: #111827;
    }
    .mini {
        color: #6b7280;
        font-size: 13px;
    }
    </style>
    """, unsafe_allow_html=True)


def main():
    aplicar_estilos()
    st.title("📂 Partida de Apertura")

    empresas = cargar_empresas()
    partidas = cargar_json_lista(PARTIDAS_FILE)
    cuentas_catalogo = opciones_cuentas_catalogo()

    if not empresas:
        st.warning("No se encontraron empresas en empresas.json.")
        return

    if not cuentas_catalogo:
        cuentas_catalogo = [
            "Caja", "Bancos", "Clientes", "Inventario",
            "Proveedores", "Capital"
        ]

    empresas_opciones = sorted(
        empresas,
        key=lambda x: str(x.get("nombre", "")).lower()
    )

    col1, col2 = st.columns([1, 1.2], gap="large")

    with col1:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="panel-title">Datos de apertura</div>', unsafe_allow_html=True)

        empresa_sel = st.selectbox(
            "Empresa",
            empresas_opciones,
            format_func=lambda x: f"{x.get('nombre', '')} ({x.get('nit', '')})"
        )

        empresa_nombre = str(empresa_sel.get("nombre", "")).strip()
        empresa_id = normalizar_nit(empresa_sel.get("nit"))

        anio = st.number_input(
            "Año de apertura",
            min_value=2000,
            max_value=2100,
            value=datetime.now().year,
            step=1
        )

        contexto_actual = f"{empresa_id}_{int(anio)}"
        contexto_anterior = st.session_state.get("apertura_contexto_actual")

        if contexto_anterior != contexto_actual:
            limpiar_formulario_apertura()
            st.session_state["apertura_contexto_actual"] = contexto_actual
            st.rerun()

        apertura_existente = obtener_partida_apertura(partidas, empresa_id, int(anio))

        if apertura_existente:
            st.info("Ya existe una partida de apertura para esta empresa y año. Puedes cargarla, modificarla y guardarla nuevamente.")

        cantidad_lineas_default = len(apertura_existente["cuentas"]) if apertura_existente else 4

        cantidad_lineas = st.number_input(
            "Cantidad de líneas",
            min_value=1,
            max_value=50,
            value=max(1, int(cantidad_lineas_default)),
            step=1,
            key="apertura_lineas_control"
        )

        cuentas_capturadas = []

        for i in range(int(cantidad_lineas)):
            valor = apertura_existente["cuentas"][i] if apertura_existente and i < len(apertura_existente["cuentas"]) else {}

            cuenta_ini = str(valor.get("cuenta", "")).strip()
            debe_ini = float(valor.get("debe", 0) or 0)
            haber_ini = float(valor.get("haber", 0) or 0)

            opciones = cuentas_catalogo.copy()
            if cuenta_ini and cuenta_ini not in opciones:
                opciones = [cuenta_ini] + opciones

            st.markdown(f"**Línea #{i + 1}**")
            c1, c2, c3 = st.columns([2.2, 1, 1])

            cuenta = c1.selectbox(
                "Cuenta",
                opciones,
                index=opciones.index(cuenta_ini) if cuenta_ini in opciones else 0,
                key=f"apertura_cuenta_{i}"
            )

            debe = c2.number_input(
                "Debe",
                min_value=0.0,
                value=float(debe_ini),
                step=0.01,
                key=f"apertura_debe_{i}"
            )

            haber = c3.number_input(
                "Haber",
                min_value=0.0,
                value=float(haber_ini),
                step=0.01,
                key=f"apertura_haber_{i}"
            )

            cuentas_capturadas.append({
                "cuenta": str(cuenta).strip(),
                "debe": float(debe or 0),
                "haber": float(haber or 0)
            })

        cuentas_finales = consolidar_movimientos(cuentas_capturadas)
        total_debe, total_haber = total_debe_haber(cuentas_finales)
        diferencia = round(total_debe - total_haber, 2)

        st.markdown(f"**Total Debe:** Q {total_debe:,.2f}")
        st.markdown(f"**Total Haber:** Q {total_haber:,.2f}")

        if abs(diferencia) > 0.009:
            st.error(f"La partida no cuadra. Diferencia: Q {diferencia:,.2f}")
        else:
            st.success("La partida cuadra correctamente.")

        guardar = st.button("💾 Guardar partida de apertura", use_container_width=True, type="primary")
        st.markdown('</div>', unsafe_allow_html=True)

        if guardar:
            if not cuentas_finales:
                st.warning("Debes ingresar al menos una cuenta con monto.")
                return

            if abs(diferencia) > 0.009:
                st.error("No se puede guardar porque la partida no cuadra.")
                return

            nueva_apertura = crear_partida_apertura(
                empresa=empresa_nombre,
                empresa_id=empresa_id,
                anio=int(anio),
                cuentas=cuentas_finales
            )

            nuevas_partidas = []
            reemplazada = False

            for p in partidas:
                if (
                    normalizar_nit(p.get("empresa_nit")) == empresa_id
                    and int(p.get("anio", 0) or 0) == int(anio)
                    and normalizar_mes(p.get("mes"), default=-1) == 0
                ):
                    nuevas_partidas.append(nueva_apertura)
                    reemplazada = True
                else:
                    nuevas_partidas.append(p)

            if not reemplazada:
                nuevas_partidas.append(nueva_apertura)

            guardar_json_lista(PARTIDAS_FILE, nuevas_partidas)
            st.success("Partida de apertura guardada correctamente.")

    with col2:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="panel-title">Vista previa</div>', unsafe_allow_html=True)

        if "cuentas_capturadas" in locals():
            st.dataframe(pd.DataFrame(cuentas_finales), use_container_width=True, height=420)

        st.markdown('<div class="mini">Se guarda en partidas.json como PDA 0 con mes = 0.</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)


if __name__ == "__main__":
    main()