import streamlit as st
import pandas as pd
import json
import os

st.set_page_config(
    page_title="Salarios mínimos",
    page_icon="💼",
    layout="wide"
)

SALARIOS_MINIMOS_FILE = "salarios_minimos.json"


def aplicar_estilos():
    st.markdown("""
    <style>
    .block-container {
        padding-top: 2rem;
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
        margin-bottom: 12px;
        color: #111827;
    }

    .small-muted {
        color: #6b7280;
        font-size: 13px;
    }

    .header-box {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 12px;
        margin-bottom: 14px;
    }

    .header-title {
        font-size: 1.6rem;
        font-weight: 800;
        color: #111827;
    }

    .header-badge {
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


def cargar_json_dict(ruta):
    if os.path.exists(ruta):
        with open(ruta, "r", encoding="utf-8") as f:
            try:
                data = json.load(f)
                return data if isinstance(data, dict) else {}
            except json.JSONDecodeError:
                return {}
    return {}


def guardar_json_dict(ruta, data):
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)


def cargar_salarios():
    return cargar_json_dict(SALARIOS_MINIMOS_FILE)


def guardar_salarios(data):
    guardar_json_dict(SALARIOS_MINIMOS_FILE, data)


def asegurar_estructura_anio(data, anio):
    anio = str(anio).strip()

    if anio not in data:
        data[anio] = {}

    if "CE1" not in data[anio]:
        data[anio]["CE1"] = {}

    if "CE2" not in data[anio]:
        data[anio]["CE2"] = {}

    for ce in ["CE1", "CE2"]:
        if "agricola" not in data[anio][ce]:
            data[anio][ce]["agricola"] = 0.0
        if "no_agricola" not in data[anio][ce]:
            data[anio][ce]["no_agricola"] = 0.0
        if "maquila" not in data[anio][ce]:
            data[anio][ce]["maquila"] = 0.0

    return data


def salarios_a_dataframe(data):
    filas = []

    for anio in sorted(data.keys(), key=lambda x: int(x)):
        for ce in ["CE1", "CE2"]:
            valores = data.get(anio, {}).get(ce, {})
            filas.append({
                "Año": int(anio),
                "Circunscripción": ce,
                "Agrícola": float(valores.get("agricola", 0) or 0),
                "No agrícola": float(valores.get("no_agricola", 0) or 0),
                "Maquila": float(valores.get("maquila", 0) or 0),
            })

    if not filas:
        return pd.DataFrame(columns=["Año", "Circunscripción", "Agrícola", "No agrícola", "Maquila"])

    return pd.DataFrame(filas)


def main():
    aplicar_estilos()

    st.markdown("""
    <div class="header-box">
        <div class="header-title">💼 Módulo de salarios mínimos</div>
        <div class="header-badge">Archivo: salarios_minimos.json</div>
    </div>
    """, unsafe_allow_html=True)

    salarios_data = cargar_salarios()

    col1, col2 = st.columns([1, 1.4], gap="large")

    with col1:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="panel-title">Registrar / editar año</div>', unsafe_allow_html=True)

        anios_existentes = sorted(
            [int(x) for x in salarios_data.keys() if str(x).isdigit()]
        )

        opciones_anio = []
        if anios_existentes:
            opciones_anio = anios_existentes.copy()

        anio_manual = st.number_input(
            "Año",
            min_value=2020,
            step=1,
            value=anios_existentes[-1] if anios_existentes else 2026
        )

        anio = str(int(anio_manual))
        salarios_data = asegurar_estructura_anio(salarios_data, anio)

        st.markdown("#### CE1")
        ce1_agricola = st.number_input(
            "CE1 - Agrícola",
            min_value=0.0,
            value=float(salarios_data[anio]["CE1"].get("agricola", 0) or 0),
            step=0.01,
            key=f"ce1_agricola_{anio}"
        )
        ce1_no_agricola = st.number_input(
            "CE1 - No agrícola",
            min_value=0.0,
            value=float(salarios_data[anio]["CE1"].get("no_agricola", 0) or 0),
            step=0.01,
            key=f"ce1_no_agricola_{anio}"
        )
        ce1_maquila = st.number_input(
            "CE1 - Maquila",
            min_value=0.0,
            value=float(salarios_data[anio]["CE1"].get("maquila", 0) or 0),
            step=0.01,
            key=f"ce1_maquila_{anio}"
        )

        st.markdown("#### CE2")
        ce2_agricola = st.number_input(
            "CE2 - Agrícola",
            min_value=0.0,
            value=float(salarios_data[anio]["CE2"].get("agricola", 0) or 0),
            step=0.01,
            key=f"ce2_agricola_{anio}"
        )
        ce2_no_agricola = st.number_input(
            "CE2 - No agrícola",
            min_value=0.0,
            value=float(salarios_data[anio]["CE2"].get("no_agricola", 0) or 0),
            step=0.01,
            key=f"ce2_no_agricola_{anio}"
        )
        ce2_maquila = st.number_input(
            "CE2 - Maquila",
            min_value=0.0,
            value=float(salarios_data[anio]["CE2"].get("maquila", 0) or 0),
            step=0.01,
            key=f"ce2_maquila_{anio}"
        )

        colb1, colb2 = st.columns(2)

        guardar = colb1.button("💾 Guardar año", use_container_width=True, type="primary")
        eliminar = colb2.button("🗑️ Eliminar año", use_container_width=True)

        if guardar:
            salarios_data[anio]["CE1"]["agricola"] = round(float(ce1_agricola or 0), 2)
            salarios_data[anio]["CE1"]["no_agricola"] = round(float(ce1_no_agricola or 0), 2)
            salarios_data[anio]["CE1"]["maquila"] = round(float(ce1_maquila or 0), 2)

            salarios_data[anio]["CE2"]["agricola"] = round(float(ce2_agricola or 0), 2)
            salarios_data[anio]["CE2"]["no_agricola"] = round(float(ce2_no_agricola or 0), 2)
            salarios_data[anio]["CE2"]["maquila"] = round(float(ce2_maquila or 0), 2)

            guardar_salarios(salarios_data)
            st.success(f"Salarios del año {anio} guardados correctamente.")

        if eliminar:
            if anio in salarios_data:
                del salarios_data[anio]
                guardar_salarios(salarios_data)
                st.success(f"Año {anio} eliminado correctamente.")
                st.rerun()
            else:
                st.warning("Ese año no existe en el archivo.")

        st.caption("Puedes usar este módulo para registrar los salarios mínimos manualmente año con año.")
        st.markdown("</div>", unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="panel-title">Tabla de salarios registrados</div>', unsafe_allow_html=True)

        df = salarios_a_dataframe(salarios_data)
        st.dataframe(df, use_container_width=True, height=420)

        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="panel-title">Vista JSON</div>', unsafe_allow_html=True)

        st.code(json.dumps(salarios_data, ensure_ascii=False, indent=4), language="json")

        st.download_button(
            "📥 Descargar salarios_minimos.json",
            data=json.dumps(salarios_data, ensure_ascii=False, indent=4),
            file_name="salarios_minimos.json",
            mime="application/json",
            use_container_width=True
        )

        st.markdown("</div>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()