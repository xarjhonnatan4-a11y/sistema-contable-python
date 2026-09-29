import streamlit as st
import pandas as pd

from utils.cuentas import cargar_cuentas, guardar_cuentas, opciones_select

st.title("📚 Catálogo de cuentas")
st.caption("Define y controla las cuentas que se usarán en Compras/Ventas. Así evitas cuentas duplicadas por errores de escritura.")

cuentas = cargar_cuentas()

# Mostrar tabla editable
df = pd.DataFrame(cuentas)
if df.empty:
    df = pd.DataFrame(columns=["codigo","nombre","tipo","naturaleza","activa"])

st.subheader("Cuentas registradas")

df_edited = st.data_editor(
    df,
    use_container_width=True,
    num_rows="dynamic",
    column_config={
        "codigo": st.column_config.TextColumn("Código", help="Ej: 1101, 4101, etc."),
        "nombre": st.column_config.TextColumn("Nombre", help="Nombre de la cuenta."),
        "tipo": st.column_config.SelectboxColumn("Tipo", options=["", "activo","pasivo","patrimonio","ingreso","gasto"]),
        "naturaleza": st.column_config.SelectboxColumn("Naturaleza", options=["", "deudora","acreedora"]),
        "activa": st.column_config.CheckboxColumn("Activa", help="Si está desactivada, no aparecerá en los selectores."),
    },
    key="cuentas_editor"
)

col1, col2 = st.columns([1,1])
with col1:
    if st.button("💾 Guardar catálogo", type="primary"):
        # validación mínima: nombre no vacío
        out = []
        for _, r in df_edited.iterrows():
            nombre = str(r.get("nombre","") or "").strip()
            codigo = str(r.get("codigo","") or "").strip()
            if not nombre and not codigo:
                continue
            out.append({
                "codigo": codigo,
                "nombre": nombre,
                "tipo": str(r.get("tipo","") or "").strip(),
                "naturaleza": str(r.get("naturaleza","") or "").strip(),
                "activa": bool(r.get("activa", True))
            })
        guardar_cuentas(out)
        st.success("✅ Catálogo guardado. Ya puedes usarlo en el Libro de Compras/Ventas.")
with col2:
    st.info(f"Activas: {sum(1 for c in cuentas if c.get('activa', True))} / Total: {len(cuentas)}")

st.divider()
st.subheader("Vista rápida (cómo se verá en los selectores)")
opts = opciones_select(cargar_cuentas())
st.write(opts[:25] + (["…"] if len(opts) > 25 else []))
