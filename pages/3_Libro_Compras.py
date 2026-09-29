import streamlit as st 
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
import io
import os
import json
from datetime import datetime
from openpyxl.styles import PatternFill
import calendar

from utils.cuentas import cargar_cuentas, opciones_nombres, asegurar_cuentas_por_defecto
from utils.encoding import normalizar_columnas_dataframe

def normalizar_nit(valor):
    import pandas as pd
    if pd.isna(valor):
        return ""
    valor = str(valor).strip()
    if valor.endswith(".0"):
        valor = valor[:-2]
    return valor



CONFIG_FILE = "empresas.json"
CLASIFICACIONES_FILE = "clasificaciones.json"

DEFAULT_CUENTAS_SEMILLA = ["OTRAS","COMBUSTIBLE","COMBUSTIBLE GASTO","ENERGÍA ELÉCTRICA","Compras","IVA Crédito","IDP","IDP Gasto","Tasa Municipal","Otros Impuestos","Bancos"]
asegurar_cuentas_por_defecto(DEFAULT_CUENTAS_SEMILLA)


def cargar_empresas():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def cargar_clasificaciones():
    if os.path.exists(CLASIFICACIONES_FILE):
        try:
            with open(CLASIFICACIONES_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError:
            return {}
    return {}

def guardar_clasificaciones(data):
    with open(CLASIFICACIONES_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

def normalizar_nit_empresa(nit):
    return str(nit or "").strip().replace("-", "").replace(" ", "").upper()

# Helper: normaliza lo que venga del data_editor a un DataFrame seguro
def normalize_editable(obj):
    if isinstance(obj, pd.DataFrame):
        return obj.copy()
    if isinstance(obj, dict):
        # Intentar como dict columnas -> listas
        try:
            df = pd.DataFrame(obj)
            df.columns = df.columns.map(str)
            return df
        except Exception:
            pass
        # Intentar como dict index->rowdict
        try:
            df = pd.DataFrame.from_dict(obj, orient='index')
            df.reset_index(drop=True, inplace=True)
            df.columns = df.columns.map(str)
            return df
        except Exception:
            pass
    if isinstance(obj, list):
        try:
            df = pd.DataFrame(obj)
            df.columns = df.columns.map(str)
            return df
        except Exception:
            pass
    # fallback vacío
    return pd.DataFrame()


def es_combustible_gasto(valor):
    """Devuelve True si la clasificación debe enviar el IDP a la cuenta IDP Gasto."""
    clasif = str(valor or "").strip().upper()
    return clasif in {"COMBUSTIBLE GASTO", "COMBUSTIBLE GASTOS"}

empresas = cargar_empresas()
clasificaciones_global = cargar_clasificaciones()

def app():
    st.title("📘 Generador de Libro Contable - Compras SAT")

    if not empresas:
        st.warning("⚠️ No hay empresas registradas. Primero registre una en el menú principal.")
        return

    empresas_nombres = [e["nombre"] for e in empresas]
    empresa_sel = st.selectbox("Seleccione la empresa", empresas_nombres)
    empresa_data = next(e for e in empresas if e["nombre"] == empresa_sel)
    empresa_nit = normalizar_nit_empresa(empresa_data.get("nit"))

    clasificaciones_empresa = clasificaciones_global.get(empresa_nit, {})

    # Compatibilidad con archivo viejo por nombre
    if not clasificaciones_empresa:
        clasificaciones_empresa = clasificaciones_global.get(empresa_sel, {})

    anio_actual = datetime.now().year
    anio = st.number_input("Año a trabajar", value=anio_actual)

    mes_actual = datetime.now().month
    mes_defecto = 12 if mes_actual == 1 else mes_actual - 1
    meses_nombre = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
                    "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

    mes_num = st.number_input("Mes a trabajar", min_value=1, max_value=12, value=mes_defecto)
    mes_num = int(mes_num)
    mes = meses_nombre[(mes_num - 1) % 12]

    ultimo_folio = st.number_input("Último folio utilizado", min_value=0, value=0)
    archivo_excel = st.file_uploader("Sube el archivo Excel descargado del SAT", type=["xls", "xlsx"])

    st.info(f"""
    **NIT:** {empresa_data.get('nit', '')}  
    **Propietario:** {empresa_data['propietario']}  
    **Régimen:** {empresa_data['regimen']}  
    **Gasolinera:** {"Sí" if empresa_data['es_gasolinera'] else "No"}  
    """)

    if archivo_excel:
        try:
            extension = os.path.splitext(archivo_excel.name)[1]
            if extension == '.xls':
                df = pd.read_excel(archivo_excel, engine='xlrd', dtype={"NIT del emisor": str})
            else:
                df = pd.read_excel(archivo_excel, dtype={"NIT del emisor": str})

            df = normalizar_columnas_dataframe(df)
            if "NIT del emisor" in df.columns:
                df["NIT del emisor"] = df["NIT del emisor"].apply(normalizar_nit)

            # Filtrar anulados de entrada
            df = df[df["Estado"].astype(str).str.upper() != "ANULADO"]

            columnas_requeridas = [
                'Fecha de emisión', 'Número de Autorización', 'Tipo de DTE (nombre)',
                'Serie', 'Número del DTE', 'Clasificación emisor', 'Exportación',
                'NIT del emisor', 'Nombre completo del emisor', 'Código de establecimiento',
                'Nombre del establecimiento', 'ID del receptor', 'Nombre completo del receptor',
                'NIT del Certificador', 'Nombre completo del Certificador', 'Estado',
                'Moneda', 'Gran Total (Moneda Original)', 'IVA (monto de este impuesto)',
                'Marca de anulado', 'Fecha de anulación', 'Petróleo (monto de este impuesto)'
            ]

            faltantes = [col for col in columnas_requeridas if col not in df.columns]
            if faltantes:
                st.error(f"❌ Faltan columnas en el archivo: {faltantes}")
                st.stop()

            df = df[columnas_requeridas]
            df['DIA'] = pd.to_datetime(df['Fecha de emisión'].astype(str).str.split('T').str[0], errors='coerce').dt.day
            df = df.sort_values(by="DIA").reset_index(drop=True)

            # Columnas nuevas
            df["IDP"] = 0.00
            df["Tasa Municipal"] = 0.00
            df["Otros Impuestos"] = 0.00
            df["Total"] = 0.00
            df["Compra Neta"] = 0.00
            df["IVA Crédito"] = 0.00
            df["Clasificación"] = ""

            # Calcula por fila (manteniendo tu lógica)
            for i, row in df.iterrows():
                tipo_dte = str(row["Tipo de DTE (nombre)"]).upper()
                total_original = float(row["Gran Total (Moneda Original)"])
                iva = float(row["IVA (monto de este impuesto)"] or 0)
                idp_archivo = float(row["Petróleo (monto de este impuesto)"] or 0)
                nit = normalizar_nit(row["NIT del emisor"])

                # CASOS ESPECIALES
                if nit == "14946211":
                    # DEOCSA → energía eléctrica
                    df.at[i, "Clasificación"] = "ENERGÍA ELÉCTRICA"
                    # Calculamos igual para que queden montos, si quieres otro tratamiento puedes ajustar
                    compra_neta = round((abs(iva) * 100) / 12, 2) 
                    df.at[i, "Compra Neta"] = compra_neta
                    df.at[i, "IVA Crédito"] = round(iva, 2)
                    diferencia = total_original - (compra_neta + iva)
                    df.at[i, "Tasa Municipal"] = diferencia if diferencia > 0 else 0.00
                    df.at[i, "IDP"] = 0.00
                    df.at[i, "Otros Impuestos"] = 0.00
                    df.at[i, "Total"] = total_original - df.at[i, "Tasa Municipal"]
                    continue

                if tipo_dte == "FPEQ":
                    df.at[i, "Clasificación"] = "PEQUEÑO CONTRIBUYENTE"
                    df.at[i, "Total"] = total_original
                    df.at[i, "Compra Neta"] = total_original
                    df.at[i, "IVA Crédito"] = 0.00
                    df.at[i, "IDP"] = 0.00
                    df.at[i, "Tasa Municipal"] = 0.00
                    df.at[i, "Otros Impuestos"] = 0.00
                    continue

                if tipo_dte == "RDON":
                    df.at[i, "Clasificación"] = "DONACIÓN"
                    df.at[i, "Total"] = total_original
                    df.at[i, "Compra Neta"] = total_original
                    df.at[i, "IVA Crédito"] = 0.00
                    df.at[i, "IDP"] = 0.00
                    df.at[i, "Tasa Municipal"] = 0.00
                    df.at[i, "Otros Impuestos"] = 0.00
                    continue

                # Compra Neta basada en IVA (IVA = 12% => neto = IVA * 100 / 12)
                compra_neta = round((abs(iva) * 100) / 12, 2) if iva != 0 else 0.00
                df.at[i, "Compra Neta"] = compra_neta
                df.at[i, "IVA Crédito"] = round(iva, 2)  # Guardar siempre el IVA real


                diferencia = total_original - (compra_neta + iva)

                # IDP (si en archivo vino Petróleo)
                df.at[i, "IDP"] = diferencia if idp_archivo > 0 and diferencia > 0 else 0.00

                # Tasa Municipal sólo para DEOCSA (ya tratado arriba), si no → 0
                df.at[i, "Tasa Municipal"] = 0.00

                # Si hay diferencia y no es IDP -> otros impuestos (>=1 quetzal)
                if diferencia >= 1 and df.at[i, "IDP"] == 0:
                    df.at[i, "Otros Impuestos"] += diferencia

                # Sumar impuestos extra si existen en el archivo
                impuestos_extra = [
                    "Turismo Hospedaje (monto de este impuesto)",
                    "Turismo Pasajes (monto de este impuesto)",
                    "Timbre de Prensa (monto de este impuesto)",
                    "Bomberos (monto de este impuesto)",
                    "Tasa Municipal (monto de este impuesto)",
                    "Bebidas alcohólicas (monto de este impuesto)",
                    "Tabaco (monto de este impuesto)",
                    "Cemento (monto de este impuesto)",
                    "Bebidas no Alcohólicas (monto de este impuesto)",
                    "Tarifa Portuaria (monto de este impuesto)"
                ]
                for impuesto in impuestos_extra:
                    if impuesto in df.columns:
                        valor = float(row.get(impuesto, 0) or 0)
                        # ignorar decimales < 1 en la suma principal (ya pusiste regla diferencia >=1)
                        df.at[i, "Otros Impuestos"] += valor

                df.at[i, "Total"] = total_original - df.at[i, "IDP"] - df.at[i, "Tasa Municipal"] - df.at[i, "Otros Impuestos"]

                # CASO NCRE → valores negativos
                if tipo_dte == "NCRE":
                    df.at[i, "Total"] = -abs(df.at[i, "Total"])
                    df.at[i, "Compra Neta"] = -abs(df.at[i, "Compra Neta"])
                    df.at[i, "IVA Crédito"] = -abs(df.at[i, "IVA Crédito"])
                    df.at[i, "Otros Impuestos"] = -abs(df.at[i, "Otros Impuestos"])

                # Clasificación final por NIT o por clasificaciones guardadas
                if nit == "14946211":
                    df.at[i, "Clasificación"] = "ENERGÍA ELÉCTRICA"
                elif nit in clasificaciones_empresa:
                    df.at[i, "Clasificación"] = clasificaciones_empresa[nit]
                elif idp_archivo > 0:
                    df.at[i, "Clasificación"] = "COMBUSTIBLE"
                else:
                    df.at[i, "Clasificación"] = "OTRAS"

            columnas_exportar = [
                "DIA", "Tipo de DTE (nombre)", "Número del DTE", "NIT del emisor",
                "Nombre completo del emisor", "Clasificación", "IDP", "Tasa Municipal", "Otros Impuestos",
                "Total", "Compra Neta", "IVA Crédito"
            ]
            df = df[columnas_exportar]

            # Guardar en session_state para editar
            st.session_state.df_editable = df.copy()

            # Mostrar editor y obtener el resultado EDITADO por el usuario
            # Crear estilos condicionales
            # Crear columna "Señal" solo para mostrar
            df_mostrar = st.session_state.df_editable.copy()
            df_mostrar["Señal"] = ""

            for i, row in df_mostrar.iterrows():
                tipo_dte = str(row["Tipo de DTE (nombre)"]).upper()
                if tipo_dte not in ["FACT", "FPEQ", "FCAM", "NCRE"]:
                    df_mostrar.at[i, "Señal"] = "⚠️"

            # 🔹 Editor interactivo
            df_edited = st.data_editor(
                df_mostrar,
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "Clasificación": st.column_config.SelectboxColumn(
                        "Cuenta contable",
                        options=opciones_nombres(cargar_cuentas()),
                        required=True,
                        help="Selecciona una cuenta del Catálogo (evita errores por escritura)."
                    ),
                    "Señal": st.column_config.TextColumn("Señal", disabled=True),
                },
                key="editor"
            )
            st.session_state.df_editable.update(normalize_editable(df_edited))
            # 🔧 Si tienes esta función, déjala; si no, coméntala
            # df_edited = normalize_editable(df_edited)

            # ✅ Validar cuentas contra catálogo (evita cuentas "inventadas")
            cuentas_validas = set(opciones_nombres(cargar_cuentas()))
            if "Clasificación" in df_edited.columns and cuentas_validas:
                invalidas = sorted({str(x).strip() for x in df_edited["Clasificación"].astype(str).unique() if str(x).strip() and str(x).strip() not in cuentas_validas})
                if invalidas:
                    st.warning("⚠️ Se detectaron cuentas que NO están en el Catálogo. Se reemplazaron por 'OTRAS'. Revisa el Catálogo si necesitas agregarlas.")
                    df_edited.loc[df_edited["Clasificación"].astype(str).isin(invalidas), "Clasificación"] = "OTRAS"

            # 🔁 Refresco inmediato si el usuario modifica algo
            if "df_prev" not in st.session_state:
                st.session_state.df_prev = df_edited.copy()
            elif not df_edited.equals(st.session_state.df_prev):
                st.session_state.df_editable = df_edited.copy()
                st.session_state.df_prev = df_edited.copy()
                st.rerun()


            # Normalizar lo que el editor devolvió (puede ser dict/list)
            df_edited = normalize_editable(df_edited)

            # Asegurar que todas las columnas que exportamos existan (si el usuario borró, las recreamos)
            for col in columnas_exportar:
                if col not in df_edited.columns:
                    df_edited[col] = ""

            # Reordenar y convertir tipos
            df_edited = df_edited[columnas_exportar].copy()
            # convertir números
            num_cols = ["IDP", "Tasa Municipal", "Otros Impuestos", "Total", "Compra Neta", "IVA Crédito"]
            for c in num_cols:
                df_edited[c] = pd.to_numeric(df_edited[c], errors='coerce').fillna(0.0)

            # Guardar el DataFrame limpio en session_state para uso posterior
            st.session_state.df_editable = df_edited.copy()

            # Actualizar / guardar clasificaciones por NIT según lo editado
            if "clasificaciones_por_nit" not in st.session_state:
                st.session_state.clasificaciones_por_nit = {}
                for nit in df_edited["NIT del emisor"].apply(normalizar_nit).unique():
                    st.session_state.clasificaciones_por_nit[nit] = df_edited.loc[df_edited["NIT del emisor"].astype(str) == nit, "Clasificación"].iloc[0]

            cambios = False
            for i, row in df_edited.iterrows():
                nit = normalizar_nit(row["NIT del emisor"])
                clasif = row["Clasificación"]

                # Añadir señal si NO es FACT, FPEQ o FCAM
                tipo_dte = str(row["Tipo de DTE (nombre)"]).upper()
                if tipo_dte not in ["FACT", "FPEQ", "FCAM"]:
                    df_edited.at[i, "Señal"] = "⚠️ Documento diferente"
                else:
                    df_edited.at[i, "Señal"] = ""

                if st.session_state.clasificaciones_por_nit.get(nit) != clasif:
                    st.session_state.clasificaciones_por_nit[nit] = clasif
                    clasificaciones_empresa[nit] = clasif
                    cambios = True

            if cambios:
                clasificaciones_global[empresa_nit] = clasificaciones_empresa

                # limpieza opcional del formato viejo por nombre
                if empresa_sel in clasificaciones_global:
                    del clasificaciones_global[empresa_sel]

                guardar_clasificaciones(clasificaciones_global)

                # reflejar cambios en el df_edited
                for nit, clasif in st.session_state.clasificaciones_por_nit.items():
                    df_edited.loc[df_edited["NIT del emisor"].apply(normalizar_nit) == nit, "Clasificación"] = clasif

                st.session_state.df_editable = df_edited.copy()

            # ------------------- Datos NCRE -------------------
            # ------------------- Resumen por tipo de documento -------------------
            tipos_normales = ["FACT", "FPEQ", "FCAM"]
            tipos_controlados = tipos_normales + ["NCRE"]

            tipos_dte = df_edited["Tipo de DTE (nombre)"].astype(str).str.upper().str.strip()

            # Facturas normales
            df_facturas = df_edited[tipos_dte.isin(tipos_normales)]
            total_facturas = len(df_facturas)

            # NCRE
            df_ncre = df_edited[tipos_dte == "NCRE"]
            cantidad_ncre = len(df_ncre)
            total_ncre = df_ncre["Compra Neta"].sum()

            # Otros documentos
            df_otros = df_edited[~tipos_dte.isin(tipos_controlados)].copy()
            total_otros_docs = len(df_otros)

            otros_resumen = (
                df_otros.assign(_tipo=tipos_dte[~tipos_dte.isin(tipos_controlados)].values)
                .groupby("_tipo")
                .size()
                .reset_index(name="Cantidad")
                .rename(columns={"_tipo": "Tipo"})
            )

            otros_tipos_texto = ", ".join(
                [f"{row['Tipo']} ({row['Cantidad']})" for _, row in otros_resumen.iterrows()]
            ) if not otros_resumen.empty else "Ninguno"

            if total_otros_docs > 0:
                st.warning(
                    f"⚠️ Se detectaron {total_otros_docs} documento(s) de tipo distinto a FACT, FPEQ, FCAM y NCRE: {otros_tipos_texto}"
                )
            else:
                st.success("✅ Solo se detectaron documentos FACT, FPEQ, FCAM y/o NCRE.")

            # ------------------- RESUMEN FINAL -------------------
            resumen_compra = df_edited.groupby("Clasificación").agg({
                "Compra Neta": "sum"
            }).reset_index()

            total_compra_neta = resumen_compra["Compra Neta"].sum() if not resumen_compra.empty else 0
            total_iva_credito = df_edited["IVA Crédito"].sum()

            mask_idp_gasto = df_edited["Clasificación"].apply(es_combustible_gasto)
            total_idp_gasto = df_edited.loc[mask_idp_gasto, "IDP"].sum()
            total_idp = df_edited.loc[~mask_idp_gasto, "IDP"].sum()

            total_tasa = df_edited["Tasa Municipal"].sum()
            total_otros_impuestos = df_edited["Otros Impuestos"].sum()

            # Sólo incluir líneas del resumen extra si su total != 0
            resumen_extra = []
            if total_iva_credito != 0:
                resumen_extra.append(["IVA Crédito", total_iva_credito])
            if total_idp != 0:
                resumen_extra.append(["IDP", total_idp])
            if total_idp_gasto != 0:
                resumen_extra.append(["IDP Gasto", total_idp_gasto])
            if total_tasa != 0:
                resumen_extra.append(["Tasa Municipal", total_tasa])
            if total_otros_impuestos != 0:
                resumen_extra.append(["Otros Impuestos", total_otros_impuestos])

            resumen_extra.append([
                "TOTAL GENERAL",
                total_compra_neta + total_iva_credito + total_idp + total_idp_gasto + total_tasa + total_otros_impuestos
            ])

            resumen_compra = pd.concat([
                resumen_compra,
                pd.DataFrame(resumen_extra, columns=resumen_compra.columns)
            ], ignore_index=True)


        

            st.markdown("### 📊 Resumen Final")
            st.dataframe(resumen_compra, use_container_width=True)

            # BOTÓN: GENERAR EXCEL
            if st.button("Generar libro contable"):
                clasificaciones_empresa.update(st.session_state.clasificaciones_por_nit)
                clasificaciones_global[empresa_nit] = clasificaciones_empresa

                # limpieza opcional del formato viejo por nombre
                if empresa_sel in clasificaciones_global:
                    del clasificaciones_global[empresa_sel]

                guardar_clasificaciones(clasificaciones_global)

                # Tomar el DataFrame que está realmente editado en session_state
                df_to_export = normalize_editable(st.session_state.df_editable)

                # Asegurar columnas (evita problemas si algo fue eliminado)
                for col in columnas_exportar:
                    if col not in df_to_export.columns:
                        df_to_export[col] = ""

                df_to_export = df_to_export[columnas_exportar].copy()

                # Eliminar filas vacías/incompletas: se considera vacía si faltan las columnas clave
                columnas_clave = ["Número del DTE", "NIT del emisor"]
                df_to_export = df_to_export.dropna(subset=columnas_clave, how="any")
                df_to_export = df_to_export.loc[df_to_export.notna().any(axis=1)].reset_index(drop=True)

                # Convertir numéricos por si acaso
                for c in num_cols:
                    df_to_export[c] = pd.to_numeric(df_to_export[c], errors='coerce').fillna(0.0)
                tipos_normales_export = ["FACT", "FPEQ", "FCAM"]
                tipos_controlados_export = tipos_normales_export + ["NCRE"]

                tipos_dte_export = df_to_export["Tipo de DTE (nombre)"].astype(str).str.upper().str.strip()

                total_facturas_export = int(tipos_dte_export.isin(tipos_normales_export).sum())

                df_ncre_export = df_to_export[tipos_dte_export == "NCRE"]
                cantidad_ncre_export = len(df_ncre_export)
                total_ncre_export = float(df_ncre_export["Compra Neta"].sum())

                df_otros_export = df_to_export[~tipos_dte_export.isin(tipos_controlados_export)].copy()
                total_otros_docs_export = len(df_otros_export)

                otros_resumen_export = (
                    df_otros_export.assign(_tipo=tipos_dte_export[~tipos_dte_export.isin(tipos_controlados_export)].values)
                    .groupby("_tipo")
                    .size()
                    .reset_index(name="Cantidad")
                    .rename(columns={"_tipo": "Tipo"})
                )

                otros_tipos_texto_export = ", ".join(
                    [f"{row['Tipo']} ({row['Cantidad']})" for _, row in otros_resumen_export.iterrows()]
                ) if not otros_resumen_export.empty else "Ninguno"

                generar_excel(
                    df_to_export,
                    empresa_data,
                    anio,
                    mes,
                    ultimo_folio,
                    columnas_exportar,
                    resumen_compra,
                    total_facturas_export,
                    cantidad_ncre_export,
                    total_ncre_export,
                    total_otros_docs_export,
                    otros_tipos_texto_export
                )
                guardar_partida_compras(empresa_data, anio, mes_num, resumen_compra, contrapartida_nombre="Bancos")

        except Exception as e:
            st.error(f"Ocurrió un error: {e}")



# Archivo donde se guardarán las partidas del libro de compras
PARTIDAS_COMPRAS_FILE = "partidascompras.json"

def guardar_partida_compras(empresa_data, anio, mes, resumen_compra, contrapartida_nombre="Bancos"):
    """
    Genera y guarda automáticamente la partida contable del libro de compras.
    Usa el DataFrame resumen_compra (columnas: 'Clasificación' y 'Compra Neta').
    Añade líneas por IVA Crédito, IDP, Tasa Municipal y Otros Impuestos si existen.
    Guarda en PARTIDAS_COMPRAS_FILE evitando duplicados por empresa/año/mes.
    """
    try:
        empresa_nombre = str(empresa_data.get("nombre", "")).strip()
        empresa_nit = str(empresa_data.get("nit", "")).strip()
        
        # Obtener último día del mes
        ultimo_dia = calendar.monthrange(int(anio), int(mes))[1]
        fecha_partida = datetime(int(anio), int(mes), ultimo_dia).strftime("%d/%m/%Y")

        

        partida_json = {
            "empresa": empresa_nombre,
            "empresa_nit": empresa_nit,
            "anio": int(anio),
            "mes": int(mes),
            "fecha": fecha_partida,
            "codigo": f"PDA {str(ultimo_dia).zfill(2)}",
            "glosa": "Registro de los egresos efectuados en el presente mes",
            "cuentas": []
        }

        total_debe = 0.0
        df_res = resumen_compra.copy()

        # Asegurarse de que las columnas existan
        if "Compra Neta" not in df_res.columns and "CompraNeta" in df_res.columns:
            df_res = df_res.rename(columns={"CompraNeta": "Compra Neta"})

        es_gasolinera = bool(empresa_data.get("es_gasolinera", False))
        cuentas_omitidas_gasolinera = {"COMBUSTIBLE", "IDP"}

        # Agregar clasificaciones normales
        for _, fila in df_res.iterrows():
            nombre = str(fila.get("Clasificación", "")).strip()
            monto = float(fila.get("Compra Neta", 0) or 0)
            nombre_upper = nombre.upper()

            if nombre_upper.startswith("TOTAL") or monto == 0:
                continue

            # Solo para el JSON: en gasolineras, COMBUSTIBLE e IDP se distribuyen manualmente
            if es_gasolinera and nombre_upper in cuentas_omitidas_gasolinera:
                continue

            partida_json["cuentas"].append({
                "nombre": nombre if nombre else "OTRAS",
                "debe": round(monto, 2),
                "haber": 0.0
            })
            total_debe += monto

        # Contrapartida (haber)
        partida_json["cuentas"].append({
            "nombre": contrapartida_nombre,
            "debe": 0.0,
            "haber": round(total_debe, 2)
        })

        # Cargar archivo existente
        if os.path.exists(PARTIDAS_COMPRAS_FILE):
            with open(PARTIDAS_COMPRAS_FILE, "r", encoding="utf-8") as f:
                try:
                    partidas_data = json.load(f)
                except json.JSONDecodeError:
                    partidas_data = []
        else:
            partidas_data = []

        # Evitar duplicados (empresa + año + mes)
        existe = any(
            (
                str(p.get("empresa_nit", "")).strip() == empresa_nit
                or str(p.get("empresa", "")).strip() == empresa_nombre
            ) and
            int(p.get("anio", 0) or 0) == int(anio) and
            int(p.get("mes", 0) or 0) == int(mes)
            for p in partidas_data
        )
        if existe:
            st.warning("ℹ️ Ya existe una partida para esta empresa, año y mes. No se guardó nuevamente.")
            return

        # Guardar nueva partida
        partidas_data.append(partida_json)
        with open(PARTIDAS_COMPRAS_FILE, "w", encoding="utf-8") as f:
            json.dump(partidas_data, f, indent=4, ensure_ascii=False)

        st.success("✅ Partida guardada correctamente en 'partidascompras.json'")

    except Exception as e:
        st.error(f"❌ Ocurrió un error al guardar la partida: {e}")

# ===== Función generar_excel con folios, bordes y VIENEN/VAN =====

def generar_excel(
    df,
    empresa_data,
    anio,
    mes,
    ultimo_folio,
    columnas_exportar,
    resumen_compra,
    total_facturas,
    cantidad_ncre,
    total_ncre,
    total_otros_docs,
    otros_tipos_texto
):
    # Si no hay filas, generar un excel con encabezado vacío
    total_lineas = len(df)
    
    lineas_por_folio = 50
    if total_lineas == 0:
        total_folios = 1
    else:
        total_folios = (total_lineas + lineas_por_folio - 1) // lineas_por_folio

    wb = Workbook()
    ws = wb.active
    ws.title = "Libro Compras"

    fila_actual = 1
    acumulado = {
        "IDP": 0.0,
        "Tasa Municipal": 0.0,
        "Otros Impuestos": 0.0,
        "Total": 0.0,
        "Compra Neta": 0.0,
        "IVA Crédito": 0.0
    }

    estilo_folio = Font(bold=True)
    estilo_encabezado = Font(bold=True)
    estilo_total = Font(bold=True)
    borde = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # columnas numéricas en el orden correcto (índices 0-based en tuple)
    idx_IDP = columnas_exportar.index("IDP")
    idx_Tasa = columnas_exportar.index("Tasa Municipal")
    idx_Otros = columnas_exportar.index("Otros Impuestos")
    idx_Total = columnas_exportar.index("Total")
    idx_Compra = columnas_exportar.index("Compra Neta")
    idx_IVA = columnas_exportar.index("IVA Crédito")

    for folio in range(total_folios):
        folio_actual = ultimo_folio + folio + 1

        # Encabezados del folio
        ws.merge_cells(start_row=fila_actual, start_column=1, end_row=fila_actual, end_column=len(columnas_exportar))
        cell = ws.cell(row=fila_actual, column=1)
        cell.value = f"{empresa_data['nombre']} | NIT: {empresa_data.get('nit', '')} | Año: {anio} | Mes: {mes} | Folio: {folio_actual}"
        cell.font = estilo_folio
        cell.alignment = Alignment(horizontal='center')
        fila_actual += 1

        ws.merge_cells(start_row=fila_actual, start_column=1, end_row=fila_actual, end_column=len(columnas_exportar))
        cell = ws.cell(row=fila_actual, column=1)
        cell.value = "CIFRAS EXPRESADAS EN QUETZALES - LIBRO DE COMPRAS"
        cell.font = Font(bold=True, italic=True)
        cell.alignment = Alignment(horizontal='center')
        fila_actual += 1

        # Encabezado de columnas
        for col_num, col_name in enumerate(columnas_exportar, 1):
            cell = ws.cell(row=fila_actual, column=col_num, value=col_name)
            cell.font = estilo_encabezado
            cell.alignment = Alignment(horizontal='center')
            cell.border = borde
        fila_actual += 1

        inicio = folio * lineas_por_folio
        fin = min(inicio + lineas_por_folio, total_lineas)
        datos_folio = df.iloc[inicio:fin]

        # Si no es el primer folio, colocar VIENEN con acumulados
        if folio > 0:
            ws.cell(row=fila_actual, column=6, value="VIENEN").font = estilo_total
            ws.cell(row=fila_actual, column=7, value=acumulado["IDP"]).number_format = '"Q"#,##0.00'
            ws.cell(row=fila_actual, column=8, value=acumulado["Tasa Municipal"]).number_format = '"Q"#,##0.00'
            ws.cell(row=fila_actual, column=9, value=acumulado["Otros Impuestos"]).number_format = '"Q"#,##0.00'
            ws.cell(row=fila_actual, column=10, value=acumulado["Total"]).number_format = '"Q"#,##0.00'
            ws.cell(row=fila_actual, column=11, value=acumulado["Compra Neta"]).number_format = '"Q"#,##0.00'
            ws.cell(row=fila_actual, column=12, value=acumulado["IVA Crédito"]).number_format = '"Q"#,##0.00'
            fila_actual += 1

        # Subtotales del folio
        subtotal = {k: 0.0 for k in acumulado.keys()}

        # Es más robusto iterar como tupla sin nombre para evitar problemas con nombres de atributos
        for fila_dato in datos_folio.itertuples(index=False, name=None):
            for col_num, valor in enumerate(fila_dato, 1):
                cell = ws.cell(row=fila_actual, column=col_num, value=valor)
                cell.border = borde
                # columnas numéricas (7..12)
                if col_num in [7, 8, 9, 10, 11, 12]:
                    cell.number_format = '"Q"#,##0.00'
            subtotal["IDP"] += float(fila_dato[idx_IDP] or 0)
            subtotal["Tasa Municipal"] += float(fila_dato[idx_Tasa] or 0)
            subtotal["Otros Impuestos"] += float(fila_dato[idx_Otros] or 0)
            subtotal["Total"] += float(fila_dato[idx_Total] or 0)
            subtotal["Compra Neta"] += float(fila_dato[idx_Compra] or 0)
            subtotal["IVA Crédito"] += float(fila_dato[idx_IVA] or 0)
            fila_actual += 1

        # Rellenar con filas en blanco hasta lineas_por_folio
        filas_datos = len(datos_folio)
        filas_a_rellenar = lineas_por_folio - filas_datos
        for _ in range(filas_a_rellenar):
            for col_num in range(1, len(columnas_exportar) + 1):
                cell = ws.cell(row=fila_actual, column=col_num, value="")
                cell.border = borde
            fila_actual += 1

        # actualizar acumulados
        for k in acumulado:
            acumulado[k] += subtotal[k]

        # Al final del folio: VAN o TOTAL GENERAL
        if folio == total_folios - 1:
            ws.cell(row=fila_actual, column=6, value="TOTAL GENERAL").font = estilo_total
        else:
            ws.cell(row=fila_actual, column=6, value="VAN").font = estilo_total

        ws.cell(row=fila_actual, column=7, value=acumulado["IDP"]).number_format = '"Q"#,##0.00'
        ws.cell(row=fila_actual, column=8, value=acumulado["Tasa Municipal"]).number_format = '"Q"#,##0.00'
        ws.cell(row=fila_actual, column=9, value=acumulado["Otros Impuestos"]).number_format = '"Q"#,##0.00'
        ws.cell(row=fila_actual, column=10, value=acumulado["Total"]).number_format = '"Q"#,##0.00'
        ws.cell(row=fila_actual, column=11, value=acumulado["Compra Neta"]).number_format = '"Q"#,##0.00'
        ws.cell(row=fila_actual, column=12, value=acumulado["IVA Crédito"]).number_format = '"Q"#,##0.00'

        # aplicar borde a la fila de VAN/TOTAL
        for col_num in range(1, len(columnas_exportar) + 1):
            ws.cell(row=fila_actual, column=col_num).border = borde
        fila_actual += 2  # espacio antes del siguiente folio

        # Resumen de documentos
    ws.cell(row=fila_actual, column=6, value="TOTAL DE FACTURAS").font = Font(bold=True)
    ws.cell(row=fila_actual, column=7, value=total_facturas).number_format = '0'
    fila_actual += 1

    ws.cell(row=fila_actual, column=6, value="NCRE - Cantidad").font = Font(bold=True)
    ws.cell(row=fila_actual, column=7, value=cantidad_ncre).number_format = '0'
    fila_actual += 1

    ws.cell(row=fila_actual, column=6, value="NCRE - Valor Neto").font = Font(bold=True)
    ws.cell(row=fila_actual, column=7, value=total_ncre).number_format = '"Q"#,##0.00'
    fila_actual += 1

    ws.cell(row=fila_actual, column=6, value="OTROS DOCUMENTOS").font = Font(bold=True)
    ws.cell(row=fila_actual, column=7, value=total_otros_docs).number_format = '0'
    fila_actual += 1

    ws.cell(row=fila_actual, column=6, value="TIPOS OTROS").font = Font(bold=True)
    ws.cell(row=fila_actual, column=7, value=otros_tipos_texto)
    fila_actual += 2

    # Resumen final al final del workbook
    ws.cell(row=fila_actual, column=6, value="Resumen Final").font = Font(bold=True)
    fila_actual += 1
    ws.cell(row=fila_actual, column=6, value="Clasificación").font = Font(bold=True)
    ws.cell(row=fila_actual, column=7, value="Compra Neta").font = Font(bold=True)
    fila_actual += 1
    for idx, row in resumen_compra.iterrows():
        ws.cell(row=fila_actual, column=6, value=row["Clasificación"])
        ws.cell(row=fila_actual, column=7, value=row["Compra Neta"]).number_format = '"Q"#,##0.00'
        fila_actual += 1


    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    nombre_archivo = f"Libro_Compras_{empresa_data['nombre']}_{mes}_{anio}.xlsx"
    st.success("Libro generado con éxito.")
    st.download_button(
        label="📥 Descargar libro de compras",
        data=output,
        file_name=nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

app()
