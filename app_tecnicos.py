
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl.utils import column_index_from_string

st.set_page_config(page_title="Análisis técnico por operador", layout="wide")
st.title("📦 Análisis de cajas completas y picking por técnico")

# Cargar archivo de valoración
valoracion_file = st.file_uploader("📂 Subí tu archivo de valoración (.xlsx)", type=["xlsx"], key="valoracion")

# Cargar archivo de referencia
referencia_file = st.file_uploader("📂 Subí el archivo de referencia de cajas (.xlsx)", type=["xlsx"], key="referencia")

if valoracion_file and referencia_file:
    try:
        df_val = pd.read_excel(valoracion_file, sheet_name="LASER", header=16)
        st.success("✅ Hoja 'LASER' encontrada correctamente en el archivo de valoración.")
    except Exception as e:
        st.error("❌ No se pudo leer la hoja 'LASER'. Verificá el archivo.")
        st.stop()

    try:
        df_ref = pd.read_excel(referencia_file, sheet_name=0, header=16)
        df_ref = df_ref[["CODIGO ADMIN", "Cajas"]].dropna()
        df_ref.columns = ["Codigo", "Unidades por Caja"]
        df_ref["Codigo"] = df_ref["Codigo"].astype(str).str.strip()
        df_ref = df_ref.drop_duplicates(subset="Codigo")
        st.success("✅ Archivo de referencia cargado correctamente.")
    except Exception as e:
        st.error(f"Error al cargar referencia de cajas: {e}")
        st.stop()

    # Definir columnas por técnico
    tecnicos_columnas = {
        "MAX": ("I", "J", "K", "AI", "AJ", "AK", "AL", "AM"),
        "ANTONIO": ("L", "M", "N", "AN", "AO", "AP", "AQ", "AR"),
        "OSCAR L.": ("O", "P", "Q", "AS", "AT", "AU", "AV", "AW"),
        "JAVIER": ("R", "S", "T", "AX", "AY", "AZ", "BA", "BB"),
        "CARLOS": ("U", "V", "W", "BC", "BD", "BE", "BF", "BG"),
        "ANDRYS": ("X", "Y", "Z", "BH", "BI", "BJ", "BK", "BL"),
    }

    st.subheader("📋 Tabla de resumen por técnico")
    resumen_total = []

    for tecnico, columnas in tecnicos_columnas.items():
        cols_buenas = columnas[:3]
        cols_malas = columnas[3:]

        idx_buenas = [column_index_from_string(c) - 1 for c in cols_buenas if column_index_from_string(c) - 1 < df_val.shape[1]]
        idx_malas = [column_index_from_string(c) - 1 for c in cols_malas if column_index_from_string(c) - 1 < df_val.shape[1]]

        if not idx_buenas:
            continue

        codigos = df_val.iloc[:, 0].astype(str).str.strip()
        buenas = pd.to_numeric(df_val.iloc[:, idx_buenas].fillna(0).sum(axis=1), errors="coerce").fillna(0).astype(int)
        malas = pd.to_numeric(df_val.iloc[:, idx_malas].fillna(0).sum(axis=1), errors="coerce").fillna(0).astype(int) if idx_malas else pd.Series([0]*len(df_val))

        data = pd.DataFrame({
            "Codigo": codigos,
            "Unidades Buenas": buenas,
            "Unidades Defectuosas": malas
        })

        data = data.merge(df_ref, how="left", on="Codigo")
        data = data[data["Unidades Buenas"] > 0]
        data["Cajas Completas"] = (data["Unidades Buenas"] // data["Unidades por Caja"]).fillna(0).astype(int)
        data["Unidades Sobrantes"] = (data["Unidades Buenas"] % data["Unidades por Caja"]).fillna(0).astype(int)
        data["Técnico"] = tecnico

        resumen_total.append(data)

    if not resumen_total:
        st.warning("No se encontraron técnicos con datos válidos.")
        st.stop()

    final_df = pd.concat(resumen_total, ignore_index=True)
    st.dataframe(final_df, use_container_width=True)

    st.subheader("📊 Gráficos por técnico")

    for tecnico in final_df["Técnico"].unique():
        df_tecnico = final_df[final_df["Técnico"] == tecnico]
        total_buenas = df_tecnico["Unidades Buenas"].sum()
        total_picking = df_tecnico["Unidades Sobrantes"].sum()
        total_def = df_tecnico["Unidades Defectuosas"].sum()
        cajas = df_tecnico["Cajas Completas"].sum()

        st.markdown(f"#### Técnico: **{tecnico}**")
        col1, col2 = st.columns([1, 3])

        with col1:
            fig, ax = plt.subplots()
            valores = [total_buenas - total_picking, total_picking]
            etiquetas = ["Cajas Completas", "Para Picking"]
            colores = ["#4CAF50", "#FFC107"]
            wedges, _ = ax.pie(valores, labels=None, startangle=90, colors=colores)
            ax.axis("equal")
            ax.legend(wedges, etiquetas, title="Clasificación", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
            st.pyplot(fig)

        with col2:
            st.metric("Unidades Buenas", int(total_buenas))
            st.metric("Unidades a Picking", int(total_picking))
            st.metric("Cajas Completas", int(cajas))
            st.metric("Unidades Defectuosas", int(total_def))
            st.metric("Total General", int(total_buenas + total_def))
