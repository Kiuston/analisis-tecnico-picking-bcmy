import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl.utils import column_index_from_string

st.set_page_config(page_title="Análisis de cajas completas y picking por técnico", layout="wide")
st.title("📦 Análisis de cajas completas y picking por técnico")

uploaded_file = st.file_uploader("📂 Subí tu archivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="LASER", header=16)
        st.success("✅ Hoja 'LASER' encontrada correctamente.")
    except Exception as e:
        st.error("❌ No se pudo leer la hoja 'LASER'.")
        st.stop()

    codigos_ref = df[["CODIGO ADMIN", "Cajas"]].dropna()
    codigos_ref.columns = ["Codigo", "Unidades por Caja"]
    codigos_ref["Codigo"] = codigos_ref["Codigo"].astype(str).str.strip().str.upper()
    codigos_ref = codigos_ref.drop_duplicates(subset="Codigo")

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

        idx_buenas = [column_index_from_string(c) - 1 for c in cols_buenas if column_index_from_string(c) - 1 < df.shape[1]]
        idx_malas = [column_index_from_string(c) - 1 for c in cols_malas if column_index_from_string(c) - 1 < df.shape[1]]

        if not idx_buenas:
            continue

        codigos = df.iloc[:, 0].astype(str).str.strip().str.upper()
        buenas = pd.to_numeric(df.iloc[:, idx_buenas].fillna(0).sum(axis=1), errors="coerce").fillna(0)
        malas = pd.to_numeric(df.iloc[:, idx_malas].fillna(0).sum(axis=1), errors="coerce").fillna(0) if idx_malas else pd.Series([0]*len(df))

        data = pd.DataFrame({
            "Codigo": codigos,
            "Unidades Buenas": buenas,
            "Unidades Defectuosas": malas
        })

        data = data.merge(codigos_ref, how="left", on="Codigo")
        data = data[data["Unidades por Caja"].notnull()]
        data = data[data["Unidades Buenas"] > 0]
        data["Cajas Completas"] = (data["Unidades Buenas"] // data["Unidades por Caja"]).astype(int)
        data["Unidades Sobrantes"] = (data["Unidades Buenas"] % data["Unidades por Caja"]).astype(int)
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
            st.metric("Cajas Completas", int(df_tecnico["Cajas Completas"].sum()))
            st.metric("Unidades Defectuosas", int(df_tecnico["Unidades Defectuosas"].sum()))
            st.metric("Total General", int(total_buenas + df_tecnico["Unidades Defectuosas"].sum()))
