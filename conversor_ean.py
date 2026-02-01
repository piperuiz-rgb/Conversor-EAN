import streamlit as st
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="RGB - Conversor Gextia", page_icon="🔄")

# Estilos minimalistas para evitar carga innecesaria
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { width: 100%; border-radius: 0px; height: 3em; background-color: #000; color: #fff; }
    </style>
    """, unsafe_allow_html=True)

st.title("🔄 Conversor Gextia a EAN")
st.info("Sube tu catálogo y el informe de ventas para obtener un Excel limpio con EANs.")

# 1. CARGA DEL CATÁLOGO
cat_file = st.file_uploader("1. Sube tu CATALOGUE.xlsx", type=['xlsx'])

@st.cache_data
def process_catalog(file):
    df = pd.read_excel(file, engine='openpyxl')
    df.columns = [str(c).strip() for c in df.columns]
    # Creamos la llave maestra: REF_COLOR_TALLA
    df['KEY'] = (df['Referencia'].astype(str).str.strip().str.upper() + "_" + 
                 df['Color'].astype(str).str.strip().str.upper() + "_" + 
                 df['Talla'].astype(str).str.strip().str.upper())
    return df[['KEY', 'EAN']]

if cat_file:
    df_cat = process_catalog(cat_file)
    st.success("✅ Catálogo cargado correctamente.")

    # 2. CARGA DEL INFORME SUCIO
    st.write("---")
    ventas_file = st.file_uploader("2. Sube el informe de VENTAS (Gextia)", type=['xlsx'])

    if ventas_file:
        df_v = pd.read_excel(ventas_file)
        cols = df_v.columns.tolist()
        
        c1, c2 = st.columns(2)
        col_txt = c1.selectbox("Columna con [REF]... (COL, TAL)", cols)
        col_cant = c2.selectbox("Columna Cantidad", cols)

        if st.button("GENERAR EXCEL LIMPIO"):
            with st.spinner("Procesando..."):
                # Función de extracción optimizada
                def parse_gextia(text):
                    text = str(text)
                    r = re.search(r'\[(.*?)\]', text)
                    s = re.findall(r'\((.*?)\)', text)
                    if r and s:
                        ref = r.group(1).strip().upper()
                        parts = s[-1].split(',')
                        if len(parts) >= 2:
                            return f"{ref}_{parts[0].strip().upper()}_{parts[1].strip().upper()}"
                    return None

                # Creamos la columna de cruce en el excel de ventas
                df_v['JOIN_KEY'] = df_v[col_txt].apply(parse_gextia)

                # Cruzamos con el catálogo (MERGE es mucho más rápido que un bucle for)
                df_final = pd.merge(df_v, df_cat, left_on='JOIN_KEY', right_on='KEY', how='inner')

                if not df_final.empty:
                    # Resultado final: EAN y Cantidad
                    resultado = df_final[['EAN', col_cant]].rename(columns={col_cant: 'Cantidad'})
                    
                    st.success(f"📦 Se han convertido {len(resultado)} líneas con éxito.")
                    st.dataframe(resultado.head(10), use_container_width=True)

                    # Preparar descarga
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        resultado.to_excel(writer, index=False)
                    
                    st.download_button(
                        label="📥 DESCARGAR EXCEL PARA PETICIONES",
                        data=output.getvalue(),
                        file_name="ean_limpios_para_peticiones.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("No se encontraron coincidencias. Verifica que el catálogo y el informe tengan las mismas Referencias/Colores/Tallas.")
                  
