import streamlit as st
import pandas as pd
import re
import io
import os

# CONFIGURACIÓN RÁPIDA
st.set_page_config(page_title="Conversor RGB", layout="centered")

# 1. CARGA DEL CATÁLOGO (Se ejecuta solo una vez al arrancar)
@st.cache_data
def load_catalog_fast():
    archivo = "catalogue.xlsx"
    if os.path.exists(archivo):
        df = pd.read_excel(archivo, engine='openpyxl')
        # Limpieza de nombres de columnas
        df.columns = [str(c).strip() for c in df.columns]
        # Crear la llave de búsqueda (REF_COL_TAL)
        df['KEY_MASTER'] = (df['Referencia'].astype(str).str.strip().str.upper() + "_" + 
                            df['Color'].astype(str).str.strip().str.upper() + "_" + 
                            df['Talla'].astype(str).str.strip().str.upper())
        # Solo nos quedamos con la llave y el EAN para que pese poco en memoria
        return df[['KEY_MASTER', 'EAN']]
    return None

df_cat = load_catalog_fast()

st.title("🔄 Conversor Gextia")

# Verificamos si el catálogo se leyó correctamente
if df_cat is not None:
    st.success(f"✅ Catálogo listo ({len(df_cat)} referencias cargadas)")
    
    st.write("---")
    st.subheader("Subir Ventas de Gextia")
    # Este es el único archivo que vas a subir tú
    archivo_sucio = st.file_uploader("Sube el Excel con descripciones largas", type=['xlsx'])

    if archivo_sucio:
        df_v = pd.read_excel(archivo_sucio)
        
        # Selección de columnas (Gextia suele llamar EAN a la descripción)
        c1, c2 = st.columns(2)
        col_txt = c1.selectbox("Columna con [REF]...", df_v.columns)
        col_cant = c2.selectbox("Columna con Cantidad", df_v.columns)

        if st.button("LIMPIAR Y CONVERTIR", type="primary"):
            # Lógica ultra-rápida para evitar que el navegador se desconecte (Axios Error)
            try:
                def extraer_key(t):
                    t = str(t)
                    ref = re.search(r'\[(.*?)\]', t)
                    specs = re.findall(r'\((.*?)\)', t)
                    if ref and specs:
                        r = ref.group(1).strip().upper()
                        partes = specs[-1].split(',')
                        if len(partes) >= 2:
                            return f"{r}_{partes[0].strip().upper()}_{partes[1].strip().upper()}"
                    return None

                # Procesamiento por bloques (vectorizado)
                df_v['JOIN_KEY'] = df_v[col_txt].apply(extraer_key)
                
                # Unir tablas (Merge)
                res = pd.merge(df_v, df_cat, left_on='JOIN_KEY', right_on='KEY_MASTER', how='inner')

                if not res.empty:
                    # Formato para la App de Peticiones
                    final = res[['EAN', col_cant]].rename(columns={col_cant: 'Cantidad'})
                    
                    st.success(f"✅ ¡Hecho! {len(final)} líneas identificadas.")
                    
                    # Preparar el botón de descarga
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        final.to_excel(writer, index=False)
                    
                    st.download_button(
                        label="📥 BAJAR EXCEL PARA PETICIONES",
                        data=output.getvalue(),
                        file_name="ean_listos.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("No se encontró ninguna coincidencia. Revisa si el catálogo tiene las mismas tallas/colores.")
            except Exception as e:
                st.error(f"Error técnico: {e}")
else:
    st.error("❌ No encuentro el archivo 'catalogue.xlsx' en GitHub. Asegúrate de que el nombre sea exacto y esté en la misma carpeta.")

