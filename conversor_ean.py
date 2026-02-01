import streamlit as st
import pandas as pd
import re
import io
import os

# Deshabilitamos elementos visuales innecesarios para ahorrar ancho de banda
st.set_page_config(page_title="Conversor Fast", layout="centered")

@st.cache_data(show_spinner=False)
def load_db():
    if os.path.exists("catalogue.xlsx"):
        df = pd.read_excel("catalogue.xlsx", engine='openpyxl')
        # Forzamos limpieza de nombres de columnas
        df.columns = [str(c).strip() for c in df.columns]
        # Crear KEY única
        df['KEY'] = (df['Referencia'].astype(str).str.strip().str.upper() + "_" + 
                     df['Color'].astype(str).str.strip().str.upper() + "_" + 
                     df['Talla'].astype(str).str.strip().str.upper())
        # Mantenemos solo lo mínimo necesario en memoria
        return df[['KEY', 'EAN']].set_index('KEY')
    return None

db = load_db()

st.title("🔄 Conversor Gextia")

if db is not None:
    # El cargador de archivos es el punto crítico del AxiosError
    archivo = st.file_uploader("Sube el Excel sucio", type=['xlsx'], key="f_up")

    if archivo:
        # Botón para disparar la lógica sin previsualizar nada antes
        if st.button("PROCESAR Y DESCARGAR"):
            try:
                # Leemos el archivo sucio
                df_v = pd.read_excel(archivo)
                
                # Identificamos columnas por posición (0=EAN sucio, 1=Cantidad)
                # Esto evita errores si los nombres de columna varían
                col_txt = df_v.columns[0]
                col_can = df_v.columns[1]

                def limpiar(t):
                    t = str(t)
                    r = re.search(r'\[(.*?)\]', t)
                    s = re.findall(r'\((.*?)\)', t)
                    if r and s:
                        p = s[-1].split(',')
                        if len(p) >= 2:
                            return f"{r.group(1).strip().upper()}_{p[0].strip().upper()}_{p[1].strip().upper()}"
                    return None

                # Creamos la columna de cruce
                df_v['JOIN'] = df_v[col_txt].apply(limpiar)
                
                # Cruce ultra-rápido usando el índice del catálogo
                # Esto es lo más eficiente en memoria/red
                df_v = df_v.join(db, on='JOIN', how='inner', rsuffix='_REAL')

                if not df_v.empty:
                    # Preparamos el resultado final
                    res = df_v[['EAN_REAL', col_can]].rename(columns={'EAN_REAL': 'EAN', col_can: 'Cantidad'})
                    
                    # Generamos el Excel en memoria
                    out = io.BytesIO()
                    with pd.ExcelWriter(out, engine='openpyxl') as writer:
                        res.to_excel(writer, index=False)
                    
                    st.success("✅ Conversión completada.")
                    st.download_button("📥 DESCARGAR AHORA", out.getvalue(), "ean_limpios.xlsx", use_container_width=True)
                else:
                    st.warning("⚠️ No se encontraron coincidencias. Revisa el catálogo.")
            
            except Exception as e:
                st.error(f"Error en proceso: {e}")
else:
    st.error("❌ No se encontró catalogue.xlsx")
    
