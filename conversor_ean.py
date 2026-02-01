import streamlit as st
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="Conversor RGB - Gextia", layout="centered")

@st.cache_data(show_spinner=False)
def load_cat():
    if os.path.exists("catalogue.xlsx"):
        df = pd.read_excel("catalogue.xlsx", engine='openpyxl')
        df.columns = [str(c).strip() for c in df.columns]
        # Creamos la llave maestra uniendo Ref, Color y Talla del catálogo
        df['KEY_MASTER'] = (df['Referencia'].astype(str).str.strip().str.upper() + "_" + 
                            df['Color'].astype(str).str.strip().str.upper() + "_" + 
                            df['Talla'].astype(str).str.strip().str.upper())
        # Nos quedamos con la llave y el EAN REAL
        return df[['KEY_MASTER', 'EAN']]
    return None

df_cat = load_cat()

st.title("🔄 Conversor Gextia")

if df_cat is not None:
    st.success(f"✅ Catálogo listo: {len(df_cat)} referencias.")
    
    archivo_v = st.file_uploader("Sube el Excel de Gextia (2 columnas: EAN y Cantidad)", type=['xlsx'])

    if archivo_v:
        df_v = pd.read_excel(archivo_v)
        
        # Forzamos a que use las columnas que me has dicho
        # Si no se llaman exactamente así, el código las busca por posición
        col_sucio = "EAN" if "EAN" in df_v.columns else df_v.columns[0]
        col_cant = "Cantidad" if "Cantidad" in df_v.columns else df_v.columns[1]

        if st.button("LIMPIAR Y CONVERTIR"):
            def procesar_texto_sucio(t):
                t = str(t)
                # Extraer [REFERENCIA]
                ref = re.search(r'\[(.*?)\]', t)
                # Extraer (COLOR, TALLA)
                specs = re.findall(r'\((.*?)\)', t)
                
                if ref and specs:
                    r = ref.group(1).strip().upper()
                    partes = specs[-1].split(',')
                    if len(partes) >= 2:
                        c = partes[0].strip().upper()
                        talla = partes[1].strip().upper()
                        return f"{r}_{c}_{talla}"
                return None

            # Aplicar limpieza rápida
            df_v['LLAVE_CRUCE'] = df_v[col_sucio].apply(procesar_texto_sucio)
            
            # Cruzar con el catálogo para obtener el EAN de verdad
            resultado = pd.merge(df_v, df_cat, left_on='LLAVE_CRUCE', right_on='KEY_MASTER', how='inner')

            if not resultado.empty:
                # El resultado final tendrá el EAN real del catálogo y la cantidad
                # Renombramos el EAN del catálogo para que no choque con el nombre de la columna sucia
                df_final = resultado[['EAN_y', col_cant]].rename(columns={'EAN_y': 'EAN', col_cant: 'Cantidad'})
                
                st.success(f"✅ Procesadas {len(df_final)} líneas con éxito.")
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False)
                
                st.download_button("📥 DESCARGAR EXCEL PARA PETICIONES", output.getvalue(), "ean_limpios.xlsx")
            else:
                st.error("No se encontraron coincidencias. Revisa si los nombres de Color/Talla en el catálogo coinciden con el texto del paréntesis.")
else:
    st.error("⚠️ No encuentro 'catalogue.xlsx' en la carpeta de la App.")
    
