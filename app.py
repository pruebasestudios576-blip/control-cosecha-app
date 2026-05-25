import streamlit as st
import pandas as pd
from datetime import datetime
import os

# Configuración visual de la App para pantallas de Tablet
st.set_page_config(page_title="Empresas Polar - Cosecha", layout="centered")

# Estilos personalizados para mejorar la visibilidad en tablet
st.markdown("""
    <style>
    .stButton>button { width: 100%; font-weight: bold; background-color: #0056b3; color: white; }
    .stNumberInput input { font-size: 18px; }
    </style>
""", unsafe_allow_html=True)

st.title("🌾 Registro de Cosecha (Tablet)")
st.write("Control de Calidad Sostenible")

# --- SECCIÓN DE IDENTIFICACIÓN ---
st.subheader("Datos del Lote")
c1, c2, c3 = st.columns(3)
with c1:
    fecha = st.date_input("Fecha", datetime.now())
with c2:
    lote = st.text_input("Número de Lote", placeholder="Ej. 1122334455")
with c3:
    turno = st.selectbox("Turno", ["I", "II", "III"])

# --- SECCIÓN DE PARÁMETROS (Réplica de tu planilla Excel) ---
st.subheader("Resultados de Laboratorio")
col_izq, col_der = st.columns(2)

with col_izq:
    h = st.number_input("HUMEDAD (<= 24.00%)", min_value=0.0, max_value=100.0, value=20.0, step=0.1)
    imp = st.number_input("IMPUREZAS (<= 5.00%)", min_value=0.0, value=4.0, step=0.1)
    g_germen = st.number_input("GRANOS GERMEN DAÑADO (<= 11.00%)", min_value=0.0, value=9.0)
    g_calor = st.number_input("GRANO DAÑADO POR CALOR (<= 3.00%)", min_value=0.0, value=2.0)
    g_infestados = st.number_input("GRANOS INFESTADOS (<= 11.00%)", min_value=0.0, value=9.0)
    g_infectados = st.number_input("GRANOS INFECTADOS (<= 11.00%)", min_value=0.0, value=9.0)

with col_der:
    g_totales = st.number_input("GRANOS DAÑADOS TOTALES (<= 11.00%)", min_value=0.0, value=9.0)
    g_partidos = st.number_input("GRANOS PARTIDOS (<= 7.00%)", min_value=0.0, value=6.0)
    p_esp = st.number_input("PESO ESPECIFICO (0.715 - 0.800 Kg/Lt)", min_value=0.0, value=0.72, step=0.01)
    aflatoxinas = st.number_input("AFLATOXINAS TOTALES (<= 20 ppb)", min_value=0.0, value=15.0)
    fumonisina = st.number_input("FUMONISINA (<= 4.00 ppm)", min_value=0.0, value=1.0)

st.write("---")

# --- BOTONES DE ACCIÓN ---
if st.button("💾 REGISTRAR EN EL HISTÓRICO"):
    if not lote:
        st.error("Por favor, ingresa el número de lote antes de guardar.")
    else:
        archivo = "cosecha_historico.xlsx"
        nuevos_datos = {
            "Fecha": [fecha.strftime("%d/%m/%Y")], "Lote": [lote], "Turno": [turno],
            "Humedad": [h], "Impurezas": [imp], "G. Germen": [g_germen], "G. Calor": [g_calor],
            "Infestados": [g_infestados], "Infectados": [g_infectados], "Dañados Totales": [g_totales],
            "Partidos": [g_partidos], "Peso Esp": [p_esp], "Aflatoxinas": [aflatoxinas], "Fumonisina": [fumonisina]
        }
        df_nuevo = pd.DataFrame(nuevos_datos)
        
        # Guardar o añadir al archivo de Excel
        if not os.path.isfile(archivo):
            df_nuevo.to_excel(archivo, index=False)
        else:
            with pd.ExcelWriter(archivo, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
                # Encontrar la última fila de forma dinámica
                start_row = writer.sheets['Sheet1'].max_row
                df_nuevo.to_excel(writer, index=False, header=False, startrow=start_row)
        
        st.success(f"¡Lote {lote} registrado exitosamente en el Excel virtual!")

# Generador de plantilla para WhatsApp
if st.button("📋 GENERAR REPORTE WHATSAPP"):
    reporte = f"""*EMPRESAS POLAR*
*Lote:* {lote}
*Fecha:* {fecha.strftime('%d/%m/%Y')}
*Turno:* {turno}
----------------------------
🌾 *H:* {h:.2f}%
🌾 *Imp:* {imp:.2f}%
🧪 *Aflatoxina:* {aflatoxinas} ppb
🧪 *Fumonisina:* {fumonisina} ppm"""
    st.text_area("Presiona sostenido en la tablet para copiar:", value=reporte, height=150)
