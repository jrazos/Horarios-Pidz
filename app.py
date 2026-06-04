import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import pytz
import io
# Importamos las herramientas de diseño nativas de Excel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="Generador de Horarios - Zorro", page_icon="🦊", layout="centered")

st.title("🦊 Generador de Horarios de Capacitación")
st.write("Sube el reporte de Excel para organizar automáticamente a tu equipo esta semana.")

archivo_subido = st.file_uploader("📂 Sube tu archivo aquí (cualquier nombre funciona)", type=['xlsx'])

if archivo_subido is not None:
    try:
        df = pd.read_excel(archivo_subido)
        df['Nombre Completo'] = df['Nombre(s)'].astype(str) + ' ' + df['Apellido(s)'].astype(str)
        df_pendientes = df[df['Progreso'].isin(['No iniciado', 'En proceso', 'No Inciado', '0%'])]
        
        resumen = df_pendientes.groupby(['Nombre Completo', 'Puesto']).agg(
            Total_Pendientes=('Nombre curso', 'count'),
            Nombres_Cursos=('Nombre curso', lambda x: ', '.join(x.astype(str)))
        ).reset_index()

        resumen = resumen.sort_values(by='Total_Pendientes', ascending=False)
        resumen['Turno_Area'] = resumen.groupby('Puesto').cumcount()
        resumen = resumen.sort_values(by=['Turno_Area', 'Total_Pendientes'], ascending=[True, False])

        cola_colaboradores = resumen.to_dict('records')
        total_colaboradores = len(cola_colaboradores)

        if total_colaboradores == 0:
            st.success("✅ Todo el personal está al 100% en sus cursos.")
        else:
            indice_colaborador = 0
            tz_mx = pytz.timezone('America/Mexico_City')
            hoy = datetime.now(tz_mx)
            fecha_actual = hoy + timedelta(days=1)
            
            dias_semana = {0: 'lunes', 1: 'martes', 2: 'miércoles', 3: 'jueves', 4: 'viernes', 5: 'sábado', 6: '
