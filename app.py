import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import pytz
import io

# Configuración de la página web
st.set_page_config(page_title="Generador de Horarios - Zorro", page_icon="🦊", layout="centered")

st.title("🦊 Generador de Horarios de Capacitación")
st.write("Sube el reporte de Excel para organizar automáticamente a tu equipo esta semana.")

# Botón para subir archivo
archivo_subido = st.file_uploader("📂 Sube tu archivo aquí (cualquier nombre funciona)", type=['xlsx'])

if archivo_subido is not None:
    try:
        # 1. Cargar datos
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
            st.success("✅ Todo el personal está al 100% en sus cursos. No se requieren horarios.")
        else:
            indice_colaborador = 0
            tz_mx = pytz.timezone('America/Mexico_City')
            hoy = datetime.now(tz_mx)
            fecha_actual = hoy + timedelta(days=1)
            
            dias_semana = {0: 'lunes', 1: 'martes', 2: 'miércoles', 3: 'jueves', 4: 'viernes', 5: 'sábado', 6: 'domingo'}
            meses = {1: 'enero', 2: 'febrero', 3: 'marzo', 4: 'abril', 5: 'mayo', 6: 'junio', 7: 'julio', 8: 'agosto', 9: 'septiembre', 10: 'octubre', 11: 'noviembre', 12: 'diciembre'}
            
            horarios = []
            dias_programados = 0
            
            # Barra de carga visual en la web
            barra_progreso = st.progress(0, text="Generando la mejor combinación de horarios...")
            
            while dias_programados < 7:
                hora_actual = fecha_actual.replace(hour=8, minute=0, second=0, microsecond=0)
                hora_limite = fecha_actual.replace(hour=20, minute=0, second=0, microsecond=0)
                inicio_bloqueo = fecha_actual.replace(hour=13, minute=30, second=0, microsecond=0)
                fin_bloqueo = fecha_actual.replace(hour=14, minute=30, second=0, microsecond=0)
                fecha_texto = f"{dias_semana[fecha_actual.weekday()]} {fecha_actual.day} de {meses[fecha_actual.month]}"
                
                while hora_actual < hora_limite:
                    hora_fin = hora_actual + timedelta(minutes=30)
                    if hora_actual >= inicio_bloqueo and hora_actual < fin_bloqueo:
                        horarios.append({
                            'Fecha y Horario': f"{fecha_texto} de {inicio_bloqueo.strftime('%H:%M')} a {fin_bloqueo.strftime('%H:%M')}",
                            'Colaborador': '⚠️ RECESO / OPERACIÓN',
                            'Puesto': '---',
                            'Cursos a avanzar': 'Espacio reservado para operación',
                            'Pendientes': 0
                        })
                        hora_actual = fin_bloqueo
                        continue 
                    
                    colaborador = cola_colaboradores[indice_colaborador]
                    horarios.append({
                        'Fecha y Horario': f"{fecha_texto} de {hora_actual.strftime('%H:%M')} a {hora_fin.strftime('%H:%M')}",
                        'Colaborador': colaborador['Nombre Completo'],
                        'Puesto': colaborador['Puesto'],
                        'Cursos a avanzar': colaborador['Nombres_Cursos'],
                        'Pendientes': colaborador['Total_Pendientes']
                    })
                    hora_actual = hora_fin
                    indice_colaborador = (indice_colaborador + 1) % total_colaboradores 
                    
                fecha_actual += timedelta(days=1)
                dias_programados += 1
                barra_progreso.progress(int((dias_programados/7)*100), text="Generando la mejor combinación de horarios...")

            df_horarios = pd.DataFrame(horarios)
            
            # Función para aplicar estilos combinados (Día + Semáforo)
            def aplicar_estilos(row):
                dia = row['Fecha y Horario'].split()[0]
                colores_dia = {
                    'lunes': 'background-color: #ffcccc', 
                    'martes': 'background-color: #ccffcc', 
                    'miércoles': 'background-color: #ffffcc', 
                    'jueves': 'background-color: #ffe6cc', 
                    'viernes': 'background-color: #cce5ff', 
                    'sábado': 'background-color: #e6ccff', 
                    'domingo': 'background-color: #e6e6e6'
                }
                
                # Asignamos el color del día a toda la fila primero
                color_base = colores_dia.get(dia, '')
                estilos = [color_base] * len(row)
                
                # Reglas especiales
                if row['Colaborador'] == '⚠️ RECESO / OPERACIÓN':
                    estilos = ['background-color: #d9d9d9'] * len(row)
                else:
                    # Encontrar en qué posición exacta está la columna 'Pendientes'
                    idx_pend = row.index.get_loc('Pendientes')
                    pendientes = int(row['Pendientes'])
                    
                    # 🚥 SEMÁFORO 🚥 (Sobrescribimos solo el color de esa celda)
                    if pendientes >= 10:
                        # ROJO (con letras blancas y en negritas para que resalte)
                        estilos[idx_pend] = 'background-color: #ff4d4d; color: white; font-weight: bold;' 
                    elif 4 <= pendientes <= 9:
                        # AMARILLO (letras negras y negritas)
                        estilos[idx_pend] = 'background-color: #ffcc00; color: black; font-weight: bold;'
                    elif 1 <= pendientes <= 3:
                        # VERDE (letras blancas y negritas)
                        estilos[idx_pend] = 'background-color: #00cc66; color: white; font-weight: bold;'
                        
                return estilos
            
            df_estilizado = df_horarios.style.apply(aplicar_estilos, axis=1)
            
            # Preparar Excel para descarga en web
            output = io.BytesIO()
            df_estilizado.to_excel(output, index=False, engine='openpyxl')
            
            st.success("✅ ¡Semana y semáforo generados al 100%!")
            st.download_button(
                label="📥 Descargar Excel de Horarios",
                data=output.getvalue(),
                file_name="Horarios_Semanales_Zorro.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"❌ Hubo un problema con el archivo. Asegúrate de que tenga las 5 columnas exactas. Detalle técnico: {e}")
