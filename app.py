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
        
        # --- 📊 CÁLCULO DEL PORCENTAJE DE AVANCE DE LA SUCURSAL ---
        total_cursos = len(df)
        # Contamos cuántos están finalizados (limpiando espacios por si acaso)
        cursos_finalizados = len(df[df['Progreso'].astype(str).str.strip().isin(['Finalizado', 'finalizado', '100%'])])
        
        if total_cursos > 0:
            porcentaje_avance = (cursos_finalizados / total_cursos) * 100
        else:
            porcentaje_avance = 0.0

        # Mostrar métrica visual en la página web
        st.subheader("📈 Estado de Capacitación de la Sucursal")
        
        # Color del indicador dependiendo del avance
        if porcentaje_avance >= 85:
            st.success(f"¡Excelente ritmo! La sucursal tiene un **{porcentaje_avance:.1f}%** de avance total ({cursos_finalizados} de {total_cursos} cursos completados).")
        elif 50 <= porcentaje_avance < 85:
            st.warning(f"Buen esfuerzo. La sucursal tiene un **{porcentaje_avance:.1f}%** de avance total ({cursos_finalizados} de {total_cursos} cursos completados). ¡A seguir impulsando!")
        else:
            st.error(f"Atención requerida. La sucursal tiene un **{porcentaje_avance:.1f}%** de avance total ({cursos_finalizados} de {total_cursos} cursos completados). Este horario ayudará a subir el indicador.")
        
        st.divider() # Línea divisoria visual

        # --- CONTINUACIÓN DEL PROCESO DE HORARIOS ---
        df['Nombre Completo'] = df['Nombre(s)'].astype(str) + ' ' + df['Apellido(s)'].astype(str)
        
        # Consideramos pendientes los que están "No iniciado", "En proceso" o variantes comunes
        df_pendientes = df[df['Progreso'].astype(str).str.strip().isin(['No iniciado', 'En proceso', 'No Inciado', '0%'])]
        
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
            st.success("🎉 ¡Felicidades! Todo el personal está al 100% en sus cursos. No se requieren horarios semanales.")
        else:
            indice_colaborador = 0
            tz_mx = pytz.timezone('America/Mexico_City')
            hoy = datetime.now(tz_mx)
            fecha_actual = hoy + timedelta(days=1)
            
            dias_semana = {0: 'lunes', 1: 'martes', 2: 'miércoles', 3: 'jueves', 4: 'viernes', 5: 'sábado', 6: 'domingo'}
            meses = {1: 'enero', 2: 'febrero', 3: 'marzo', 4: 'abril', 5: 'mayo', 6: 'junio', 7: 'julio', 8: 'agosto', 9: 'septiembre', 10: 'octubre', 11: 'noviembre', 12: 'diciembre'}
            
            horarios = []
            dias_programados = 0
            
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
                barra_progreso.progress(int((dias_programados/7)*100), text="Aplicando diseño y formatos...")

            df_horarios = pd.DataFrame(horarios)
            
            def aplicar_estilos(row):
                dia = row['Fecha y Horario'].split()[0]
                colores_dia = {
                    'lunes': 'background-color: #ffcccc', 'martes': 'background-color: #ccffcc', 
                    'miércoles': 'background-color: #ffffcc', 'jueves': 'background-color: #ffe6cc', 
                    'viernes': 'background-color: #cce5ff', 'sábado': 'background-color: #e6ccff', 
                    'domingo': 'background-color: #e6e6e6'
                }
                estilos = [colores_dia.get(dia, '')] * len(row)
                
                if row['Colaborador'] == '⚠️ RECESO / OPERACIÓN':
                    estilos = ['background-color: #d9d9d9'] * len(row)
                else:
                    idx_pend = row.index.get_loc('Pendientes')
                    pendientes = int(row['Pendientes'])
                    if pendientes >= 10:
                        estilos[idx_pend] = 'background-color: #ff4d4d; color: white; font-weight: bold;' 
                    elif 4 <= pendientes <= 9:
                        estilos[idx_pend] = 'background-color: #ffcc00; color: black; font-weight: bold;'
                    elif 1 <= pendientes <= 3:
                        estilos[idx_pend] = 'background-color: #00cc66; color: white; font-weight: bold;'
                return estilos
            
            df_estilizado = df_horarios.style.apply(aplicar_estilos, axis=1)
            
            # --- 🎨 PROCESO DE DISEÑO PROFESIONAL EN EXCEL ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_estilizado.to_excel(writer, index=False, sheet_name='Horarios')
                workbook = writer.book
                worksheet = writer.sheets['Horarios']
                
                header_fill = PatternFill(start_color="203764", end_color="203764", fill_type="solid")
                header_font = Font(color="FFFFFF", bold=True, size=12)
                borde_delgado = Border(left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'), 
                                       top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF'))
                
                # 1. Aplicar diseño a los títulos
                for cell in worksheet[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = borde_delgado

                # 2. Aplicar bordes y alineación al resto de la tabla
                for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
                    for cell in row:
                        cell.border = borde_delgado
                        if cell.column == 4:
                            cell.alignment = Alignment(vertical="center", wrap_text=True)
                        elif cell.column == 5:
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                        else:
                            cell.alignment = Alignment(vertical="center")

                # 3. Ancho de columnas
                worksheet.column_dimensions['A'].width = 30  
                worksheet.column_dimensions['B'].width = 35  
                worksheet.column_dimensions['C'].width = 25  
                worksheet.column_dimensions['D'].width = 45  
                worksheet.column_dimensions['E'].width = 15  

                # 4. Alto de filas a 30 puntos
                for row_idx in range(1, worksheet.max_row + 1):
                    worksheet.row_dimensions[row_idx].height = 30

            st.success("✅ ¡Formato Premium generado con éxito!")
            st.download_button(
                label="📥 Descargar Horario Premium",
                data=output.getvalue(),
                file_name="Horarios_Semanales_Premium.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"❌ Hubo un problema con el archivo. Asegúrate de que la columna se llame 'Progreso' y tenga los estatus correctos. Detalle técnico: {e}")