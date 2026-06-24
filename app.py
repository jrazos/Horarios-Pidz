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
        cursos_finalizados = len(df[df['Progreso'].astype(str).str.strip().isin(['Finalizado', 'finalizado', '100%'])])
        
        if total_cursos > 0:
            porcentaje_avance = (cursos_finalizados / total_cursos) * 100
        else:
            porcentaje_avance = 0.0

        st.subheader("📈 Estado de Capacitación de la Sucursal")
        
        # Definimos color y mensaje para la web y el excel
        color_hex = ""
        mensaje_web = ""
        
        if porcentaje_avance >= 85:
            color_hex = "00cc66" # Verde
            st.success(f"¡Excelente ritmo! Avance: **{porcentaje_avance:.1f}%**")
        elif 50 <= porcentaje_avance < 85:
            color_hex = "ffcc00" # Amarillo
            st.warning(f"Buen esfuerzo. Avance: **{porcentaje_avance:.1f}%**")
        else:
            color_hex = "ff4d4d" # Rojo
            st.error(f"Atención requerida. Avance: **{porcentaje_avance:.1f}%**")
        
        st.divider()

        # --- GENERACIÓN DE LA LÓGICA DE HORARIOS ---
        df['Nombre Completo'] = df['Nombre(s)'].astype(str) + ' ' + df['Apellido(s)'].astype(str)
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
            st.success("🎉 ¡Felicidades! Todo el personal está al 100%.")
        else:
            indice_colaborador = 0
            tz_mx = pytz.timezone('America/Mexico_City')
            hoy = datetime.now(tz_mx)
            fecha_actual = hoy + timedelta(days=1)
            
            dias_semana = {0: 'lunes', 1: 'martes', 2: 'miércoles', 3: 'jueves', 4: 'viernes', 5: 'sábado', 6: 'domingo'}
            meses = {1: 'enero', 2: 'febrero', 3: 'marzo', 4: 'abril', 5: 'mayo', 6: 'junio', 7: 'julio', 8: 'agosto', 9: 'septiembre', 10: 'octubre', 11: 'noviembre', 12: 'diciembre'}
            
            horarios = []
            dias_programados = 0
            barra_progreso = st.progress(0, text="Generando horarios...")
            
            while dias_programados < 7:
                hora_actual = fecha_actual.replace(hour=8, minute=0, second=0, microsecond=0)
                hora_limite = fecha_actual.replace(hour=20, minute=0, second=0, microsecond=0)
                inicio_bloqueo = fecha_actual.replace(hour=13, minute=30, second=0, microsecond=0)
                fin_bloqueo = fecha_actual.replace(hour=14, minute=30, second=0, microsecond=0)
                fecha_texto = f"{dias_semana[fecha_actual.weekday()]} {fecha_actual.day} de {meses[fecha_actual.month]}"
                
                while hora_actual < hora_limite:
                    hora_fin = hora_actual + timedelta(minutes=30)
                    if hora_actual >= inicio_bloqueo and hora_actual < fin_bloqueo:
                        horarios.append({'Fecha y Horario': f"{fecha_texto} de {inicio_bloqueo.strftime('%H:%M')} a {fin_bloqueo.strftime('%H:%M')}", 'Colaborador': '⚠️ RECESO / OPERACIÓN', 'Puesto': '---', 'Cursos a avanzar': 'Espacio reservado para operación', 'Pendientes': 0})
                        hora_actual = fin_bloqueo
                        continue 
                    
                    colaborador = cola_colaboradores[indice_colaborador]
                    horarios.append({
                        'Fecha y Horario': f"{fecha_texto} de {hora_actual.strftime('%H:%M')} a {hora_fin.strftime('%H:%M')}",
                        'Colaborador': colaborador['Nombre Completo'], 'Puesto': colaborador['Puesto'],
                        'Cursos a avanzar': colaborador['Nombres_Cursos'], 'Pendientes': colaborador['Total_Pendientes']
                    })
                    hora_actual = hora_fin
                    indice_colaborador = (indice_colaborador + 1) % total_colaboradores 
                fecha_actual += timedelta(days=1)
                dias_programados += 1
                barra_progreso.progress(int((dias_programados/7)*100))

            df_horarios = pd.DataFrame(horarios)
            
            # Estilos de filas
            def aplicar_estilos(row):
                dia = row['Fecha y Horario'].split()[0]
                colores_dia = {'lunes': 'background-color: #ffcccc', 'martes': 'background-color: #ccffcc', 'miércoles': 'background-color: #ffffcc', 'jueves': 'background-color: #ffe6cc', 'viernes': 'background-color: #cce5ff', 'sábado': 'background-color: #e6ccff', 'domingo': 'background-color: #e6e6e6'}
                estilos = [colores_dia.get(dia, '')] * len(row)
                if row['Colaborador'] == '⚠️ RECESO / OPERACIÓN':
                    estilos = ['background-color: #d9d9d9'] * len(row)
                else:
                    idx_pend = row.index.get_loc('Pendientes')
                    pendientes = int(row['Pendientes'])
                    if pendientes >= 10: estilos[idx_pend] = 'background-color: #ff4d4d; color: white; font-weight: bold;' 
                    elif 4 <= pendientes <= 9: estilos[idx_pend] = 'background-color: #ffcc00; color: black; font-weight: bold;'
                    elif 1 <= pendientes <= 3: estilos[idx_pend] = 'background-color: #00cc66; color: white; font-weight: bold;'
                return estilos

            df_estilizado = df_horarios.style.apply(aplicar_estilos, axis=1)

            # --- 🎨 EXCEL CON TÍTULO DE AVANCE ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Escribimos los datos empezando en la fila 4 (dejamos 3 libres arriba)
                df_estilizado.to_excel(writer, index=False, sheet_name='Horarios', startrow=3)
                
                workbook = writer.book
                worksheet = writer.sheets['Horarios']
                
                # 1. 🏆 CREAR EL TÍTULO DE AVANCE PRINCIPAL
                worksheet.merge_cells('A1:E2') # Combinar las celdas de arriba
                celda_titulo = worksheet['A1']
                celda_titulo.value = f"REPORTE SEMANAL DE CAPACITACIÓN | AVANCE DE SUCURSAL: {porcentaje_avance:.1f}%"
                
                # Estilo del Título Principal
                celda_titulo.font = Font(color="FFFFFF" if porcentaje_avance < 50 or porcentaje_avance >= 85 else "000000", bold=True, size=16)
                celda_titulo.fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
                celda_titulo.alignment = Alignment(horizontal="center", vertical="center")

                # 2. DISEÑO DE ENCABEZADOS DE TABLA (Fila 4)
                header_fill = PatternFill(start_color="203764", end_color="203764", fill_type="solid")
                header_font = Font(color="FFFFFF", bold=True, size=12)
                borde = Border(left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'), top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF'))

                for cell in worksheet[4]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = borde

                # 3. FORMATO GENERAL Y ALTO DE FILAS
                for row_idx in range(1, worksheet.max_row + 1):
                    worksheet.row_dimensions[row_idx].height = 30
                    if row_idx >= 5: # Celdas de datos
                        for cell in worksheet[row_idx]:
                            cell.border = borde
                            cell.alignment = Alignment(vertical="center", wrap_text=True if cell.column == 4 else False)

                # Ancho de columnas
                worksheet.column_dimensions['A'].width = 30
                worksheet.column_dimensions['B'].width = 35
                worksheet.column_dimensions['C'].width = 25
                worksheet.column_dimensions['D'].width = 45
                worksheet.column_dimensions['E'].width = 15

            st.success("✅ ¡Horario con Título de Avance generado!")
            st.download_button(label="📥 Descargar Horario con Avance", data=output.getvalue(), file_name="Horarios_Zorro_Con_Avance.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"❌ Error técnico: {e}")