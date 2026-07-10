import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import pytz
import io
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="Generador de Horarios - Zorro", page_icon="🦊", layout="wide")

st.title("🦊 Generador de Horarios de Capacitación - Grupo Zorro")
st.write("Sube el reporte de Excel para organizar automáticamente a tu equipo esta semana.")

archivo_subido = st.file_uploader("📂 Sube tu archivo aquí (cualquier nombre funciona)", type=['xlsx'])

if archivo_subido is not None:
    try:
        df = pd.read_excel(archivo_subido)
        
        # --- LIMPIEZA DE DATOS ---
        df['Progreso_clean'] = df['Progreso'].astype(str).str.strip().str.lower()
        df['Nombre Completo'] = df['Nombre(s)'].astype(str) + ' ' + df['Apellido(s)'].astype(str)
        df['Estatus_Original'] = df['Progreso'].astype(str).str.strip()
        
        terminos_pendientes = ['no iniciado', 'no inciado', 'en proceso', 'en progreso', '0%']
        
        # --- CÁLCULO ESTRICTO DEL PORCENTAJE DE AVANCE ---
        total_cursos = len(df)
        cursos_finalizados = len(df[df['Progreso_clean'] == 'finalizado'])
        
        filtro_pendientes = df['Progreso_clean'].isin(terminos_pendientes)
        
        if total_cursos > 0:
            porcentaje_avance = (cursos_finalizados / total_cursos) * 100
        else:
            porcentaje_avance = 0.0

        st.subheader("📈 Estado de Capacitación de la Sucursal")
        
        color_hex = ""
        if porcentaje_avance >= 85:
            color_hex = "00cc66" 
            st.success(f"¡Excelente ritmo! Avance de la Sucursal: **{porcentaje_avance:.1f}%**")
        elif 50 <= porcentaje_avance < 85:
            color_hex = "ffcc00" 
            st.warning(f"Buen esfuerzo. Avance de la Sucursal: **{porcentaje_avance:.1f}%**")
        else:
            color_hex = "ff4d4d" 
            st.error(f"Atención requerida. Avance de la Sucursal: **{porcentaje_avance:.1f}%**")
        
        st.divider()

        # --- ESTADÍSTICAS PARA EL TOP 10 ---
        stats = df.groupby('Nombre Completo').agg(
            Pendientes=('Progreso_clean', lambda x: x.isin(terminos_pendientes).sum()),
            Finalizados=('Progreso_clean', lambda x: (x == 'finalizado').sum())
        ).reset_index()
        
        top_10_peor = stats.sort_values(by='Pendientes', ascending=False).head(10)
        top_10_mejor = stats.sort_values(by=['Pendientes', 'Finalizados'], ascending=[True, False]).head(10)

        # --- GENERACIÓN DE LA LÓGICA DE HORARIOS ---
        df_pendientes = df[filtro_pendientes].copy()
        
        df_pendientes['Curso_Detalle'] = df_pendientes['Nombre curso'].astype(str) + " (" + df_pendientes['Estatus_Original'] + ")"
        
        resumen = df_pendientes.groupby(['Nombre Completo', 'Puesto']).agg(
            Total_Pendientes=('Curso_Detalle', 'count'),
            Nombres_Cursos=('Curso_Detalle', lambda x: ', '.join(x))
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
            barra_progreso = st.progress(0, text="Generando horarios semanales...")
            
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
                            'Cursos Pendientes (Detalle)': 'Espacio reservado para operación', 
                            'Total': 0
                        })
                        hora_actual = fin_bloqueo
                        continue 
                    
                    colaborador = cola_colaboradores[indice_colaborador]
                    horarios.append({
                        'Fecha y Horario': f"{fecha_texto} de {hora_actual.strftime('%H:%M')} a {hora_fin.strftime('%H:%M')}",
                        'Colaborador': colaborador['Nombre Completo'], 
                        'Puesto': colaborador['Puesto'],
                        'Cursos Pendientes (Detalle)': colaborador['Nombres_Cursos'], 
                        'Total': colaborador['Total_Pendientes']
                    })
                    hora_actual = hora_fin
                    indice_colaborador = (indice_colaborador + 1) % total_colaboradores 
                fecha_actual += timedelta(days=1)
                dias_programados += 1
                barra_progreso.progress(int((dias_programados/7)*100))

            df_horarios = pd.DataFrame(horarios)
            
            def categorizar_puesto(puesto):
                p_lower = str(puesto).lower()
                if p_lower == '---':
                    return 'Z_Recesos' 
                elif any(rol in p_lower for rol in ['gerente', 'subgerente', 'comodin', 'comodín', 'administrativa', 'administrativo']):
                    return 'A_Gerencia'
                elif 'caja' in p_lower or 'cajer' in p_lower:
                    return 'B_Cajas'
                elif 'cremer' in p_lower or 'perecedero' in p_lower:
                    return 'C_Cremería y Perecederos'
                elif 'autoservicio' in p_lower or 'surtidor' in p_lower:
                    return 'D_Autoservicio'
                else:
                    return f"E_{str(puesto).title()}"
            
            df_horarios['Categoria_Orden'] = df_horarios['Puesto'].apply(categorizar_puesto)
            df_horarios = df_horarios.sort_values(by=['Categoria_Orden', 'Colaborador', 'Fecha y Horario']).reset_index(drop=True)
            df_horarios = df_horarios.drop(columns=['Categoria_Orden'])

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
                    idx_pend = row.index.get_loc('Total')
                    pendientes = int(row['Total'])
                    if pendientes >= 10: estilos[idx_pend] = 'background-color: #ff4d4d; color: white; font-weight: bold;' 
                    elif 4 <= pendientes <= 9: estilos[idx_pend] = 'background-color: #ffcc00; color: black; font-weight: bold;'
                    elif 1 <= pendientes <= 3: estilos[idx_pend] = 'background-color: #00cc66; color: white; font-weight: bold;'
                return estilos

            df_estilizado = df_horarios.style.apply(aplicar_estilos, axis=1)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_estilizado.to_excel(writer, index=False, sheet_name='Horarios', startrow=3)
                workbook = writer.book
                worksheet = writer.sheets['Horarios']
                
                borde = Border(left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'), top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF'))

                worksheet.merge_cells('A1:E2')
                celda_titulo = worksheet['A1']
                celda_titulo.value = f"REPORTE SEMANAL DE CAPACITACIÓN | AVANCE DE SUCURSAL: {porcentaje_avance:.1f}%"
                celda_titulo.font = Font(color="FFFFFF" if porcentaje_avance < 50 or porcentaje_avance >= 85 else "000000", bold=True, size=16)
                celda_titulo.fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
                celda_titulo.alignment = Alignment(horizontal="center", vertical="center")
                
                for row_t in range(1, 3):
                    for col_t in range(1, 6):
                        worksheet.cell(row=row_t, column=col_t).border = borde

                header_fill = PatternFill(start_color="203764", end_color="203764", fill_type="solid")
                header_font = Font(color="FFFFFF", bold=True, size=12)
                for cell in worksheet[4][:5]: 
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = borde

                for row_idx in range(1, worksheet.max_row + 1):
                    worksheet.row_dimensions[row_idx].height = 30
                    if row_idx >= 5: 
                        for cell in worksheet[row_idx][:5]: 
                            cell.border = borde
                            cell.alignment = Alignment(vertical="center", wrap_text=True if cell.column == 4 else False)

                worksheet.column_dimensions['A'].width = 30
                worksheet.column_dimensions['B'].width = 35
                worksheet.column_dimensions['C'].width = 25
                worksheet.column_dimensions['D'].width = 45
                worksheet.column_dimensions['E'].width = 15
                
                worksheet.column_dimensions['F'].width = 3
                worksheet.column_dimensions['G'].width = 35
                worksheet.column_dimensions['H'].width = 25
                
                # --- TOP 10 PEOR ---
                worksheet.merge_cells('G4:H4')
                titulo_peor = worksheet['G4']
                titulo_peor.value = "⚠️ TOP 10 - MAYOR REZAGO"
                titulo_peor.fill = PatternFill(start_color="ff4d4d", end_color="ff4d4d", fill_type="solid")
                titulo_peor.font = Font(color="FFFFFF", bold=True)
                titulo_peor.alignment = Alignment(horizontal="center", vertical="center")
                worksheet.cell(row=4, column=7).border = borde
                worksheet.cell(row=4, column=8).border = borde
                
                fill_lista_peor = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid") 
                font_lista_peor = Font(color="9C0006") 
                
                for i, (idx, fila) in enumerate(top_10_peor.iterrows(), start=5):
                    celda_nombre = worksheet.cell(row=i, column=7, value=fila['Nombre Completo'])
                    celda_cursos = worksheet.cell(row=i, column=8, value=f"{fila['Pendientes']} cursos por hacer")
                    for cell in [celda_nombre, celda_cursos]:
                        cell.fill = fill_lista_peor
                        cell.font = font_lista_peor
                        cell.border = borde
                        cell.alignment = Alignment(vertical="center")
                    celda_cursos.alignment = Alignment(horizontal="center", vertical="center")

                # --- TOP 10 MEJOR ---
                worksheet.merge_cells('G18:H18')
                titulo_mejor = worksheet['G18']
                titulo_mejor.value = "🌟 TOP 10 - MEJOR APROVECHAMIENTO"
                titulo_mejor.fill = PatternFill(start_color="00cc66", end_color="00cc66", fill_type="solid")
                titulo_mejor.font = Font(color="FFFFFF", bold=True)
                titulo_mejor.alignment = Alignment(horizontal="center", vertical="center")
                worksheet.cell(row=18, column=7).border = borde
                worksheet.cell(row=18, column=8).border = borde
                
                fill_lista_mejor = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid") 
                font_lista_mejor = Font(color="006100") 
                
                # ¡AQUÍ ESTÁ LA CORRECCIÓN! Agregué 'value=' para que escriba los datos
                for i, (idx, fila) in enumerate(top_10_mejor.iterrows(), start=19):
                    celda_nombre = worksheet.cell(row=i, column=7, value=fila['Nombre Completo'])
                    celda_cursos = worksheet.cell(row=i, column=8, value=f"{fila['Pendientes']} cursos por hacer")
                    for cell in [celda_nombre, celda_cursos]:
                        cell.fill = fill_lista_mejor
                        cell.font = font_lista_mejor
                        cell.border = borde
                        cell.alignment = Alignment(vertical="center")
                    celda_cursos.alignment = Alignment(horizontal="center", vertical="center")
                    
                for r in range(4, 29):
                    worksheet.row_dimensions[r].height = 30

            st.success("✅ ¡Horario generado y agrupado por áreas con éxito!")
            st.download_button(label="📥 Descargar Horario Final", data=output.getvalue(), file_name="Horarios_Zorro_Agrupados.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"❌ Error técnico: {e}")