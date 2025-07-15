import streamlit as st
from google import genai
import pandas as pd

# Configuración de la página
st.set_page_config(page_title="Gestor de Juicios Evaluativos", layout="wide")
# Forzar tema oscuro en la app (requiere config.toml en .streamlit)
st.title("📊 Gestor de Juicios Evaluativos")
st.markdown("""
Esta aplicación te permite analizar, visualizar y exportar los resultados de los juicios evaluativos de aprendices a partir de un archivo "Reporte de Juicios Evaluativos". 

Carga el archivo de juicios evaluativos generado por el SENA y explora diferentes análisis: distribución de estados de los aprendices, porcentajes de aprobación por resultado de aprendizaje, mapas de calor, análisis por funcionario y más. Puedes filtrar los resultados por resultado de aprendizaje y descargar un reporte personalizado en Excel con formato especial.



""")

# Botón para subir archivo Excel
df = None
archivo = st.file_uploader("Sube el archivo Excel de Juicios Evaluativos:", type=["xls", "xlsx"])
if archivo is not None:
    try:
        # --- Mostrar información del reporte (filas 1 a 12, solo columnas 0 y 2) ---
        # Leer las primeras 12 filas como información del reporte
        info_reporte_full = pd.read_excel(archivo, nrows=12, header=None, engine="xlrd" if archivo.name.endswith(".xls") else None)
        info_reporte = info_reporte_full[[0, 2]]
        # Mostrar la información del reporte en Streamlit (solo columnas 0 y 2)
        # Leer el resto del archivo como datos (saltando las primeras 12 filas)
        df = pd.read_excel(archivo, skiprows=12, engine="xlrd" if archivo.name.endswith(".xls") else None)
        print(df.info())  # Mostrar información del DataFrame para depuración
        # Eliminar columnas vacías
        df = df.dropna(axis=1, how="all")
        # --- Filtro en la barra lateral para Resultados de Aprendizaje ---
        resultados_unicos = df['Resultado de Aprendizaje'].dropna().unique().tolist()
        resultados_unicos.sort()
        seleccion_resultados = st.sidebar.multiselect(
            'Filtrar por Resultados de Aprendizaje:',
            options=['Todos'] + resultados_unicos,
            default=['Todos']
        )
        if 'Todos' in seleccion_resultados or not seleccion_resultados:
            df_filtrado = df
        else:
            df_filtrado = df[df['Resultado de Aprendizaje'].isin(seleccion_resultados)]
        # Si existe la columna 'Fecha y Hora del Juicio Evaluativo', mantenerla en el DataFrame filtrado
        if 'Fecha y Hora del Juicio Evaluativo' in df_filtrado.columns:
            # No eliminar ni modificar, solo asegurar que esté presente en df_filtrado
            pass
        # Asegurar estudiantes únicos por Número de Documento y mostrar nombre y apellidos
        df_filtrado = df_filtrado.drop_duplicates(subset=["Número de Documento", "Resultado de Aprendizaje"])
        # Crear tabla pivote: filas=estudiantes únicos por Número de Documento, columnas=Resultados de Aprendizaje, valores=Juicio de Evaluación
        # Agregar la columna Estado al índice para que aparezca en la tabla exportada
        tabla_pivote = pd.pivot_table(
            df_filtrado,
            index=["Número de Documento", "Nombre", "Apellidos", "Estado"],
            columns="Resultado de Aprendizaje",
            values="Juicio de Evaluación",
            aggfunc=lambda x: ', '.join(x.astype(str))
        )
        # Crear tabla pivote personalizada para exportar
        def map_juicio(val):
            if isinstance(val, str):
                if "aprobado" in val.lower():
                    return "X"
                elif "no aprobado" in val.lower():
                    return "N"
                elif "por evaluar" in val.lower():
                    return ""
            return ""

        # Crear MultiIndex para columnas: primer nivel=Competencia, segundo nivel=Resultado de Aprendizaje
        competencias = df_filtrado.set_index("Resultado de Aprendizaje")["Competencia"].to_dict()
        multi_cols = pd.MultiIndex.from_tuples(
            [(competencias.get(col, ""), col) for col in tabla_pivote.columns],
            names=["Competencia", "Resultado de Aprendizaje"]
        )
        tabla_export = pd.DataFrame(tabla_pivote.values, index=tabla_pivote.index, columns=multi_cols)
        # Mapear los valores de Juicio de Evaluación
        for col in tabla_export.columns:
            tabla_export[col] = tabla_export[col].apply(map_juicio)

        # Calcular porcentaje de aprobado por estudiante
        def calcular_porcentaje_aprobado(row):
            total = 0
            aprobados = 0
            for col in tabla_export.columns:
                val = row[col]
                if val in ("X", "N", ""):
                    total += 1
                    if val == "X":
                        aprobados += 1
            return round((aprobados / total) * 100, 2) if total > 0 else 0

        tabla_export[('','% Aprobado')] = tabla_export.apply(calcular_porcentaje_aprobado, axis=1)

        # Calcular porcentaje de aprobación del grupo
        total_celdas = 0
        total_aprobadas = 0
        for idx, row in tabla_export.iterrows():
            for col in tabla_export.columns[:-1]:  # Excluir columna de porcentaje
                val = row[col]
                if val in ("X", "N", ""):
                    total_celdas += 1
                    if val == "X":
                        total_aprobadas += 1
        porcentaje_grupo = round((total_aprobadas / total_celdas) * 100, 2) if total_celdas > 0 else 0
       
        # --- Botón para exportar a Excel con formato organizado y encabezados multinivel ---
        st.subheader("Exportar tabla personalizada a Excel")
        import io
        import xlsxwriter
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # --- Agregar hoja de información del reporte (solo columnas 0 y 2) ---
            info_reporte.to_excel(writer, sheet_name='Reporte', index=False, header=False)
            workbook  = writer.book
            worksheet_reporte = writer.sheets['Reporte']
            # Autoajustar ancho de columnas 0 y 1 en la hoja 'Reporte'
            for i, col in enumerate(info_reporte.columns):
                max_len = max([len(str(val)) for val in info_reporte[col].values] + [len(str(col))])
                worksheet_reporte.set_column(i, i, max_len + 2)
            # --- Agregar hoja de resultados ---
            worksheet = writer.book.add_worksheet('Resultados')
            writer.sheets['Resultados'] = worksheet
            ncols = tabla_export.reset_index().shape[1]
            header_format = workbook.add_format({'bold': True, 'bg_color': '#F4B942', 'font_color': '#1A2930', 'border': 1, 'align': 'center', 'text_wrap': True})
            subheader_format = workbook.add_format({'bold': True, 'bg_color': '#4FC3F7', 'font_color': '#1A2930', 'border': 1, 'align': 'center', 'text_wrap': True})
            index_format = workbook.add_format({'bg_color': '#ECECEC', 'border': 1, 'align': 'center'})
            header_format_vcenter = workbook.add_format({'bold': True, 'bg_color': '#F4B942', 'font_color': '#1A2930', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
            formato_verde = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'align': 'center', 'border': 1})
            formato_rojo = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'align': 'center', 'border': 1})
            formato_blanco = workbook.add_format({'bg_color': '#FFFFFF', 'align': 'center', 'border': 1})
            formato_rosa_fila = workbook.add_format({'bg_color': '#F8CBAD', 'align': 'center', 'border': 1})
            formato_dorado = workbook.add_format({'bg_color': '#FFF9C4', 'align': 'center', 'border': 1})
            formato_gris = workbook.add_format({'bg_color': '#E0E0E0', 'align': 'center', 'border': 1})
            num_index_cols = len(tabla_export.index.names)
            for col in range(num_index_cols):
                col_letter = xlsxwriter.utility.xl_col_to_name(col)
                worksheet.merge_range(0, col, 1, col, tabla_export.reset_index().columns[col][0], header_format_vcenter)
            from collections import OrderedDict
            competencias_cols = OrderedDict()
            for col in range(num_index_cols, ncols):
                competencia = tabla_export.reset_index().columns[col][0]
                if competencia not in competencias_cols:
                    competencias_cols[competencia] = [col, col]
                else:
                    competencias_cols[competencia][1] = col
            for competencia, (col_start, col_end) in competencias_cols.items():
                worksheet.merge_range(0, col_start, 0, col_end, competencia, header_format)
            for col in range(num_index_cols, ncols):
                worksheet.write(1, col, tabla_export.reset_index().columns[col][1], subheader_format)
            worksheet.set_row(0, 40, header_format)
            worksheet.set_row(1, 20, subheader_format)
            for col_num, col in enumerate(tabla_export.reset_index().columns):
                if col_num < 4 or col_num == ncols - 1:
                    max_len = max(
                        [len(str(val)) for val in tabla_export.reset_index()[col].values] + [len(str(col))]
                    )
                    worksheet.set_column(col_num, col_num, max_len + 2)
                else:
                    worksheet.set_column(col_num, col_num, 4)
            nrows = tabla_export.shape[0]
            for col in range(len(tabla_export.index.names), ncols):
                col_letter = xlsxwriter.utility.xl_col_to_name(col)
                worksheet.conditional_format(f'{col_letter}3:{col_letter}{nrows+2}', {'type': 'cell', 'criteria': '==', 'value': '"X"', 'format': formato_verde})
                worksheet.conditional_format(f'{col_letter}3:{col_letter}{nrows+2}', {'type': 'cell', 'criteria': '==', 'value': '"N"', 'format': formato_rojo})
                worksheet.conditional_format(f'{col_letter}3:{col_letter}{nrows+2}', {'type': 'cell', 'criteria': '==', 'value': '""', 'format': formato_blanco})
            tabla_reset = tabla_export.reset_index()
            for row in range(nrows):
                estado = str(tabla_reset.iloc[row, 3]).strip().upper()
                fila_rosa = estado != "EN FORMACION"
                for col in range(ncols):
                    valor = tabla_reset.iloc[row, col]
                    if fila_rosa:
                        cell_format = formato_rosa_fila
                    else:
                        if col >= len(tabla_export.index.names):
                            if valor == "X":
                                cell_format = formato_verde
                            elif valor == "N":
                                cell_format = formato_rojo
                            else:
                                cell_format = formato_blanco
                        else:
                            cell_format = index_format
                    worksheet.write(row+2, col, valor, cell_format)
            for row in range(nrows):
                for col in range(len(tabla_export.index.names), ncols):
                    valor = tabla_reset.iloc[row, col]
                    if valor in ("X", "N"):
                        num_doc = tabla_reset.iloc[row, 0]
                        resultado = tabla_reset.columns[col][1]
                        filtro = (
                            (df_filtrado["Número de Documento"] == num_doc) &
                            (df_filtrado["Resultado de Aprendizaje"] == resultado)
                        )
                        funcionario = df_filtrado.loc[filtro, "Funcionario que registro el juicio evaluativo"].astype(str).unique()
                        comentario = f"Funcionario: {funcionario[0]}" if len(funcionario) > 0 else "Funcionario: N/A"
                        worksheet.write_comment(row+2, col, comentario)
        output.seek(0)
        st.download_button(
            label="Descargar Excel personalizado",
            data=output.getvalue(),
            file_name="Tabla_Estudiantes_vs_Resultados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        # --- Analíticas y gráficos en Streamlit ---
        st.divider()
        
        # --- Tabs principales ---
        tabs = st.tabs([
            "Información general",
            "Distribución de estados",
            "Aprobación por resultado",
            "Mapa de calor",
            "Porcentaje por estudiante",
            "Juicios por funcionario",
            "Cobertura por funcionario"
        ])

        with tabs[0]:
            st.subheader("Información del reporte")
            st.dataframe(info_reporte, hide_index=True, use_container_width=True)
            st.metric("Porcentaje de aprobación del grupo", f"{porcentaje_grupo}%")

        with tabs[1]:
            st.subheader("Distribución de estados de los estudiantes (únicos)")
            st.markdown("""Este gráfico muestra la cantidad de estudiantes en cada estado (por ejemplo, 'En formación', 'Retirado', etc.). Permite identificar la distribución general del grupo según su estado académico actual.""")
            estados_unicos = df_filtrado.drop_duplicates(subset=["Número de Documento"])[["Número de Documento", "Estado"]]
            estado_counts = estados_unicos["Estado"].value_counts()
            st.bar_chart(estado_counts, use_container_width=True)

        with tabs[2]:
            st.subheader("Porcentaje de aprobación por resultado de aprendizaje")
            st.markdown("""Aquí puedes ver el porcentaje de estudiantes que aprobaron cada resultado de aprendizaje. Es útil para identificar cuáles resultados presentan mayores dificultades o logros dentro del grupo.""")
            aprob_por_resultado = df_filtrado.groupby('Resultado de Aprendizaje')['Juicio de Evaluación'].apply(
                lambda x: (x.str.lower().str.contains('aprobado').sum() / len(x)) * 100
            ).sort_values(ascending=False)
            st.bar_chart(aprob_por_resultado, use_container_width=True)

        with tabs[3]:
            st.subheader("Mapa de calor de aprobaciones por estudiante y resultado")
            st.markdown("""El mapa de calor permite visualizar rápidamente qué estudiantes han aprobado o no cada resultado de aprendizaje. El verde indica aprobado, el rojo no aprobado y el blanco por evaluar. Es útil para detectar patrones o estudiantes con dificultades específicas.""")
            import seaborn as sns
            import matplotlib.pyplot as plt
            heatmap_data = tabla_pivote.copy()
            def map_heat(val):
                if isinstance(val, str):
                    if "aprobado" in val.lower():
                        return 1
                    elif "no aprobado" in val.lower():
                        return 0
                return float('nan')
            heatmap_data = heatmap_data.applymap(map_heat)
            heatmap_data_small = heatmap_data.head(30)
            from matplotlib.colors import ListedColormap
            cmap = ListedColormap(["#FFFFFF", "#FFC7CE", "#C6EFCE"])
            fig_heat, ax_heat = plt.subplots(figsize=(min(20, 1+len(heatmap_data_small.columns)*0.7), 1+len(heatmap_data_small)*0.3))
            sns.heatmap(heatmap_data_small, cmap=cmap, cbar_kws={'label': '1=Aprobado (verde), 0=No aprobado (rojo)'}, linewidths=0.5, linecolor='gray', annot=False, vmin=0, vmax=1, ax=ax_heat)
            ax_heat.set_xlabel("Resultado de Aprendizaje")
            ax_heat.set_ylabel("Estudiante")
            ax_heat.set_title("Mapa de calor de aprobaciones (Verde=Aprobado, Rojo=No aprobado, Blanco=Por evaluar)")
            st.pyplot(fig_heat)

        with tabs[4]:
            st.subheader("Porcentaje de aprobación por estudiante")
            st.markdown("""Este gráfico muestra el porcentaje de aprobación de cada estudiante respecto a todos los resultados de aprendizaje evaluados. Permite identificar a los estudiantes con mejor y peor desempeño global.""")
            porcentaje_estudiantes = tabla_export[('', '% Aprobado')]
            nombres = tabla_export.index.get_level_values('Nombre')
            apellidos = tabla_export.index.get_level_values('Apellidos')
            nombres_completos = nombres + ' ' + apellidos
            porcentaje_df = pd.DataFrame({'Estudiante': nombres_completos, 'Porcentaje': porcentaje_estudiantes.values})
            porcentaje_df = porcentaje_df.sort_values('Porcentaje', ascending=False)
            import matplotlib.pyplot as plt
            fig, ax = plt.subplots(figsize=(8, min(0.4*len(porcentaje_df), 30)))
            ax.barh(porcentaje_df['Estudiante'], porcentaje_df['Porcentaje'], color='#4FC3F7')
            ax.set_xlabel('% Aprobado respecto a todos los resultados de aprendizaje')
            ax.set_ylabel('Estudiante')
            ax.set_title('Porcentaje de aprobación por estudiante')
            ax.set_xlim(0, 100)
            for i, (valor, y) in enumerate(zip(porcentaje_df['Porcentaje'], ax.patches)):
                ax.text(valor + 1, y.get_y() + y.get_height()/2, f'{valor:.2f}%', va='center', fontsize=8)
            ax.invert_yaxis()
            st.pyplot(fig)

        with tabs[5]:
            if 'Fecha y Hora del Juicio Evaluativo' in df_filtrado.columns:
                st.subheader("Juicios emitidos por funcionario (filtro por fecha de inicio)")
                st.markdown("""Este gráfico apila la cantidad de juicios 'Aprobado' y 'No aprobado' emitidos por cada funcionario. Permite analizar la carga y tendencia de evaluación de cada funcionario a lo largo del tiempo.""")
                fechas_func = pd.to_datetime(df_filtrado['Fecha y Hora del Juicio Evaluativo'], errors='coerce')
                fecha_min_func = fechas_func.min()
                fecha_max_func = fechas_func.max()
                filtro_todo_func = st.checkbox('Mostrar todos los juicios (funcionarios)', value=True, key='func_todo')
                if not filtro_todo_func and pd.notnull(fecha_min_func) and pd.notnull(fecha_max_func):
                    fecha_inicio = st.date_input(
                        'Selecciona la fecha de inicio para los juicios de funcionarios',
                        value=fecha_min_func.date(),
                        min_value=fecha_min_func.date(),
                        max_value=fecha_max_func.date(),
                        key='func_fecha_inicio'
                    )
                    fechas_func = pd.to_datetime(df_filtrado['Fecha y Hora del Juicio Evaluativo'], errors='coerce')
                    df_func_filtrado = df_filtrado[fechas_func.dt.date >= fecha_inicio]
                else:
                    df_func_filtrado = df_filtrado.copy()
                df_func_filtrado['Juicio de Evaluación'] = df_func_filtrado['Juicio de Evaluación'].astype(str).str.strip().str.lower()
                df_func_filtrado = df_func_filtrado[df_func_filtrado['Juicio de Evaluación'].isin(['aprobado', 'no aprobado'])]
                df_func_filtrado = df_func_filtrado[df_func_filtrado['Funcionario que registro el juicio evaluativo'].notna()]
                juicios_count2 = df_func_filtrado.groupby(['Funcionario que registro el juicio evaluativo', 'Juicio de Evaluación']).size().unstack(fill_value=0)
                import plotly.graph_objects as go
                if not juicios_count2.empty:
                    fig5 = go.Figure()
                    if 'aprobado' in juicios_count2.columns:
                        fig5.add_trace(go.Bar(
                            x=juicios_count2.index,
                            y=juicios_count2['aprobado'],
                            name='Aprobado',
                            marker_color='#4FC3F7',
                            hovertemplate='Funcionario: %{x}<br>Aprobado: %{y}<extra></extra>'
                        ))
                    if 'no aprobado' in juicios_count2.columns:
                        fig5.add_trace(go.Bar(
                            x=juicios_count2.index,
                            y=juicios_count2['no aprobado'],
                            name='No aprobado',
                            marker_color='#FFC7CE',
                            hovertemplate='Funcionario: %{x}<br>No aprobado: %{y}<extra></extra>'
                        ))
                    fig5.update_layout(barmode='stack', xaxis_title='Funcionario', yaxis_title='Cantidad de juicios', title='Juicios emitidos por funcionario (filtro por fecha de inicio)', legend_title='Tipo de Juicio', xaxis_tickangle=-30)
                    st.plotly_chart(fig5, use_container_width=True)
                else:
                    st.info("No hay juicios 'Aprobado' o 'No aprobado' registrados por los funcionarios en el rango seleccionado.")

        with tabs[6]:
            if 'Fecha y Hora del Juicio Evaluativo' in df_filtrado.columns:
                st.subheader("Porcentaje de cobertura de evaluación por funcionario")
                st.markdown("""Aquí se muestra el porcentaje de cobertura de evaluación de cada funcionario, es decir, qué proporción de los juicios posibles (estudiantes x resultados) ha registrado cada uno. Es útil para identificar la participación y carga de trabajo de los funcionarios evaluadores.""")
                try:
                    df_func_cobertura = df_func_filtrado.copy()
                except NameError:
                    df_func_cobertura = df_filtrado.copy()
                df_func_cobertura = df_func_cobertura[df_func_cobertura['Juicio de Evaluación'].isin(['aprobado', 'no aprobado'])]
                df_func_cobertura = df_func_cobertura[df_func_cobertura['Funcionario que registro el juicio evaluativo'].notna()]
                total_estudiantes = df_filtrado['Número de Documento'].nunique()
                total_resultados = df_filtrado['Resultado de Aprendizaje'].nunique()
                total_posibles = total_estudiantes * total_resultados
                juicios_por_funcionario = df_func_cobertura.groupby('Funcionario que registro el juicio evaluativo').size()
                porcentaje_cobertura = (juicios_por_funcionario / total_posibles * 100).fillna(0).sort_values(ascending=False)
                import matplotlib.pyplot as plt
                fig_cob, ax_cob = plt.subplots(figsize=(8, min(0.4*len(porcentaje_cobertura), 12)))
                ax_cob.barh(porcentaje_cobertura.index, porcentaje_cobertura.values, color='#FFD600')
                ax_cob.set_xlabel('% de cobertura de evaluación (sobre todos los estudiantes y resultados)')
                ax_cob.set_ylabel('Funcionario')
                ax_cob.set_title('Porcentaje de cobertura de evaluación por funcionario')
              
                for i, (valor, y) in enumerate(zip(porcentaje_cobertura.values, ax_cob.patches)):
                    ax_cob.text(valor + 1, y.get_y() + y.get_height()/2, f'{valor:.2f}%', va='center', fontsize=8)
                ax_cob.invert_yaxis()
                st.pyplot(fig_cob)
    except Exception as e:
        st.error(f"Error al leer el archivo Excel: {e}")

# Mostrar créditos al pie de página
st.markdown("<hr style='margin-top:40px;margin-bottom:10px;'>", unsafe_allow_html=True)
st.markdown("<div style='text-align:center; color:gray; font-size:14px;'>Desarrollado por: Jhon Fredy Valencia Gómez TIC- Electrónica</div>", unsafe_allow_html=True)
