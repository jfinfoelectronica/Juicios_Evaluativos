import streamlit as st
import pandas as pd
import gc
import warnings
warnings.filterwarnings('ignore')

# Configuración de la página
st.set_page_config(page_title="Gestor de Juicios Evaluativos", layout="wide")
# Forzar tema oscuro en la app (requiere config.toml en .streamlit)
st.title("📊 Gestor de Juicios Evaluativos")
st.markdown("""
Esta aplicación te permite analizar, visualizar y exportar los resultados de los juicios evaluativos de aprendices a partir de un archivo "Reporte de Juicios Evaluativos". 

Carga el archivo de juicios evaluativos generado por el SENA y explora diferentes análisis: distribución de estados de los aprendices, porcentajes de aprobación por resultado de aprendizaje, mapas de calor, análisis por funcionario y más. Puedes filtrar los resultados por resultado de aprendizaje y descargar un reporte personalizado en Excel con formato especial.

**✅ Optimizado para archivos grandes (más de 4000 filas)**

""")

def leer_excel_robusto(archivo, skiprows=0, nrows=None):
    """
    Función robusta para leer archivos Excel grandes con múltiples estrategias.
    Soporta .xls, .xlsx, .xlsb y maneja archivos de gran tamaño.
    """
    nombre_archivo = archivo.name.lower()
    
    # Estrategias de lectura ordenadas por eficiencia para archivos grandes
    estrategias = []
    
    if nombre_archivo.endswith('.xlsb'):
        estrategias.append(('pyxlsb', 'pyxlsb'))
    elif nombre_archivo.endswith('.xlsx'):
        estrategias.extend([
            ('openpyxl', 'openpyxl'),
            ('calamine', 'calamine'),
            ('xlrd', None)
        ])
    elif nombre_archivo.endswith('.xls'):
        estrategias.extend([
            ('xlrd', 'xlrd'),
            ('calamine', 'calamine')
        ])
    
    # Agregar estrategia por defecto
    estrategias.append((None, None))
    
    for i, (engine_name, engine) in enumerate(estrategias):
        try:             
            # Configurar parámetros según el motor
            kwargs = {
                'skiprows': skiprows,
                'header': None if nrows else 0
            }
            
            if nrows is not None:
                kwargs['nrows'] = nrows
                
            if engine:
                kwargs['engine'] = engine
            
            # Leer archivo
            df = pd.read_excel(archivo, **kwargs)
            
            # Validar que se leyeron datos
            if df.empty:
                raise ValueError("El archivo está vacío o no contiene datos válidos")
            
            
            return df
            
        except Exception as e:
            error_msg = str(e)
            st.warning(f"⚠️ Error con {engine_name or 'motor por defecto'}: {error_msg}")
            
            # Si es el último intento, mostrar error detallado
            if i == len(estrategias) - 1:
                st.error(f"""
                ❌ **No se pudo leer el archivo después de intentar todos los métodos disponibles.**
                
                **Posibles soluciones:**
                1. Verifica que el archivo no esté corrupto
                2. Intenta guardar el archivo en formato .xlsx desde Excel
                3. Si el archivo es muy grande (>100MB), considera dividirlo en partes más pequeñas
                4. Verifica que el archivo tenga el formato esperado del SENA
                
                **Error técnico:** {error_msg}
                """)
                return None
            
            continue
    
    return None

def optimizar_memoria_dataframe(df):
    """
    Optimiza el uso de memoria del DataFrame convirtiendo tipos de datos.
    """
    if df is None:
        return None
    
    # Columnas que NO deben convertirse a categóricas para evitar errores de concatenación
    columnas_excluidas = ['Nombre', 'Apellidos', 'Número de Documento']
        
    # Optimizar tipos de datos para reducir memoria
    for col in df.columns:
        if df[col].dtype == 'object':
            # Intentar convertir a categoría si hay muchos valores repetidos
            # pero excluir columnas críticas que se usan en concatenaciones
            if (df[col].nunique() / len(df) < 0.5 and 
                col not in columnas_excluidas):  # Si menos del 50% son únicos y no está excluida
                try:
                    df[col] = df[col].astype('category')
                except:
                    pass
        elif df[col].dtype == 'int64':
            # Optimizar enteros
            try:
                df[col] = pd.to_numeric(df[col], downcast='integer')
            except:
                pass
        elif df[col].dtype == 'float64':
            # Optimizar flotantes
            try:
                df[col] = pd.to_numeric(df[col], downcast='float')
            except:
                pass
    
    return df

def leer_archivo_por_chunks(archivo, skiprows=12, chunk_size=5000):
    """
    Lee archivos extremadamente grandes por chunks para evitar problemas de memoria.
    """
    try:
        st.info("🔄 Archivo muy grande detectado. Leyendo por chunks para optimizar memoria...")
        
        chunks = []
        chunk_count = 0
        
        # Intentar leer por chunks
        nombre_archivo = archivo.name.lower()
        engine = None
        
        if nombre_archivo.endswith('.xlsb'):
            engine = 'pyxlsb'
        elif nombre_archivo.endswith('.xlsx'):
            engine = 'openpyxl'
        elif nombre_archivo.endswith('.xls'):
            engine = 'xlrd'
        
        # Leer el archivo completo primero para obtener el número total de filas
        try:
            df_temp = pd.read_excel(archivo, skiprows=skiprows, engine=engine, nrows=1)
            total_cols = df_temp.shape[1]
            del df_temp
            gc.collect()
        except:
            total_cols = None
        
        # Leer por chunks
        current_row = skiprows
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        while True:
            try:
                chunk = pd.read_excel(
                    archivo, 
                    skiprows=current_row, 
                    nrows=chunk_size,
                    engine=engine,
                    header=0 if chunk_count == 0 else None
                )
                
                if chunk.empty:
                    break
                
                # Si no es el primer chunk, usar las columnas del primer chunk
                if chunk_count > 0 and chunks:
                    chunk.columns = chunks[0].columns
                
                chunks.append(chunk)
                chunk_count += 1
                current_row += chunk_size
                
                # Actualizar progreso
                status_text.text(f"Procesando chunk {chunk_count}: {len(chunk)} filas leídas")
                
                # Limitar memoria liberando chunks antiguos si hay demasiados
                if len(chunks) > 10:  # Procesar en lotes de 10 chunks
                    df_partial = pd.concat(chunks, ignore_index=True)
                    chunks = [df_partial]
                    gc.collect()
                
            except Exception as e:
                if "No data" in str(e) or chunk.empty:
                    break
                else:
                    raise e
        
        progress_bar.progress(100)
        
        if chunks:
            st.info(f"📊 Combinando {chunk_count} chunks...")
            df_final = pd.concat(chunks, ignore_index=True)
            
            # Limpiar memoria
            del chunks
            gc.collect()
            
            st.success(f"✅ Archivo leído exitosamente por chunks: {df_final.shape[0]:,} filas")
            return df_final
        else:
            st.error("❌ No se pudieron leer datos del archivo")
            return None
            
    except Exception as e:
        st.error(f"❌ Error al leer archivo por chunks: {str(e)}")
        return None

# Botón para subir archivo Excel
df = None
archivo = st.file_uploader(
    "Sube el archivo Excel de Juicios Evaluativos:", 
    type=["xls", "xlsx", "xlsb"],
    help="Formatos soportados: .xls, .xlsx, .xlsb (recomendado para archivos grandes)"
)

if archivo is not None:
    try:
        # Mostrar información del archivo
        file_size = len(archivo.getvalue()) / (1024 * 1024)  # MB
        st.info(f"📁 **Archivo:** {archivo.name} ({file_size:.1f} MB)")
        
        if file_size > 50:
            st.warning("⚠️ Archivo grande detectado. El procesamiento puede tomar varios minutos.")
        
        # --- Mostrar información del reporte (filas 1 a 12, solo columnas 0 y 2) ---
        with st.spinner("Leyendo información del reporte..."):
            info_reporte_full = leer_excel_robusto(archivo, nrows=12)
            
            if info_reporte_full is not None and info_reporte_full.shape[1] >= 3:
                info_reporte = info_reporte_full[[0, 2]]
            else:
                st.warning("No se pudo leer la información del reporte. Continuando con los datos principales...")
                info_reporte = pd.DataFrame()
        
        # --- Leer el resto del archivo como datos (saltando las primeras 12 filas) ---
        with st.spinner("Leyendo datos principales del archivo..."):
            # Intentar lectura normal primero
            df = leer_excel_robusto(archivo, skiprows=12)
            
            # Si falla la lectura normal y el archivo es grande, intentar por chunks
            if df is None and file_size > 10:  # Si es mayor a 10MB
                st.warning("⚠️ Lectura normal falló. Intentando lectura por chunks...")
                df = leer_archivo_por_chunks(archivo, skiprows=12)
            
            if df is None:
                st.error("❌ No se pudo leer el archivo con ningún método disponible.")
                st.stop()
        
        # Optimizar memoria
        with st.spinner("Optimizando uso de memoria..."):
            df = optimizar_memoria_dataframe(df)
            
            # Eliminar columnas vacías
            df = df.dropna(axis=1, how="all")
            
            # Forzar recolección de basura
            gc.collect()
        
        # Validar estructura del archivo
        columnas_requeridas = [
            'Resultado de Aprendizaje', 
            'Juicio de Evaluación', 
            'Número de Documento',
            'Nombre', 
            'Apellidos', 
            'Estado'
        ]
        
        columnas_faltantes = [col for col in columnas_requeridas if col not in df.columns]
        
        if columnas_faltantes:
            st.error(f"""
            ❌ **El archivo no tiene la estructura esperada.**
            
            **Columnas faltantes:** {', '.join(columnas_faltantes)}
            
            **Columnas encontradas:** {', '.join(df.columns.tolist())}
            
            Por favor, verifica que el archivo sea un reporte de juicios evaluativos del SENA con el formato correcto.
            """)
            st.stop()
        
        # Mostrar estadísticas del archivo procesado
        st.success(f"🎉 **Archivo procesado exitosamente!**")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("📊 Total de filas", f"{len(df):,}")
        with col2:
            st.metric("👥 Estudiantes únicos", f"{df['Número de Documento'].nunique():,}")
        with col3:
            st.metric("📚 Resultados de aprendizaje", f"{df['Resultado de Aprendizaje'].nunique():,}")
        with col4:
            memoria_mb = df.memory_usage(deep=True).sum() / (1024 * 1024)
            st.metric("💾 Memoria utilizada", f"{memoria_mb:.1f} MB")
        
        # Usar el DataFrame completo sin filtro de Resultados de Aprendizaje
        df_filtrado = df
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
            
            # Crear rangos de merge solo si no se superponen
            for competencia, (col_start, col_end) in competencias_cols.items():
                try:
                    if col_start == col_end:
                        # Si es una sola columna, usar write en lugar de merge_range
                        worksheet.write(0, col_start, competencia, header_format)
                    else:
                        # Solo hacer merge si hay múltiples columnas
                        worksheet.merge_range(0, col_start, 0, col_end, competencia, header_format)
                except Exception as e:
                    # Si hay error en el merge, escribir en cada celda individualmente
                    for col in range(col_start, col_end + 1):
                        worksheet.write(0, col, competencia, header_format)
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
            # Convertir a string para evitar errores con tipos categóricos
            nombres_completos = nombres.astype(str) + ' ' + apellidos.astype(str)
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
    except MemoryError:
        st.error("""
        ❌ **Error de memoria insuficiente**
        
        El archivo es demasiado grande para procesar con la memoria disponible.
        
        **Soluciones recomendadas:**
        1. **Convertir a formato .xlsb** (más eficiente para archivos grandes)
        2. **Dividir el archivo** en partes más pequeñas (máximo 50,000 filas por archivo)
        3. **Cerrar otras aplicaciones** para liberar memoria
        4. **Filtrar datos en Excel** antes de cargar (eliminar columnas innecesarias)
        """)
    except pd.errors.EmptyDataError:
        st.error("❌ El archivo está vacío o no contiene datos válidos.")
    except (ValueError, KeyError) as e:
        if "Excel file format" in str(e) or "not supported" in str(e):
            st.error("❌ El archivo no es un archivo Excel válido o está corrupto.")
        else:
            raise e
    except Exception as e:
        error_msg = str(e)
        st.error(f"""
        ❌ **Error inesperado al procesar el archivo**
        
        **Error técnico:** {error_msg}
        
        **Posibles soluciones:**
        1. Verifica que el archivo no esté abierto en Excel
        2. Intenta guardar el archivo en formato .xlsx o .xlsb
        3. Verifica que el archivo tenga el formato esperado del SENA
        4. Si el archivo es muy grande (>100MB), considera dividirlo en partes
        
        **Formatos recomendados para archivos grandes:**
        - **.xlsb** (binario de Excel, más rápido)
        - **.xlsx** (estándar, compatible)
        """)
        
        # Mostrar información técnica adicional en un expander
        with st.expander("🔧 Información técnica detallada"):
            st.code(f"""
Tipo de error: {type(e).__name__}
Mensaje: {error_msg}
Archivo: {archivo.name if archivo else 'N/A'}
Tamaño: {file_size:.1f} MB
            """)
            
        # Sugerir alternativas
        st.info("""
        💡 **¿Necesitas ayuda?**
        
        Si continúas teniendo problemas:
        1. Intenta con un archivo más pequeño para probar la aplicación
        2. Contacta al administrador del sistema
        3. Verifica que tengas la versión más reciente de la aplicación
        """)

# Mostrar créditos al pie de página
st.markdown("<hr style='margin-top:40px;margin-bottom:10px;'>", unsafe_allow_html=True)
st.markdown("<div style='text-align:center; color:gray; font-size:14px;'>Desarrollado por: Jhon Fredy Valencia Gómez TIC- Electrónica</div>", unsafe_allow_html=True)
