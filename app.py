import streamlit as st
import pandas as pd
import plotly.express as px
import io

# Configuración de la página con estética moderna y limpia
st.set_page_config(
    page_title="BrainLabs Report Builder",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilo CSS personalizado para mejorar la estética (Premium Feel)
st.markdown("""
<style>
    /* Estilo para los títulos principales */
    .main-title {
        font-family: 'Outfit', 'Inter', sans-serif;
        color: #1E293B;
        font-weight: 800;
        font-size: 2.5rem;
        margin-bottom: 0.2rem;
        background: linear-gradient(90deg, #3B82F6 0%, #8B5CF6 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    
    .subtitle {
        font-family: 'Inter', sans-serif;
        color: #64748B;
        font-size: 1.1rem;
        margin-bottom: 2rem;
    }
    
    /* Tarjetas contenedoras */
    .card {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 1.5rem;
        margin-bottom: 1.5rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    
    /* Ajustes generales */
    .stTabs [data-baseweb="tab-list"] {
        gap: 24px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        font-weight: 600;
        font-size: 1.1rem;
        border-radius: 4px 4px 0px 0px;
    }
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# Inicializar estados de la aplicación en st.session_state
if "raw_df" not in st.session_state:
    st.session_state.raw_df = None
if "custom_cols" not in st.session_state:
    st.session_state.custom_cols = []
if "file_name" not in st.session_state:
    st.session_state.file_name = ""

# Cabecera de la aplicación
st.markdown('<h1 class="main-title">📊 BrainLabs Report Builder</h1>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Carga tus datos, crea fórmulas personalizadas y diseña gráficos interactivos al instante.</p>', unsafe_allow_html=True)

# ----------------- SIDEBAR: CARGA DE ARCHIVO -----------------
with st.sidebar:
    st.image("https://img.icons8.com/clouds/150/000000/combo-chart.png", width=100)
    st.markdown("### 📥 Cargar Datos")
    uploaded_file = st.file_uploader(
        "Sube tu archivo Excel (.xlsx, .xls) o CSV",
        type=["csv", "xlsx", "xls"],
        help="El archivo debe contener una fila de encabezados en la parte superior."
    )
    
    # Procesar archivo cargado
    if uploaded_file is not None:
        # Evitar recargar constantemente si es el mismo archivo
        if st.session_state.file_name != uploaded_file.name:
            st.session_state.file_name = uploaded_file.name
            st.session_state.custom_cols = [] # Reiniciar columnas calculadas al cambiar de archivo
            
            try:
                if uploaded_file.name.endswith(".csv"):
                    # Detectar encoding común
                    try:
                        df = pd.read_csv(uploaded_file, encoding='utf-8')
                    except UnicodeDecodeError:
                        df = pd.read_csv(uploaded_file, encoding='latin1')
                else:
                    df = pd.read_excel(uploaded_file)
                
                # Limpiar nombres de columnas (quitar espacios sobrantes)
                df.columns = [str(col).strip() for col in df.columns]
                st.session_state.raw_df = df
            except Exception as e:
                st.error(f"Error al leer el archivo: {e}")
    else:
        st.session_state.raw_df = None
        st.session_state.file_name = ""
        st.session_state.custom_cols = []

    st.markdown("---")
    st.markdown("### ⚙️ Información de Datos")
    if st.session_state.raw_df is not None:
        rows, cols = st.session_state.raw_df.shape
        st.success("✅ Archivo cargado con éxito")
        st.metric(label="Filas", value=rows)
        st.metric(label="Columnas Originales", value=cols)
    else:
        st.info("Por favor, sube un archivo para comenzar.")

# ----------------- PROCESAMIENTO DE COLUMNAS CALCULADAS -----------------
def apply_custom_columns(df):
    """Aplica las operaciones calculadas al DataFrame en orden."""
    if df is None:
        return None
    
    # Copia para no modificar el original
    working_df = df.copy()
    
    for c_col in st.session_state.custom_cols:
        name = c_col["name"]
        col1 = c_col["col1"]
        op = c_col["op"]
        operand_type = c_col["operand_type"]
        
        # Validar que col1 exista y sea numérica
        if col1 not in working_df.columns:
            continue
        
        v1 = pd.to_numeric(working_df[col1], errors='coerce')
        
        # Obtener el segundo operando (columna o valor constante)
        if operand_type == "column":
            col2 = c_col["col2"]
            if col2 not in working_df.columns:
                continue
            v2 = pd.to_numeric(working_df[col2], errors='coerce')
        else:
            v2 = float(c_col["constant_val"])
            
        # Realizar la operación matemática
        if op == "Multiplicar (*)":
            working_df[name] = v1 * v2
        elif op == "Dividir (/)":
            # Evitar división por cero
            working_df[name] = v1.div(v2).replace([float('inf'), float('-inf')], None)
        elif op == "Sumar (+)":
            working_df[name] = v1 + v2
        elif op == "Restar (-)":
            working_df[name] = v1 - v2
            
    return working_df

# Obtener DataFrame con columnas calculadas
if st.session_state.raw_df is not None:
    processed_df = apply_custom_columns(st.session_state.raw_df)
else:
    processed_df = None


# ----------------- VISTA DE TABS -----------------
if processed_df is not None:
    tab1, tab2 = st.tabs(["📋 Generador de Reportes", "📈 Visualización de Gráficos"])
    
    # ----------------------------------------------------
    # TAB 1: GENERADOR DE REPORTES
    # ----------------------------------------------------
    with tab1:
        st.markdown("### 🔧 Configura tu Reporte")
        
        col_setup, col_preview = st.columns([1, 2])
        
        with col_setup:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown("#### ➕ Agregar Columna Calculada")
            
            # Contenedor para crear columna personalizada (sin st.form para permitir vista previa en tiempo real)
            with st.container():
                # Validar columnas numéricas para las operaciones
                numeric_cols = processed_df.select_dtypes(include=['number']).columns.tolist()
                
                new_col_name = st.text_input(
                    "Nombre de la nueva columna",
                    placeholder="Ej. Ingreso Neto, Margen, etc."
                ).strip()
                
                col1_selected = st.selectbox(
                    "Selecciona Columna A",
                    options=numeric_cols,
                    help="Esta columna debe tener valores numéricos."
                )
                
                operator = st.selectbox(
                    "Operación",
                    options=["Multiplicar (*)", "Dividir (/)", "Sumar (+)", "Restar (-)"]
                )
                
                op2_type = st.radio(
                    "Segundo operando",
                    options=["Otra columna", "Valor constante"],
                    horizontal=True
                )
                
                if op2_type == "Otra columna":
                    col2_selected = st.selectbox("Selecciona Columna B", options=numeric_cols)
                    constant_value = 0.0
                else:
                    col2_selected = None
                    constant_value = st.number_input("Ingresa el valor constante", value=1.0, format="%f")
                
                # --- VISTA PREVIA EN TIEMPO REAL ---
                if col1_selected:
                    try:
                        v1 = pd.to_numeric(processed_df[col1_selected], errors='coerce')
                        if op2_type == "Otra columna" and col2_selected:
                            v2 = pd.to_numeric(processed_df[col2_selected], errors='coerce')
                        else:
                            v2 = float(constant_value)
                        
                        if operator == "Multiplicar (*)":
                            preview_series = v1 * v2
                        elif operator == "Dividir (/)":
                            preview_series = v1.div(v2).replace([float('inf'), float('-inf')], None)
                        elif operator == "Sumar (+)":
                            preview_series = v1 + v2
                        elif operator == "Restar (-)":
                            preview_series = v1 - v2
                            
                        st.markdown("**🔍 Vista previa del cálculo (primeras 3 filas):**")
                        preview_df = pd.DataFrame({
                            col1_selected: v1.head(3),
                            "Operador": operator.split()[0],
                            "Operando B": v2.head(3) if op2_type == "Otra columna" else v2,
                            "=": "=",
                            new_col_name if new_col_name else "Resultado": preview_series.head(3)
                        })
                        st.dataframe(preview_df, use_container_width=True)
                    except Exception as e:
                        pass
                
                submit_button = st.button("Confirmar y Crear Columna", use_container_width=True)
                
                if submit_button:
                    # Validaciones
                    if not new_col_name:
                        st.error("Por favor, ingresa un nombre para la nueva columna.")
                    elif new_col_name in processed_df.columns:
                        st.error(f"Ya existe una columna llamada '{new_col_name}'.")
                    elif not col1_selected:
                        st.error("Debes seleccionar al menos una columna numérica de origen.")
                    else:
                        # Guardar configuración en la lista
                        new_col_data = {
                            "name": new_col_name,
                            "col1": col1_selected,
                            "op": operator,
                            "operand_type": "column" if op2_type == "Otra columna" else "constant",
                            "col2": col2_selected,
                            "constant_val": constant_value
                        }
                        st.session_state.custom_cols.append(new_col_data)
                        st.success(f"¡Columna '{new_col_name}' creada exitosamente!")
                        st.rerun()
            
            st.markdown('</div>', unsafe_allow_html=True)
            
            # Listado de columnas calculadas creadas
            if st.session_state.custom_cols:
                st.markdown('<div class="card">', unsafe_allow_html=True)
                st.markdown("#### ⚙️ Columnas Calculadas Activas")
                for index, c_col in enumerate(st.session_state.custom_cols):
                    col_item, col_btn = st.columns([4, 1])
                    with col_item:
                        if c_col["operand_type"] == "column":
                            expr = f"`{c_col['col1']}` {c_col['op'].split()[-1]} `{c_col['col2']}`"
                        else:
                            expr = f"`{c_col['col1']}` {c_col['op'].split()[-1]} {c_col['constant_val']}"
                        st.markdown(f"**{c_col['name']}** = {expr}")
                    with col_btn:
                        # Botón para borrar columna usando key única
                        if st.button("🗑️", key=f"del_{index}_{c_col['name']}"):
                            st.session_state.custom_cols.pop(index)
                            st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)
        
        with col_preview:
            st.markdown('<div class="card" style="height: 100%;">', unsafe_allow_html=True)
            st.markdown("#### ⚙️ Seleccionar Columnas a Incluir")
            
            # Selector de columnas final para el reporte
            all_cols = processed_df.columns.tolist()
            selected_columns = st.multiselect(
                "Selecciona las columnas que quieres ver en tu reporte final:",
                options=all_cols,
                default=all_cols
            )
            
            st.markdown("#### 📋 Vista Previa del Reporte")
            
            if len(selected_columns) > 0:
                # Filtrar el DataFrame final con las columnas seleccionadas
                final_df = processed_df[selected_columns]
                
                # Mostrar tabla interactiva de Streamlit
                st.dataframe(final_df, use_container_width=True, height=350)
                
                # Fila con opciones de exportación/descarga
                st.markdown("#### 📥 Descargar Reporte")
                dl_col1, dl_col2, _ = st.columns([1, 1, 2])
                
                with dl_col1:
                    # Exportar a CSV
                    csv_data = final_df.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="Descargar CSV 📄",
                        data=csv_data,
                        file_name="reporte_personalizado.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                
                with dl_col2:
                    # Exportar a Excel (.xlsx) usando openpyxl
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        final_df.to_excel(writer, index=False, sheet_name='Reporte')
                    excel_data = output.getvalue()
                    
                    st.download_button(
                        label="Descargar Excel 📁",
                        data=excel_data,
                        file_name="reporte_personalizado.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            else:
                st.warning("Selecciona al menos una columna para visualizar e iniciar la descarga.")
            st.markdown('</div>', unsafe_allow_html=True)

    # ----------------------------------------------------
    # TAB 2: VISUALIZACIÓN DE GRÁFICOS
    # ----------------------------------------------------
    with tab2:
        st.markdown("### 📈 Visualizador de Gráficos Interactivos")
        
        col_chart_setup, col_chart_render = st.columns([1, 2])
        
        with col_chart_setup:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown("#### ⚙️ Configurar Gráfico")
            
            chart_type = st.selectbox(
                "Tipo de Gráfico",
                options=["Barras", "Líneas", "Dispersión (Scatter)", "Torta (Pie)"]
            )
            
            # Lista de todas las columnas procesadas (incluyendo calculadas)
            all_processed_cols = processed_df.columns.tolist()
            numeric_cols = processed_df.select_dtypes(include=['number']).columns.tolist()
            
            x_axis = st.selectbox(
                "Eje X (Categorías, Fechas o Grupos)",
                options=all_processed_cols
            )
            
            y_axis = st.selectbox(
                "Eje Y (Valores numéricos)",
                options=numeric_cols,
                help="El eje Y requiere valores numéricos para poder graficar correctamente."
            )
            
            color_by = st.selectbox(
                "Agrupar o Colorear por (Opcional)",
                options=["Ninguno"] + all_processed_cols
            )
            
            chart_title = st.text_input(
                "Título del Gráfico",
                value=f"Gráfico de {y_axis} por {x_axis}"
            )
            
            st.markdown('</div>', unsafe_allow_html=True)
            
            with st.expander("🔍 Ver datos seleccionados para el gráfico"):
                cols_to_preview = [x_axis, y_axis]
                if color_by != "Ninguno" and color_by not in cols_to_preview:
                    cols_to_preview.append(color_by)
                st.dataframe(processed_df[cols_to_preview].head(5), use_container_width=True)
                
            st.info("💡 Consejo: Los gráficos creados con Plotly son completamente interactivos. Puedes hacer zoom, seleccionar áreas o pasar el mouse por encima para ver los valores. Usa la barra de herramientas flotante del gráfico para descargarlo como imagen (icono de cámara).")
            
        with col_chart_render:
            st.markdown('<div class="card" style="height: 100%;">', unsafe_allow_html=True)
            st.markdown(f"#### 📊 {chart_title}")
            
            # Limpiar datos nulos en las columnas seleccionadas para evitar fallos de Plotly
            cols_to_plot = [x_axis, y_axis]
            color_column = None if color_by == "Ninguno" else color_by
            if color_column:
                cols_to_plot.append(color_column)
                
            plot_df = processed_df[cols_to_plot].dropna()
            
            if not plot_df.empty:
                try:
                    # Crear gráfico correspondiente
                    if chart_type == "Barras":
                        fig = px.bar(
                            plot_df, 
                            x=x_axis, 
                            y=y_axis, 
                            color=color_column,
                            title=chart_title,
                            template="plotly_white"
                        )
                    elif chart_type == "Líneas":
                        fig = px.line(
                            plot_df, 
                            x=x_axis, 
                            y=y_axis, 
                            color=color_column,
                            title=chart_title,
                            template="plotly_white"
                        )
                    elif chart_type == "Dispersión (Scatter)":
                        fig = px.scatter(
                            plot_df, 
                            x=x_axis, 
                            y=y_axis, 
                            color=color_column,
                            title=chart_title,
                            template="plotly_white"
                        )
                    elif chart_type == "Torta (Pie)":
                        # El gráfico de torta agrupa la suma si hay duplicados en el eje X
                        pie_df = plot_df.groupby(x_axis)[y_axis].sum().reset_index()
                        fig = px.pie(
                            pie_df, 
                            names=x_axis, 
                            values=y_axis, 
                            title=chart_title,
                            template="plotly_white"
                        )
                    
                    # Personalizar diseño para estética premium
                    fig.update_layout(
                        font_family="Inter, sans-serif",
                        title_font_family="Outfit, sans-serif",
                        title_font_size=20,
                        legend_title_font_family="Outfit, sans-serif",
                        paper_bgcolor='rgba(0,0,0,0)',
                        plot_bgcolor='rgba(0,0,0,0)'
                    )
                    
                    # Mostrar gráfico en Streamlit
                    st.plotly_chart(fig, use_container_width=True)
                    
                except Exception as e:
                    st.error(f"Error al generar el gráfico: {e}")
            else:
                st.warning("No hay suficientes datos válidos (sin valores nulos) para generar el gráfico.")
            st.markdown('</div>', unsafe_allow_html=True)
else:
    # Estado inicial: Pantalla vacía pidiendo carga
    st.markdown('<div class="card" style="text-align: center; padding: 3rem;">', unsafe_allow_html=True)
    st.image("https://img.icons8.com/clouds/200/000000/cloud-upload.png", width=150)
    st.markdown("### ¡Comencemos!")
    st.markdown("Sube un archivo Excel o CSV desde la barra lateral izquierda para cargar la información, empezar a construir tu reporte y graficar.")
    st.markdown('</div>', unsafe_allow_html=True)
