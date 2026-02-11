import streamlit as st
import pandas as pd
import io
import xlsxwriter

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Agrupador Excel Pro", page_icon="🗂️", layout="wide")

# --- ESTILOS VISUALES ---
CABIFY_PURPLE = '#7145D6'
CABIFY_LIGHT  = '#F3F0FA'
CABIFY_ACCENT = '#B89EF7'
WHITE         = '#FFFFFF'

st.markdown(f"""
    <style>
    .stApp {{ background-color: #F4F6F9; }}
    div.stButton > button {{
        background: linear-gradient(90deg, {CABIFY_PURPLE} 0%, #5633A8 100%);
        color: white; border: none; padding: 15px; border-radius: 10px; width: 100%; font-weight: bold;
    }}
    .header {{ text-align: center; color: #4A4A4A; padding: 20px; }}
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# LÓGICA DE AGRUPACIÓN (AQUÍ ESTÁ LA MAGIA DE LOS DETALLES)
# ==============================================================================

def generar_excel_jerarquico(df, col_g1, col_g2, cols_sum, expandir_todo):
    """
    Genera un Excel donde CADA FILA DE DETALLE SE CONSERVA.
    Se insertan filas de totales entre medio.
    """
    output = io.BytesIO()
    
    # 1. PREPARAR DATOS
    # Ordenamos para que los grupos queden juntos
    df = df.sort_values(by=[col_g1, col_g2])
    
    # Identificar columnas que no son ni grupos ni sumas (texto descriptivo)
    cols_extra = [c for c in df.columns if c not in cols_sum and c not in [col_g1, col_g2]]
    
    # Orden final de columnas en el Excel
    cols_export = [col_g1, col_g2] + cols_extra + cols_sum
    
    # INICIO DE ESCRITURA CON ENGINE XLSXWRITER
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Reporte Detallado")
        
        # --- ESTILOS ---
        # Estilo Cabecera
        fmt_header = workbook.add_format({
            'bold': True, 'fg_color': CABIFY_PURPLE, 'font_color': WHITE, 
            'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        
        # Estilo Nivel 1 (Total Grupo Mayor - Ej: EMPRESA)
        fmt_total_g1 = workbook.add_format({
            'bold': True, 'bg_color': CABIFY_ACCENT, 'font_color': WHITE, 
            'num_format': '#,##0.00', 'border': 1
        })
        
        # Estilo Nivel 2 (Total Grupo Menor - Ej: TIPO GASTO)
        fmt_total_g2 = workbook.add_format({
            'bold': True, 'bg_color': CABIFY_LIGHT, 'font_color': '#333333', 
            'num_format': '#,##0.00', 'border': 1
        })
        
        # Estilo Detalle (Filas normales)
        fmt_detalle_txt = workbook.add_format({'border': 1})
        fmt_detalle_num = workbook.add_format({'border': 1, 'num_format': '#,##0.00'})

        # --- ESCRIBIR CABECERAS ---
        for i, col in enumerate(cols_export):
            worksheet.write(0, i, col, fmt_header)
        
        # Ajustar anchos
        worksheet.set_column(0, 1, 25) # Grupos anchos
        worksheet.set_column(2, len(cols_export)-1, 15) # Resto normal

        # Índices de columnas numéricas para saber dónde aplicar formato moneda
        indices_num = [cols_export.index(c) for c in cols_sum]

        # --- ITERACIÓN DE ESCRITURA FILA POR FILA ---
        current_row = 1 # Empezamos en la fila 1 (la 0 es header)

        # 1. Loop Grupo Mayor (Ej: Empresa A, Empresa B...)
        for nombre_g1, df_g1 in df.groupby(col_g1, sort=False):
            
            # 2. Loop Grupo Menor (Ej: Gasto Viaje, Gasto Nomina...)
            for nombre_g2, df_g2 in df_g1.groupby(col_g2, sort=False):
                
                # A. ESCRIBIR LOS DETALLES (LAS FILAS REALES)
                # ---------------------------------------------------------
                # Aquí recorremos cada fila original del Excel subido
                for _, row_data in df_g2.iterrows():
                    
                    # Definimos el nivel de agrupamiento (Level 2 = Detalle más profundo)
                    # Si 'expandir_todo' es False, estas filas estarán ocultas (hidden=True)
                    worksheet.set_row(current_row, None, None, {'level': 2, 'hidden': not expandir_todo})
                    
                    for col_idx, col_name in enumerate(cols_export):
                        valor = row_data[col_name]
                        # Usar formato moneda si es columna numérica, sino texto normal
                        estilo = fmt_detalle_num if col_idx in indices_num else fmt_detalle_txt
                        worksheet.write(current_row, col_idx, valor, estilo)
                    
                    current_row += 1 # Avanzamos cursor
                
                # B. ESCRIBIR SUBTOTAL NIVEL 2 (Total del subgrupo)
                # ---------------------------------------------------------
                worksheet.set_row(current_row, None, None, {'level': 1, 'hidden': False, 'collapsed': not expandir_todo})
                
                # Escribimos etiquetas
                worksheet.write(current_row, 0, nombre_g1, fmt_total_g2) # Repetimos nombre grupo 1
                worksheet.write(current_row, 1, f"TOTAL {nombre_g2}", fmt_total_g2)
                
                # Rellenamos vacíos en medio con formato gris
                for k in range(2, len(cols_export)):
                    worksheet.write(current_row, k, "", fmt_total_g2)
                
                # Escribimos las sumas
                for col_sum in cols_sum:
                    idx = cols_export.index(col_sum)
                    suma = df_g2[col_sum].sum()
                    worksheet.write(current_row, idx, suma, fmt_total_g2)
                
                current_row += 1 # Avanzamos cursor

            # C. ESCRIBIR TOTAL NIVEL 1 (Total del Grupo Mayor)
            # ---------------------------------------------------------
            worksheet.set_row(current_row, None, None, {'level': 0, 'collapsed': False})
            
            worksheet.write(current_row, 0, f"TOTAL {nombre_g1}", fmt_total_g1)
            
            # Rellenamos vacíos
            for k in range(1, len(cols_export)):
                worksheet.write(current_row, k, "", fmt_total_g1)
                
            # Escribimos las sumas mayores
            for col_sum in cols_sum:
                idx = cols_export.index(col_sum)
                suma = df_g1[col_sum].sum()
                worksheet.write(current_row, idx, suma, fmt_total_g1)
            
            current_row += 1 # Avanzamos cursor

        # D. GRAN TOTAL FINAL
        # ---------------------------------------------------------
        worksheet.write(current_row, 0, "GRAN TOTAL GLOBAL", fmt_header)
        for k in range(1, len(cols_export)):
             worksheet.write(current_row, k, "", fmt_header)
             
        for col_sum in cols_sum:
            idx = cols_export.index(col_sum)
            suma_total = df[col_sum].sum()
            worksheet.write(current_row, idx, suma_total, fmt_header)

        # Color de la pestaña
        worksheet.set_tab_color(CABIFY_PURPLE)

    return output.getvalue()

# ==============================================================================
# INTERFAZ DE USUARIO (APP)
# ==============================================================================

st.markdown("<div class='header'><h1>🗂️ Agrupador de Excel con Detalles</h1><p>Organiza tus datos manteniendo visible cada fila.</p></div>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Sube tu archivo Excel (.xlsx)", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.success(f"✅ Archivo cargado: {len(df)} filas encontradas.")
        st.markdown("---")

        col1, col2, col3 = st.columns(3)
        
        # Detectar columnas numéricas automáticamente
        cols_numericas = df.select_dtypes(include=['float', 'int']).columns.tolist()
        cols_todas = df.columns.tolist()

        with col1:
            st.markdown("### 1. Agrupación Mayor")
            g1 = st.selectbox("Ej: Empresa, Pais", cols_todas, index=0)
            
        with col2:
            st.markdown("### 2. Agrupación Menor")
            # Intentar seleccionar la segunda columna por defecto
            idx_def = 1 if len(cols_todas) > 1 else 0
            g2 = st.selectbox("Ej: Categoría, Cuenta, Vendedor", cols_todas, index=idx_def)
            
        with col3:
            st.markdown("### 3. Valores a Sumar")
            c_sum = st.multiselect("Ej: Total, Saldo, Cantidad", cols_todas, default=cols_numericas)

        st.markdown("### 4. Opciones de Visualización")
        expandir = st.checkbox("📂 ¿Descargar con todos los detalles abiertos?", value=False, help="Si no marcas esto, el Excel vendrá con los grupos cerrados y tendrás que darle al (+) para ver los detalles.")

        st.markdown("###")
        
        if st.button("🚀 GENERAR REPORTE AGRUPADO"):
            if not c_sum:
                st.error("⚠️ Debes seleccionar al menos una columna numérica para sumar.")
            elif g1 == g2:
                st.warning("⚠️ El Grupo Mayor y Menor son iguales. Por favor elige columnas distintas.")
            else:
                with st.spinner("Creando estructura y escribiendo filas..."):
                    
                    excel_data = generar_excel_jerarquico(df, g1, g2, c_sum, expandir)
                    
                    st.success("¡Reporte listo!")
                    st.download_button(
                        label="📥 DESCARGAR EXCEL (CON DETALLES)",
                        data=excel_data,
                        file_name="Reporte_Agrupado_Detallado.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    except Exception as e:
        st.error(f"❌ Error al leer el archivo: {e}")

else:
    st.info("👆 Carga un archivo Excel para ver las opciones.")
