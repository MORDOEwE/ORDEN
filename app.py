import streamlit as st
import pandas as pd
import io
import xlsxwriter

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Agrupador Pro (ERP)", page_icon="🧹", layout="wide")

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
    .warning-box {{ background-color: #FFF3CD; padding: 10px; border-radius: 5px; border: 1px solid #FFEEBA; color: #856404; }}
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# LÓGICA DE LIMPIEZA (NUEVO)
# ==============================================================================

def limpiar_datos_erp(df, col_agrupacion):
    """
    1. Rellena los huecos vacíos hacia abajo (para ERPs que dejan celdas en blanco).
    2. Elimina las filas que ya son totales en el archivo original.
    """
    df_clean = df.copy()
    
    # 1. Rellenar hacia abajo (Forward Fill)
    # Esto sirve cuando el ERP pone el nombre de la cuenta solo en la primera fila del grupo
    df_clean[col_agrupacion] = df_clean[col_agrupacion].ffill()
    
    # 2. Eliminar filas basura (Totales nativos del ERP)
    # Buscamos filas donde la columna de agrupación empiece por "Total" o "Saldo"
    mascara_totales = df_clean[col_agrupacion].astype(str).str.contains(r'^(TOTAL|Total|Sum|Saldo)', na=False, regex=True)
    
    # Invertimos la máscara para quedarnos con lo que NO es total
    df_clean = df_clean[~mascara_totales]
    
    return df_clean

# ==============================================================================
# LÓGICA DE AGRUPACIÓN Y EXCEL
# ==============================================================================

def generar_excel_jerarquico(df, col_g1, col_g2, cols_sum, expandir_todo):
    output = io.BytesIO()
    
    # Aseguramos que no haya NaNs en las columnas de agrupación para evitar errores
    df[col_g1] = df[col_g1].fillna("SIN CLASIFICAR")
    df[col_g2] = df[col_g2].fillna("SIN CLASIFICAR")

    df = df.sort_values(by=[col_g1, col_g2])
    
    cols_extra = [c for c in df.columns if c not in cols_sum and c not in [col_g1, col_g2]]
    cols_export = [col_g1, col_g2] + cols_extra + cols_sum
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Reporte Detallado")
        
        # Estilos
        fmt_header = workbook.add_format({'bold': True, 'fg_color': CABIFY_PURPLE, 'font_color': WHITE, 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        fmt_total_g1 = workbook.add_format({'bold': True, 'bg_color': CABIFY_ACCENT, 'font_color': WHITE, 'num_format': '#,##0.00', 'border': 1})
        fmt_total_g2 = workbook.add_format({'bold': True, 'bg_color': CABIFY_LIGHT, 'font_color': '#333333', 'num_format': '#,##0.00', 'border': 1})
        fmt_detalle_txt = workbook.add_format({'border': 1})
        fmt_detalle_num = workbook.add_format({'border': 1, 'num_format': '#,##0.00'})

        # Headers
        for i, col in enumerate(cols_export):
            worksheet.write(0, i, col, fmt_header)
        
        worksheet.set_column(0, 1, 30)
        worksheet.set_column(2, len(cols_export)-1, 15)

        indices_num = [cols_export.index(c) for c in cols_sum]
        current_row = 1

        # LOGICA DE ESCRITURA
        for nombre_g1, df_g1 in df.groupby(col_g1, sort=False):
            for nombre_g2, df_g2 in df_g1.groupby(col_g2, sort=False):
                
                # A. DETALLES
                for _, row_data in df_g2.iterrows():
                    worksheet.set_row(current_row, None, None, {'level': 2, 'hidden': not expandir_todo})
                    for col_idx, col_name in enumerate(cols_export):
                        valor = row_data[col_name]
                        estilo = fmt_detalle_num if col_idx in indices_num else fmt_detalle_txt
                        worksheet.write(current_row, col_idx, valor, estilo)
                    current_row += 1 
                
                # B. SUBTOTAL G2
                worksheet.set_row(current_row, None, None, {'level': 1, 'hidden': False, 'collapsed': not expandir_todo})
                worksheet.write(current_row, 0, nombre_g1, fmt_total_g2)
                worksheet.write(current_row, 1, f"TOTAL {str(nombre_g2)}", fmt_total_g2)
                for k in range(2, len(cols_export)): worksheet.write(current_row, k, "", fmt_total_g2)
                for col_sum in cols_sum:
                    idx = cols_export.index(col_sum)
                    worksheet.write(current_row, idx, df_g2[col_sum].sum(), fmt_total_g2)
                current_row += 1

            # C. SUBTOTAL G1
            worksheet.set_row(current_row, None, None, {'level': 0, 'collapsed': False})
            worksheet.write(current_row, 0, f"TOTAL {str(nombre_g1)}", fmt_total_g1)
            for k in range(1, len(cols_export)): worksheet.write(current_row, k, "", fmt_total_g1)
            for col_sum in cols_sum:
                idx = cols_export.index(col_sum)
                worksheet.write(current_row, idx, df_g1[col_sum].sum(), fmt_total_g1)
            current_row += 1

        # GRAN TOTAL
        worksheet.write(current_row, 0, "GRAN TOTAL", fmt_header)
        for k in range(1, len(cols_export)): worksheet.write(current_row, k, "", fmt_header)
        for col_sum in cols_sum:
            idx = cols_export.index(col_sum)
            worksheet.write(current_row, idx, df[col_sum].sum(), fmt_header)

        worksheet.set_tab_color(CABIFY_PURPLE)

    return output.getvalue()

# ==============================================================================
# INTERFAZ DE USUARIO
# ==============================================================================

st.markdown("<div class='header'><h1>🧹 Agrupador Inteligente (Modo ERP)</h1><p>Limpia totales basura, rellena espacios y agrupa correctamente.</p></div>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Sube tu archivo Excel (.xlsx)", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Cargar ignorando las primeras filas si es necesario (header=0 usualmente está bien)
        df_raw = pd.read_excel(uploaded_file)
        
        st.info(f"📂 Archivo cargado con {len(df_raw)} filas.")

        col1, col2 = st.columns(2)
        
        cols_todas = df_raw.columns.tolist()
        cols_numericas = df_raw.select_dtypes(include=['float', 'int']).columns.tolist()

        with col1:
            st.subheader("1. Configuración de Grupos")
            g1 = st.selectbox("Agrupación Principal (Nivel 1)", cols_todas, index=0, help="Ej: Cuenta o Empresa")
            g2 = st.selectbox("Agrupación Secundaria (Nivel 2)", cols_todas, index=1 if len(cols_todas)>1 else 0, help="Ej: Tercero o Detalle")
            
        with col2:
            st.subheader("2. Valores y Limpieza")
            c_sum = st.multiselect("Columnas a Sumar", cols_todas, default=cols_numericas)
            
            st.markdown("---")
            st.markdown("**🧹 Opciones de Limpieza:**")
            usar_limpieza = st.checkbox("Activar Limpieza ERP (Recomendado)", value=True, help="Elimina filas que dicen 'Total' y rellena celdas vacías hacia abajo.")
            expandir = st.checkbox("Descargar Expandido", value=False, help="Ver todos los detalles abiertos por defecto.")

        if st.button("🚀 PROCESAR ARCHIVO"):
            if g1 == g2:
                st.error("⚠️ Elige columnas diferentes para Grupo 1 y 2.")
            elif not c_sum:
                st.error("⚠️ Elige qué columnas sumar.")
            else:
                with st.spinner("Limpiando y reestructurando datos..."):
                    
                    df_procesado = df_raw.copy()
                    
                    # --- APLICAR LIMPIEZA SI ESTÁ MARCADO ---
                    if usar_limpieza:
                        filas_antes = len(df_procesado)
                        # 1. Rellenar huecos en la columna principal (ej. Cuenta)
                        df_procesado = limpiar_datos_erp(df_procesado, g1)
                        filas_despues = len(df_procesado)
                        st.caption(f"✅ Se eliminaron {filas_antes - filas_despues} filas de totales 'basura' del archivo original.")
                    
                    # Generar Excel
                    excel_data = generar_excel_jerarquico(df_procesado, g1, g2, c_sum, expandir)
                    
                    st.success("¡Archivo transformado correctamente!")
                    st.download_button(
                        label="📥 DESCARGAR REPORTE LIMPIO Y AGRUPADO",
                        data=excel_data,
                        file_name="Reporte_Agrupado_Clean.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    except Exception as e:
        st.error(f"❌ Error: {e}")
