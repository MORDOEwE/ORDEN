import streamlit as st
import pandas as pd
import io
import xlsxwriter

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Organizador Excel", page_icon="📑", layout="wide")

# --- ESTILOS VISUALES (CABIFY THEME) ---
CABIFY_PURPLE = '#7145D6'
CABIFY_LIGHT  = '#F3F0FA'
CABIFY_ACCENT = '#B89EF7'
WHITE         = '#FFFFFF'

st.markdown(f"""
    <style>
    .stApp {{ background-color: #F4F6F9; }}
    div.stButton > button {{
        background: linear-gradient(90deg, {CABIFY_PURPLE} 0%, #5633A8 100%);
        color: white; border: none; padding: 12px 24px; border-radius: 8px; width: 100%;
    }}
    </style>
""", unsafe_allow_html=True)

# ==========================================
# LÓGICA DE PROCESAMIENTO (MOTOR)
# ==========================================

def procesar_excel_agrupado(df, col_g1, col_g2, cols_sum):
    """Ordena, calcula subtotales y genera la estructura para el Excel."""
    if df.empty: return df
    
    # 1. Ordenar datos para que la agrupación funcione
    df_sorted = df.sort_values(by=[col_g1, col_g2]).copy()
    
    # Definir columnas a mostrar (Grupos + Texto restante + Sumas)
    cols_texto = [c for c in df.columns if c not in cols_sum and c not in [col_g1, col_g2]]
    cols_finales = [col_g1, col_g2] + cols_texto + cols_sum
    
    rows_buffer = []

    # 2. Iterar por Grupo Mayor (Nivel 1)
    for nombre_g1, df_g1 in df_sorted.groupby(col_g1, sort=False):
        
        # 3. Iterar por Grupo Menor (Nivel 2)
        for nombre_g2, df_g2 in df_g1.groupby(col_g2, sort=False):
            
            # A. Filas de Detalle
            temp = df_g2[cols_finales].copy()
            temp['__META__'] = 'DETALLE'
            rows_buffer.append(temp)
            
            # B. Subtotal Nivel 2
            sub = pd.Series(index=cols_finales).fillna('')
            sub[col_g1] = nombre_g1 # Mantener el padre visible
            sub[col_g2] = f"TOTAL {str(nombre_g2).upper()}"
            sub['__META__'] = 'SUBTOTAL_N2'
            for c in cols_sum: sub[c] = df_g2[c].sum()
            rows_buffer.append(sub.to_frame().T)

        # C. Subtotal Nivel 1
        tot = pd.Series(index=cols_finales).fillna('')
        tot[col_g1] = f"TOTAL {str(nombre_g1).upper()}"
        tot['__META__'] = 'SUBTOTAL_N1'
        for c in cols_sum: tot[c] = df_g1[c].sum()
        rows_buffer.append(tot.to_frame().T)

    # D. Gran Total
    df_fin = pd.concat(rows_buffer, ignore_index=True)
    grand = pd.Series(index=cols_finales).fillna('')
    grand[col_g1] = "GRAN TOTAL GLOBAL"
    grand['__META__'] = 'GRAN_TOTAL'
    for c in cols_sum: grand[c] = df_fin[df_fin['__META__'] == 'DETALLE'][c].sum()
    
    return pd.concat([df_fin, grand.to_frame().T], ignore_index=True), cols_finales

def generar_excel_estilizado(df_final, cols_export, cols_sum, col_g1, col_g2):
    """Escribe el Excel con formato condicional y esquema (Outlining)."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        sheet_name = "Reporte Agrupado"
        # Exportar datos crudos (sin columna META)
        df_export = df_final[cols_export]
        df_export.to_excel(writer, sheet_name=sheet_name, index=False)
        
        wb = writer.book
        ws = writer.sheets[sheet_name]
        
        # --- DEFINICIÓN DE ESTILOS ---
        # Header
        fmt_head = wb.add_format({'bold': True, 'fg_color': CABIFY_PURPLE, 'font_color': WHITE, 'border': 1, 'align': 'center'})
        
        # Nivel 2 (Subtotal Menor) - Gris claro / Lila suave
        fmt_n2_txt = wb.add_format({'bold': True, 'bg_color': CABIFY_LIGHT, 'border': 1})
        fmt_n2_num = wb.add_format({'bold': True, 'bg_color': CABIFY_LIGHT, 'num_format': '#,##0.00', 'border': 1})
        
        # Nivel 1 (Subtotal Mayor) - Lila Intenso
        fmt_n1_txt = wb.add_format({'bold': True, 'bg_color': CABIFY_ACCENT, 'font_color': WHITE, 'border': 1})
        fmt_n1_num = wb.add_format({'bold': True, 'bg_color': CABIFY_ACCENT, 'font_color': WHITE, 'num_format': '#,##0.00', 'border': 1})
        
        # Gran Total - Morado Cabify
        fmt_tot_txt = wb.add_format({'bold': True, 'bg_color': CABIFY_PURPLE, 'font_color': WHITE, 'border': 1})
        fmt_tot_num = wb.add_format({'bold': True, 'bg_color': CABIFY_PURPLE, 'font_color': WHITE, 'num_format': '#,##0.00', 'border': 1})

        # Aplicar formato a cabeceras
        for i, col in enumerate(cols_export):
            ws.write(0, i, col, fmt_head)
        
        # Ajustar ancho columnas
        ws.set_column(0, 1, 30) # Las dos primeras (agrupadoras) más anchas
        ws.set_column(2, len(cols_export)-1, 15)

        # Índices de columnas numéricas
        idx_num = [df_export.columns.get_loc(c) for c in cols_sum]
        
        # --- BUCLE DE FORMATO POR FILA ---
        for i, row in df_final.iterrows():
            r = i + 1 # Excel index (1-based offset for header)
            meta = row['__META__']
            data = row[cols_export]
            
            if meta == 'DETALLE':
                # Nivel 2: Oculto por defecto
                ws.set_row(r, None, None, {'level': 2, 'hidden': True})
                # Formato simple moneda para detalles
                for c_idx in idx_num:
                    ws.write_number(r, c_idx, data.iloc[c_idx], wb.add_format({'num_format': '#,##0.00'}))

            elif meta == 'SUBTOTAL_N2':
                # Nivel 1: Visible pero colapsado
                ws.set_row(r, None, None, {'level': 1, 'hidden': False, 'collapsed': True})
                _pintar_fila(ws, r, data, idx_num, fmt_n2_txt, fmt_n2_num)

            elif meta == 'SUBTOTAL_N1':
                # Nivel 0: Siempre visible
                ws.set_row(r, None, None, {'level': 0, 'collapsed': False})
                _pintar_fila(ws, r, data, idx_num, fmt_n1_txt, fmt_n1_num)

            elif meta == 'GRAN_TOTAL':
                _pintar_fila(ws, r, data, idx_num, fmt_tot_txt, fmt_tot_num)
                
        ws.set_tab_color(CABIFY_PURPLE)
        
    return output.getvalue()

def _pintar_fila(ws, r, data, idxs, ft, fn):
    for c, val in enumerate(data):
        ws.write(r, c, val, fn if c in idxs else ft)

# ==========================================
# INTERFAZ DE USUARIO (FRONTEND)
# ==========================================

st.title("📑 Organizador de Excel Inteligente")
st.markdown("Sube cualquier archivo, elige cómo agruparlo y descárgalo formateado.")

uploaded_file = st.file_uploader("Sube tu archivo Excel (.xlsx)", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Leer archivo
        df = pd.read_excel(uploaded_file)
        st.success("✅ Archivo cargado correctamente")
        
        st.markdown("### ⚙️ Configuración del Reporte")
        
        col1, col2, col3 = st.columns(3)
        
        # Selección de Columnas
        all_cols = df.columns.tolist()
        num_cols = df.select_dtypes(include=['float', 'int']).columns.tolist()
        
        with col1:
            g1 = st.selectbox("1. Agrupar Principal (Nivel 1)", options=all_cols, index=0)
        
        with col2:
            # Intentar seleccionar una diferente por defecto
            idx_g2 = 1 if len(all_cols) > 1 else 0
            g2 = st.selectbox("2. Agrupar Detalle (Nivel 2)", options=all_cols, index=idx_g2)
            
        with col3:
            # Pre-seleccionar numéricas
            cols_sum = st.multiselect("3. Columnas a Sumar", options=all_cols, default=num_cols)

        if st.button("🚀 ORGANIZAR Y DESCARGAR"):
            if g1 == g2:
                st.warning("⚠️ El nivel 1 y nivel 2 son iguales. Se recomienda elegir columnas distintas.")
            
            with st.spinner("Procesando agrupación..."):
                # Ejecutar lógica
                df_procesado, cols_finales = procesar_excel_agrupado(df, g1, g2, cols_sum)
                
                # Generar Excel
                excel_data = generar_excel_estilizado(df_procesado, cols_finales, cols_sum, g1, g2)
                
                st.markdown("###")
                st.download_button(
                    label="📥 DESCARGAR EXCEL ORGANIZADO",
                    data=excel_data,
                    file_name="Reporte_Organizado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.balloons()

    except Exception as e:
        st.error(f"Error leyendo el archivo: {e}")

else:
    st.info("👆 Sube un archivo para comenzar.")
