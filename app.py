import streamlit as st
import pandas as pd
import io
import engine  # Importamos nuestro motor lógico

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Conciliador Fiscal", page_icon="🟣", layout="wide")

# --- CSS PERSONALIZADO ---
st.markdown("""
    <style>
    .stApp { background-color: #F4F6F9; }
    /* Estilos para inputs y alertas */
    div[data-testid="stFileUploader"] { background-color: white; padding: 20px; border-radius: 10px; border: 1px solid #ddd;}
    label p { color: #333 !important; font-weight: bold; }
    div[data-baseweb="notification"] { color: #333 !important; }
    
    /* Botón */
    div.stButton > button {
        background: linear-gradient(90deg, #7145D6 0%, #5633A8 100%);
        color: white; border: none; padding: 15px 32px;
        font-weight: bold; border-radius: 12px; width: 100%;
        transition: transform 0.2s;
    }
    div.stButton > button:hover { transform: scale(1.02); color: white !important; }
    
    /* Header */
    .main-header {
        background: linear-gradient(135deg, #7145D6 0%, #9C7FE4 100%);
        padding: 30px; border-radius: 0 0 20px 20px; text-align: center; color: white; margin-bottom: 20px;
    }
    </style>
    
    <div class="main-header">
        <h1>🟣 Conciliador Fiscal</h1>
        <p>Cruce inteligente: DIAN vs ERP Netsuite</p>
    </div>
""", unsafe_allow_html=True)

# --- LAYOUT ---
col1, col2 = st.columns(2, gap="large")

with col1:
    st.markdown("### 📂 Documentos Fiscales (DIAN)")
    file_dian = st.file_uploader("Cargar Excel DIAN", type=["xlsx", "xls"], key="dian")

with col2:
    st.markdown("### 📊 Documentos Internos (Netsuite)")
    file_cont = st.file_uploader("Cargar Contabilidad Unificada", type=["xlsx", "xls"], key="cont")

# --- PROCESO ---
st.markdown("###")
if st.button("🚀 EJECUTAR CONCILIACIÓN"):
    if not file_dian or not file_cont:
        st.error("⚠️ Faltan archivos obligatorios.")
    else:
        status = st.status("Procesando...", expanded=True)
        try:
            # 1. LECTURA
            status.write("📖 Leyendo archivos...")
            df_dian_raw = engine.leer_dian(file_dian)
            df_dian_raw = engine.crear_llave_conciliacion(df_dian_raw)
            
            df_cont_full = engine.leer_contabilidad_completa(file_cont)
            
            if df_cont_full is None:
                status.update(label="Error en contabilidad", state="error")
                st.stop()

            # 2. FILTRADO
            status.write("🔍 Aplicando filtros inteligentes...")
            
            # Gastos
            df_dian_gastos = engine.filtrar_dian_gastos(df_dian_raw)
            df_cont_gastos = engine.filtrar_solo_gastos(df_cont_full)
            
            # Ingresos
            df_dian_ingresos = engine.filtrar_dian_ingresos(df_dian_raw)
            df_cont_ingresos = engine.filtrar_solo_ingresos(df_cont_full)

            # 3. CRUCE (MATCHING)
            status.write("⚙️ Cruzando bases de datos...")
            # Cruce Gastos
            cg, sdg, scg = engine.ejecutar_conciliacion_universal(df_dian_gastos, df_cont_gastos)
            # Cruce Ingresos
            ci, sdi, sci = engine.ejecutar_conciliacion_universal(df_dian_ingresos, df_cont_ingresos)

            # 4. PREPARACIÓN DE DATOS UNIFICADOS
            status.write("📊 Unificando datos para reporte...")
            
            # Detectar columnas dinámicas de la DIAN para mapeo
            cols_dian = df_dian_raw.columns
            mapa_cols = {
                'total': next((c for c in cols_dian if 'total' in c), 'total'),
                'emisor': next((c for c in cols_dian if 'nombre_emisor' in c), 'emisor'),
                'iva': next((c for c in cols_dian if 'iva' in c or 'impuesto' in c), None)
            }
            
            df_final_gastos = engine.preparar_datos_unificados(cg, sdg, scg, mapa_cols)
            # Para ingresos invertimos el mapa (el emisor soy yo, me importa el receptor)
            mapa_ing = mapa_cols.copy()
            mapa_ing['emisor'] = next((c for c in cols_dian if 'nombre_receptor' in c), 'receptor')
            df_final_ingresos = engine.preparar_datos_unificados(ci, sdi, sci, mapa_ing)

            # 5. GENERACIÓN EXCEL
            status.write("📝 Generando Excel con agrupación dinámica...")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                
                # CONFIGURACIÓN DINÁMICA DE COLUMNAS
                config_reporte = {
                    'col_grupo_1': 'EMPRESA_GRUPO',
                    'col_grupo_2': 'TIPO',
                    'cols_suma': ['VALOR_DIAN', 'VALOR_CONT', 'DIFERENCIA'],
                    'cols_texto': ['NIT', 'EMPRESA', 'LLAVE', 'CUENTA']
                }
                
                engine.generar_reporte_agrupado(writer, df_final_gastos, '1. Gastos', config_reporte)
                engine.generar_reporte_agrupado(writer, df_final_ingresos, '2. Ingresos', config_reporte)
                
                # Bases puras
                engine.formatear_hoja_base(writer, 'Base DIAN', df_dian_raw)
                df_dian_raw.to_excel(writer, sheet_name='Base DIAN', index=False)
                
                engine.formatear_hoja_base(writer, 'Base Contable', df_cont_full)
                df_cont_full.to_excel(writer, sheet_name='Base Contable', index=False)

            status.update(label="✅ ¡Proceso completado!", state="complete", expanded=False)
            
            st.success("Reporte generado exitosamente.")
            st.download_button(
                label="📥 DESCARGAR REPORTE CONCILIADO",
                data=output.getvalue(),
                file_name="Conciliacion_Fiscal_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            status.update(label="❌ Ocurrió un error", state="error")
            st.error(f"Detalle del error: {e}")
