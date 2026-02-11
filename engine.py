import pandas as pd
import numpy as np
import re
import streamlit as st

# --- CONSTANTES VISUALES ---
CABIFY_PURPLE = '#7145D6'
CABIFY_LIGHT  = '#F3F0FA'
CABIFY_ACCENT = '#B89EF7'
WHITE         = '#FFFFFF'

# =================================================================
# 1. UTILIDADES DE LIMPIEZA
# =================================================================

def normalize_col_name(col_name):
    """Normaliza nombres de columnas (quita espacios, caracteres raros, minusculas)."""
    return re.sub(r'[^\w]+', '_', str(col_name)).lower().strip('_')

def clean_nit_numeric(nit_series):
    if nit_series is None or nit_series.empty:
        return pd.Series([''] * len(nit_series), dtype=str)
    return nit_series.astype(str).str.replace(r'[^0-9]+', '', regex=True)

def standardize_company_name(name_series):
    if name_series is None or name_series.empty:
        return pd.Series(['SIN NOMBRE'] * len(name_series), dtype=str)
    return (
        name_series.astype(str)
        .str.upper()
        .str.replace(r'[^A-Z0-9\s]+', '', regex=True)
        .str.replace(r'\s+', ' ', regex=True)
        .str.strip()
        .str.replace(r'\b(S A S|SAS|S A|SA|LTDA|LTDA|BIC|B I C)\b', '', regex=True) 
        .str.strip()
    )

def limpiar_moneda_colombia(valor):
    if pd.isna(valor) or str(valor).strip() == '': return 0.0
    s = str(valor).replace('$', '').replace(' ', '')
    try:
        if '.' in s and ',' in s:
            if s.rfind(',') > s.rfind('.'): s = s.replace('.', '').replace(',', '.')
            else: s = s.replace(',', '')
        elif ',' in s: 
            parts = s.split(',')
            if len(parts[-1]) == 2: s = s.replace(',', '.') 
            else: s = s.replace(',', '') 
        return float(s)
    except: return 0.0

# =================================================================
# 2. LECTURA DE ARCHIVOS (CACHÉ + CALAMINE)
# =================================================================

@st.cache_data(ttl=3600, show_spinner=False)
def leer_dian(file_obj):
    if file_obj is None: return None
    try:
        # Intenta usar calamine (muy rápido)
        df = pd.read_excel(file_obj, engine="calamine", dtype=str)
    except:
        # Fallback a openpyxl
        file_obj.seek(0)
        df = pd.read_excel(file_obj, dtype=str)
    
    col_map = {col: normalize_col_name(col) for col in df.columns}
    df.rename(columns=col_map, inplace=True)
    return df

@st.cache_data(ttl=3600, show_spinner=False)
def leer_contabilidad_completa(file_obj):
    if file_obj is None: return None
    try:
        file_obj.seek(0)
        try: df_preview = pd.read_excel(file_obj, nrows=20, header=None, engine='openpyxl')
        except: df_preview = pd.read_excel(file_obj, nrows=20, header=None)
        
        header_row = 0
        for i, row in df_preview.iterrows():
            row_str = row.astype(str).values
            if 'Cuenta' in row_str and 'Fecha' in row_str:
                header_row = i; break
        
        file_obj.seek(0)
        df = pd.read_excel(file_obj, header=header_row, engine='openpyxl')
        
        # Limpieza básica Netsuite
        df['Cuenta'] = df['Cuenta'].astype(str).replace(['nan', 'None', ''], np.nan)
        condicion_cabecera = df['Cuenta'].str.match(r'^\d', na=False)
        df['CUENTA_COMPLETA'] = df['Cuenta'].where(condicion_cabecera, other=np.nan).ffill()
        df = df[~df['Cuenta'].astype(str).str.startswith('Total', na=False)] 
        df = df[df['Fecha'].notna()] 
        df['Cuenta'] = df['CUENTA_COMPLETA']
        
        df['CODIGO_CUENTA'] = df['Cuenta'].str.strip().str.extract(r'^(\d+)')
        
        col_deb = next((c for c in df.columns if 'Déb' in c), None)
        col_cred = next((c for c in df.columns if 'Créd' in c), None)
        val_deb = df[col_deb].apply(limpiar_moneda_colombia) if col_deb else 0.0
        val_cred = df[col_cred].apply(limpiar_moneda_colombia) if col_cred else 0.0
        df['SALDO_NETO_CALCULADO'] = val_deb - val_cred
        
        # Renombrar a estándar interno
        col_ref_orig = next((c for c in df.columns if 'mero de doc' in c or 'Nro' in c), 'Número de documento')
        col_nit_orig = next((c for c in df.columns if 'Identifi' in c or 'Nit' in c), 'Número Identificación')
        col_nom_orig = next((c for c in df.columns if 'Nombre' in c), 'Nombre')
        
        df_renamed = df.rename(columns={
            col_ref_orig: 'u_ref', 
            col_nit_orig: 'u_infoco01',
            col_nom_orig: 'u_cardname', 
            'Cuenta': 'u_acctname'
        })
        df_renamed['u_saldo_f'] = df['SALDO_NETO_CALCULADO']
        if 'u_infoco01' in df_renamed.columns:
            df_renamed['u_infoco01'] = df_renamed['u_infoco01'].astype(str).str.replace(r'\.0$', '', regex=True)
            
        return df_renamed
    except Exception as e:
        print(f"Error leyendo contabilidad: {e}")
        return None

# =================================================================
# 3. LÓGICA DE FILTRADO Y CRUCE
# =================================================================

def crear_llave_conciliacion(df):
    cols = df.columns
    prefijo = next((c for c in cols if 'prefijo' in c), None)
    folio = next((c for c in cols if 'folio' in c), None)
    if not prefijo or not folio: return df
    df['LLAVE_DIAN'] = (
        df[prefijo].astype(str).str.strip() + 
        df[folio].astype(str).str.strip()
    ).str.replace(r'[^\w]+', '', regex=True).str.upper()
    return df

def filtrar_solo_gastos(df):
    df = df[df['CODIGO_CUENTA'].str.startswith('5', na=False)]
    return df[~df['u_acctname'].str.contains('DIFERENCIA EN CAMBIO|DEPRECIACI', case=False, na=False)]

def filtrar_solo_ingresos(df):
    df = df[df['CODIGO_CUENTA'].str.startswith('4', na=False)]
    df = df[~df['u_acctname'].str.contains('DIFERENCIA EN CAMBIO', case=False, na=False)]
    df['u_saldo_f'] = df['u_saldo_f'] * -1 # Invertir signo ingresos
    return df

def filtrar_dian_gastos(df):
    col_grupo = next((c for c in df.columns if 'grupo' in c), None)
    if not col_grupo: return df
    return df[df[col_grupo].astype(str).str.lower().str.contains('recibido')].copy()

def filtrar_dian_ingresos(df):
    col_grupo = next((c for c in df.columns if 'grupo' in c), None)
    if not col_grupo: return df
    return df[df[col_grupo].astype(str).str.lower().str.contains('emitido')].copy()

def ejecutar_conciliacion_universal(df_dian, df_cont):
    """Realiza el cruce y devuelve 3 DataFrames: Coincidencias, Sobra DIAN, Sobra Contabilidad"""
    if df_cont.empty or df_dian.empty: return pd.DataFrame(), df_dian, df_cont

    # Crear llaves
    df_cont['LLAVE_CONT'] = df_cont['u_ref'].astype(str).str.strip().str.replace(r'[^\w]+', '', regex=True).str.upper()
    
    # Agrupar contabilidad por documento (puede haber multiples lineas por factura)
    agg_dict = {'u_saldo_f': 'sum', 'u_infoco01': 'first', 'u_cardname': 'first', 'u_acctname': 'first'}
    df_cont_agg = df_cont.groupby('LLAVE_CONT').agg(agg_dict).reset_index()

    if 'LLAVE_DIAN' not in df_dian.columns: return pd.DataFrame(), df_dian, df_cont

    # MERGE (CRUCE)
    # 1. Coincidencias
    df_coinc = pd.merge(df_dian, df_cont_agg, left_on='LLAVE_DIAN', right_on='LLAVE_CONT', how='inner', suffixes=('_DIAN', '_CONT'))
    
    # 2. Sobrante DIAN
    df_left = pd.merge(df_dian, df_cont_agg[['LLAVE_CONT']], left_on='LLAVE_DIAN', right_on='LLAVE_CONT', how='left', indicator=True)
    df_sob_dian = df_left[df_left['_merge'] == 'left_only'].drop(columns=['LLAVE_CONT', '_merge'])
    
    # 3. Sobrante Contabilidad (Expandido a sus lineas originales)
    df_right = pd.merge(df_cont_agg, df_dian[['LLAVE_DIAN']], left_on='LLAVE_CONT', right_on='LLAVE_DIAN', how='left', indicator=True)
    lista_llaves_sobrantes = df_right[df_right['_merge'] == 'left_only']['LLAVE_CONT']
    df_sob_cont = df_cont[df_cont['LLAVE_CONT'].isin(lista_llaves_sobrantes)]
    
    return df_coinc, df_sob_dian, df_sob_cont

# =================================================================
# 4. PREPARACIÓN DE DATOS UNIFICADOS
# =================================================================

def preparar_datos_unificados(coinc, sob_dian, sob_cont, cols_dian_map):
    """
    Toma los 3 resultados del cruce y devuelve UN solo DataFrame estandarizado
    listo para ser procesado por el generador de reportes.
    """
    lista_dfs = []
    
    col_total_dian = cols_dian_map.get('total', 'total_bruto')
    col_emisor_dian = cols_dian_map.get('emisor', 'nombre_emisor')
    col_iva_dian = cols_dian_map.get('iva', 'iva')

    # 1. COINCIDENCIAS
    if not coinc.empty:
        t = coinc.copy()
        t['NIT'] = clean_nit_numeric(t['u_infoco01'])
        t['EMPRESA'] = t[col_emisor_dian] if col_emisor_dian in t.columns else t['u_cardname']
        t['EMPRESA_GRUPO'] = standardize_company_name(t['EMPRESA'])
        
        # Valores
        val_d = pd.to_numeric(t[col_total_dian], errors='coerce').fillna(0)
        # Restar IVA si existe columna en DIAN para obtener subtotal
        if col_iva_dian and col_iva_dian in t.columns:
            val_d -= pd.to_numeric(t[col_iva_dian], errors='coerce').fillna(0)
            
        t['VALOR_DIAN'] = val_d
        t['VALOR_CONT'] = t['u_saldo_f']
        t['DIFERENCIA'] = t['VALOR_DIAN'] - t['VALOR_CONT']
        t['TIPO'] = 'COINCIDENCIA'
        t['LLAVE'] = t['LLAVE_DIAN']
        t['CUENTA'] = t['u_acctname']
        lista_dfs.append(t)

    # 2. SOBRANTE DIAN
    if not sob_dian.empty:
        t = sob_dian.copy()
        col_nit = next((c for c in t.columns if 'nit' in c or 'identificaci' in c), None)
        t['NIT'] = clean_nit_numeric(t[col_nit]) if col_nit else ''
        t['EMPRESA'] = t[col_emisor_dian] if col_emisor_dian in t.columns else 'DESCONOCIDO'
        t['EMPRESA_GRUPO'] = standardize_company_name(t['EMPRESA'])
        
        val_d = pd.to_numeric(t[col_total_dian], errors='coerce').fillna(0)
        if col_iva_dian and col_iva_dian in t.columns:
            val_d -= pd.to_numeric(t[col_iva_dian], errors='coerce').fillna(0)
            
        t['VALOR_DIAN'] = val_d
        t['VALOR_CONT'] = 0
        t['DIFERENCIA'] = val_d
        t['TIPO'] = 'SOBRANTE DIAN'
        t['LLAVE'] = t['LLAVE_DIAN']
        t['CUENTA'] = ''
        lista_dfs.append(t)

    # 3. SOBRANTE CONTABILIDAD
    if not sob_cont.empty:
        t = sob_cont.copy()
        t['NIT'] = clean_nit_numeric(t['u_infoco01'])
        t['EMPRESA'] = t['u_cardname']
        t['EMPRESA_GRUPO'] = standardize_company_name(t['u_cardname'])
        
        t['VALOR_DIAN'] = 0
        t['VALOR_CONT'] = t['u_saldo_f']
        t['DIFERENCIA'] = -t['u_saldo_f']
        t['TIPO'] = 'SOBRANTE CONT'
        t['LLAVE'] = t['LLAVE_CONT'] if 'LLAVE_CONT' in t.columns else t['u_ref']
        t['CUENTA'] = t['u_acctname']
        lista_dfs.append(t)

    if not lista_dfs: return pd.DataFrame()
    
    df_final = pd.concat(lista_dfs, ignore_index=True)
    # Filtrar basura muy pequeña
    df_final = df_final[(df_final['VALOR_DIAN'].abs() > 1) | (df_final['VALOR_CONT'].abs() > 1)]
    return df_final

# =================================================================
# 5. GENERADOR EXCEL DINÁMICO
# =================================================================

def generar_reporte_agrupado(writer, df, sheet_name, config):
    """Genera hoja Excel con agrupación y colores dinámicos."""
    if df.empty: return

    # Extraer Configuración
    g1, g2 = config.get('col_grupo_1'), config.get('col_grupo_2')
    cols_suma = [c for c in config.get('cols_suma', []) if c in df.columns]
    cols_texto = [c for c in config.get('cols_texto', []) if c in df.columns]
    cols_finales = [g1, g2] + cols_texto + cols_suma

    # Procesar Datos
    df_sorted = df.sort_values(by=[g1, g2]).copy()
    rows_buffer = []

    for nombre_g1, df_g1 in df_sorted.groupby(g1, sort=False):
        for nombre_g2, df_g2 in df_g1.groupby(g2, sort=False):
            # Detalle
            temp = df_g2[cols_finales].copy()
            temp['__META__'] = 'DETALLE'
            rows_buffer.append(temp)
            
            # Subtotal N2
            sub = pd.Series(index=cols_finales).fillna('')
            sub[g1] = nombre_g1
            sub[g2] = f"SUBTOTAL {str(nombre_g2).upper()}"
            sub['__META__'] = 'SUBTOTAL_N2'
            for c in cols_suma: sub[c] = df_g2[c].sum()
            rows_buffer.append(sub.to_frame().T)

        # Subtotal N1
        tot = pd.Series(index=cols_finales).fillna('')
        tot[g1] = f"TOTAL {str(nombre_g1).upper()}"
        tot['__META__'] = 'SUBTOTAL_N1'
        for c in cols_suma: tot[c] = df_g1[c].sum()
        rows_buffer.append(tot.to_frame().T)

    # Gran Total
    df_fin = pd.concat(rows_buffer, ignore_index=True)
    grand = pd.Series(index=cols_finales).fillna('')
    grand[g1] = "GRAN TOTAL GLOBAL"
    grand['__META__'] = 'GRAN_TOTAL'
    for c in cols_suma: grand[c] = df_fin[df_fin['__META__'] == 'DETALLE'][c].sum()
    df_fin = pd.concat([df_fin, grand.to_frame().T], ignore_index=True)

    # Escritura Excel
    df_export = df_fin[cols_finales]
    df_export.to_excel(writer, sheet_name=sheet_name, index=False)
    
    wb, ws = writer.book, writer.sheets[sheet_name]
    
    # Estilos
    fmt_head = wb.add_format({'bold': True, 'fg_color': CABIFY_PURPLE, 'font_color': WHITE, 'border': 1})
    fmt_n2_txt = wb.add_format({'bold': True, 'bg_color': CABIFY_LIGHT})
    fmt_n2_num = wb.add_format({'bold': True, 'bg_color': CABIFY_LIGHT, 'num_format': '#,##0.00'})
    fmt_n1_txt = wb.add_format({'bold': True, 'bg_color': CABIFY_ACCENT, 'font_color': WHITE})
    fmt_n1_num = wb.add_format({'bold': True, 'bg_color': CABIFY_ACCENT, 'font_color': WHITE, 'num_format': '#,##0.00'})
    fmt_tot_txt = wb.add_format({'bold': True, 'bg_color': CABIFY_PURPLE, 'font_color': WHITE})
    fmt_tot_num = wb.add_format({'bold': True, 'bg_color': CABIFY_PURPLE, 'font_color': WHITE, 'num_format': '#,##0.00'})

    # Header
    for i, col in enumerate(df_export.columns): ws.write(0, i, col, fmt_head)
    ws.set_column(0, len(cols_finales)-1, 18)
    
    idx_num = [df_export.columns.get_loc(c) for c in cols_suma]

    for i, row in df_fin.iterrows():
        r = i + 1
        meta = row['__META__']
        data = row[cols_finales]
        
        if meta == 'DETALLE':
            ws.set_row(r, None, None, {'level': 2, 'hidden': True})
            for c_idx in idx_num: ws.write_number(r, c_idx, data.iloc[c_idx], wb.add_format({'num_format': '#,##0.00'}))
        elif meta == 'SUBTOTAL_N2':
            ws.set_row(r, None, None, {'level': 1, 'hidden': False, 'collapsed': True})
            _pintar(ws, r, data, idx_num, fmt_n2_txt, fmt_n2_num)
        elif meta == 'SUBTOTAL_N1':
            ws.set_row(r, None, None, {'level': 0, 'collapsed': False})
            _pintar(ws, r, data, idx_num, fmt_n1_txt, fmt_n1_num)
        elif meta == 'GRAN_TOTAL':
            _pintar(ws, r, data, idx_num, fmt_tot_txt, fmt_tot_num)
            
    ws.set_tab_color(CABIFY_PURPLE)

def _pintar(ws, r, data, idxs, ft, fn):
    for c, val in enumerate(data):
        ws.write(r, c, val, fn if c in idxs else ft)

def formatear_hoja_base(writer, sheet_name, df):
    if df.empty: return
    ws = writer.sheets[sheet_name]
    ws.set_tab_color('gray')
    fmt = writer.book.add_format({'bold': True, 'bg_color': '#DDDDDD', 'border': 1})
    for i, col in enumerate(df.columns): ws.write(0, i, col, fmt)
