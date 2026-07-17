import streamlit as st
import pandas as pd
import io

# Configuración visual de la aplicación
st.set_page_config(page_title="Informe Concentrado Semanal", page_icon="📋", layout="wide")

st.title("📋 Informe Concentrado Semanal - Tienda Curicó")
st.markdown("""
### Proceso de Carga y Consolidación de Datos:
1. **Carga los tres archivos independientes** (AX365, Stock de Máquina y Stock Físico).
2. El sistema filtrará automáticamente los concentrados *MCI TINTER, TWIST y Transparentes*.
3. El sistema **ordenará automáticamente los lotes por fecha de ingreso** (de la más antigua a la más nueva).
4. Al procesar, el sistema analizará qué lote se consumió según su orden cronológico de entrada en AX.
""")

# --- SECCIÓN DE CARGA DE ARCHIVOS INDEPENDIENTES ---
st.subheader("📁 1. Accesos para Carga de Datos")
col1, col2, col3 = st.columns(3)

with col1:
    archivo_ax = st.file_uploader("📥 Datos de AX365 (.xlsx)", type=["xlsx"], help="Base principal con códigos, lotes y fechas del ERP")
with col2:
    archivo_maquina = st.file_uploader("📥 Stock de la Máquina (.xlsx, .csv)", type=["xlsx", "csv"])
with col3:
    archivo_bodega = st.file_uploader("📥 Stock Físico Bodega (.xlsx, .csv)", type=["xlsx", "csv"])

if archivo_ax is not None:
    try:
        excel_book = pd.ExcelFile(archivo_ax)
        hojas = excel_book.sheet_names
        hoja_seleccionada = st.selectbox("Selecciona la pestaña del reporte AX365:", hojas, index=0)
        
        # Leer la planilla original de AX365
        df_original = pd.read_excel(archivo_ax, sheet_name=hoja_seleccionada)
        df_original.columns = df_original.columns.str.strip()
        
        # --- DETECCIÓN AUTOMÁTICA DE COLUMNAS ---
        def buscar_columna(df, palabras_clave):
            for col in df.columns:
                if any(p in str(col).lower() for p in palabras_clave):
                    return col
            return None

        col_codigo = buscar_columna(df_original, ['código ax', 'codigo ax', 'artículo', 'articulo', 'item'])
        col_concentrado = buscar_columna(df_original, ['concentrado', 'nombre', 'producto'])
        col_lote = buscar_columna(df_original, ['lotes consumidos', 'lote', 'batch', 'número de lote'])
        col_fecha = buscar_columna(df_original, ['fecha', 'ingreso', 'creación', 'creacion', 'registro', 'f. ingreso'])
        
        if not col_codigo:
            st.error("🚨 Error: No se encontró la columna de 'Código AX' en el archivo de AX365.")
            st.stop()
            
        if not col_concentrado:
            st.error("🚨 Error: No se encontró la columna de descripción del producto ('Concentrado') en AX365.")
            st.stop()

        # Limpieza inicial de la base de datos
        df_limpio = df_original.dropna(subset=[col_codigo, col_concentrado]).copy()
        df_limpio[col_codigo] = df_limpio[col_codigo].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        
        # --- PROCESAMIENTO DE FECHAS ---
        if col_fecha:
            df_limpio[col_fecha] = pd.to_datetime(df_limpio[col_fecha], errors='coerce')
            df_limpio = df_limpio.sort_values(by=[col_fecha], ascending=True, na_position='last')
            st.toast("📅 Lotes ordenados por fecha de ingreso correctamente.", icon="⏳")
        else:
            st.warning("⚠️ No se detectó una columna de 'Fecha' o 'Ingreso'. El sistema usará el orden original del archivo.")
        
        # --- FILTRO ESTRICTO DE CONCENTRADOS SOLICITADOS ---
        df_limpio['Concentrado_Minuscula'] = df_limpio[col_concentrado].astype(str).str.lower().str.strip()
        
        condicion_mci = df_limpio['Concentrado_Minuscula'].str.contains('mci tinter', na=False)
        condicion_twist = df_limpio['Concentrado_Minuscula'].str.contains('twist', na=False)
        
        condicion_rojo_ox = (df_limpio['Concentrado_Minuscula'].str.contains('transparente rojo', na=False) | 
                             (df_limpio['Concentrado_Minuscula'].str.contains('transparente', na=False) & df_limpio['Concentrado_Minuscula'].str.contains('rojo', na=False)))
                             
        condicion_amarillo_ox = (df_limpio['Concentrado_Minuscula'].str.contains('transparente amarillo', na=False) | 
                                 (df_limpio['Concentrado_Minuscula'].str.contains('transparente', na=False) & df_limpio['Concentrado_Minuscula'].str.contains('amarillo', na=False)))
        
        df_filtrado = df_limpio[condicion_mci | condicion_twist | condicion_rojo_ox | condicion_amarillo_ox].copy()
        df_filtrado = df_filtrado.drop(columns=['Concentrado_Minuscula'])
        
        if df_filtrado.empty:
            st.warning("⚠️ No se encontraron concentrados válidos en la pestaña seleccionada.")
            st.stop()
        
        # --- PROCESAMIENTO DE CARGAS DE STOCK ---
        dict_maquina = {}
        dict_bodega = {}
        
        def procesar_archivo_externo(archivo):
            if archivo.name.endswith('.csv'):
                df_ext = pd.read_csv(archivo)
            else:
                df_ext = pd.read_excel(archivo)
            df_ext.columns = df_ext.columns.str.strip()
            
            c_cod = buscar_columna(df_ext, ['código ax', 'codigo ax', 'artículo', 'articulo', 'item'])
            c_cant = buscar_columna(df_ext, ['cantidad', 'stock', 'físico', 'fisico', 'total', 'unidades'])
            
            if c_cod and c_cant:
                df_ext[c_cod] = df_ext[c_cod].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                df_ext[c_cant] = pd.to_numeric(df_ext[c_cant], errors='coerce').fillna(0)
                return df_ext.groupby(c_cod)[c_cant].sum().to_dict()
            return {}

        if archivo_maquina is not None:
            dict_maquina = procesar_archivo_externo(archivo_maquina)
            st.toast("✅ Datos de la Máquina vinculados", icon="⚙️")
            
        if archivo_bodega is not None:
            dict_bodega = procesar_archivo_externo(archivo_bodega)
            st.toast("✅ Datos de Bodega vinculados", icon="📦")

        # --- CRUCE AUTOMÁTICO DE DATOS ---
        df_filtrado['Inventario Físico Bodega (A mano)'] = df_filtrado[col_codigo].map(dict_bodega).fillna(0.0)
        df_filtrado['Inventario Máquina (A mano)'] = df_filtrado[col_codigo].map(dict_maquina).fillna(0.0)
        df_filtrado['Consumo Registrado Semanal'] = 0.0

        # --- MÓDULO INTERACTIVO DE PANTALLA ---
        st.subheader(f"📝 2. Módulo de Validación y Ajustes ({len(df_filtrado)} filas ordenadas por antigüedad)")
        st.info("La tabla ya viene ordenada con los registros más antiguos arriba. Digita los consumos de la semana.")

        columnas_pantalla = [col_codigo, col_concentrado]
        if col_lote: columnas_pantalla.append(col_lote)
        if col_fecha: columnas_pantalla.append(col_fecha)
        columnas_pantalla.extend(['Inventario Físico Bodega (A mano)', 'Inventario Máquina (A mano)', 'Consumo Registrado Semanal'])

        config_columnas = {}
        if col_fecha:
            config_columnas[col_fecha] = st.column_config.DatetimeColumn(format="DD/MM/YYYY")

        df_ingresado = st.data_editor(
            df_filtrado[columnas_pantalla],
            use_container_width=True,
            num_rows="dynamic",
            disabled=[col_codigo, col_concentrado, col_lote, col_fecha] if col_lote and col_fecha else [col_codigo, col_concentrado],
            column_config=config_columnas,
            key="tabla_maestra_cronologica"
        )

        # --- ANÁLISIS DE CONSUMO POR CRONOLOGÍA ---
        if st.button("⚡ 3. Calcular Informe de Concentrado", type="primary"):
            with st.spinner("Analizando consumos cronológicos de AX365..."):
                df_proc = df_ingresado.copy()
                df_proc['Inventario Físico Bodega (A mano)'] = pd.to_numeric(df_proc['Inventario Físico Bodega (A mano)'], errors='coerce').fillna(0)
                df_proc['Inventario Máquina (A mano)'] = pd.to_numeric(df_proc['Inventario Máquina (A mano)'], errors='coerce').fillna(0)
                df_proc['Consumo Registrado Semanal'] = pd.to_numeric(df_proc['Consumo Registrado Semanal'], errors='coerce').fillna(0)
                
                df_proc['stock sistema en inventory (unid)'] = df_proc['Inventario Físico Bodega (A mano)'] + df_proc['Inventario Máquina (A mano)']
                
                dict_consumos = df_proc.groupby(col_codigo)['Consumo Registrado Semanal'].sum().to_dict()
                reporte_bajas = []
                alertas_quiebre = []

                for codigo_ax, gasto_total in dict_consumos.items():
                    if gasto_total <= 0 or str(codigo_ax).lower() == 'nan':
                        continue
                        
                    gasto_restante = gasto_total
                    indices_lotes = df_proc[df_proc[col_codigo] == codigo_ax].index 
                    
                    for idx in indices_lotes:
                        if gasto_restante <= 0:
                            break
                            
                        stock_lote_actual = df_proc.at[idx, 'stock sistema en inventory (unid)']
                        
                        if stock_lote_actual > 0:
                            descuento = min(stock_lote_actual, gasto_restante)
                            df_proc.at[idx, 'stock sistema en inventory (unid)'] -= descuento
                            gasto_restante -= descuento
                            
                            fecha_str = df_proc.at[idx, col_fecha].strftime('%d/%m/%Y') if col_fecha and pd.notna(df_proc.at[idx, col_fecha]) else "No registrada"
                            
                            reporte_bajas.append({
                                'Código AX': codigo_ax,
                                'Concentrado': df_proc.at[idx, col_concentrado],
                                'Lote': df_proc.at[idx, col_lote] if col_lote else "N/A",

                            

