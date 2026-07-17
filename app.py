import streamlit as st
import pandas as pd
import io

# Configuración visual de la aplicación
st.set_page_config(page_title="Informe Concentrado Semanal", page_icon="📋", layout="wide")

st.title("📋 Informe Concentrado Semanal - Tienda Curicó")
st.markdown("""
### Proceso de Carga y Consolidación:
1. **Sube tu archivo base de AX365** (Filtra automáticamente *MCI TINTER, TWIST y Transparentes*).
2. **Opcional:** Sube las planillas de inventario de **Bodega** y/o **Máquina** para cargar los datos automáticamente.
3. **Corrige o completa a mano** cualquier dato directamente en la tabla interactiva si es necesario.
4. El sistema calculará el inventario total y procesará el descuento de lotes mediante el método FIFO.
""")

# --- SECCIÓN DE CARGA DE ARCHIVOS ---
st.subheader("📁 1. Carga de Documentos")
col1, col2, col3 = st.columns(3)

with col1:
    archivo_ax = st.file_uploader("Base de AX365 (.xlsx)", type=["xlsx"])
with col2:
    archivo_bodega = st.file_uploader("Opcional: Stock Físico Bodega (.xlsx, .csv)", type=["xlsx", "csv"])
with col3:
    archivo_maquina = st.file_uploader("Opcional: Stock de la Máquina (.xlsx, .csv)", type=["xlsx", "csv"])

if archivo_ax is not None:
    try:
        excel_book = pd.ExcelFile(archivo_ax)
        hojas = excel_book.sheet_names
        hoja_seleccionada = st.selectbox("Selecciona la pestaña del reporte AX365:", hojas, index=0)
        
        # Leer la planilla original del ERP
        df_original = pd.read_excel(archivo_ax, sheet_name=hoja_seleccionada)
        df_original.columns = df_original.columns.str.strip()
        
        # --- DETECCIÓN AUTOMÁTICA DE DATOS AX365 ---
        def buscar_columna(df, palabras_clave):
            for col in df.columns:
                if any(p in str(col).lower() for p in palabras_clave):
                    return col
            return None

        col_codigo = buscar_columna(df_original, ['código ax', 'codigo ax', 'artículo', 'articulo', 'item'])
        col_concentrado = buscar_columna(df_original, ['concentrado', 'nombre', 'producto'])
        col_lote = buscar_columna(df_original, ['lotes consumidos', 'lote', 'batch', 'número de lote'])
        
        if not col_codigo:
            st.error("🚨 Error: No se encontró la columna de 'Código AX' en el archivo.")
            st.stop()
            
        if not col_concentrado:
            st.error("🚨 Error: No se encontró la columna de descripción del producto ('Concentrado').")
            st.stop()

        # Limpieza de la base AX365
        df_limpio = df_original.dropna(subset=[col_codigo, col_concentrado]).copy()
        df_limpio[col_codigo] = df_limpio[col_codigo].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        
        # --- FILTRO ESTRICTO DE CONCENTRADOS ---
        df_limpio['Concentrado_Minuscula'] = df_limpio[col_concentrado].astype(str).str.lower().str.strip()
        
        condicion_mci = df_limpio['Concentrado_Minuscula'].str.contains('mci tinter', na=False)
        condicion_twist = df_limpio['Concentrado_Minuscula'].str.contains('twist', na=False)
        condicion_rojo_ox = df_limpio['Concentrado_Minuscula'].str.contains('transparente rojo', na=False) | (df_limpio['Concentrado_Minuscula'].str.contains('transparente', na=False) & df_limpio['Concentrado_Minuscula'].str.contains('rojo', na=False))
        condicion_amarillo_ox = df_limpio['Concentrado_Minuscula'].str.contains('transparente amarillo', na=False) | (df_limpio['Concentrado_Minuscula'].str.contains('transparente', na=False) & df_limpio['Concentrado_Minuscula'].str.contains('amarillo', na=False))
        
        df_filtrado = df_limpio[condicion_mci | condicion_twist | condicion_rojo_ox | condicion_amarillo_ox].copy()
        df_filtrado = df_filtrado.drop(columns=['Concentrado_Minuscula'])
        
        if df_filtrado.empty:
            st.warning("⚠️ No se encontraron los concentrados solicitados en la pestaña seleccionada.")
            st.stop()
        
        # --- PROCESAMIENTO DE CARGAS ADICIONALES (BODEGA Y MÁQUINA) ---
        dict_bodega = {}
        dict_maquina = {}
        
        def procesar_archivo_extern(archivo):
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

        if archivo_bodega is not None:
            dict_bodega = procesar_archivo_extern(archivo_bodega)
            st.toast("✅ Datos de Bodega cargados con éxito", icon="📦")
            
        if archivo_maquina is not None:
            dict_maquina = procesar_archivo_extern(archivo_maquina)
            st.toast("✅ Datos de la Máquina cargados con éxito", icon="⚙️")

        # --- INYECCIÓN Y CRUCE DE DATOS ---
        df_filtrado['Inventario Físico Bodega (A mano)'] = df_filtrado[col_codigo].map(dict_bodega).fillna(0.0)
        df_filtrado['Inventario Máquina (A mano)'] = df_filtrado[col_codigo].map(dict_maquina).fillna(0.0)
        df_filtrado['Consumo Registrado Semanal'] = 0.0

        # --- MÓDULO INTERACTIVO DE PANTALLA ---
        st.subheader(f"📝 2. Módulo de Revisión y Entrada Manual ({len(df_filtrado)} filas)")
        st.info("Los datos de los archivos subidos se cruzaron automáticamente. Puedes modificarlos o completarlos abajo.")

        columnas_pantalla = [col_codigo, col_concentrado]
        if col_lote: 
            columnas_pantalla.append(col_lote)
        columnas_pantalla.extend(['Inventario Físico Bodega (A mano)', 'Inventario Máquina (A mano)', 'Consumo Registrado Semanal'])

        df_ingresado = st.data_editor(
            df_filtrado[columnas_pantalla],
            use_container_width=True,
            num_rows="dynamic",
            disabled=[col_codigo, col_concentrado, col_lote] if col_lote else [col_codigo, col_concentrado],
            key="tabla_maestra"
        )

        # --- PROCESAMIENTO MATEMÁTICO (FIFO) ---
        if st.button("⚡ 3. Generar Informe y Aplicar FIFO", type="primary"):
            with st.spinner("Ejecutando algoritmo FIFO..."):
                df_proc = df_ingresado.copy()
                df_proc['Inventario Físico Bodega (A mano)'] = pd.to_numeric(df_proc['Inventario Físico Bodega (A mano)'], errors='coerce').fillna(0)
                df_proc['Inventario Máquina (A mano)'] = pd.to_numeric(df_proc['Inventario Máquina (A mano)'], errors='coerce').fillna(0)
                df_proc['Consumo Registrado Semanal'] = pd.to_numeric(df_proc['Consumo Registrado Semanal'], errors='coerce').fillna(0)
                
                df_proc['stock sistema en inventario (unid)'] = df_proc['Inventario Físico Bodega (A mano)'] + df_proc['Inventario Máquina (A mano)']
                
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
                            
                        stock_lote_actual = df_proc.at[idx, 'stock sistema en inventario (unid)']
                        
                        if stock_lote_actual > 0:
                            descuento = min(stock_lote_actual, gasto_restante)
                            df_proc.at[idx, 'stock sistema en inventario (unid)'] -= descuento
                            gasto_restante -= descuento
                            
                            reporte_bajas.append({
                                'Código AX': codigo_ax,
                                'Concentrado': df_proc.at[idx, col_concentrado],
                                'Lote Afectado': df_proc.at[idx, col_lote] if col_lote else f"Fila {idx+2}",
                                'Cantidad Descontada': round(descuento, 4),
                                'Stock Restante en Lote': round(df_proc.at[idx, 'stock sistema en inventario (unid)'], 4)
                            })
                            
                    if gasto_restante > 0:
                        alertas_quiebre.append(f"🚨 Inventario Insuficiente: Código AX {codigo_ax} requiere {gasto_total} unidades, faltaron {gasto_restante:.2f}.")

                st.success("✨ ¡Informe Calculado Exitosamente!")
                
                # Preparar descarga Excel
                salida_excel = io.BytesIO()
                with pd.ExcelWriter(salida_excel, engine='openpyxl') as writer:
                    df_proc.to_excel(writer, sheet_name='Informe Planilla Semanal', index=False)
                    if reporte_bajas:
                        pd.DataFrame(reporte_bajas).to_excel(writer, sheet_name='Detalle Lotes Consumidos', index=False)
                salida_excel.seek(0)

                st.download_button(
                    label="📥 Descargar Informe Concentrado Semanal (.xlsx)",
                    data=salida_excel,
                    file_name="Informe_Concentrado_Semanal.xlsx",
