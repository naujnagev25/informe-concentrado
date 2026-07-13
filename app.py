import streamlit as st
import pandas as pd
import io

# Configuración visual de la aplicación
st.set_page_config(page_title="Informe Concentrado Semanal", page_icon="📋", layout="wide")

st.title("📋 Informe Concentrado Semanal - Tienda Curicó")
st.markdown("""
### Proceso de Carga Filtrado:
1. **Sube tu archivo exportado de AX365**. El sistema filtrará automáticamente dejando **solo** los concentrados *MCI TINTER, TWIST y Transparentes (Rojo/Amarillo Óxido)*.
2. **Digita a mano** en la tabla interactiva el Inventario Físico de Bodega, el de la Máquina y el Consumo Semanal.
3. El sistema sumará ambos valores y procesará el consumo restando del lote más antiguo al más nuevo.
""")

# Área para cargar el archivo Excel base de AX365
archivo_ax = st.file_uploader("1. Sube la exportación de AX365 (.xlsx)", type=["xlsx"])

if archivo_ax is not None:
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
        st.error("🚨 Error: No se encontró la columna de descripción del producto ('Concentrado') para aplicar los filtros.")
        st.stop()

    # Filtrar filas vacías de la exportación de AX365 para dejar limpia la tabla
    df_limpio = df_original.dropna(subset=[col_codigo, col_concentrado]).copy()
    df_limpio[col_codigo] = df_limpio[col_codigo].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    
    # --- FILTRO ESTRICTO DE CONCENTRADOS SOLICITADOS ---
    df_limpio['Concentrado_Minuscula'] = df_limpio[col_concentrado].astype(str).str.lower().str.strip()
    
    condicion_mci = df_limpio['Concentrado_Minuscula'].str.contains('mci tinter', na=False)
    condicion_twist = df_limpio['Concentrado_Minuscula'].str.contains('twist', na=False)
    condicion_rojo_ox = df_limpio['Concentrado_Minuscula'].str.contains('transparente rojo', na=False) | (df_limpio['Concentrado_Minuscula'].str.contains('transparente', na=False) & df_limpio['Concentrado_Minuscula'].str.contains('rojo', na=False))
    condicion_amarillo_ox = df_limpio['Concentrado_Minuscula'].str.contains('transparente amarillo', na=False) | (df_limpio['Concentrado_Minuscula'].str.contains('transparente', na=False) & df_limpio['Concentrado_Minuscula'].str.contains('amarillo', na=False))
    
    # Unir todas las condiciones (Filtra solo lo solicitado)
    df_filtrado = df_limpio[condicion_mci | condicion_twist | condicion_rojo_ox | condicion_amarillo_ox].copy()
    df_filtrado = df_filtrado.drop(columns=['Concentrado_Minuscula']) # Eliminar columna temporal
    
    if df_filtrado.empty:
        st.warning("⚠️ No se encontraron productos MCI TINTER, TWIST o Transparentes Óxido en la pestaña seleccionada. Revisa el archivo.")
        st.stop()
    
    # --- INYECCIÓN DE COLUMNAS PARA INGRESO MANUAL ---
    df_filtrado['Inventario Físico Bodega (A mano)'] = 0.0
    df_filtrado['Inventario Máquina (A mano)'] = 0.0
    df_filtrado['Consumo Registrado Semanal'] = 0.0

    st.subheader(f"📝 2. Módulo de Entrada Manual ({len(df_filtrado)} filas filtradas)")
    st.info("La tabla ahora muestra SOLO tus concentrados específicos. Haz doble clic para registrar tus datos a mano.")

    # Definir las columnas visibles en la cuadrícula de la pantalla
    columnas_pantalla = [col_codigo, col_concentrado]
    if col_lote: columnas_pantalla.append(col_lote)
    columnas_pantalla.extend(['Inventario Físico Bodega (A mano)', 'Inventario Máquina (A mano)', 'Consumo Registrado Semanal'])

    # Renderizar la tabla interactiva filtrada
    df_ingresado = st.data_editor(
        df_filtrado[columnas_pantalla],
        use_container_width=True,
        num_rows="dynamic",
        disabled=[col_codigo, col_concentrado, col_lote] if col_lote else [col_codigo, col_concentrado]
    )

    # --- PROCESAMIENTO MATEMÁTICO AL PRESIONAR EL BOTÓN ---
    if st.button("⚡ 3. Generar Informe y Aplicar FIFO", type="primary"):
        with st.spinner("Procesando datos filtrados y ejecutando descuentos por antigüedad..."):
            
            # Saneamiento de las entradas manuales
            df_ingresado['Inventario Físico Bodega (A mano)'] = pd.to_numeric(df_ingresado['Inventario Físico Bodega (A mano)'], errors='coerce').fillna(0)
            df_ingresado['Inventario Máquina (A mano)'] = pd.to_numeric(df_ingresado['Inventario Máquina (A mano)'], errors='coerce').fillna(0)
            df_ingresado['Consumo Registrado Semanal'] = pd.to_numeric(df_ingresado['Consumo Registrado Semanal'], errors='coerce').fillna(0)
            
            # Cálculo automático de Stock Físico Sistema
            df_ingresado['stock sistema en inventario (unid)'] = df_ingresado['Inventario Físico Bodega (A mano)'] + df_ingresado['Inventario Máquina (A mano)']
            
            df_final = df_ingresado.copy()
            
            # Consolidar cuánto se consumió en total por cada código único de producto
            dict_consumos = df_ingresado.groupby(col_codigo)['Consumo Registrado Semanal'].sum().to_dict()
            
            reporte_bajas = []
            alertas_quiebre = []

            # --- ALGORITMO FIFO ---
            for codigo_ax, gasto_total in dict_consumos.items():
                if gasto_total <= 0 or codigo_ax == 'nan':
                    continue
                    
                gasto_restante = gasto_total
                indices_lotes = df_final[df_final[col_codigo] == codigo_ax].index
                
                for idx in indices_lotes:
                    if gasto_restante <= 0:
                        break
                        
                    stock_lote_actual = df_final.at[idx, 'stock sistema en inventario (unid)']
                    
                    if stock_lote_actual > 0:
                        descuento = min(stock_lote_actual, gasto_restante)
                        
                        df_final.at[idx, 'stock sistema en inventario (unid)'] -= descuento
                        gasto_restante -= descuento
                        
                        reporte_bajas.append({
                            'Código AX': codigo_ax,
                            'Concentrado': df_final.at[idx, col_concentrado],
                            'Lote Afectado': df_final.at[idx, col_lote] if col_lote else f"Fila {idx+2}",
                            'Cantidad Descontada': round(descuento, 4),
                            'Stock Restante en Lote': round(df_final.at[idx, 'stock sistema en inventario (unid)'], 4)
                        })
                        
                if gasto_restante > 0:
                    alertas_quiebre.append(f"🚨 Inventario Insuficiente: Código AX {codigo_ax} requiere {gasto_total} unidades, pero faltaron {gasto_restante:.2f} unidades por falta de stock.")

            st.success("✨ ¡Informe Concentrado Semanal calculado de forma exitosa!")
            
            # Estructurar la exportación Excel final limpia
            salida_excel = io.BytesIO()
            with pd.ExcelWriter(salida_excel, engine='openpyxl') as writer:
                df_final.to_excel(writer, sheet_name='Informe Planilla Semanal', index=False)
                if reporte_bajas:
                    pd.DataFrame(reporte_bajas).to_excel(writer, sheet_name='Detalle Lotes Consumidos', index=False)
            salida_excel.seek(0)

            st.download_button(
                label="📥 Descargar Informe Concentrado Semanal (.xlsx)",
                data=salida_excel,
                file_name="Informe_Concentrado_Semanal.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

            if alertas_quiebre:
                st.subheader("⚠️ Alertas de Diferencias")
                for al in alertas_quiebre:
                    st.warning(al)

            # Desplegar visualización de resultados
            tab1, tab2 = st.tabs(["📊 Ver Lotes Consumidos (FIFO)", "🗄️ Ver Planilla de Stock Final"])
            with tab1:
                if reporte_bajas:
                    st.dataframe(pd.DataFrame(reporte_bajas), use_container_width=True)
                else:
                    st.info("No se registraron movimientos.")
            with tab2:
                columnas_finales = [col_codigo, col_concentrado, 'stock sistema en inventario (unid)', 'Consumo Registrado Semanal']
                st.dataframe(df_final[columnas_finales], use_container_width=True)
