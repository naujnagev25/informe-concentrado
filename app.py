import streamlit as st
import pandas as pd
import io

# Ajustes de la interfaz
st.set_page_config(page_title="Conector AX365 FIFO", page_icon="⚙️", layout="wide")
st.title("⚙️ Conector FIFO Automático para Dynamics 365 (AX365)")
st.markdown("Carga directamente el reporte exportado de AX365. El sistema limpiará las filas corruptas y procesará el inventario.")

# Cargador del archivo crudo de AX365
archivo_ax365 = st.file_uploader("Sube el archivo Excel exportado de AX365 (.xlsx)", type=["xlsx"])

if archivo_ax365 is not None:
    try:
        # 1. Leer las hojas del libro de Excel
        excel_book = pd.ExcelFile(archivo_ax365)
        hojas_disponibles = excel_book.sheet_names
        
        # Intentar emparejar automáticamente las pestañas por nombres comunes
        detectar_inv = "Inventario" if "Inventario" in hojas_disponibles else hojas_disponibles[0]
        detectar_cons = "Consumos" if "Consumos" in hojas_disponibles else (hojas_disponibles[1] if len(hojas_disponibles) > 1 else hojas_disponibles[0])
        
        st.sidebar.header("📁 Pestañas del Sistema ERP")
        hoja_inventario = st.sidebar.selectbox("Pestaña de Stock AX365:", hojas_disponibles, index=hojas_disponibles.index(detectar_inv))
        hoja_consumos = st.sidebar.selectbox("Pestaña de Consumos AX365:", hojas_disponibles, index=hojas_disponibles.index(detectar_cons))
        
        # --- PROCESO ULTRA-AUTOMÁTICO ---
        with st.spinner("🔧 Limpiando formatos y cabeceras de AX365..."):
            
            # Carga inicial
            df_inv = pd.read_excel(archivo_ax365, sheet_name=hoja_inventario)
            df_cons = pd.read_excel(archivo_ax365, sheet_name=hoja_consumos)
            
            # FILTRO CRUCIAL AX365: Eliminar filas vacías o subtotales que genera el ERP
            df_inv = df_inv.dropna(subset=['Código AX']).copy()
            df_cons = df_cons.dropna(subset=['Código AX']).copy()
            
            # Normalizar los formatos de códigos a Texto para evitar que se rompa la búsqueda
            df_inv['Código AX'] = df_inv['Código AX'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            df_cons['Código AX'] = df_cons['Código AX'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            
            # Identificar celdas de stock basándose en variaciones de nombre de AX
            col_stock = [c for c in df_inv.columns if any(p in c.lower() for p in ['stock', 'inventario', 'físico', 'physical', 'on-hand'])]
            nombre_col_stock = col_stock[0] if col_stock else df_inv.columns[4]
            
            # Identificar celdas de consumo basándose en variaciones de nombre de AX
            col_gasto = [c for c in df_cons.columns if any(p in c.lower() for p in ['consumo', 'gasto', 'cantidad', 'quantity'])]
            nombre_col_gasto = col_gasto[0] if col_gasto else df_cons.columns[-1]
            
            # EVITAR ERRORES ¡#VALOR!: Convertir textos corruptos a ceros numéricos de forma segura
            df_inv[nombre_col_stock] = pd.to_numeric(df_inv[nombre_col_stock], errors='coerce').fillna(0)
            df_cons[nombre_col_gasto] = pd.to_numeric(df_cons[nombre_col_gasto], errors='coerce').fillna(0)
            
            reporte_bajas = []
            alertas_quiebre = []
            
            # --- ALGORITMO FIFO SOBRE LISTADO ERP ---
            for _, fila_c in df_cons.iterrows():
                codigo_ax = fila_c['Código AX']
                cantidad_a_descontar = fila_c[nombre_col_gasto]
                
                if cantidad_a_descontar <= 0:
                    continue
                
                # Agrupar lotes manteniendo rigurosamente el orden de bajada de AX365 (Antiguo -> Nuevo)
                lotes_producto = df_inv[df_inv['Código AX'] == codigo_ax].copy()
                
                if lotes_producto.empty:
                    continue
                
                for idx, lote in lotes_producto.iterrows():
                    if cantidad_a_descontar <= 0:
                        break
                    
                    stock_lote_actual = df_inv.at[idx, nombre_col_stock]
                    
                    if stock_lote_actual > 0:
                        descuento = min(stock_lote_actual, cantidad_a_descontar)
                        
                        df_inv.at[idx, nombre_col_stock] -= descuento
                        cantidad_a_descontar -= descuento
                        
                        reporte_bajas.append({
                            'Código AX': codigo_ax,
                            'Concentrado/Producto': lote.get('Concentrado', lote.get('Nombre del producto', 'N/A')),
                            'Lote Consumido': lote.get('Número de lote', 'N/A'),
                            'Cantidad Descontada': round(descuento, 4),
                            'Inventario Restante Lote': round(df_inv.at[idx, nombre_col_stock], 4)
                        })
                
                if cantidad_a_descontar > 0:
                    alertas_quiebre.append(f"🚨 Insuficiencia en ERP: El código {codigo_ax} requiere {cantidad_a_descontar:.3f} unidades más de las disponibles.")
            
            # --- IMPRESIÓN DE INTERFAZ WEB AUTOMÁTICA ---
            st.success("✨ ¡Reporte de AX365 procesado y limpiado correctamente!")
            
            # Conversión final del documento limpio a memoria lista para transferir
            salida_excel = io.BytesIO()
            with pd.ExcelWriter(salida_excel, engine='openpyxl') as writer:
                df_inv.to_excel(writer, sheet_name='Inventario Actualizado', index=False)
                if reporte_bajas:
                    pd.DataFrame(reporte_bajas).to_excel(writer, sheet_name='Trazabilidad de Bajas', index=False)
            salida_excel.seek(0)
            
            # Botón principal de guardado
            st.download_button(
                label="📥 Descargar Excel Corregido sin errores",
                data=salida_excel,
                file_name="Reporte_FIFO_AX365_Listo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
            if alertas_quiebre:
                st.subheader("⚠️ Descalces Registrados contra AX365")
                for al in alertas_quiebre:
                    st.warning(al)
            
            tab1, tab2 = st.tabs(["📋 Historial de Lotes Afectados", "📊 Datos Consolidados"])
            with tab1:
                if reporte_bajas:
                    st.dataframe(pd.DataFrame(reporte_bajas), use_container_width=True)
                else:
                    st.info("No se hallaron registros válidos de consumo en esta carga.")
            with tab2:
                st.dataframe(df_inv, use_container_width=True)
                
    except Exception as e:
        st.error(f"❌ Error de lectura en la estructura AX365: {e}. Confirma que el archivo mantenga la columna 'Código AX'.")
