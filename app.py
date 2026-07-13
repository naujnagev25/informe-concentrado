import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="FIFO Curicó", page_icon="🎨", layout="wide")
st.title("🎨 Sistema FIFO - Tienda CURICÓ")
st.markdown("Herramienta automática para descontar consumos semanales priorizando el lote más antiguo.")

archivo = st.file_uploader("Sube tu archivo Excel (.xlsx)", type=["xlsx"])

if archivo:
    try:
        excel = pd.ExcelFile(archivo)
        col1, col2 = st.columns(2)
        with col1:
            hoja_inv = st.selectbox("Pestaña de Inventario Actual:", excel.sheet_names)
        with col2:
            hoja_cons = st.selectbox("Pestaña de Consumos Semanales:", excel.sheet_names)
            
        if st.button("🚀 Procesar Descuentos por Lote", type="primary"):
            df_inv = pd.read_excel(archivo, sheet_name=hoja_inv)
            df_cons = pd.read_excel(archivo, sheet_name=hoja_cons)
            
            # Ajuste estricto a las columnas de tu documento de Curicó
            df_inv['Código AX'] = df_inv['Código AX'].astype(str).str.strip()
            df_inv['Inventario físico'] = pd.to_numeric(df_inv['Inventario físico'], errors='coerce').fillna(0)
            df_cons['Código AX'] = df_cons['Código AX'].astype(str).str.strip()
            df_cons['Cantidad Consumida'] = pd.to_numeric(df_cons['Cantidad Consumida'], errors='coerce').fillna(0)
            
            bajas = []
            
            for _, fila in df_cons.iterrows():
                cod = fila['Código AX']
                gasto = fila['Cantidad Consumida']
                
                lotes = df_inv[df_inv['Código AX'] == cod].copy()
                for idx, lote in lotes.iterrows():
                    if gasto <= 0: break
                    stock = df_inv.at[idx, 'Inventario físico']
                    if stock > 0:
                        restar = min(stock, gasto)
                        df_inv.at[idx, 'Inventario físico'] -= restar
                        gasto -= restar
                        bajas.append({
                            'Código AX': cod,
                            'Producto': lote.get('Concentrado', 'N/A'),
                            'Lote Consumido': lote.get('Número de lote', 'N/A'),
                            'Cantidad Restada': restar
                        })
                if gasto > 0:
                    st.error(f"🚨 Alerta: ¡Quiebre de stock para el código {cod}! Faltaron {gasto:.2f} unidades.")
            
            st.success("✨ ¡Inventario recalculado correctamente!")
            
            # Creación del nuevo archivo para descarga inmediata
            salida = io.BytesIO()
            with pd.ExcelWriter(salida, engine='openpyxl') as writer:
                df_inv.to_excel(writer, sheet_name='Inventario Actualizado', index=False)
                if bajas:
                    pd.DataFrame(bajas).to_excel(writer, sheet_name='Reporte de Bajas FIFO', index=False)
            salida.seek(0)
            
            st.download_button(
                label="📥 Descargar Excel con Lotes Actualizados",
                data=salida,
                file_name="Inventario_FIFO_Curico.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Error al leer las columnas: Asegúrate de que coincidan con el formato original de la tienda. ({e})")
