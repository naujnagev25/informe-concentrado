import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Configuración visual de la aplicación web
st.set_page_config(page_title="Informe Concentrado Semanal", page_icon="📋", layout="wide")

st.title("📋 Informe Concentrado Semanal - Tienda Curicó")
st.markdown("""
### Instrucciones de Uso:
1. **Sube el archivo base de AX365** en el acceso de abajo.
2. El sistema filtrará y ordenará los productos automáticamente: Primero **MCI TINTER**, luego **Transparentes** y después **TWIST** (de los más antiguos a los más nuevos por Fecha de fabricación).
3. **Digita a mano** el Stock de Bodega, el Stock de la Máquina y el Consumo Registrado Semanal directamente en la tabla interactiva.
4. Haz clic en **⚡ Calcular y Procesar Reporte** para obtener el resumen ejecutivo de consumos y costos en CLP.
""")

# --- DICCIONARIO MAESTRO DE COSTOS BASE EN PESOS CHILENOS (CLP) ---
COSTOS_BASE_MAESTROS = {
    "13342919": 31942, "13344019": 31317, "13344319": 82671, "13344419": 43103,
    "13344519": 75731, "13344619": 28040, "13344719": 58265, "13344819": 43768,
    "13344919": 51315, "13345019": 112916, "13345119": 112919, "13345219": 112921,
    "13345319": 112922, "13345419": 112924, "13345519": 112920, "13345619": 112922,
    "13345719": 112924,
    "11901104": 6278,  "11901204": 6280,  "11424004": 6278,  "11424104": 6282,
    "11424204": 6286,  "11424304": 6289,  "11424404": 6290,  "11424504": 6292,
    "11424604": 6295,  "11424704": 6297,  "11424804": 6299,  "11424904": 6298,
    "11425104": 6291,  "11425204": 6292
}

# --- ACCESO ÚNICO PARA CARGA DEL ARCHIVO AX365 ---
st.subheader("📁 1. Carga de Documento Base")
archivo_ax = st.file_uploader("📥 Datos de AX365 (.xlsx)", type=["xlsx"])

# Función para detectar columnas basada en palabras clave
def detectar_columna_ax(df, palabras_clave):
    for col in df.columns:
        if any(p in str(col).lower() for p in palabras_clave):
            return col
    return None

if archivo_ax is not None:
    excel_book = pd.ExcelFile(archivo_ax)
    hojas = excel_book.sheet_names
    hoja_seleccionada = st.selectbox("Selecciona la pestaña del reporte AX365:", hojas, index=0)

    df_original = pd.read_excel(archivo_ax, sheet_name=hoja_seleccionada)
    df_original.columns = df_original.columns.str.strip()

    # --- DETECCIÓN DE COLUMNAS ORIGINALES ---
    col_codigo = detectar_columna_ax(df_original, ['código ax', 'codigo ax', 'artículo'])
    col_lote = detectar_columna_ax(df_original, ['número de lote', 'numero de lote', 'lote'])
    col_fecha = detectar_columna_ax(df_original, ['fecha de fabricación', 'fecha de fabricacion', 'fabricación'])
    col_concentrado = detectar_columna_ax(df_original, ['concentrado', 'nombre', 'producto', 'descripción'])
    col_inv_ax = detectar_columna_ax(df_original, ['inventario físico', 'inventario fisico', 'físico disponible'])

    # Estandarizar nombres de columnas internos
    if col_codigo: df_original = df_original.rename(columns={col_codigo: 'Código AX'})
    if col_lote: df_original = df_original.rename(columns={col_lote: 'Número de lote'})
    if col_fecha: df_original = df_original.rename(columns={col_fecha: 'Fecha de fabricación'})
    if col_concentrado: df_original = df_original.rename(columns={col_concentrado: 'Concentrado'})
    if col_inv_ax: df_original = df_original.rename(columns={col_inv_ax: 'Inventario físico AX'})

    col_codigo = 'Código AX'
    col_lote = 'Número de lote' if 'Número de lote' in df_original.columns else None
    col_fecha = 'Fecha de fabricación' if 'Fecha de fabricación' in df_original.columns else None
    col_concentrado = 'Concentrado'
    col_inv_ax = 'Inventario físico AX' if 'Inventario físico AX' in df_original.columns else None

    if col_codigo not in df_original.columns or col_concentrado not in df_original.columns:
        st.error("🚨 Error: No se encontraron las columnas esenciales ('Código AX' y 'Concentrado') en el archivo.")
    else:
        # --- FILTRADO Y LIMPIEZA INICIAL ---
        df_limpio = df_original.dropna(subset=[col_codigo, col_concentrado]).copy()
        df_limpio[col_codigo] = df_limpio[col_codigo].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

        if col_inv_ax:
            df_limpio[col_inv_ax] = pd.to_numeric(df_limpio[col_inv_ax], errors='coerce').fillna(0.0)

        df_limpio['Costo unitario (CLP)'] = df_limpio[col_codigo].map(COSTOS_BASE_MAESTROS).fillna(0).astype(int)

        if col_fecha:
            df_limpio[col_fecha] = pd.to_datetime(df_limpio[col_fecha], errors='coerce')

        # --- CLASIFICACIÓN Y ORDENAMIENTO EN GRUPOS SOLICITADOS ---
        df_limpio['Concentrado_Minuscula'] = df_limpio[col_concentrado].astype(str).str.lower().str.strip()
        
        condicion_mci = df_limpio['Concentrado_Minuscula'].str.contains('mci tinter', na=False)
        condicion_twist = df_limpio['Concentrado_Minuscula'].str.contains('twist', na=False)
        condicion_transparente = (df_limpio['Concentrado_Minuscula'].str.contains('transparente', na=False)) & \
                                 (df_limpio['Concentrado_Minuscula'].str.contains('rojo', na=False) | df_limpio['Concentrado_Minuscula'].str.contains('amarillo', na=False))

        # Asignar un índice numérico para controlar el orden estricto de las categorías
        df_limpio['Orden_Categoria'] = 99
        df_limpio.loc[condicion_mci, 'Orden_Categoria'] = 1          # 1° MCI TINTER
        df_limpio.loc[condicion_transparente, 'Orden_Categoria'] = 2   # 2° Transparentes
        df_limpio.loc[condicion_twist, 'Orden_Categoria'] = 3          # 3° TWIST

        # Filtrar para dejar únicamente los concentrados requeridos
        df_filtrado = df_limpio[df_limpio['Orden_Categoria'] != 99].copy()
        
        # Ordenar: Primero por categoría (1, 2, 3) y luego por fecha antigua de fabricación
        if col_fecha:
            df_filtrado = df_filtrado.sort_values(by=['Orden_Categoria', col_fecha], ascending=[True, True], na_position='last')
        else:
            df_filtrado = df_filtrado.sort_values(by=['Orden_Categoria'], ascending=True)

        df_filtrado = df_filtrado.drop(columns=['Concentrado_Minuscula', 'Orden_Categoria'])

        if df_filtrado.empty:
            st.warning("⚠️ No se encontraron concentrados válidos con las reglas estipuladas.")
        else:
            # --- INICIALIZAR COLUMNAS PARA REGISTRO MANUAL ---
            if 'df_maestro_tabla' not in st.session_state:
                df_filtrado['Inventario Físico Bodega (A mano)'] = 0.0
                df_filtrado['Inventario Máquina (A mano)'] = 0.0
                df_filtrado['Consumo Registrado Semanal'] = 0.0
                st.session_state['df_maestro_tabla'] = df_filtrado.copy()

            st.subheader("📝 2. Módulo de Entrada Manual (Stocks y Consumos)")
            st.info("La tabla está organizada en orden: MCI TINTER ➡️ TRANSPARENTES ➡️ TWIST. Digita los valores recolectados a mano.")

            columnas_pantalla = [col_codigo, col_concentrado]
            if col_lote: columnas_pantalla.append(col_lote)
            if col_fecha: columnas_pantalla.append(col_fecha)
            if col_inv_ax: columnas_pantalla.append(col_inv_ax)
            columnas_pantalla.extend(['Costo unitario (CLP)', 'Inventario Físico Bodega (A mano)', 'Inventario Máquina (A mano)', 'Consumo Registrado Semanal'])

            config_columnas = {}
            if col_fecha:
                config_columnas[col_fecha] = st.column_config.DatetimeColumn(format="DD/MM/YYYY")
            config_columnas['Costo unitario (CLP)'] = st.column_config.NumberColumn(format="$ %d")

            columnas_deshabilitadas = [col_codigo, col_concentrado]
            if col_lote: columnas_deshabilitadas.append(col_lote)
            if col_fecha: columnas_deshabilitadas.append(col_fecha)
            if col_inv_ax: columnas_deshabilitadas.append(col_inv_ax)
            columnas_deshabilitadas.append('Costo unitario (CLP)')

            # Renderizar editor de datos interactivo
            df_ingresado = st.data_editor(
                st.session_state['df_maestro_tabla'][columnas_pantalla],
                use_container_width=True,
                num_rows="dynamic",
                disabled=columnas_deshabilitadas,
                column_config=config_columnas,
                key="tabla_maestra_cronologica"
            )
            
            # Guardar persistencia en memoria interna
            st.session_state['df_maestro_tabla'].update(df_ingresado)

            # --- BOTÓN MAESTRO DE PROCESAMIENTO ---
            st.markdown("---")
            if st.button("⚡ Calcular y Procesar Reporte", type="primary", use_container_width=True):
                with st.spinner("Procesando cálculos financieros y rebajas por lotes cronológicos..."):
                    df_proc = st.session_state['df_maestro_tabla'].copy()
                    
                    df_proc['Inventario Físico Bodega (A mano)'] = pd.to_numeric(df_proc['Inventario Físico Bodega (A mano)'], errors='coerce').fillna(0)
                    df_proc['Inventario Máquina (A mano)'] = pd.to_numeric(df_proc['Inventario Máquina (A mano)'], errors='coerce').fillna(0)
                    df_proc['Consumo Registrado Semanal'] = pd.to_numeric(df_proc['Consumo Registrado Semanal'], errors='coerce').fillna(0)
                    df_proc['Costo unitario (CLP)'] = pd.to_numeric(df_proc['Costo unitario (CLP)'], errors='coerce').fillna(0).astype(int)
                    
                    # Stock Inicial disponible en Tienda
                    df_proc['stock sistema en inventory (unid)'] = df_proc['Inventario Físico Bodega (A mano)'] + df_proc['Inventario Máquina (A mano)']
                    
                    dict_consumos = df_proc.groupby(col_codigo)['Consumo Registrado Semanal'].sum().to_dict()
                    
                    reporte_bajas = []
