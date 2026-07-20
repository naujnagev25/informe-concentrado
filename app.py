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
1. **Carga los tres archivos independientes** (AX365, Stock de Máquina y Stock Físico).
2. El sistema filtrará automáticamente los concentrados *MCI TINTER, TWIST y Transparentes*.
3. **Ingresa los consumos semanales** directamente en la tabla interactiva.
4. Presiona el botón **⚡ Calcular y Procesar Reporte** para ejecutar el algoritmo cronológico y generar tu Excel corporativo en CLP.
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

# --- SECCIÓN DE CARGA DE ARCHIVOS INDEPENDIENTES ---
st.subheader("📁 1. Accesos para Carga de Datos")
col1, col2, col3 = st.columns(3)

with col1:
    archivo_ax = st.file_uploader("📥 Datos de AX365 (.xlsx)", type=["xlsx"])
with col2:
    archivo_maquina = st.file_uploader("📥 Stock de la Máquina (.xlsx, .csv)", type=["xlsx", "csv"])
with col3:
    archivo_bodega = st.file_uploader("📥 Stock Físico Bodega (.xlsx, .csv)", type=["xlsx", "csv"])

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

    # --- DETECCIÓN DE COLUMNAS ---
    col_codigo = detectar_columna_ax(df_original, ['código ax', 'codigo ax', 'artículo'])
    col_lote = detectar_columna_ax(df_original, ['número de lote', 'numero de lote', 'lote'])
    col_fecha = detectar_columna_ax(df_original, ['fecha de fabricación', 'fecha de fabricacion', 'fabricación'])
    col_concentrado = detectar_columna_ax(df_original, ['concentrado', 'nombre', 'producto', 'descripción'])
    col_inv_ax = detectar_columna_ax(df_original, ['inventario físico', 'inventario fisico', 'físico disponible'])

    # Estandarizar nombres de columnas
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
            df_limpio = df_limpio.sort_values(by=[col_fecha], ascending=True, na_position='last')

        # Filtros de Concentrados Solicitados
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
            st.warning("⚠️ No se encontraron concentrados válidos con los criterios de búsqueda.")
        else:
            # --- PROCESAMIENTO DE PLANILLAS EXTERNAS (MÁQUINA Y BODEGA) ---
            dict_maquina = {}
            dict_bodega = {}

            def procesar_archivo_externo(archivo):
                df_ext = pd.read_csv(archivo) if archivo.name.endswith('.csv') else pd.read_excel(archivo)
                df_ext.columns = df_ext.columns.str.strip()
                c_cod = detectar_columna_ax(df_ext, ['código ax', 'codigo ax', 'artículo', 'item'])
                c_cant = detectar_columna_ax(df_ext, ['cantidad', 'stock', 'físico', 'fisico', 'total'])
                if c_cod and c_cant:
                    df_ext[c_cod] = df_ext[c_cod].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                    df_ext[c_cant] = pd.to_numeric(df_ext[c_cant], errors='coerce').fillna(0)
                    return df_ext.groupby(c_cod)[c_cant].sum().to_dict()
                return {}

            if archivo_maquina is not None: dict_maquina = procesar_archivo_externo(archivo_maquina)
            if archivo_bodega is not None: dict_bodega = procesar_archivo_externo(archivo_bodega)

            # Inyectar cruces automatizados
            if 'df_maestro_tabla' not in st.session_state:
                df_filtrado['Inventario Físico Bodega (A mano)'] = df_filtrado[col_codigo].map(dict_bodega).fillna(0.0)
                df_filtrado['Inventario Máquina (A mano)'] = df_filtrado[col_codigo].map(dict_maquina).fillna(0.0)
                df_filtrado['Consumo Registrado Semanal'] = 0.0
                st.session_state['df_maestro_tabla'] = df_filtrado.copy()

            st.subheader("📝 2. Módulo de Revisión y Entrada Manual")
            st.info("Digita a mano el 'Consumo Registrado Semanal' en la columna correspondiente y luego haz clic abajo en Calcular.")

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

            # Renderizar editor de datos vinculando el session_state para mantener la memoria intacta
            df_ingresado = st.data_editor(
                st.session_state['df_maestro_tabla'][columnas_pantalla],
                use_container_width=True,
                num_rows="dynamic",
                disabled=columnas_deshabilitadas,
                column_config=config_columnas,
                key="tabla_maestra_cronologica"
            )
            
            # Sincronizar cambios manuales de inmediato en el estado interno
            st.session_state['df_maestro_tabla'].update(df_ingresado)

            # --- BOTÓN DE INICIO PARA PROCESAR EL CÁLCULO ---
            st.markdown("---")
            if st.button("⚡ Calcular y Procesar Reporte", type="primary", use_container_width=True):
                with st.spinner("Ejecutando algoritmo matemático de rebaja por lotes antiguos..."):
                    df_proc = st.session_state['df_maestro_tabla'].copy()
                    
