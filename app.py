import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Cargador de Facturas de Combustible", layout="wide")

st.title("⛽ Importador Masivo de Facturas de Combustible para Finnegans")
st.markdown("Herramienta de validación de Choferes, Patentes y Centros de Costo")

# --- SECCIÓN 1: CARGA DE ARCHIVOS ---
st.sidebar.header("📁 Carga de Archivos")

file_maestro = st.sidebar.file_uploader("1. Maestro de Choferes (.xlsx)", type=["xlsx"])
file_factura = st.sidebar.file_uploader("2. Detalle de Consumos / Factura (.xlsx o .csv)", type=["xlsx", "csv"])

if file_maestro and file_factura:
    # Cargar Maestro de Choferes
    df_maestro = pd.read_excel(file_maestro)
    # Limpieza de patente en el maestro
    df_maestro['PATENTE_CLEAN'] = df_maestro['PATENTE'].astype(str).str.replace(" ", "").str.upper()
    
    # Cargar Detalle de Consumos
    if file_factura.name.endswith('.xlsx'):
        df_consumos = pd.read_excel(file_factura)
    else:
        df_consumos = pd.read_csv(file_factura)

    # Limpieza de patente en consumos si viene esa columna
    if 'PATENTE' in df_consumos.columns:
        df_consumos['PATENTE_CLEAN'] = df_consumos['PATENTE'].astype(str).str.replace(" ", "").str.upper()

    # --- SECCIÓN 2: DATOS DE CABECERA DE LA FACTURA ---
    st.subheader("📝 Datos de Cabecera del Comprobante")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        nro_interno = st.text_input("Número Interno (Agrupador)", "13393")
    with c2:
        fecha_fc = st.text_input("Fecha Comprobante", "14/08/2026")
    with c3:
        proveedor_cod = st.text_input("Código Proveedor (CUIT)", "30646766369")
    with c4:
        nro_comprobante = st.text_input("N° Comprobante", "A-00098-00040851")

    # --- SECCIÓN 3: CRUCE Y VALIDACIÓN ---
    st.markdown("---")
    st.subheader("🔍 Validando Choferes y Centros de Costo")

    # Cruce por Chofer o Patente segun disponibilidad
    if 'CHOFER' in df_consumos.columns:
        df_merged = pd.merge(df_consumos, df_maestro[['CHOFER', 'REPARTO', 'PATENTE_CLEAN']], on='CHOFER', how='left')
    elif 'PATENTE_CLEAN' in df_consumos.columns:
        df_merged = pd.merge(df_consumos, df_maestro[['REPARTO', 'PATENTE_CLEAN']].drop_duplicates(), on='PATENTE_CLEAN', how='left')
    else:
        df_merged = df_consumos.copy()
        df_merged['REPARTO'] = None

    # Detectar Faltantes / Errores
    faltantes = df_merged[df_merged['REPARTO'].isna()]

    if not faltantes.empty:
        st.error(f"⚠️ Atención: Se encontraron {len(faltantes)} renglones sin Centro de Costo / Reparto asignado.")
        st.write("Registros que requieren corrección antes de exportar:")
        st.dataframe(faltantes)
    else:
        st.success("✅ Excelente: Todos los ítems tienen su Centro de Costo asignado correctamente.")

    # --- SECCIÓN 4: ARMAR PLANTILLA FINNEGANS ---
    st.markdown("---")
    st.subheader("📊 Vista Previa del Excel Final (Formato Finnegans)")

    # Construcción del dataframe que va al Excel de Finnegans
    df_finnegans = pd.DataFrame()
    df_finnegans['Número'] = [nro_interno] * len(df_merged)
    df_finnegans['Fecha'] = [fecha_fc] * len(df_merged)
    df_finnegans['Proveedor'] = [proveedor_cod] * len(df_merged)
    df_finnegans['Comprobante'] = [nro_comprobante] * len(df_merged)
    df_finnegans['Condición de Pago'] = ['CUENTA CORRIENTE 7 DÍAS'] * len(df_merged)
    df_finnegans['Moneda'] = ['ARS'] * len(df_merged)
    df_finnegans['Producto'] = ['COMB'] * len(df_merged)
    df_finnegans['Cantidad'] = df_merged['Cantidad'] if 'Cantidad' in df_merged.columns else 1.0
    df_finnegans['Precio'] = df_merged['Precio'] if 'Precio' in df_merged.columns else df_merged['Bruto']
    df_finnegans['Dimensión'] = 'DIMCTC'
    df_finnegans['Valor de dimensión'] = df_merged['REPARTO']

    st.dataframe(df_finnegans)

    # --- SECCIÓN 5: DESCARGAR EXCEL ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_finnegans.to_excel(writer, index=False, sheet_name='Factura_Importar')
    
    st.download_button(
        label="📥 Descargar Excel Listo para Finnegans",
        data=output.getvalue(),
        file_name=f"Importacion_Finnegans_{nro_comprobante}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("👈 Por favor, subí el Maestro de Choferes y el archivo de la Factura desde la barra lateral izquierda para empezar.")
