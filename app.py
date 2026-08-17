import streamlit as st
import pandas as pd
import pypdf
import re
import io

st.set_page_config(page_title="Cargador de Facturas de Combustible", layout="wide")

st.title("⛽ Importador Masivo de Facturas de Combustible para Finnegans")
st.markdown("Herramienta de validación de Choferes, Patentes y Centros de Costo")

# --- SECCIÓN 1: CARGA DE ARCHIVOS ---
st.sidebar.header("📁 Carga de Archivos")

file_maestro = st.sidebar.file_uploader("1. Maestro de Choferes (.xlsx)", type=["xlsx"])
file_factura = st.sidebar.file_uploader("2. Detalle de Consumos / Factura (.xlsx, .csv o .pdf)", type=["xlsx", "csv", "pdf"])

# Variables para cabecera por defecto
auto_nro_int = "13393"
auto_fecha = "14/08/2026"
auto_proveedor = "30646766369"
auto_comprobante = "A-00098-00040851"

df_consumos = None

if file_factura:
    if file_factura.name.endswith('.pdf'):
        # Extracción automática desde PDF
        reader = pypdf.PdfReader(file_factura)
        pdf_text = "\n".join([page.extract_text() for page in reader.pages if page.extract_text()])
        
        # Extraer cabecera del PDF
        comp_m = re.search(r'([A-Z]-\d{5}-\d{8})', pdf_text)
        if comp_m: auto_comprobante = comp_m.group(1)
            
        fecha_m = re.search(r'(\d{2}/\d{2}/\d{4})', pdf_text)
        if fecha_m: auto_fecha = fecha_m.group(1)
            
        num_m = re.search(r'Numero Interno:\s*(\d+)', pdf_text)
        if num_m: auto_nro_int = num_m.group(1)
            
        cuit_m = re.search(r'C\.U\.I\.T\.\s*(\d{2}-\d{8}-\d{1})', pdf_text)
        if cuit_m: auto_proveedor = cuit_m.group(1).replace('-', '')

        # Extraer ítems de consumo
        pattern = r'(COMBUSTIBLES\s+REPARTO.*?)\s+Litros\s*(\d+(?:\.\d+)?)\s+([\d\.,]+)\s+([\d\.,]+)'
        matches = re.findall(pattern, pdf_text)
        
        rows = []
        for m in matches:
            rows.append({
                'Descripción': m[0].strip(),
                'Cantidad': float(m[1]),
                'Precio': float(m[2].replace('.', '').replace(',', '.')),
                'Bruto': float(m[3].replace('.', '').replace(',', '.'))
            })
        df_consumos = pd.DataFrame(rows)
        
    elif file_factura.name.endswith('.xlsx'):
        df_consumos = pd.read_excel(file_factura)
    else:
        df_consumos = pd.read_csv(file_factura)

if file_maestro and df_consumos is not None:
    # Cargar y limpiar Maestro
    df_maestro = pd.read_excel(file_maestro)
    df_maestro['PATENTE_CLEAN'] = df_maestro['PATENTE'].astype(str).str.replace(" ", "").str.upper()

    if 'PATENTE' in df_consumos.columns:
        df_consumos['PATENTE_CLEAN'] = df_consumos['PATENTE'].astype(str).str.replace(" ", "").str.upper()

    # --- SECCIÓN 2: DATOS DE CABECERA DE LA FACTURA ---
    st.subheader("📝 Datos de Cabecera del Comprobante")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        nro_interno = st.text_input("Número Interno (Agrupador)", auto_nro_int)
    with c2:
        fecha_fc = st.text_input("Fecha Comprobante", auto_fecha)
    with c3:
        proveedor_cod = st.text_input("Código Proveedor (CUIT)", auto_proveedor)
    with c4:
        nro_comprobante = st.text_input("N° Comprobante", auto_comprobante)

    # --- SECCIÓN 3: CRUCE Y VALIDACIÓN ---
    st.markdown("---")
    st.subheader("🔍 Validando Choferes y Centros de Costo")

    if 'CHOFER' in df_consumos.columns:
        df_merged = pd.merge(df_consumos, df_maestro[['CHOFER', 'REPARTO', 'PATENTE_CLEAN']], on='CHOFER', how='left')
    elif 'PATENTE_CLEAN' in df_consumos.columns:
        df_merged = pd.merge(df_consumos, df_maestro[['REPARTO', 'PATENTE_CLEAN']].drop_duplicates(), on='PATENTE_CLEAN', how='left')
    else:
        df_merged = df_consumos.copy()
        df_merged['REPARTO'] = None

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
    st.info("👈 Por favor, subí el Maestro de Choferes y la Factura (.pdf, .xlsx o .csv) para comenzar.")
