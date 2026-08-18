import streamlit as st
import pandas as pd
import pypdf
import re
import io

st.set_page_config(page_title="Cargador de Facturas de Combustible", layout="wide")

st.title("⛽ Importador Masivo de Facturas de Combustible para Finnegans")
st.markdown("Generador automático compatible con la Plantilla Oficial de Finnegans")

# --- SECCIÓN 1: CARGA DE ARCHIVOS ---
st.sidebar.header("📁 Carga de Archivos")

file_maestro = st.sidebar.file_uploader("1. Maestro de Choferes (.xlsx)", type=["xlsx"])
file_factura = st.sidebar.file_uploader("2. Factura A - Petroeste (.pdf)", type=["pdf"])
file_remitos = st.sidebar.file_uploader("3. Listado de Remitos (.pdf, .xlsx, .csv)", type=["pdf", "xlsx", "csv"])

# Precios Unitarios NETOS de Factura (Columna Amarilla)
precios_unitarios_netos = {
    'DIESEL 500': 1627.88,
    'INFINIA DIESEL': 1921.82,
    'NAFTA SUPER': 1441.45,
    'NAFTA INFINIA': 1634.92
}

auto_nro_int = "13393"
auto_fecha = "14/08/2026"
auto_proveedor = "30646766369"
auto_comprobante = "A-00098-00040851"
auto_condicion_pago = "CC7"

# 1. Leer Datos de la Factura A
if file_factura:
    reader_fc = pypdf.PdfReader(file_factura)
    txt_fc = "\n".join([page.extract_text() for page in reader_fc.pages if page.extract_text()])
    
    comp_m = re.search(r'([A-Z]-\d{4,5}-\d{8})', txt_fc)
    if comp_m: auto_comprobante = comp_m.group(1)
        
    fecha_m = re.search(r'(\d{2}/\d{2}/\d{4})', txt_fc)
    if fecha_m: auto_fecha = fecha_m.group(1)
        
    num_m = re.search(r'Numero Interno:\s*(\d+)', txt_fc)
    if num_m: auto_nro_int = num_m.group(1)
        
    cuit_m = re.search(r'C\.U\.I\.T\.\s*(\d{2}-\d{8}-\d{1})', txt_fc)
    if cuit_m: auto_proveedor = cuit_m.group(1).replace('-', '')

df_consumos = None

# 2. Leer Remitos (Litros y Patentes)
if file_remitos:
    if file_remitos.name.endswith('.pdf'):
        reader_rm = pypdf.PdfReader(file_remitos)
        pdf_text = "\n".join([page.extract_text() for page in reader_rm.pages if page.extract_text()])
        
        lines = pdf_text.split('\n')
        records = []
        current_chofer, current_patente = "", ""

        for line in lines:
            line_s = line.strip()
            if re.match(r'^\d{2}/\d{2}\s+\d{2}:\d{2}\s+\d{4}-\d{8}', line_s):
                fuel_m = re.search(r'(DIESEL\s+\d+|INFINIA\s+DIESEL|NAFTA\s+SUPER|NAFTA\s+INFINIA)', line_s)
                if fuel_m:
                    fuel = fuel_m.group(1)
                    before_fuel = line_s[:fuel_m.start()]
                    after_fuel = line_s[fuel_m.start():]
                    clean_before = re.sub(r'^\d{2}/\d{2}\s+\d{2}:\d{2}\s+\d{4}-\d{8}\s+', '', before_fuel).strip()
                    
                    nums = re.findall(r'[\d\.]+\,\d+', after_fuel)
                    qty = float(nums[0].replace('.', '').replace(',', '.')) if len(nums) >= 1 else 1.0
                    
                    pat_m = re.search(r'([A-Z]{2,3}\s*\d{3}\s*[A-Z]{0,2}|[A-Z]{3}\s*\d{3})', clean_before)
                    patente = pat_m.group(1).replace(" ", "") if pat_m else ""
                    chofer = clean_before[:pat_m.start()].strip() if pat_m else clean_before
                    
                    current_chofer, current_patente = chofer, patente
                    records.append({'CHOFER': chofer, 'PATENTE': patente, 'Artículo': fuel, 'Litros': qty})
            else:
                fuel_m = re.search(r'^(DIESEL\s+\d+|INFINIA\s+DIESEL|NAFTA\s+SUPER|NAFTA\s+INFINIA)', line_s)
                if fuel_m:
                    fuel = fuel_m.group(1)
                    nums = re.findall(r'[\d\.]+\,\d+', line_s)
                    qty = float(nums[0].replace('.', '').replace(',', '.')) if len(nums) >= 1 else 1.0
                    records.append({'CHOFER': current_chofer, 'PATENTE': current_patente, 'Artículo': fuel, 'Litros': qty})

        df_consumos = pd.DataFrame(records)
    elif file_remitos.name.endswith('.xlsx'):
        df_consumos = pd.read_excel(file_remitos)
    else:
        df_consumos = pd.read_csv(file_remitos)

if file_maestro and df_consumos is not None and not df_consumos.empty:
    df_maestro = pd.read_excel(file_maestro)
    df_maestro['PATENTE_CLEAN'] = df_maestro['PATENTE'].astype(str).str.replace(" ", "").str.upper()

    if 'PATENTE' in df_consumos.columns:
        df_consumos['PATENTE_CLEAN'] = df_consumos['PATENTE'].astype(str).str.replace(" ", "").str.upper().str.rstrip('A')

    # --- SECCIÓN 2: DATOS DE CABECERA ---
    st.subheader("📝 Datos de Cabecera del Comprobante")
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: nro_interno = st.text_input("Número Interno (NUMERO)", auto_nro_int)
    with c2: fecha_fc = st.text_input("Fecha (FECHA)", auto_fecha)
    with c3: proveedor_cod = st.text_input("Proveedor (CUIT)", auto_proveedor)
    with c4: nro_comprobante = st.text_input("Comprobante", auto_comprobante)
    with c5: cond_pago = st.text_input("Condición de Pago", auto_condicion_pago)

    # --- SECCIÓN 3: VALIDACIÓN Y CRUCE ---
    st.markdown("---")
    st.subheader("🔍 Validando Choferes, Precios Unitarios Netos y Centros de Costo")

    df_consumos['Precio_Unitario_Neto'] = df_consumos['Artículo'].map(precios_unitarios_netos).fillna(1.0)

    if 'PATENTE_CLEAN' in df_consumos.columns:
        df_merged = pd.merge(df_consumos, df_maestro[['REPARTO', 'PATENTE_CLEAN']].drop_duplicates(), on='PATENTE_CLEAN', how='left')
    else:
        df_merged = df_consumos.copy()
        df_merged['REPARTO'] = None

    faltantes = df_merged[df_merged['REPARTO'].isna()]

    if not faltantes.empty:
        st.error(f"⚠️ Atención: Se encontraron {len(faltantes)} renglones sin Centro de Costo asignado.")
        st.dataframe(faltantes[['CHOFER', 'PATENTE', 'Artículo', 'Litros', 'Precio_Unitario_Neto']])
    else:
        st.success("✅ Excelente: Todos los consumos tienen asignado el precio unitario neto y su centro de costo.")

    # --- SECCIÓN 4: GENERACIÓN DE LA PLANTILLA ESTRUCTURADA ---
    st.markdown("---")
    st.subheader("📊 Vista Previa del Excel de Importación Masiva (Formato Oficial Finnegans)")

    # Construcción de las 36 columnas oficiales del Exportador Masivo
    df_finnegans = pd.DataFrame({
        'NUMERO': [nro_interno] * len(df_merged),
        'FECHA': [fecha_fc] * len(df_merged),
        'PROVEEDOR': [proveedor_cod] * len(df_merged),
        'COMPROBANTE': [nro_comprobante] * len(df_merged),
        'CONDICIONPAGO': [cond_pago] * len(df_merged),
        'SUCURSAL': None,
        'DESCRIPCION': 'Factura Combustibles Repartos',
        'PRODUCTO': ['COMB'] * len(df_merged),
        'DESCRIPCIONITEM': df_merged['Artículo'],
        'CANTIDAD': df_merged['Litros'],
        'PRECIO': df_merged['Precio_Unitario_Neto'],
        'PRECIOSOBRE': None,
        'MONEDA_COTIZACION': [1] * len(df_merged),
        'COTIZACION': [1] * len(df_merged),
        'MONEDA': ['ARS'] * len(df_merged),
        'WORKFLOW': None,
        'FECHACOMPROBANTE': [fecha_fc] * len(df_merged),
        'FECHABASEVENCIMIENTO': None,
        'DESTINATARIO': None,
        'PROVINCIA_DESTINO': None,
        'PROVINCIA_DESTINO_ITEM': None,
        'FECHAPROXIMOPASO': None,
        'DIMENSION': ['DIMCTC'] * len(df_merged),
        'DIMENSIONVALOR': df_merged['REPARTO'],
        'DIMENSION2': None,
        'DIMENSIONVALOR2': None,
        'DIMENSION3': None,
        'DIMENSIONVALOR3': None,
        'DIMENSIONVTO': None,
        'DIMENSIONVALORVTO': None,
        'DIMENSIONVTO2': None,
        'DIMENSIONVALORVTO2': None,
        'IMPORTE_CONTROL': None,
        'EQUIPO_SOLICITANTE': None,
        'TIPO_ITEM': None,
        'FECHAHASTA': None
    })

    st.dataframe(df_finnegans)

    # --- SECCIÓN 5: DESCARGA EN FORMATO .XLS / .XLSX ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_finnegans.to_excel(writer, index=False, sheet_name='Sheet0')
    
    st.download_button(
        label="📥 Descargar Plantilla Oficial para Finnegans (.xlsx)",
        data=output.getvalue(),
        file_name=f"Importacion_Finnegans_{nro_comprobante}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("👈 Subí el Maestro de Choferes, la Factura A y el Listado de Remitos para generar la plantilla.")
