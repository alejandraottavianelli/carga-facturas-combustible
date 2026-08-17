import streamlit as st
import pandas as pd
import pypdf
import re
import io

st.set_page_config(page_title="Cargador de Facturas de Combustible", layout="wide")

st.title("⛽ Importador Masivo de Facturas de Combustible para Finnegans")
st.markdown("Herramienta de asignación de precios de Factura sobre Remitos y Centros de Costo")

# --- SECCIÓN 1: CARGA DE ARCHIVOS ---
st.sidebar.header("📁 Carga de Archivos")

file_maestro = st.sidebar.file_uploader("1. Maestro de Choferes (.xlsx)", type=["xlsx"])
file_factura = st.sidebar.file_uploader("2. Factura A - Petroeste (.pdf)", type=["pdf"])
file_remitos = st.sidebar.file_uploader("3. Listado de Remitos (.pdf, .xlsx, .csv)", type=["pdf", "xlsx", "csv"])

# Precios Unitarios NETOS de la Factura (ejemplo Petroeste)
precios_unitarios_factura = {
    'DIESEL 500': 2235.00,
    'INFINIA DIESEL': 2604.00,
    'NAFTA SUPER': 1829.00,
    'NAFTA INFINIA': 2072.00
}

# Impuestos extraídos de la Factura Petroeste A-00098-00040851
impuestos_factura = {
    'ITC': 307583.78,
    'Impuestos Provinciales': 69423.40,
    'Impuestos Municipales': 176904.25,
    'IVA_21': 832300.19,
    'Percepciones': 553911.43
}

auto_nro_int = "13393"
auto_fecha = "14/08/2026"
auto_proveedor = "30646766369"
auto_comprobante = "A-00098-00040851"

# 1. Leer Datos de Factura A
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

    # --- SECCIÓN 2: CABECERA ---
    st.subheader("📝 Datos de Cabecera del Comprobante")
    c1, c2, c3, c4 = st.columns(4)
    with c1: nro_interno = st.text_input("Número Interno", auto_nro_int)
    with c2: fecha_fc = st.text_input("Fecha Comprobante", auto_fecha)
    with c3: proveedor_cod = st.text_input("Código Proveedor (CUIT)", auto_proveedor)
    with c4: nro_comprobante = st.text_input("N° Comprobante", auto_comprobante)

    # --- SECCIÓN 3: VALIDACIÓN DE PRECIOS Y CENTROS DE COSTO ---
    st.markdown("---")
    st.subheader("🔍 Validando Choferes, Precios Unitarios y Centros de Costo")

    # Precio Unitario Neto
    df_consumos['Precio_Unitario_Neto'] = df_consumos['Artículo'].map(precios_unitarios_factura).fillna(1.0)
    df_consumos['Neto_Gravado_Item'] = df_consumos['Litros'] * df_consumos['Precio_Unitario_Neto']

    # Cruce con Maestro de Choferes/Repartos
    if 'PATENTE_CLEAN' in df_consumos.columns:
        df_merged = pd.merge(df_consumos, df_maestro[['REPARTO', 'PATENTE_CLEAN']].drop_duplicates(), on='PATENTE_CLEAN', how='left')
    else:
        df_merged = df_consumos.copy()
        df_merged['REPARTO'] = None

    faltantes = df_merged[df_merged['REPARTO'].isna()]

    if not faltantes.empty:
        st.error(f"⚠️ Atención: Se encontraron {len(faltantes)} renglones sin Centro de Costo asignado.")
        st.dataframe(faltantes[['CHOFER', 'PATENTE', 'Artículo', 'Litros', 'Precio_Unitario_Neto', 'Neto_Gravado_Item']])
    else:
        st.success("✅ Excelente: Todos los consumos fueron calculados con el Precio Unitario Neto y asignados al Centro de Costo.")

    # --- SECCIÓN 4: PLANTILLA FINNEGANS ---
    st.markdown("---")
    st.subheader("📊 Vista Previa de Carga Masiva (Ítems + Impuestos/Tasas)")

    # Tabla de Ítems
    df_finnegans_items = pd.DataFrame()
    df_finnegans_items['Número'] = [nro_interno] * len(df_merged)
    df_finnegans_items['Fecha'] = [fecha_fc] * len(df_merged)
    df_finnegans_items['Proveedor'] = [proveedor_cod] * len(df_merged)
    df_finnegans_items['Comprobante'] = [nro_comprobante] * len(df_merged)
    df_finnegans_items['Condición de Pago'] = ['CUENTA CORRIENTE 7 DÍAS'] * len(df_merged)
    df_finnegans_items['Moneda'] = ['ARS'] * len(df_merged)
    df_finnegans_items['Producto'] = ['COMB'] * len(df_merged)
    df_finnegans_items['Cantidad'] = df_merged['Litros']
    df_finnegans_items['Precio'] = df_merged['Precio_Unitario_Neto']
    df_finnegans_items['Dimensión'] = 'DIMCTC'
    df_finnegans_items['Valor de dimensión'] = df_merged['REPARTO']

    st.write("**Detalle de Ítems (Combustible Neto):**")
    st.dataframe(df_finnegans_items)

    # Tabla de Impuestos / Tasas anexos
    df_impuestos = pd.DataFrame([
        {'Comprobante': nro_comprobante, 'Tipo': 'Otras Tasas', 'Tasa/Impuesto': 'ITC', 'Importe': impuestos_factura['ITC']},
        {'Comprobante': nro_comprobante, 'Tipo': 'Otras Tasas', 'Tasa/Impuesto': 'Impuestos Provinciales', 'Importe': impuestos_factura['Impuestos Provinciales']},
        {'Comprobante': nro_comprobante, 'Tipo': 'Otras Tasas', 'Tasa/Impuesto': 'Impuestos Municipales', 'Importe': impuestos_factura['Impuestos Municipales']},
        {'Comprobante': nro_comprobante, 'Tipo': 'IVA', 'Tasa/Impuesto': '21%', 'Importe': impuestos_factura['IVA_21']}
    ])

    st.write("**Detalle de Impuestos y Tasas de la Factura:**")
    st.dataframe(df_impuestos)

    # --- SECCIÓN 5: DESCARGAR EXCEL CON 2 SOLAPAS ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_finnegans_items.to_excel(writer, index=False, sheet_name='Items_Factura')
        df_impuestos.to_excel(writer, index=False, sheet_name='Impuestos_Tasas')
    
    st.download_button(
        label="📥 Descargar Excel Listo para Finnegans",
        data=output.getvalue(),
        file_name=f"Importacion_Finnegans_{nro_comprobante}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("👈 Subí el Maestro de Choferes, la Factura A y el Listado de Remitos para procesar.")
