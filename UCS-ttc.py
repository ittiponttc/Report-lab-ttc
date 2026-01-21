import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import font_manager
import io
from datetime import datetime, date
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter # เพิ่ม Import นี้สำหรับแก้บั๊ก Excel
from openpyxl.chart import ScatterChart, Reference, Series
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import base64

# Page config
st.set_page_config(
    page_title="UCS Test Report Generator",
    page_icon="🧪",
    layout="wide"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2rem;
        font-weight: bold;
        color: #1e3a5f;
        text-align: center;
        padding: 1rem;
        background: linear-gradient(90deg, #e8f4f8 0%, #d1e8f0 100%);
        border-radius: 10px;
        margin-bottom: 2rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #2c5282;
        border-bottom: 2px solid #3182ce;
        padding-bottom: 0.5rem;
        margin-top: 1.5rem;
    }
</style>
""", unsafe_allow_html=True)

# Title
st.markdown('<div class="main-header">🧪 Unconfined Compression Test (ASTM D2166)<br>Report Generator</div>', unsafe_allow_html=True)
st.markdown('<p style="text-align: center; color: #666;">King Mongkut\'s University of Technology North Bangkok</p>', unsafe_allow_html=True)

# Initialize session state
if 'calculated' not in st.session_state:
    st.session_state.calculated = False
if 'results' not in st.session_state:
    st.session_state.results = None

# Sidebar for project information
st.sidebar.markdown("## 📋 Project Information")

with st.sidebar:
    project_name = st.text_input("Project Name / ชื่อโครงการ", "")
    specimen_from = st.text_input("Specimen From / ตัวอย่างจาก", "")
    location = st.text_input("Location / สถานที่", "")
    column_no = st.text_input("Column No. / หมายเลขเสาเข็ม", "")
    depth = st.text_input("Depth (m) / ความลึก", "7.50-8.50")
    sample_number = st.number_input("Sample Number / หมายเลขตัวอย่าง", min_value=1, value=1)
    
    st.markdown("---")
    st.markdown("### 📅 Dates")
    date_of_jetting = st.text_input("Date of Jetting / วันที่ฉีด (Text)", str(date.today()))
    date_of_testing = st.date_input("Date of Testing / วันที่ทดสอบ", date.today())
    curing_time = st.number_input("Curing Time (days) / เวลาบ่ม", min_value=1, value=28)
    
    st.markdown("---")
    st.markdown("### 👤 Personnel")
    tested_by = st.text_input("Tested by / ผู้ทดสอบ", "Asst.Prof.Dr.Ittipon Meepon")
    controlled_by = st.text_input("Controlled Test by", "")
    notified_by = st.text_input("Notified by", "")

# Main content
col1, col2 = st.columns([1, 1])

with col1:
    st.markdown('<p class="sub-header">📊 Sample Properties</p>', unsafe_allow_html=True)
    
    subcol1, subcol2 = st.columns(2)
    with subcol1:
        diameter = st.number_input("Diameter, D (cm)", min_value=0.1, value=5.53, step=0.01, format="%.2f")
        height = st.number_input("Height, H (cm)", min_value=0.1, value=11.90, step=0.01, format="%.2f")
        weight = st.number_input("Weight of Sample, W (g)", min_value=0.1, value=621.07, step=0.01, format="%.2f")
    
    with subcol2:
        shearing_rate = st.text_input("Shearing Rate", "1% of sample height per minute")
        mixed_jet_mixing = st.text_input("Mixed Jet Mixing (kg/m³)", "-")
    
    st.markdown('<p class="sub-header">💧 Water Content Determination</p>', unsafe_allow_html=True)
    
    wcol1, wcol2 = st.columns(2)
    with wcol1:
        container_no = st.text_input("Container No.", "9")
        weight_wet_soil_can = st.number_input("Weight of Wet Soil + Can (g)", min_value=0.0, value=406.14, step=0.01)
        weight_dry_soil_can = st.number_input("Weight of Dry Soil + Can (g)", min_value=0.0, value=244.31, step=0.01)
    with wcol2:
        weight_can = st.number_input("Weight of Can (g)", min_value=0.0, value=26.05, step=0.01)

with col2:
    st.markdown('<p class="sub-header">📁 Upload Test Data</p>', unsafe_allow_html=True)
    
    st.markdown("""
    <div class="info-box">
    <strong>📌 Required Format:</strong><br>
    Excel file with columns: <code>Load (kg)</code> and <code>Vertical displacement (mm)</code>
    </div>
    """, unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls'])
    
    proving_ring_capacity = st.number_input("Proving Ring Capacity (kN)", min_value=0.1, value=20.0, step=0.1)
    factor_k = st.number_input("Factor K (kg/division)", min_value=0.0001, value=1.8661, step=0.0001, format="%.4f")

# Calculate derived values
area = np.pi * (diameter / 2) ** 2  # cm²
volume = area * height  # cm³
wet_unit_weight = weight / volume if volume > 0 else 0  # g/cm³

# Water content calculation
weight_water = weight_wet_soil_can - weight_dry_soil_can
weight_dry_soil = weight_dry_soil_can - weight_can
water_content = (weight_water / weight_dry_soil * 100) if weight_dry_soil > 0 else 0
dry_unit_weight = wet_unit_weight / (1 + water_content/100) if water_content > 0 else wet_unit_weight

# Display calculated properties
st.markdown('<p class="sub-header">📐 Calculated Sample Properties</p>', unsafe_allow_html=True)
prop_col1, prop_col2, prop_col3, prop_col4 = st.columns(4)
with prop_col1:
    st.metric("Area, A (cm²)", f"{area:.2f}")
with prop_col2:
    st.metric("Volume, V (cm³)", f"{volume:.2f}")
with prop_col3:
    st.metric("Wet Unit Weight (g/cm³)", f"{wet_unit_weight:.2f}")
with prop_col4:
    st.metric("Dry Unit Weight (g/cm³)", f"{dry_unit_weight:.2f}")

# Process uploaded data
if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        
        # Check for required columns
        required_cols = ['Load (kg)', 'Vertical displacement (mm)']
        if not all(col in df.columns for col in required_cols):
            st.error(f"❌ Required columns not found. Please ensure file has: {required_cols}")
        else:
            st.success("✅ Data loaded successfully!")
            
            # Calculate stress-strain data
            df['Deformation, R (mm)'] = df['Vertical displacement (mm)']
            df['Axial Load, P (kg)'] = df['Load (kg)']
            df['Axial Strain, ξ (R/10H)'] = df['Deformation, R (mm)'] / (10 * height)
            df['Axial Strain (%)'] = df['Axial Strain, ξ (R/10H)'] * 100
            df['Corrected Area, Ac (cm²)'] = area / (1 - df['Axial Strain, ξ (R/10H)'])
            df['Axial Stress, σ (ksc)'] = (df['Axial Load, P (kg)'] / df['Corrected Area, Ac (cm²)'])
            
            # Find maximum stress (UCS)
            max_stress_idx = df['Axial Stress, σ (ksc)'].idxmax()
            ucs = df.loc[max_stress_idx, 'Axial Stress, σ (ksc)']
            failure_strain = df.loc[max_stress_idx, 'Axial Strain (%)']
            undrained_shear_strength = ucs / 2
            
            # Calculate Modulus of Elasticity (E50)
            half_ucs = ucs / 2
            df_before_peak = df[df.index <= max_stress_idx]
            idx_e50 = (df_before_peak['Axial Stress, σ (ksc)'] - half_ucs).abs().idxmin()
            strain_at_half = df.loc[idx_e50, 'Axial Strain (%)']
            modulus_e50 = (half_ucs / (strain_at_half / 100)) if strain_at_half > 0 else 0
            
            # Display results
            st.markdown('<p class="sub-header">🎯 Test Results</p>', unsafe_allow_html=True)
            
            res_col1, res_col2, res_col3, res_col4 = st.columns(4)
            with res_col1:
                st.metric("Unconfined Compressive Strength", f"{ucs:.2f} ksc")
            with res_col2:
                st.metric("Undrained Shear Strength", f"{undrained_shear_strength:.2f} ksc")
            with res_col3:
                st.metric("Failure Strain", f"{failure_strain:.2f} %")
            with res_col4:
                st.metric("Modulus E50", f"{modulus_e50:.2f} ksc")
            
            # Plot - wide format for bottom of page
            fig, ax = plt.subplots(figsize=(10, 4)) # Wide and short
            ax.plot(df['Axial Strain (%)'], df['Axial Stress, σ (ksc)'], 'b-o', markersize=3, linewidth=1)
            ax.set_xlabel('Axial Strain, ε (%)', fontsize=9)
            ax.set_ylabel('Axial Stress, σ (ksc)', fontsize=9)
            ax.grid(True, linestyle='--', which='both', alpha=0.6)
            ax.minorticks_on()
            
            # Save plot to buffer
            img_buffer = io.BytesIO()
            fig.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight')
            img_buffer.seek(0)
            
            st.pyplot(fig)
            
            # Download section
            st.markdown('<p class="sub-header">💾 Download Reports</p>', unsafe_allow_html=True)
            
            download_col1, download_col2 = st.columns(2)
            
            # --- HELPER FOR WORD BORDERS ---
            def set_cell_border(cell, **kwargs):
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                for edge in ('top', 'start', 'bottom', 'end', 'left', 'right'):
                    if edge in kwargs:
                        edge_data = kwargs[edge]
                        tag = 'w:{}'.format(edge)
                        element = tcPr.find(qn(tag))
                        if element is None:
                            element = OxmlElement(tag)
                            tcPr.append(element)
                        for key in ["sz", "val", "color", "space", "shadow"]:
                            if key in edge_data:
                                element.set(qn('w:{}'.format(key)), str(edge_data[key]))

            with download_col1:
                # Excel Export
                def create_excel_report():
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "UCS Test Report"
                    
                    # (Keep existing Excel Logic, simplified for brevity but crucial parts retained)
                    ws.append(["KING MONGKUT'S UNIVERSITY OF TECHNOLOGY NORTH BANGKOK"])
                    ws.append(["Unconfined Compression Test (ASTM D2166)"])
                    ws.append([])
                    
                    # Data dump
                    headers = ['R (mm)', 'P (kg)', 'ξ', 'Ac (cm²)', 'ε (%)', 'σ (ksc)']
                    ws.append(headers)
                    for row in df[['Deformation, R (mm)', 'Axial Load, P (kg)', 
                                   'Axial Strain, ξ (R/10H)', 'Corrected Area, Ac (cm²)',
                                   'Axial Strain (%)', 'Axial Stress, σ (ksc)']].values:
                        ws.append(list(row))
                        
                    # FIX: Use get_column_letter for MergedCell error
                    from openpyxl.utils import get_column_letter
                    for column_cells in ws.columns:
                        length = max(len(str(cell.value) if cell.value else "") for cell in column_cells)
                        col_idx = column_cells[0].column
                        col_letter = get_column_letter(col_idx) # Correct way
                        ws.column_dimensions[col_letter].width = length + 2
                    
                    excel_buffer = io.BytesIO()
                    wb.save(excel_buffer)
                    excel_buffer.seek(0)
                    return excel_buffer
                
                excel_data = create_excel_report()
                st.download_button(
                    label="📥 Download Excel Report",
                    data=excel_data,
                    file_name=f"UCS_Data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with download_col2:
                # Word Export - THE ONE PAGE LAYOUT
                def create_word_report():
                    doc = Document()
                    
                    # 1. SETUP MARGINS (NARROW)
                    section = doc.sections[0]
                    section.top_margin = Cm(1.27)
                    section.bottom_margin = Cm(1.27)
                    section.left_margin = Cm(1.27)
                    section.right_margin = Cm(1.27)
                    
                    style = doc.styles['Normal']
                    style.font.name = 'Times New Roman'
                    style.font.size = Pt(9)

                    # 2. HEADER
                    header_p = doc.add_paragraph()
                    header_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run1 = header_p.add_run("KING MONGKUT'S UNIVERSITY OF TECHNOLOGY NORTH BANGKOK\n")
                    run1.bold = True
                    run1.font.size = Pt(11)
                    run2 = header_p.add_run("DEPARTMENT OF TEACHER TRAINING IN CIVIL ENGINEERING\n")
                    run2.bold = True
                    run2.font.color.rgb = RGBColor(255, 0, 0)
                    run3 = header_p.add_run("Unconfined Compression Test (ASTM D2166)")
                    run3.bold = True
                    run3.font.size = Pt(10)

                    # 3. PROJECT INFO (Compact Table)
                    info_table = doc.add_table(rows=4, cols=4)
                    info_table.style = 'Table Grid'
                    info_table.autofit = False 
                    
                    # Set widths roughly
                    for row in info_table.rows:
                        row.cells[0].width = Cm(2.5)
                        row.cells[1].width = Cm(6.5)
                        row.cells[2].width = Cm(3.0)
                        row.cells[3].width = Cm(6.5)

                    # Fill Info
                    info_map = [
                        (0, "Project Name :", project_name, "Tested by :", tested_by),
                        (1, "Location :", location, "Date of Jetting :", date_of_jetting),
                        (2, "Column No. :", column_no, "Date of Testing :", str(date_of_testing)),
                        (3, "Depth :", depth, "Curing Time :", f"{curing_time} days")
                    ]
                    
                    for r, l1, v1, l2, v2 in info_map:
                        info_table.cell(r,0).text = l1
                        info_table.cell(r,1).text = str(v1)
                        info_table.cell(r,2).text = l2
                        info_table.cell(r,3).text = str(v2)
                    
                    doc.add_paragraph() # Spacer

                    # 4. MAIN BODY (Split into Left: Data, Right: Props/Results)
                    # Create a master table with 1 row, 2 columns (invisible borders)
                    body_table = doc.add_table(rows=1, cols=2)
                    body_table.autofit = False
                    body_table.columns[0].width = Cm(8.0)  # Left: Data
                    body_table.columns[1].width = Cm(10.5) # Right: Props & Results & Photo

                    # --- LEFT COLUMN: DATA TABLE ---
                    left_cell = body_table.cell(0, 0)
                    left_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
                    
                    # Add data table inside left cell
                    p_data = left_cell.paragraphs[0]
                    p_data.add_run("Test Data").bold = True
                    
                    # Calculate how many rows fit (approx 35-40 for one page)
                    display_rows = len(df)
                    
                    dt = left_cell.add_table(rows=display_rows + 2, cols=5)
                    dt.style = 'Table Grid'
                    dt.autofit = False
                    
                    headers = ['R\n(mm)', 'P\n(kg)', 'Strain\n(%)', 'Area\n(cm²)', 'Stress\n(ksc)']
                    for i, h in enumerate(headers):
                        cell = dt.cell(0, i)
                        cell.text = h
                        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        cell.paragraphs[0].runs[0].font.size = Pt(8)
                        cell.paragraphs[0].runs[0].font.bold = True
                        dt.column_cells(i)[0].width = Cm(1.5)

                    for r_idx, row_data in enumerate(df[['Deformation, R (mm)', 'Axial Load, P (kg)', 
                                                         'Axial Strain (%)', 'Corrected Area, Ac (cm²)',
                                                         'Axial Stress, σ (ksc)']].values):
                        row = dt.rows[r_idx + 1]
                        for c_idx, val in enumerate(row_data):
                            cell = row.cells[c_idx]
                            cell.text = f"{val:.2f}"
                            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                            cell.paragraphs[0].runs[0].font.size = Pt(8)
                            
                    # --- RIGHT COLUMN: PROPERTIES & RESULTS ---
                    right_cell = body_table.cell(0, 1)
                    right_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
                    
                    # 4.1 Properties Table
                    p_prop = right_cell.paragraphs[0]
                    p_prop.add_run("Sample Properties").bold = True
                    
                    prop_t = right_cell.add_table(rows=8, cols=3)
                    prop_t.style = 'Table Grid'
                    prop_list = [
                        ("Diameter, D", f"{diameter:.2f}", "cm"),
                        ("Height, H", f"{height:.2f}", "cm"),
                        ("Area, A", f"{area:.2f}", "cm²"),
                        ("Volume, V", f"{volume:.2f}", "cm³"),
                        ("Weight, W", f"{weight:.2f}", "g"),
                        ("Wet Unit Wt.", f"{wet_unit_weight:.2f}", "g/cm³"),
                        ("Dry Unit Wt.", f"{dry_unit_weight:.2f}", "g/cm³"),
                        ("Water Content", f"{water_content:.2f}", "%")
                    ]
                    for i, (l, v, u) in enumerate(prop_list):
                        prop_t.cell(i,0).text = l
                        prop_t.cell(i,1).text = v
                        prop_t.cell(i,1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        prop_t.cell(i,2).text = u
                    
                    right_cell.add_paragraph() # Spacer
                    
                    # 4.2 Results Table
                    p_res = right_cell.add_paragraph()
                    p_res.add_run("Test Results").bold = True
                    
                    res_t = right_cell.add_table(rows=4, cols=3)
                    res_t.style = 'Table Grid'
                    res_list = [
                        ("UCS (qu)", f"{ucs:.2f}", "ksc"),
                        ("Shear Strength (su)", f"{undrained_shear_strength:.2f}", "ksc"),
                        ("Failure Strain", f"{failure_strain:.2f}", "%"),
                        ("Modulus E50", f"{modulus_e50:.2f}", "ksc")
                    ]
                    for i, (l, v, u) in enumerate(res_list):
                        res_t.cell(i,0).text = l
                        res_t.cell(i,1).text = v
                        res_t.cell(i,1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        res_t.cell(i,1).paragraphs[0].runs[0].font.bold = True
                        res_t.cell(i,2).text = u

                    right_cell.add_paragraph()
                    
                    # 4.3 Remarks
                    p_rem = right_cell.add_paragraph()
                    p_rem.add_run("Remarks:").bold = True
                    right_cell.add_paragraph(f"- Proving Ring: {proving_ring_capacity} kN")
                    right_cell.add_paragraph(f"- Factor K: {factor_k}")
                    
                    # 5. GRAPH (Bottom spanning full width)
                    doc.add_paragraph()
                    chart_p = doc.add_paragraph()
                    chart_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    # Use the buffer from earlier
                    chart_run = chart_p.add_run()
                    chart_run.add_picture(img_buffer, width=Cm(14)) # Width to fit nicely
                    
                    # 6. SIGNATURES
                    sig_table = doc.add_table(rows=3, cols=3)
                    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                    sig_table.cell(0,0).text = "Test by:"
                    sig_table.cell(0,1).text = "Controlled Test by:"
                    sig_table.cell(0,2).text = "Notified by:"
                    
                    sig_table.cell(2,0).text = f"({tested_by})"
                    sig_table.cell(2,1).text = f"({controlled_by if controlled_by else '.......................'})"
                    sig_table.cell(2,2).text = f"({notified_by if notified_by else '.......................'})"
                    
                    for row in sig_table.rows:
                        for cell in row.cells:
                            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

                    # Save to buffer
                    word_buffer = io.BytesIO()
                    doc.save(word_buffer)
                    word_buffer.seek(0)
                    return word_buffer
                
                word_data = create_word_report()
                st.download_button(
                    label="📥 Download Word Report",
                    data=word_data,
                    file_name=f"UCS_Report_{project_name}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            
            plt.close()
            
    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        st.exception(e)

else:
    st.info("👆 Please upload an Excel file with test data (Load and Vertical displacement) to generate the report.")

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666; padding: 1rem;">
    <p>🏫 King Mongkut's University of Technology North Bangkok</p>
    <p>Department of Teacher Training in Civil Engineering</p>
    <p>© 2024 - UCS Test Report Generator</p>
</div>
""", unsafe_allow_html=True)
