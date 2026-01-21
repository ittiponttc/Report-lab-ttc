import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import io
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter  # Import for column letter fix
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

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
    .info-box {
        background-color: #f0f9ff;
        border-left: 4px solid #3182ce;
        padding: 1rem;
        border-radius: 0 8px 8px 0;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Title
st.markdown('<div class="main-header">🧪 UCS Report Generator (ASTM D2166)</div>', unsafe_allow_html=True)

# Initialize session state
if 'calculated' not in st.session_state:
    st.session_state.calculated = False
if 'results' not in st.session_state:
    st.session_state.results = None

# Sidebar for project information
st.sidebar.markdown("## 📋 Project Information")

with st.sidebar:
    job_number = st.text_input("Job Number / เลขที่งาน", "685/2564")
    project_name = st.text_input("Project Name", "")
    specimen_from = st.text_input("Specimen from", "")
    location = st.text_input("Location", "")
    column_no = st.text_input("Column No.", "")
    depth = st.text_input("Depth (m)", "7.50-8.50")
    sample_number = st.number_input("Sample Number", min_value=1, value=1)
    
    st.markdown("---")
    st.markdown("### 📅 Dates & Personnel")
    date_of_jetting = st.text_input("Date of Jetting", "-")
    date_of_testing = st.date_input("Date of Testing", date.today())
    curing_time = st.text_input("Curing Time (days)", "28")
    tested_by = st.text_input("Tested by", "Asst.Prof.Dr.Ittipon Meepon")
    controlled_by = st.text_input("Controlled Test by", "")
    notified_by = st.text_input("Notified by", "")
    
    st.markdown("---")
    st.markdown("### 📷 Sample Photo")
    sample_image = st.file_uploader("Upload Sample Image", type=['png', 'jpg', 'jpeg'])

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
    <strong>📌 Required Excel Columns:</strong><br>
    <code>Load (kg)</code> and <code>Vertical displacement (mm)</code>
    </div>
    """, unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader("Upload Data Excel File", type=['xlsx', 'xls'])
    
    proving_ring_capacity = st.number_input("Proving Ring Capacity (kN)", min_value=0.1, value=20.0, step=0.1)
    factor_k = st.number_input("Factor K (kg/division)", min_value=0.0001, value=1.8661, step=0.0001, format="%.4f")

# Calculations
area = np.pi * (diameter / 2) ** 2
volume = area * height
wet_unit_weight = weight / volume if volume > 0 else 0

weight_water = weight_wet_soil_can - weight_dry_soil_can
weight_dry_soil = weight_dry_soil_can - weight_can
water_content = (weight_water / weight_dry_soil * 100) if weight_dry_soil > 0 else 0
dry_unit_weight = wet_unit_weight / (1 + water_content/100) if water_content > 0 else wet_unit_weight

# Display properties
st.markdown('<p class="sub-header">📐 Calculated Properties</p>', unsafe_allow_html=True)
prop_col1, prop_col2, prop_col3, prop_col4 = st.columns(4)
prop_col1.metric("Area (cm²)", f"{area:.2f}")
prop_col2.metric("Volume (cm³)", f"{volume:.2f}")
prop_col3.metric("Wet Unit Wt (g/cm³)", f"{wet_unit_weight:.2f}")
prop_col4.metric("Water Content (%)", f"{water_content:.2f}")

# Process Data
if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        required_cols = ['Load (kg)', 'Vertical displacement (mm)']
        
        if not all(col in df.columns for col in required_cols):
            st.error(f"❌ Required columns missing. Need: {required_cols}")
        else:
            # Calculations
            df['Deformation, R (mm)'] = df['Vertical displacement (mm)']
            df['Axial Load, P (kg)'] = df['Load (kg)']
            df['Axial Strain, ξ (R/10H)'] = df['Deformation, R (mm)'] / (10 * height)
            df['Axial Strain (%)'] = df['Axial Strain, ξ (R/10H)'] * 100
            df['Corrected Area, Ac (cm²)'] = area / (1 - df['Axial Strain, ξ (R/10H)'])
            df['Axial Stress, σ (ksc)'] = (df['Axial Load, P (kg)'] / df['Corrected Area, Ac (cm²)'])
            
            # Results
            max_stress_idx = df['Axial Stress, σ (ksc)'].idxmax()
            ucs = df.loc[max_stress_idx, 'Axial Stress, σ (ksc)']
            failure_strain = df.loc[max_stress_idx, 'Axial Strain (%)']
            undrained_shear_strength = ucs / 2
            
            # E50
            half_ucs = ucs / 2
            df_before_peak = df[df.index <= max_stress_idx]
            idx_e50 = (df_before_peak['Axial Stress, σ (ksc)'] - half_ucs).abs().idxmin()
            strain_at_half = df.loc[idx_e50, 'Axial Strain (%)']
            modulus_e50 = (half_ucs / (strain_at_half / 100)) if strain_at_half > 0 else 0
            
            st.success("✅ Calculations Complete")
            
            # Plot
            fig, ax = plt.subplots(figsize=(8, 6))
            ax.plot(df['Axial Strain (%)'], df['Axial Stress, σ (ksc)'], 'b-o', markersize=4, linewidth=1)
            ax.set_xlabel('Axial Strain, ε, %')
            ax.set_ylabel('Axial Stress, σ, ksc')
            ax.grid(True, linestyle='--', which='both', alpha=0.6)
            ax.minorticks_on()
            ax.grid(which='minor', linestyle=':', linewidth='0.5', color='gray')
            
            img_buffer = io.BytesIO()
            fig.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight')
            img_buffer.seek(0)
            
            # --- WORD REPORT GENERATION ---
            def set_cell_border(cell, **kwargs):
                """
                Set cell`s border
                Usage:
                set_cell_border(
                    cell,
                    top={"sz": 12, "val": "single", "color": "FF0000", "space": "0"},
                    bottom={"sz": 12, "color": "00FF00", "val": "single"},
                    start={"sz": 24, "val": "dashed", "shadow": "true"},
                    end={"sz": 12, "val": "dashed"},
                )
                """
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

            def create_word_report():
                doc = Document()
                
                # Setup Margins (Narrow)
                sections = doc.sections
                for section in sections:
                    section.top_margin = Cm(1.27)
                    section.bottom_margin = Cm(1.27)
                    section.left_margin = Cm(1.27)
                    section.right_margin = Cm(1.27)

                # Default Style
                style = doc.styles['Normal']
                font = style.font
                font.name = 'Times New Roman'
                font.size = Pt(10)

                # --- HEADER ---
                # Using a table for the Header to manage alignment
                header_table = doc.add_table(rows=3, cols=1)
                header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                
                # Row 1: University Name
                c1 = header_table.cell(0, 0)
                p1 = c1.paragraphs[0]
                run1 = p1.add_run("KING MONGKUT'S UNIVERSITY OF TECHNOLOGY NORTH BANGKOK")
                run1.bold = True
                run1.font.size = Pt(12)
                p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Row 2: Department (Red)
                c2 = header_table.cell(1, 0)
                p2 = c2.paragraphs[0]
                run2 = p2.add_run("DEPARTMENT OF TEACHER TRAINING IN CIVIL ENGINEERING")
                run2.bold = True
                run2.font.color.rgb = RGBColor(255, 0, 0)
                run2.font.size = Pt(11)
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Row 3: Address
                c3 = header_table.cell(2, 0)
                p3 = c3.paragraphs[0]
                p3.add_run("1518 Pracharat 1 Road, Wongsawang, Bangsue, Bangkok 10800, THAILAND\n")
                p3.add_run("Telephone and Fax :02-555-2000 ext 3247,3253, Fax. 02-587-8260")
                p3.alignment = WD_ALIGN_PARAGRAPH.CENTER

                doc.add_paragraph() # Spacer

                # --- JOB TITLE ---
                title_table = doc.add_table(rows=1, cols=2)
                title_table.autofit = False
                title_table.columns[0].width = Inches(6.0)
                title_table.columns[1].width = Inches(2.0)
                
                t_cell_1 = title_table.cell(0, 0)
                p_title = t_cell_1.paragraphs[0]
                run_title = p_title.add_run("Unconfined Compression Test (ASTM D2166)")
                run_title.bold = True
                run_title.font.size = Pt(11)
                p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                t_cell_2 = title_table.cell(0, 1)
                p_job = t_cell_2.paragraphs[0]
                run_job = p_job.add_run(f"เลขที่งาน {job_number}")
                run_job.bold = True
                run_job.font.color.rgb = RGBColor(0, 0, 255) # Blue in image looks nice, user asked for red before but image has blue text mostly? Let's use Blue for job no.
                p_job.alignment = WD_ALIGN_PARAGRAPH.RIGHT

                # --- PROJECT INFO ---
                # Grid of info
                info_table = doc.add_table(rows=4, cols=4)
                info_table.style = 'Table Grid'
                
                # Setup specific widths if needed, but auto might work
                # Data mapping
                info_data = [
                    ("Specimen from :", specimen_from, "Tested by :", tested_by),
                    ("Project Name :", project_name, "Date of Jetting :", date_of_jetting),
                    ("Location :", location, "Date of Testing :", str(date_of_testing)),
                    ("Shearing Rate :", shearing_rate, "Curing Time :", f"{curing_time} days"),
                    ("Column No. :", column_no, "Sample Number :", str(sample_number)),
                    ("Depth :", depth, "", "") 
                ]
                
                # Re-create table with exact rows needed (6 rows based on image logic)
                # Note: The image has a mix of merged cells. Let's do a simple 4-col layout
                # To match image exactly requires a lot of border hacking. 
                # We will use a clean table for info.
                
                doc.add_paragraph() # Spacer
                
                p_info = doc.add_table(rows=3, cols=4) # 3 rows for main info
                # Row 1
                p_info.cell(0,0).text = f"Specimen from : {specimen_from}"
                p_info.cell(0,2).text = f"Tested by : {tested_by}"
                # Row 2
                p_info.cell(1,0).text = f"Project Name : {project_name}"
                p_info.cell(1,2).text = f"Date of Jetting : {date_of_jetting}"
                # Row 3
                p_info.cell(2,0).text = f"Location : {location}"
                p_info.cell(2,2).text = f"Date of Testing : {date_of_testing}"
                
                # Add more rows for extra info
                r4 = p_info.add_row()
                r4.cells[0].text = f"Shearing Rate : {shearing_rate}"
                r4.cells[2].text = f"Curing Time : {curing_time} days"
                
                r5 = p_info.add_row()
                r5.cells[0].text = f"Column No. : {column_no}"
                r5.cells[1].text = f"Depth : {depth}"
                r5.cells[2].text = f"Sample Number : {sample_number}"

                # --- MAIN DATA BODY (The Complex Part) ---
                doc.add_paragraph()
                
                # We create one big table to hold: Data(Left), Props(Middle), Photo(Right)
                # Columns: 
                # 0: Deformation
                # 1: Load
                # 2: Strain (ratio)
                # 3: Area
                # 4: Stress
                # 5: Spacer
                # 6: Prop Labels
                # 7: Prop Values
                # 8: Prop Unit
                # 9: Photo placeholder
                
                # Determine rows needed
                data_rows = len(df) + 2 # + headers
                prop_rows = 20 # fixed slots for props
                total_rows = max(data_rows, prop_rows)
                
                main_table = doc.add_table(rows=total_rows + 5, cols=10) # +5 buffer
                main_table.style = 'Table Grid'
                main_table.autofit = False
                
                # Set Column Widths (Approximate for A4)
                widths = [1.2, 1.2, 1.2, 1.2, 1.2, 0.2, 2.5, 1.5, 1.0, 4.0]
                for i, width in enumerate(widths):
                    for row in main_table.rows:
                        row.cells[i].width = Cm(width)

                # --- FILL LEFT SIDE (DATA) ---
                # Headers
                headers = ['Deformation\nR', 'Axial Load\nP', 'Axial Strain\nξ', 'Corrected Area\nAc', 'Axial Stress\nσ']
                units = ['(mm)', '(kg)', '(R/10H)', '(cm²)', '(ksc)']
                
                for i, (h, u) in enumerate(zip(headers, units)):
                    cell = main_table.cell(0, i)
                    cell.text = h
                    cell.paragraphs[0].runs[0].font.bold = True
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    cell2 = main_table.cell(1, i)
                    cell2.text = u
                    cell2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

                # Data
                for r_idx, row_data in enumerate(df[['Deformation, R (mm)', 'Axial Load, P (kg)', 
                                                     'Axial Strain, ξ (R/10H)', 'Corrected Area, Ac (cm²)',
                                                     'Axial Stress, σ (ksc)']].values):
                    row_idx_table = r_idx + 2
                    for c_idx, val in enumerate(row_data):
                        cell = main_table.cell(row_idx_table, c_idx)
                        cell.text = f"{val:.2f}"
                        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        cell.paragraphs[0].runs[0].font.size = Pt(9)

                # --- FILL MIDDLE SIDE (PROPERTIES) ---
                # Helper to write to specific cell in main table
                def write_prop(r, c, text, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT):
                    cell = main_table.cell(r, c)
                    cell.text = text
                    if bold: cell.paragraphs[0].runs[0].font.bold = True
                    cell.paragraphs[0].alignment = align
                    cell.paragraphs[0].runs[0].font.size = Pt(9)

                # Cement Mixing
                write_prop(0, 6, "Cement Mixing Descriptions", bold=True)
                main_table.cell(0, 6).merge(main_table.cell(0, 8))
                
                write_prop(1, 6, "Mixed Jet Mixing")
                write_prop(1, 7, f"{mixed_jet_mixing}")
                write_prop(1, 8, "kg/m³")

                # Sample Properties
                start_r = 2
                write_prop(start_r, 6, "Soil-Jet Mixing Sample", bold=True)
                main_table.cell(start_r, 6).merge(main_table.cell(start_r, 8))
                
                props_list = [
                    ("Diameter, D", f"{diameter:.2f}", "cm."),
                    ("Height, H", f"{height:.2f}", "cm."),
                    ("Area, A", f"{area:.2f}", "cm²"),
                    ("Volume, V", f"{volume:.2f}", "cm³"),
                    ("Weight of Sample, W", f"{weight:.2f}", "g"),
                    ("Wet Unit Weight", f"{wet_unit_weight:.2f}", "g/cm³"),
                    ("Dry Unit Weight", f"{dry_unit_weight:.2f}", "g/cm³"),
                ]
                
                for i, (lbl, val, unit) in enumerate(props_list):
                    r = start_r + 1 + i
                    write_prop(r, 6, lbl)
                    write_prop(r, 7, val, align=WD_ALIGN_PARAGRAPH.CENTER)
                    write_prop(r, 8, unit)

                # Water Content
                wc_start = start_r + 1 + len(props_list)
                write_prop(wc_start, 6, "Water Content Determination", bold=True)
                main_table.cell(wc_start, 6).merge(main_table.cell(wc_start, 8))
                
                wc_list = [
                    ("Container No.", container_no, ""),
                    ("Weight of Wet Soil + Can", f"{weight_wet_soil_can:.2f}", "g"),
                    ("Weight of Dry Soil + Can", f"{weight_dry_soil_can:.2f}", "g"),
                    ("Weight of Water", f"{weight_water:.2f}", "g"),
                    ("Weight of Can", f"{weight_can:.2f}", "g"),
                    ("Weight of Dry Soil", f"{weight_dry_soil:.2f}", "g"),
                    ("Water Content, w", f"{water_content:.2f}", "%"),
                ]
                
                for i, (lbl, val, unit) in enumerate(wc_list):
                    r = wc_start + 1 + i
                    write_prop(r, 6, lbl)
                    write_prop(r, 7, val, align=WD_ALIGN_PARAGRAPH.CENTER)
                    write_prop(r, 8, unit)

                # --- RIGHT SIDE (PHOTO & REMARKS) ---
                # Header for Photo
                write_prop(0, 9, "Failure of Soil-Jet\nMixing Sample", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
                
                # Merge cells for photo (Row 1 to 12 approx)
                photo_cell = main_table.cell(1, 9)
                photo_bottom_cell = main_table.cell(12, 9)
                photo_cell.merge(photo_bottom_cell)
                
                # Insert Image if exists
                if sample_image:
                    paragraph = photo_cell.paragraphs[0]
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = paragraph.add_run()
                    run.add_picture(sample_image, width=Cm(4.5))
                else:
                    photo_cell.text = "[Place Photo Here]"

                # Remarks Section (Below Photo)
                rem_start = 13
                rem_cell = main_table.cell(rem_start, 9)
                rem_end_cell = main_table.cell(18, 9)
                rem_cell.merge(rem_end_cell)
                
                rem_text = rem_cell.paragraphs[0]
                rem_text.add_run("Remarks :\n").bold = True
                rem_text.add_run(f"- Proving Ring Capacity {proving_ring_capacity} kN\n")
                rem_text.add_run(f"- Factor K = {factor_k} kg/division")

                # Remove borders for the Spacer Column (5)
                for row in main_table.rows:
                    set_cell_border(row.cells[5], top={}, bottom={}, left={}, right={})

                # --- CHART SECTION ---
                doc.add_paragraph()
                # Create a table for Chart (Left) and Results (Right)
                # But looking at image, Chart is full width, Results are inset or below
                # Image: Chart is large, Results are in a small table overlaid or below.
                # Let's put Chart and then Results Table below it.
                
                chart_p = doc.add_paragraph()
                chart_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                chart_run = chart_p.add_run()
                chart_run.add_picture(img_buffer, width=Cm(12)) # Adjust width
                
                # Results Table (Small, Bottom Right aligned in image usually, but let's center it)
                res_table = doc.add_table(rows=5, cols=3)
                res_table.style = 'Table Grid'
                res_table.alignment = WD_TABLE_ALIGNMENT.RIGHT # Move to right
                
                res_data = [
                    ("Unconfined Compressive Strength, (qu)", f"{ucs:.2f}", "ksc."),
                    ("Undrained Shear Strength, (su)", f"{undrained_shear_strength:.2f}", "ksc."),
                    ("Failure Strain, (ef)", f"{failure_strain:.2f}", "%"),
                    ("Modulus of Elasticity", f"{modulus_e50:.2f}", "ksc.")
                ]
                
                for i, (lbl, val, unit) in enumerate(res_data):
                    res_table.cell(i, 0).text = lbl
                    res_table.cell(i, 1).text = val
                    res_table.cell(i, 2).text = unit
                
                # --- FOOTER / SIGNATURES ---
                doc.add_paragraph()
                doc.add_paragraph("Remarks :")
                doc.add_paragraph("1. The testing results are good only for those specimens tested.")
                doc.add_paragraph("2. Not valid unless signed and sealed.")
                
                sig_table = doc.add_table(rows=2, cols=3)
                sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                sig_table.autofit = True
                
                sig_table.cell(0,0).text = "Test by:"
                sig_table.cell(0,1).text = "Controlled Test by:"
                sig_table.cell(0,2).text = "Notified by:"
                
                sig_table.cell(1,0).text = f"( {tested_by} )"
                sig_table.cell(1,1).text = f"( {controlled_by if controlled_by else '..........................................'} )"
                sig_table.cell(1,2).text = f"( {notified_by if notified_by else '..........................................'} )"
                
                # Align signatures
                for row in sig_table.rows:
                    for cell in row.cells:
                        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

                # Save
                out_buffer = io.BytesIO()
                doc.save(out_buffer)
                out_buffer.seek(0)
                return out_buffer

            # EXCEL Report (Fixing the MergedCell error here)
            def create_excel_report():
                wb = Workbook()
                ws = wb.active
                ws.title = "UCS Test Report"
                
                # ... (Keep existing Excel logic, just applying the fix below) ...
                # Simple dump for Excel to save space in code, focused on Word
                ws.append(["Project", project_name])
                ws.append(["Test Date", date_of_testing])
                ws.append([])
                
                # Headers
                headers = ['R (mm)', 'P (kg)', 'Strain', 'Area', 'Stress (ksc)']
                ws.append(headers)
                
                for row in df[['Deformation, R (mm)', 'Axial Load, P (kg)', 
                              'Axial Strain, ξ (R/10H)', 'Corrected Area, Ac (cm²)', 
                              'Axial Stress, σ (ksc)']].values:
                    ws.append(list(row))
                
                # FIX: Use get_column_letter
                for column_cells in ws.columns:
                    length = max(len(str(cell.value) if cell.value else "") for cell in column_cells)
                    col_letter = get_column_letter(column_cells[0].column)
                    ws.column_dimensions[col_letter].width = length + 2
                    
                excel_buffer = io.BytesIO()
                wb.save(excel_buffer)
                excel_buffer.seek(0)
                return excel_buffer

            # Download Buttons
            st.markdown('<p class="sub-header">💾 Download Reports</p>', unsafe_allow_html=True)
            dcol1, dcol2 = st.columns(2)
            
            with dcol1:
                word_file = create_word_report()
                st.download_button(
                    label="📄 Download Word Report (Layout Fixed)",
                    data=word_file,
                    file_name=f"UCS_Report_{job_number.replace('/','-')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            
            with dcol2:
                excel_file = create_excel_report()
                st.download_button(
                    label="📊 Download Excel Data",
                    data=excel_file,
                    file_name=f"UCS_Data_{job_number.replace('/','-')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Error: {str(e)}")
        st.exception(e)

st.markdown("---")
st.caption("KMUTNB Civil Engineering - UCS Report Generator")
