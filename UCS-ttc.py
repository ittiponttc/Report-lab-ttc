import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import io
from datetime import date
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Page config
st.set_page_config(page_title="UCS Test Report Generator", page_icon="🧪", layout="wide")

# Custom CSS
st.markdown("""
<style>
    .main-header { font-size: 1.8rem; font-weight: bold; color: #1e3a5f; text-align: center; margin-bottom: 1rem; }
    .sub-header { font-size: 1.1rem; color: #2c5282; border-bottom: 2px solid #3182ce; padding-bottom: 0.5rem; margin-top: 1rem; }
</style>
""", unsafe_allow_html=True)

# Title
st.markdown('<div class="main-header">🧪 Unconfined Compression Test (ASTM D2166)</div>', unsafe_allow_html=True)

# --- Sidebar Inputs ---
st.sidebar.markdown("## 📋 Project Info")
with st.sidebar:
    project_name = st.text_input("Project Name", "Renovation Project")
    specimen_from = st.text_input("Specimen From", "BH-1")
    location = st.text_input("Location", "Bangkok")
    column_no = st.text_input("Column No.", "C1")
    depth = st.text_input("Depth (m)", "7.50-8.50")
    sample_number = st.text_input("Sample No.", "1")
    
    st.markdown("---")
    date_of_jetting = st.text_input("Date of Jetting", str(date.today()))
    date_of_testing = st.date_input("Date of Testing", date.today())
    curing_time = st.number_input("Curing Time (days)", value=28)
    
    st.markdown("---")
    tested_by = st.text_input("Tested by", "Asst.Prof.Dr.Ittipon Meepon")
    controlled_by = st.text_input("Controlled by", "")
    notified_by = st.text_input("Notified by", "")

# --- Main Layout ---
col1, col2 = st.columns([1, 1])

with col1:
    st.markdown('<p class="sub-header">📊 Sample Properties</p>', unsafe_allow_html=True)
    sc1, sc2 = st.columns(2)
    with sc1:
        diameter = st.number_input("Diameter (cm)", value=5.53, step=0.01)
        height = st.number_input("Height (cm)", value=11.90, step=0.01)
        weight = st.number_input("Weight (g)", value=621.07, step=0.01)
    with sc2:
        shearing_rate = st.text_input("Shearing Rate", "1% / min")
        mixed_jet_mixing = st.text_input("Mixed Jet (kg/m3)", "-")
    
    st.markdown('<p class="sub-header">💧 Water Content</p>', unsafe_allow_html=True)
    wc1, wc2 = st.columns(2)
    with wc1:
        w_wet_soil = st.number_input("Wt. Wet Soil+Can (g)", value=406.14)
        w_dry_soil = st.number_input("Wt. Dry Soil+Can (g)", value=244.31)
    with wc2:
        w_can = st.number_input("Wt. Can (g)", value=26.05)

with col2:
    st.markdown('<p class="sub-header">📁 Data & Photo</p>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("1. Upload Excel Data (Load/Disp)", type=['xlsx', 'xls'])
    
    # --- New: Image Uploader ---
    uploaded_photo = st.file_uploader("2. Upload Specimen Photo (รูปตัวอย่าง)", type=['png', 'jpg', 'jpeg'])
    if uploaded_photo:
        st.image(uploaded_photo, caption="Specimen Preview", width=200)

    proving_ring = st.number_input("Proving Ring (kN)", value=20.0)
    factor_k = st.number_input("Factor K", value=1.8661, format="%.4f")

# Calculations
area = np.pi * (diameter / 2) ** 2
volume = area * height
wet_unit_weight = weight / volume if volume > 0 else 0
w_water = w_wet_soil - w_dry_soil
w_dry = w_dry_soil - w_can
water_content = (w_water / w_dry * 100) if w_dry > 0 else 0
dry_unit_weight = wet_unit_weight / (1 + water_content/100) if water_content > 0 else 0

# Process Data
if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        # Standardize columns
        cols = df.columns
        if 'Load (kg)' not in cols or 'Vertical displacement (mm)' not in cols:
             # Try to guess if names are slightly different or just take 1st/2nd cols
             df.columns = ['Vertical displacement (mm)', 'Load (kg)'] + list(df.columns[2:])
        
        # Calculation Logic
        df['R (mm)'] = df['Vertical displacement (mm)']
        df['P (kg)'] = df['Load (kg)']
        df['Strain'] = df['R (mm)'] / (10 * height) # Unitless
        df['Strain (%)'] = df['Strain'] * 100
        df['Ac (cm2)'] = area / (1 - df['Strain'])
        df['Stress (ksc)'] = df['P (kg)'] / df['Ac (cm2)']
        
        # Find Results
        max_idx = df['Stress (ksc)'].idxmax()
        ucs = df.loc[max_idx, 'Stress (ksc)']
        failure_strain = df.loc[max_idx, 'Strain (%)']
        su = ucs / 2
        
        # E50
        half_ucs = ucs / 2
        df_pre_peak = df.iloc[:max_idx+1]
        idx_e50 = (df_pre_peak['Stress (ksc)'] - half_ucs).abs().idxmin()
        strain_e50 = df.loc[idx_e50, 'Strain (%)']
        e50 = (half_ucs / (strain_e50/100)) if strain_e50 > 0 else 0

        # Show Results on Screen
        st.success(f"Calculated: UCS = {ucs:.2f} ksc")
        
        # --- Generate Graph ---
        fig, ax = plt.subplots(figsize=(6, 4))
        ax.plot(df['Strain (%)'], df['Stress (ksc)'], 'b-o', markersize=3, linewidth=1)
        ax.set_xlabel('Axial Strain (%)')
        ax.set_ylabel('Axial Stress (ksc)')
        ax.grid(True, linestyle='--', alpha=0.6)
        ax.set_title(f"UCS: {ucs:.2f} ksc")
        plt.tight_layout()
        
        graph_buffer = io.BytesIO()
        fig.savefig(graph_buffer, format='png', dpi=150)
        graph_buffer.seek(0)
        plt.close()

        # --- Helper for Word Borders ---
        def set_border(cell, top=False, bottom=False, left=False, right=False):
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            for edge, toggle in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
                if toggle:
                    tag = f'w:{edge}'
                    element = tcPr.find(qn(tag))
                    if element is None:
                        element = OxmlElement(tag)
                        tcPr.append(element)
                    element.set(qn('w:val'), 'single')
                    element.set(qn('w:sz'), '4')
                    element.set(qn('w:space'), '0')
                    element.set(qn('w:color'), 'auto')

        # --- Word Report Generation (The 1-Page Logic) ---
        def create_word():
            doc = Document()
            
            # 1. Page Setup (Narrow Margins for 1 Page fit)
            section = doc.sections[0]
            section.top_margin = Cm(1.0)
            section.bottom_margin = Cm(1.0)
            section.left_margin = Cm(1.0)
            section.right_margin = Cm(1.0)
            
            style = doc.styles['Normal']
            style.font.name = 'Times New Roman'
            style.font.size = Pt(9)

            # 2. Header
            h_p = doc.add_paragraph()
            h_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = h_p.add_run("KING MONGKUT'S UNIVERSITY OF TECHNOLOGY NORTH BANGKOK\n")
            run.bold = True; run.font.size = Pt(11)
            run = h_p.add_run("DEPARTMENT OF TEACHER TRAINING IN CIVIL ENGINEERING\n")
            run.bold = True; run.font.size = Pt(10); run.font.color.rgb = RGBColor(255, 0, 0)
            run = h_p.add_run("Unconfined Compression Test (ASTM D2166)")
            run.bold = True; run.font.size = Pt(10)

            # 3. Project Info Table
            t_info = doc.add_table(rows=4, cols=4)
            t_info.style = 'Table Grid'
            info_data = [
                ("Project Name :", project_name, "Tested by :", tested_by),
                ("Location :", location, "Date of Jetting :", date_of_jetting),
                ("Column No. :", column_no, "Date of Testing :", str(date_of_testing)),
                ("Depth :", depth, "Specimen No. :", sample_number)
            ]
            for r, (l1, v1, l2, v2) in enumerate(info_data):
                row = t_info.rows[r]
                row.cells[0].text = l1
                row.cells[1].text = str(v1)
                row.cells[2].text = l2
                row.cells[3].text = str(v2)
                row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
                row.height = Cm(0.5)

            doc.add_paragraph().add_run().font.size = Pt(2) # Spacer

            # 4. Main Layout (Left: Data, Right: Props+Img+Graph)
            # Create a borderless table to act as columns
            main_table = doc.add_table(rows=1, cols=2)
            main_table.autofit = False
            main_table.columns[0].width = Cm(6.5) # Left Col Width
            main_table.columns[1].width = Cm(12.0) # Right Col Width

            # --- LEFT COLUMN: DATA ---
            left_cell = main_table.cell(0, 0)
            left_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
            
            # Header for data
            p = left_cell.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run("Test Data").bold = True
            
            # Data Table
            rows_to_show = min(len(df), 45) # Limit rows to fit page
            t_data = left_cell.add_table(rows=rows_to_show+1, cols=4)
            t_data.style = 'Table Grid'
            
            # Headers
            headers = ['R\n(mm)', 'P\n(kg)', 'Strain\n(%)', 'Stress\n(ksc)']
            for i, h in enumerate(headers):
                cell = t_data.cell(0, i)
                cell.text = h
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                cell.paragraphs[0].runs[0].font.size = Pt(8)
                cell.paragraphs[0].runs[0].font.bold = True

            # Fill Data
            for r in range(rows_to_show):
                row_vals = df.iloc[r]
                vals = [row_vals['R (mm)'], row_vals['P (kg)'], row_vals['Strain (%)'], row_vals['Stress (ksc)']]
                t_row = t_data.rows[r+1]
                t_row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
                t_row.height = Cm(0.4) # Tight rows
                for c, val in enumerate(vals):
                    cell = t_row.cells[c]
                    cell.text = f"{val:.2f}"
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    cell.paragraphs[0].runs[0].font.size = Pt(8)

            # --- RIGHT COLUMN: EVERYTHING ELSE ---
            right_cell = main_table.cell(0, 1)
            right_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
            
            # Use a nested table for Properties + Image
            # Row 1: Properties Header | Image Header
            # Row 2: Properties Data   | Image Placeholder
            
            sub_t = right_cell.add_table(rows=1, cols=2)
            sub_t.autofit = False
            sub_t.columns[0].width = Cm(7.0) # Props
            sub_t.columns[1].width = Cm(5.0) # Image
            
            # -- Properties Section (Left side of Right Col) --
            cell_props = sub_t.cell(0, 0)
            p_prop = cell_props.add_paragraph()
            p_prop.add_run("Soil-Jet Mixing Sample Properties").bold = True
            
            t_prop = cell_props.add_table(rows=9, cols=3)
            t_prop.style = 'Table Grid'
            props_list = [
                ("Diameter", f"{diameter:.2f}", "cm"),
                ("Height", f"{height:.2f}", "cm"),
                ("Area", f"{area:.2f}", "cm2"),
                ("Volume", f"{volume:.2f}", "cm3"),
                ("Weight", f"{weight:.2f}", "g"),
                ("Wet Unit Wt.", f"{wet_unit_weight:.2f}", "g/cm3"),
                ("Dry Unit Wt.", f"{dry_unit_weight:.2f}", "g/cm3"),
                ("Water Content", f"{water_content:.2f}", "%"),
                ("Shearing Rate", shearing_rate, "")
            ]
            for i, (l, v, u) in enumerate(props_list):
                row = t_prop.rows[i]
                row.height = Cm(0.45)
                row.cells[0].text = l
                row.cells[1].text = v
                row.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                row.cells[2].text = u
            
            # -- Image Section (Right side of Right Col) --
            cell_img = sub_t.cell(0, 1)
            cell_img.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
            p_img = cell_img.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.add_run("Failure of Sample").bold = True
            
            if uploaded_photo:
                # Add picture to paragraph
                run_img = cell_img.add_paragraph().add_run()
                run_img.add_picture(uploaded_photo, width=Cm(4.5))
            else:
                cell_img.add_paragraph("\n[No Image Uploaded]\n").alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Remarks Box below image
            cell_img.add_paragraph().add_run("\nRemarks:").bold = True
            cell_img.add_paragraph(f"- Ring: {proving_ring} kN")
            cell_img.add_paragraph(f"- Factor K: {factor_k}")

            right_cell.add_paragraph() # Spacer

            # -- Results Table --
            t_res = right_cell.add_table(rows=4, cols=3)
            t_res.style = 'Table Grid'
            res_data = [
                ("Unconfined Compressive Strength (qu)", f"{ucs:.2f}", "ksc"),
                ("Undrained Shear Strength (su)", f"{su:.2f}", "ksc"),
                ("Failure Strain", f"{failure_strain:.2f}", "%"),
                ("Modulus of Elasticity (E50)", f"{e50:.2f}", "ksc")
            ]
            for i, (l, v, u) in enumerate(res_data):
                t_res.rows[i].cells[0].text = l
                t_res.rows[i].cells[1].text = v
                t_res.rows[i].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                t_res.rows[i].cells[1].paragraphs[0].runs[0].bold = True
                t_res.rows[i].cells[2].text = u
            
            right_cell.add_paragraph() # Spacer
            
            # -- Graph --
            p_graph = right_cell.add_paragraph()
            p_graph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_graph.add_run().add_picture(graph_buffer, width=Cm(11.0))

            # 5. Signatures (Footer)
            doc.add_paragraph()
            t_sig = doc.add_table(rows=2, cols=3)
            t_sig.rows[0].cells[0].text = "Test by:"
            t_sig.rows[0].cells[1].text = "Controlled by:"
            t_sig.rows[0].cells[2].text = "Notified by:"
            
            t_sig.rows[1].cells[0].text = f"({tested_by})"
            t_sig.rows[1].cells[1].text = f"({controlled_by if controlled_by else '.......................'})"
            t_sig.rows[1].cells[2].text = f"({notified_by if notified_by else '.......................'})"
            
            for row in t_sig.rows:
                for cell in row.cells:
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Save
            f = io.BytesIO()
            doc.save(f)
            f.seek(0)
            return f

        # Download Button
        st.markdown("### 📥 Download Report")
        word_file = create_word()
        st.download_button(
            label="Download Word Report (A4 One Page)",
            data=word_file,
            file_name=f"UCS_Report_{project_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
            
    except Exception as e:
        st.error(f"Error: {e}")
