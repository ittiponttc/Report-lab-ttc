import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document

# =========================================================
# Functions
# =========================================================
def concrete_mix_design(
    wc_ratio,
    max_agg_mm,
    sg_cement,
    sg_fine,
    sg_coarse,
    air_content,
    unit_weight_coarse
):
    # ACI 211 typical values
    if max_agg_mm == 20:
        water = 185
        vol_coarse = 0.62
    elif max_agg_mm == 25:
        water = 175
        vol_coarse = 0.64
    else:
        water = 165
        vol_coarse = 0.68

    cement = water / wc_ratio
    weight_coarse = vol_coarse * unit_weight_coarse

    vol_water = water / 1000
    vol_cement = cement / (sg_cement * 1000)
    vol_coarse_abs = weight_coarse / (sg_coarse * 1000)

    vol_fine = 1 - (vol_water + vol_cement + vol_coarse_abs + air_content)
    weight_fine = vol_fine * sg_fine * 1000

    return {
        "Water": water,
        "Cement": cement,
        "Fine Aggregate": weight_fine,
        "Coarse Aggregate": weight_coarse,
        "vol_coarse": vol_coarse
    }


def moisture_correction(weight_ssd, mc, absorption):
    delta_water = weight_ssd * (mc - absorption) / 100
    batch_weight = weight_ssd * (1 + (mc - absorption) / 100)
    return delta_water, batch_weight


def generate_word_report(data):
    doc = Document()
    doc.add_heading("Concrete Mix Design Report (ACI 211)", level=1)

    doc.add_heading("1. Project Information", level=2)
    doc.add_paragraph(f"Water–Cement Ratio (w/c) = {data['wc_ratio']}")
    doc.add_paragraph(f"Maximum Aggregate Size = {data['max_agg_mm']} mm")
    doc.add_paragraph(f"Air Content = {data['air_content']*100:.1f} %")

    doc.add_heading("2. Design Assumptions (ACI 211)", level=2)
    doc.add_paragraph(
        "• Mix design is based on ACI 211.1 (Absolute Volume Method).\n"
        "• Concrete is assumed to be non-air-entrained.\n"
        "• Aggregates are assumed to be in SSD condition for initial proportioning."
    )

    doc.add_heading("3. Step-by-Step Mix Design Calculation", level=2)
    doc.add_paragraph(f"Water content selected = {data['water']} kg/m³")
    doc.add_paragraph(
        f"Cement content = Water / (w/c) = "
        f"{data['water']} / {data['wc_ratio']} = {data['cement']:.1f} kg/m³"
    )
    doc.add_paragraph(
        f"Volume of coarse aggregate = {data['vol_coarse']} × unit weight"
    )

    doc.add_heading("4. SSD Mix Proportion", level=2)
    table = doc.add_table(rows=1, cols=2)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Material"
    hdr_cells[1].text = "Quantity (kg/m³)"

    for k, v in data["ssd_mix"].items():
        row = table.add_row().cells
        row[0].text = k
        row[1].text = f"{v:.1f}"

    doc.add_heading("5. Moisture Correction", level=2)
    doc.add_paragraph(
        "Moisture correction is applied according to ACI recommendations:\n"
        "ΔW = Wagg × (MC − Absorption)"
    )
    doc.add_paragraph(
        f"Fine Aggregate ΔW = {data['dw_fine']:.1f} kg/m³\n"
        f"Coarse Aggregate ΔW = {data['dw_coarse']:.1f} kg/m³"
    )

    doc.add_heading("6. Final Batching Quantities", level=2)
    doc.add_paragraph(
        f"Corrected Mixing Water = {data['corrected_water']:.1f} kg/m³"
    )
    doc.add_paragraph(
        f"Fine Aggregate (Batch Weight) = {data['batch_fine']:.1f} kg/m³\n"
        f"Coarse Aggregate (Batch Weight) = {data['batch_coarse']:.1f} kg/m³"
    )

    return doc


# =========================================================
# Streamlit UI
# =========================================================
st.set_page_config(page_title="Concrete Mix Design ACI", layout="centered")
st.title("🧱 Concrete Mix Design (ACI 211)")
st.caption("พร้อม Moisture Correction และ Export รายงาน Word")

st.sidebar.header("📥 Mix Design Input")
wc_ratio = st.sidebar.number_input("w/c", 0.35, 0.70, 0.50, 0.01)
max_agg_mm = st.sidebar.selectbox("Max Aggregate Size (mm)", [20, 25, 40])
sg_cement = st.sidebar.number_input("SG Cement", value=3.15)
sg_fine = st.sidebar.number_input("SG Fine Aggregate", value=2.65)
sg_coarse = st.sidebar.number_input("SG Coarse Aggregate", value=2.70)
air_content = st.sidebar.slider("Air Content (%)", 1.0, 6.0, 2.0) / 100
unit_weight_coarse = st.sidebar.number_input(
    "Unit Weight Coarse Agg (kg/m³)", value=1600
)

st.sidebar.header("💧 Moisture Data")
mc_fine = st.sidebar.number_input("Fine MC (%)", value=5.0)
abs_fine = st.sidebar.number_input("Fine Absorption (%)", value=2.0)
mc_coarse = st.sidebar.number_input("Coarse MC (%)", value=1.0)
abs_coarse = st.sidebar.number_input("Coarse Absorption (%)", value=0.5)

# =========================================================
# Calculation
# =========================================================
if st.button("🧮 Calculate & Generate Report"):

    mix = concrete_mix_design(
        wc_ratio, max_agg_mm, sg_cement, sg_fine,
        sg_coarse, air_content, unit_weight_coarse
    )

    dw_fine, batch_fine = moisture_correction(
        mix["Fine Aggregate"], mc_fine, abs_fine
    )
    dw_coarse, batch_coarse = moisture_correction(
        mix["Coarse Aggregate"], mc_coarse, abs_coarse
    )

    corrected_water = mix["Water"] - (dw_fine + dw_coarse)

    df = pd.DataFrame({
        "Material": ["Water", "Cement", "Fine Agg.", "Coarse Agg."],
        "SSD Quantity (kg/m³)": [
            mix["Water"], mix["Cement"],
            mix["Fine Aggregate"], mix["Coarse Aggregate"]
        ]
    })

    st.subheader("📊 SSD Mix Proportion")
    st.dataframe(df, use_container_width=True)

    # -------- Word Report --------
    report_data = {
        "wc_ratio": wc_ratio,
        "max_agg_mm": max_agg_mm,
        "air_content": air_content,
        "water": mix["Water"],
        "cement": mix["Cement"],
        "vol_coarse": mix["vol_coarse"],
        "ssd_mix": {
            "Water": mix["Water"],
            "Cement": mix["Cement"],
            "Fine Aggregate": mix["Fine Aggregate"],
            "Coarse Aggregate": mix["Coarse Aggregate"]
        },
        "dw_fine": dw_fine,
        "dw_coarse": dw_coarse,
        "batch_fine": batch_fine,
        "batch_coarse": batch_coarse,
        "corrected_water": corrected_water
    }

    doc = generate_word_report(report_data)
    file_name = "Concrete_Mix_Design_ACI.docx"
    doc.save(file_name)

    with open(file_name, "rb") as f:
        st.download_button(
            "📄 Download Word Report",
            f,
            file_name=file_name
        )

    st.success("สร้างรายงาน Word เรียบร้อย ✔")
