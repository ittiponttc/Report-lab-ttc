import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
from io import BytesIO

st.set_page_config(page_title="UCS Report (ASTM D2166)", layout="wide")
st.title("📊 Unconfined Compression Test Report")

uploaded = st.file_uploader("Upload MHT.24-68.xlsx", type=["xlsx"])

if uploaded:
    sheet = "Ver.2"
    df_raw = pd.read_excel(uploaded, sheet_name=sheet, header=None)

    # ===============================
    # Read general information
    # ===============================
    project = df_raw.iloc[6, 3]
    location = df_raw.iloc[7, 3]
    rate = df_raw.iloc[8, 3]

    cement = df_raw.iloc[15, 5]
    diameter = df_raw.iloc[16, 5]     # mm
    height = df_raw.iloc[17, 5]       # mm
    depth = df_raw.iloc[18, 5]        # m

    st.subheader("📌 Project Information")
    st.write(f"**Project:** {project}")
    st.write(f"**Location:** {location}")
    st.write(f"**Shearing Rate:** {rate}")
    st.write(f"**Cement Content:** {cement} %")
    st.write(f"**Depth:** {depth} m")

    # ===============================
    # Read test data
    # ===============================
    data = df_raw.iloc[22:52, 1:5]
    data.columns = [
        "Vertical displacement (mm)",
        "Raw strain",
        "Load (kg)",
        "Raw stress"
    ]
    data = data.dropna()

    # ===============================
    # Engineering calculation
    # ===============================
    area_cm2 = np.pi * (diameter / 10) ** 2 / 4

    data["Axial Strain (%)"] = (
        data["Vertical displacement (mm)"] / height * 100
    )

    data["Axial Stress (ksc)"] = (
        data["Load (kg)"] / area_cm2
    )

    st.subheader("📐 Calculated Results")
    st.dataframe(data)

    # ===============================
    # Plot
    # ===============================
    fig, ax = plt.subplots(figsize=(6,5))
    ax.plot(
        data["Axial Strain (%)"],
        data["Axial Stress (ksc)"],
        marker="o"
    )
    ax.set_xlabel("Axial Strain (%)")
    ax.set_ylabel("Axial Stress (ksc)")
    ax.grid(True)

    st.subheader("📈 Stress–Strain Curve")
    st.pyplot(fig)

    # ===============================
    # Export Excel
    # ===============================
    excel_buf = BytesIO()
    with pd.ExcelWriter(excel_buf, engine="xlsxwriter") as writer:
        data.to_excel(writer, index=False, sheet_name="UCS_Result")

    st.download_button(
        "⬇️ Download Excel",
        excel_buf.getvalue(),
        "UCS_Result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ===============================
    # Export Word
    # ===============================
    doc = Document()
    doc.add_heading("Unconfined Compression Test Report", 1)

    doc.add_paragraph(f"Project: {project}")
    doc.add_paragraph(f"Location: {location}")
    doc.add_paragraph(f"Cement content: {cement} %")
    doc.add_paragraph(f"Depth: {depth} m")
    doc.add_paragraph(f"Diameter: {diameter} mm")
    doc.add_paragraph(f"Height: {height} mm")

    img = BytesIO()
    fig.savefig(img, dpi=300)
    img.seek(0)
    doc.add_picture(img, width=4000000)

    word_buf = BytesIO()
    doc.save(word_buf)

    st.download_button(
        "⬇️ Download Word Report",
        word_buf.getvalue(),
        "UCS_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
