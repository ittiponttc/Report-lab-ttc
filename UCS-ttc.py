import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
from io import BytesIO

st.set_page_config("UCS Report", layout="wide")
st.title("🧪 UCS / Axial Compression Test Report")

uploaded = st.file_uploader("Upload Excel file", type=["xlsx"])

def find_value(df, keyword, col_offset=1, row_range=5):
    for i in range(len(df)):
        for j in range(len(df.columns)):
            cell = df.iloc[i, j]
            if isinstance(cell, str) and keyword.lower() in cell.lower():
                for r in range(row_range):
                    try:
                        val = df.iloc[i + r, j + col_offset]
                        if pd.notna(val):
                            return val
                    except:
                        pass
    return None

if uploaded:
    xls = pd.ExcelFile(uploaded)
    sheet = st.selectbox("📑 Select worksheet", xls.sheet_names)
    df = pd.read_excel(uploaded, sheet_name=sheet, header=None)

    st.subheader("📌 General Information (Auto-detected)")

    project = find_value(df, "project")
    location = find_value(df, "location")
    cement = find_value(df, "cement")
    diameter = find_value(df, "diameter")
    height = find_value(df, "height")
    depth = find_value(df, "depth")

    info = {
        "Project": project,
        "Location": location,
        "Cement (%)": cement,
        "Diameter (mm)": diameter,
        "Height (mm)": height,
        "Depth (m)": depth,
    }

    for k, v in info.items():
        if v is None:
            st.warning(f"⚠ {k} not found – please check Excel")
        else:
            st.write(f"**{k}:** {v}")

    # ===========================
    # Detect data table automatically
    # ===========================
    st.subheader("📊 Test Data")

    data_start = None
    for i in range(len(df)):
        if "load" in str(df.iloc[i].values).lower():
            data_start = i + 1
            break

    if data_start is None:
        st.error("❌ Cannot detect data table (Load / Displacement)")
        st.stop()

    data = df.iloc[data_start:data_start+50, :]
    data = data.dropna(how="all", axis=1)
    data = data.dropna(how="all", axis=0)

    data.columns = ["Disp (mm)", "Strain_raw", "Load (kg)", "Stress_raw"][:len(data.columns)]

    st.dataframe(data)

    # ===========================
    # Calculation
    # ===========================
    if diameter and height:
        area_cm2 = np.pi * (diameter/10)**2 / 4
        data["Axial Strain (%)"] = data["Disp (mm)"] / height * 100
        data["Axial Stress (ksc)"] = data["Load (kg)"] / area_cm2

        # ===========================
        # Plot
        # ===========================
        fig, ax = plt.subplots()
        ax.plot(data["Axial Strain (%)"], data["Axial Stress (ksc)"], marker="o")
        ax.set_xlabel("Axial Strain (%)")
        ax.set_ylabel("Axial Stress (ksc)")
        ax.grid(True)

        st.subheader("📈 Stress–Strain Curve")
        st.pyplot(fig)

        # ===========================
        # Export Excel
        # ===========================
        excel_buf = BytesIO()
        with pd.ExcelWriter(excel_buf, engine="xlsxwriter") as writer:
            data.to_excel(writer, index=False, sheet_name="Result")

        st.download_button(
            "⬇ Download Excel",
            excel_buf.getvalue(),
            "UCS_Result.xlsx"
        )

        # ===========================
        # Export Word
        # ===========================
        doc = Document()
        doc.add_heading("Unconfined Compression Test Report", 1)

        for k, v in info.items():
            doc.add_paragraph(f"{k}: {v}")

        img = BytesIO()
        fig.savefig(img, dpi=300)
        img.seek(0)
        doc.add_picture(img, width=4000000)

        word_buf = BytesIO()
        doc.save(word_buf)

        st.download_button(
            "⬇ Download Word",
            word_buf.getvalue(),
            "UCS_Report.docx"
        )
