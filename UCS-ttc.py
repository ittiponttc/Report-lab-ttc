import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
from io import BytesIO

# =====================================================
# Page setup
# =====================================================
st.set_page_config(page_title="UCS Lab Report", layout="wide")
st.title("🧪 Unconfined Compression Test (UCS) – Lab System")

# =====================================================
# Helper functions
# =====================================================
def find_value_multi(df, keywords, col_offset=1, row_range=4):
    for key in keywords:
        for i in range(len(df)):
            for j in range(len(df.columns)):
                cell = df.iloc[i, j]
                if isinstance(cell, str) and key.lower() in cell.lower():
                    for r in range(row_range):
                        try:
                            val = df.iloc[i + r, j + col_offset]
                            if pd.notna(val):
                                return val
                        except:
                            pass
    return None


def warn(label, value, teaching):
    if value is None:
        st.warning(f"⚠ Missing: {label}")
        if teaching:
            st.info(f"👉 ตรวจสอบหัวข้อ `{label}` ในไฟล์ Excel")


# =====================================================
# UI controls
# =====================================================
uploaded = st.file_uploader("📤 Upload UCS Excel file", type=["xlsx"])

batch_mode = st.checkbox("📂 Batch mode (หลาย specimen)")
teaching_mode = st.checkbox("🧑‍🎓 Teaching mode (แนะนำเด็ก)")

# =====================================================
# Main process
# =====================================================
if uploaded:
    xls = pd.ExcelFile(uploaded)

    if not batch_mode:
        # -------------------------------------------------
        # SINGLE SPECIMEN MODE
        # -------------------------------------------------
        sheet = st.selectbox("📑 Select worksheet", xls.sheet_names)
        df = pd.read_excel(uploaded, sheet_name=sheet, header=None)

        st.subheader("📌 General Information (Auto-detected)")

        project = find_value_multi(df, ["project"])
        location = find_value_multi(df, ["location"])
        cement = find_value_multi(df, ["cement"])
        diameter = find_value_multi(df, ["diameter"])
        height = find_value_multi(df, ["height"])
        depth = find_value_multi(df, ["depth"])

        # fallback (MHT standard)
        if project is None:
            project = df.iloc[6, 3]
        if cement is None:
            cement = df.iloc[15, 5]
        if diameter is None:
            diameter = df.iloc[16, 5]
        if height is None:
            height = df.iloc[17, 5]
        if depth is None:
            depth = df.iloc[18, 5]

        info = {
            "Project": project,
            "Location": location,
            "Cement (%)": cement,
            "Diameter (mm)": diameter,
            "Height (mm)": height,
            "Depth (m)": depth,
        }

        for k, v in info.items():
            warn(k, v, teaching_mode)
            if v is not None:
                st.write(f"**{k}:** {v}")

        # -------------------------------------------------
        # Detect test data table
        # -------------------------------------------------
        data_start = None
        for i in range(len(df)):
            row_text = " ".join([str(x).lower() for x in df.iloc[i].values])
            if "load" in row_text:
                data_start = i + 1
                break

        if data_start is None:
            st.error("❌ Cannot detect Load data table")
            st.stop()

        data = df.iloc[data_start:data_start + 50, :]
        data = data.dropna(how="all")
        data = data.dropna(axis=1, how="all")
        data = data.iloc[:, :4]
        data.columns = [
            "Vertical displacement (mm)",
            "Strain_raw",
            "Load (kg)",
            "Stress_raw"
        ]

        st.subheader("📊 Raw Test Data")
        st.dataframe(data)

        # -------------------------------------------------
        # Engineering calculation
        # -------------------------------------------------
        area_cm2 = np.pi * (diameter / 10) ** 2 / 4
        data["Axial Strain (%)"] = data["Vertical displacement (mm)"] / height * 100
        data["Axial Stress (ksc)"] = data["Load (kg)"] / area_cm2

        qu = data["Axial Stress (ksc)"].max()
        idx_peak = data["Axial Stress (ksc)"].idxmax()
        strain_peak = data.loc[idx_peak, "Axial Strain (%)"]

        q50 = 0.5 * qu
        eps50 = np.interp(q50,
                          data["Axial Stress (ksc)"],
                          data["Axial Strain (%)"])
        E50 = q50 / eps50

        st.subheader("🔢 Key Results")
        col1, col2, col3 = st.columns(3)
        col1.metric("UCS (qu)", f"{qu:.2f} ksc")
        col2.metric("Strain at Peak", f"{strain_peak:.2f} %")
        col3.metric("E₅₀", f"{E50:.2f} ksc/%")

        # -------------------------------------------------
        # Plot
        # -------------------------------------------------
        fig, ax = plt.subplots(figsize=(6, 5))
        ax.plot(data["Axial Strain (%)"], data["Axial Stress (ksc)"], marker="o")
        ax.plot(strain_peak, qu, "ro", label="Peak (UCS)")
        ax.plot([0, eps50], [0, q50], "--", label="E50 secant")
        ax.set_xlabel("Axial Strain (%)")
        ax.set_ylabel("Axial Stress (ksc)")
        ax.grid(True)
        ax.legend()

        st.subheader("📈 Stress–Strain Curve")
        st.pyplot(fig)

        # -------------------------------------------------
        # Export Excel
        # -------------------------------------------------
        excel_buf = BytesIO()
        with pd.ExcelWriter(excel_buf, engine="xlsxwriter") as writer:
            data.to_excel(writer, index=False, sheet_name="UCS_Result")

        st.download_button(
            "⬇️ Download Excel",
            excel_buf.getvalue(),
            "UCS_Result.xlsx"
        )

        # -------------------------------------------------
        # Export Word (Thai–English)
        # -------------------------------------------------
        doc = Document()
        doc.add_heading("รายงานผลการทดสอบแรงอัดไม่ถูกควบคุม", 1)
        doc.add_paragraph(f"กำลังรับแรงอัดสูงสุด (UCS, qu) = {qu:.2f} ksc")
        doc.add_paragraph(f"โมดูลัส E₅₀ = {E50:.2f} ksc/%")

        doc.add_heading("Unconfined Compression Test Report", 1)
        doc.add_paragraph(f"Unconfined Compressive Strength (qu) = {qu:.2f} ksc")
        doc.add_paragraph(f"Secant Modulus E50 = {E50:.2f} ksc/%")

        img = BytesIO()
        fig.savefig(img, dpi=300)
        img.seek(0)
        doc.add_picture(img, width=4000000)

        word_buf = BytesIO()
        doc.save(word_buf)

        st.download_button(
            "⬇️ Download Word Report",
            word_buf.getvalue(),
            "UCS_Report.docx"
        )

    else:
        # -------------------------------------------------
        # BATCH MODE
        # -------------------------------------------------
        data = df.iloc[start:start + 50, :]
data = data.dropna(how="all", axis=0)
data = data.dropna(how="all", axis=1)

ncol = data.shape[1]

if ncol < 2:
    continue  # ข้าม sheet นี้ (ข้อมูลไม่พอ)

# ตั้งชื่อ column ตามจำนวนจริง
if ncol == 2:
    data.columns = ["Disp (mm)", "Load (kg)"]
elif ncol == 3:
    data.columns = ["Disp (mm)", "Load (kg)", "Extra"]
else:
    data = data.iloc[:, :4]
    data.columns = ["Disp (mm)", "Strain_raw", "Load (kg)", "Stress_raw"]
