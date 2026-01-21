import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
from io import BytesIO

# =====================================================
# Page setup
# =====================================================
st.set_page_config(page_title="UCS Lab System", layout="wide")
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


def detect_data_table(df):
    """
    หาแถวเริ่มต้นของตารางทดสอบจากคำว่า Load
    """
    for i in range(len(df)):
        row_text = " ".join([str(x).lower() for x in df.iloc[i].values])
        if "load" in row_text:
            return i + 1
    return None


def prepare_data(df, start_row):
    """
    เตรียมตารางข้อมูลแบบ error-proof
    """
    data = df.iloc[start_row:start_row + 60, :]
    data = data.dropna(how="all", axis=0)
    data = data.dropna(how="all", axis=1)

    ncol = data.shape[1]

    if ncol < 2:
        return None

    # ตั้งชื่อ column ตามจำนวนจริง
    if ncol == 2:
        data.columns = ["Disp (mm)", "Load (kg)"]
    elif ncol == 3:
        data.columns = ["Disp (mm)", "Load (kg)", "Extra"]
    else:
        data = data.iloc[:, :4]
        data.columns = [
            "Disp (mm)",
            "Strain_raw",
            "Load (kg)",
            "Stress_raw"
        ]

    return data


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

    # =================================================
    # SINGLE SPECIMEN MODE
    # =================================================
    if not batch_mode:
        sheet = st.selectbox("📑 Select worksheet", xls.sheet_names)
        df = pd.read_excel(uploaded, sheet_name=sheet, header=None)

        st.subheader("📌 General Information (Auto-detected)")

        project = find_value_multi(df, ["project"])
        location = find_value_multi(df, ["location"])
        cement = find_value_multi(df, ["cement"])
        diameter = find_value_multi(df, ["diameter"])
        height = find_value_multi(df, ["height"])
        depth = find_value_multi(df, ["depth"])

        # fallback ตามฟอร์ม MHT
        if project is None:
            project = df.iloc[6, 3] if df.shape[0] > 6 else None
        if cement is None:
            cement = df.iloc[15, 5] if df.shape[0] > 15 else None
        if diameter is None:
            diameter = df.iloc[16, 5] if df.shape[0] > 16 else None
        if height is None:
            height = df.iloc[17, 5] if df.shape[0] > 17 else None
        if depth is None:
            depth = df.iloc[18, 5] if df.shape[0] > 18 else None

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

        # =================================================
        # Detect and prepare test data
        # =================================================
        start_row = detect_data_table(df)
        if start_row is None:
            st.error("❌ Cannot detect Load data table")
            st.stop()

        data = prepare_data(df, start_row)
        if data is None:
            st.error("❌ Data table not valid")
            st.stop()

        st.subheader("📊 Raw Test Data")
        st.dataframe(data)

        # =================================================
        # Engineering calculation
        # =================================================
        if diameter is None or height is None:
            st.error("❌ Diameter or Height missing")
            st.stop()

        area_cm2 = np.pi * (diameter / 10) ** 2 / 4

        data["Axial Strain (%)"] = data["Disp (mm)"] / height * 100
        data["Axial Stress (ksc)"] = data["Load (kg)"] / area_cm2

        # UCS & Peak
        qu = data["Axial Stress (ksc)"].max()
        idx_peak = data["Axial Stress (ksc)"].idxmax()
        strain_peak = data.loc[idx_peak, "Axial Strain (%)"]

        # E50
        q50 = 0.5 * qu
        eps50 = np.interp(
            q50,
            data["Axial Stress (ksc)"],
            data["Axial Strain (%)"]
        )
        E50 = q50 / eps50

        st.subheader("🔢 Key Results")
        c1, c2, c3 = st.columns(3)
        c1.metric("UCS (qu)", f"{qu:.2f} ksc")
        c2.metric("Strain at Peak", f"{strain_peak:.2f} %")
        c3.metric("E₅₀", f"{E50:.2f} ksc/%")

        # =================================================
        # Plot
        # =================================================
        fig, ax = plt.subplots(figsize=(6, 5))
        ax.plot(
            data["Axial Strain (%)"],
            data["Axial Stress (ksc)"],
            marker="o",
            label="Stress–Strain"
        )
        ax.plot(strain_peak, qu, "ro", label="Peak (UCS)")
        ax.plot([0, eps50], [0, q50], "--", label="E50 secant")

        ax.set_xlabel("Axial Strain (%)")
        ax.set_ylabel("Axial Stress (ksc)")
        ax.grid(True)
        ax.legend()

        st.subheader("📈 Stress–Strain Curve")
        st.pyplot(fig)

        # =================================================
        # Export Excel
        # =================================================
        excel_buf = BytesIO()
        with pd.ExcelWriter(excel_buf, engine="xlsxwriter") as writer:
            data.to_excel(writer, index=False, sheet_name="UCS_Result")

        st.download_button(
            "⬇️ Download Excel",
            excel_buf.getvalue(),
            "UCS_Result.xlsx"
        )

        # =================================================
        # Export Word (Thai–English)
        # =================================================
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

    # =================================================
    # BATCH MODE
    # =================================================
    else:
        st.subheader("📂 Batch Mode Summary")

        summary = []

        for sheet in xls.sheet_names:
            df = pd.read_excel(uploaded, sheet_name=sheet, header=None)

            start_row = detect_data_table(df)
            if start_row is None:
                if teaching_mode:
                    st.warning(f"⚠ {sheet}: ไม่พบตาราง Load")
                continue

            data = prepare_data(df, start_row)
            if data is None or "Load (kg)" not in data.columns:
                if teaching_mode:
                    st.warning(f"⚠ {sheet}: ตารางข้อมูลไม่สมบูรณ์")
                continue

            diameter = find_value_multi(df, ["diameter"])
            height = find_value_multi(df, ["height"])

            if diameter is None or height is None:
                if teaching_mode:
                    st.warning(f"⚠ {sheet}: ไม่มี Diameter/Height")
                continue

            area_cm2 = np.pi * (diameter / 10) ** 2 / 4
            data["Axial Stress (ksc)"] = data["Load (kg)"] / area_cm2

            qu = data["Axial Stress (ksc)"].max()

            summary.append({
                "Specimen (Sheet)": sheet,
                "UCS (ksc)": round(qu, 2)
            })

        summary_df = pd.DataFrame(summary)
        st.dataframe(summary_df)
