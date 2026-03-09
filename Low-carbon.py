import streamlit as st
import pandas as pd
import numpy as np

st.title("Low Carbon Concrete – Full Score Simulator")

# =============================================
# Emission Factor (ตามกติกา)
# =============================================
EF = {
    "Cement": 0.81,
    "Micro silica": 0.0416,
    "Gypsum": 0.002536,
    "Limestone": 0.01577,
    "Fly ash": 0.004,
    "Aggregates": 0.00747,
    "Water": 0.000541,
    "Superplasticizer": 1.88
}

# =============================================
# Helper Functions
# =============================================

def get_strength_score(fc_MPa):
    """
    หมวด 1: กำลังอัดคอนกรีต (35 คะแนน)
    f'c ที่ 1 วัน (MPa)
    เกณฑ์ขั้นต่ำ ≥ 15 MPa (ไม่ตัดสิทธิ์ แต่ได้ 0 คะแนน)
    """
    if fc_MPa >= 30:
        return 35
    elif 27 <= fc_MPa < 30:
        return 32
    elif 24 <= fc_MPa < 27:
        return 28
    elif 21 <= fc_MPa < 24:
        return 23
    elif 18 <= fc_MPa < 21:
        return 18
    elif 15 <= fc_MPa < 18:
        return 8
    else:  # < 15 MPa
        return 0

def get_carbon_score(co2):
    """
    หมวด 2: การปลดปล่อยคาร์บอนรวม (35 คะแนน)
    CO₂ (kgCO₂e/m³)
    """
    if co2 <= 240:
        return 35
    elif co2 <= 270:
        return 32
    elif co2 <= 310:
        return 28
    elif co2 <= 350:
        return 23
    elif co2 <= 400:
        return 16
    else:  # > 400
        return 8

def get_efficiency_score(index):
    """
    หมวด 3: ประสิทธิภาพเชิงวิศวกรรม – Strength-to-Carbon Ratio (20 คะแนน)
    Index = f'c (MPa) / CO₂ (kgCO₂e/m³)
    """
    if index >= 0.16:
        return 20
    elif 0.13 <= index < 0.16:
        return 16
    elif 0.10 <= index < 0.13:
        return 12
    elif 0.07 <= index < 0.10:
        return 8
    else:  # < 0.07
        return 4

def get_workability_score(slump_mm):
    """
    หมวด 4: Workability (10 คะแนน)
    เกณฑ์: Slump 7–20 cm (70–200 mm) = ผ่าน
    """
    slump_cm = slump_mm / 10
    if 7 <= slump_cm <= 20:
        return 10
    elif (5 <= slump_cm < 7) or (20 < slump_cm <= 22):
        return 6
    elif (3 <= slump_cm < 5) or (22 < slump_cm <= 25):
        return 3
    else:
        return 0

# =============================================
# 1. กรอกส่วนผสมต่อ 1 m³
# =============================================
st.write("### 1️⃣ กรอกส่วนผสมต่อ 1 m³ (kg)")

mix = {}
cols = st.columns(2)
mat_list = list(EF.keys())
for i, mat in enumerate(mat_list):
    with cols[i % 2]:
        mix[mat] = st.number_input(mat, min_value=0.0, value=0.0, step=1.0, key=f"mix_{mat}")

# คำนวณ Embodied Carbon
carbon = sum(mix[m] * EF[m] for m in mix)

st.info(f"🌱 **Embodied Carbon = {carbon:.2f} kgCO₂e/m³**")

# แสดงรายละเอียด Embodied Carbon
with st.expander("📊 ดูรายละเอียด Embodied Carbon แต่ละวัสดุ"):
    detail_data = []
    for mat in mix:
        qty = mix[mat]
        ef = EF[mat]
        ec = qty * ef
        if qty > 0:
            detail_data.append({
                "วัสดุ": mat,
                "ปริมาณ (kg/m³)": qty,
                "EF (kgCO₂e/kg)": ef,
                "Carbon (kgCO₂e/m³)": round(ec, 4)
            })
    if detail_data:
        df_detail = pd.DataFrame(detail_data)
        st.dataframe(df_detail, use_container_width=True, hide_index=True)
    else:
        st.write("กรุณากรอกปริมาณวัสดุ")

# =============================================
# 2. กำลังอัด และ Workability
# =============================================
st.write("### 2️⃣ ผลทดสอบคอนกรีต")
col1, col2 = st.columns(2)
with col1:
    strength_MPa = st.number_input("f'c ที่ 1 วัน (MPa)", value=25.0, min_value=0.0, step=0.5)
    strength_ksc = strength_MPa * 10
    st.caption(f"= {strength_ksc:.0f} ksc")

with col2:
    slump = st.number_input("Slump (mm)", value=100, min_value=0, step=5)

# ไม่มีการตัดสิทธิ์ — ต่ำกว่า 15 MPa ได้ 0 คะแนนหมวดกำลังอัด
if strength_MPa < 15:
    st.warning(f"⚠️ f'c = {strength_MPa:.1f} MPa ({strength_MPa*10:.0f} ksc) < 15 MPa — ได้ 0 คะแนนหมวดกำลังอัด")
elif strength_MPa < 18:
    st.info(f"f'c = {strength_MPa:.1f} MPa ({strength_MPa*10:.0f} ksc) — 15–17 MPa = 8 คะแนน")
else:
    st.success(f"✅ f'c = {strength_MPa:.1f} MPa ({strength_MPa*10:.0f} ksc)")

# =============================================
# 3. คำนวณคะแนนของทีมนี้
# =============================================
st.write("### 3️⃣ ผลคะแนนของทีมคุณ")

if carbon > 0:
    sc1 = get_strength_score(strength_MPa)
    sc2 = get_carbon_score(carbon)
    index_val = strength_MPa / carbon if carbon > 0 else 0
    sc3 = get_efficiency_score(index_val)
    sc4 = get_workability_score(slump)
    total = sc1 + sc2 + sc3 + sc4

    col_a, col_b, col_c, col_d, col_e = st.columns(5)
    col_a.metric("กำลังอัด\n(35 คะแนน)", f"{sc1}")
    col_b.metric("CO₂ Emission\n(35 คะแนน)", f"{sc2}")
    col_c.metric("Strength/Carbon\n(20 คะแนน)", f"{sc3}")
    col_d.metric("Workability\n(10 คะแนน)", f"{sc4}")
    col_e.metric("รวม (100)", f"{total}")

    st.write("#### 📋 สรุปโครงสร้างคะแนน")
    score_table = pd.DataFrame({
        "หมวด": [
            "กำลังอัดคอนกรีต",
            "CO₂ Emission",
            "Strength / Carbon Index",
            "Workability",
            "รวมทั้งหมด"
        ],
        "ค่าที่วัดได้": [
            f"{strength_MPa:.1f} MPa ({strength_MPa*10:.0f} ksc)",
            f"{carbon:.2f} kgCO₂e/m³",
            f"{index_val:.4f}",
            f"{slump} mm",
            "-"
        ],
        "คะแนนเต็ม": [35, 35, 20, 10, 100],
        "คะแนนที่ได้": [sc1, sc2, sc3, sc4, total]
    })
    st.dataframe(score_table, use_container_width=True, hide_index=True)

else:
    st.warning("⚠️ กรุณากรอกส่วนผสมเพื่อคำนวณ Embodied Carbon")

# =============================================
# 4. เปรียบเทียบหลายทีม
# =============================================
st.write("### 4️⃣ เปรียบเทียบหลายทีม")

st.caption("กรอกข้อมูลทีมอื่นเพื่อเปรียบเทียบ (f'c MPa, CO₂ kgCO₂e/m³, Slump mm)")

n_teams = st.number_input("จำนวนทีมที่ต้องการเปรียบเทียบ (รวมทีมคุณ)", min_value=1, max_value=10, value=3)

team_data = []

# ทีมคุณ
team_data.append({
    "ทีม": "Your Team",
    "f'c (MPa)": strength_MPa,
    "CO₂ (kgCO₂e/m³)": carbon if carbon > 0 else 999,
    "Slump (mm)": slump,
})

# ทีมอื่น
if n_teams > 1:
    st.write("กรอกข้อมูลทีมอื่น:")
    for i in range(1, int(n_teams)):
        c1, c2, c3, c4 = st.columns([2, 2, 2, 2])
        with c1:
            tname = st.text_input(f"ชื่อทีม {i+1}", value=f"Team {chr(64+i)}", key=f"tname_{i}")
        with c2:
            tfc = st.number_input(f"f'c MPa", value=28.0, min_value=0.0, step=0.5, key=f"tfc_{i}")
        with c3:
            tco2 = st.number_input(f"CO₂ kgCO₂e/m³", value=300.0, min_value=0.0, step=1.0, key=f"tco2_{i}")
        with c4:
            tslump = st.number_input(f"Slump mm", value=100, min_value=0, step=5, key=f"tslump_{i}")
        team_data.append({
            "ทีม": tname,
            "f'c (MPa)": tfc,
            "CO₂ (kgCO₂e/m³)": tco2,
            "Slump (mm)": tslump,
        })

if st.button("🔢 คำนวณคะแนนเปรียบเทียบ"):
    df = pd.DataFrame(team_data)

    # คำนวณ Index
    df["Index"] = df["f'c (MPa)"] / df["CO₂ (kgCO₂e/m³)"]

    # คำนวณคะแนน (ไม่มีตัดสิทธิ์)
    df["คะแนนกำลังอัด (35)"] = df["f'c (MPa)"].apply(get_strength_score)
    df["คะแนน CO₂ (35)"] = df["CO₂ (kgCO₂e/m³)"].apply(get_carbon_score)
    df["คะแนน Index (20)"] = df["Index"].apply(get_efficiency_score)
    df["คะแนน Workability (10)"] = df["Slump (mm)"].apply(get_workability_score)
    df["คะแนนรวม"] = (
        df["คะแนนกำลังอัด (35)"] +
        df["คะแนน CO₂ (35)"] +
        df["คะแนน Index (20)"] +
        df["คะแนน Workability (10)"]
    )

    df["อันดับ"] = df["คะแนนรวม"].rank(ascending=False, method="min").astype(int)
    df = df.sort_values("อันดับ")

    # แสดงตารางสรุป
    display_cols = [
        "อันดับ", "ทีม",
        "f'c (MPa)", "CO₂ (kgCO₂e/m³)", "Index",
        "คะแนนกำลังอัด (35)", "คะแนน CO₂ (35)",
        "คะแนน Index (20)", "คะแนน Workability (10)",
        "คะแนนรวม"
    ]
    st.dataframe(df[display_cols].round(4), use_container_width=True, hide_index=True)

    # ไฮไลต์ผู้ชนะ
    winner = df.nlargest(1, "คะแนนรวม")
    if not winner.empty:
        st.success(f"🏆 ผู้ชนะ: **{winner.iloc[0]['ทีม']}** ด้วยคะแนน **{winner.iloc[0]['คะแนนรวม']:.0f}** คะแนน")

# =============================================
# 5. ตารางอ้างอิงเกณฑ์การให้คะแนน
# =============================================
with st.expander("📖 ตารางอ้างอิงเกณฑ์การให้คะแนน"):
    st.write("#### หมวด 1: กำลังอัดคอนกรีต (35 คะแนน) — f'c ที่ 1 วัน")
    st.write("เกณฑ์ขั้นต่ำ: f'c ≥ 15 MPa (ไม่ตัดสิทธิ์ แต่ < 15 MPa ได้ 0 คะแนน)")
    df_s = pd.DataFrame({
        "f'c ที่ 1 วัน": ["≥ 30 MPa (≥ 300 ksc)", "27–29 MPa (270–290 ksc)", "24–26 MPa (240–260 ksc)", "21–23 MPa (210–230 ksc)", "18–20 MPa (180–200 ksc)", "15–17 MPa (150–170 ksc)", "< 15 MPa (< 150 ksc)"],
        "คะแนน": [35, 32, 28, 23, 18, 8, 0]
    })
    st.dataframe(df_s, hide_index=True, use_container_width=False)

    st.write("#### หมวด 2: CO₂ Emission (35 คะแนน) — kgCO₂e/m³")
    df_c = pd.DataFrame({
        "CO₂ (kgCO₂e/m³)": ["≤ 240", "241–270", "271–310", "311–350", "351–400", "> 400"],
        "คะแนน": [35, 32, 28, 23, 16, 8]
    })
    st.dataframe(df_c, hide_index=True, use_container_width=False)

    st.write("#### หมวด 3: Strength-to-Carbon Index (20 คะแนน) — f'c(MPa)/CO₂(kgCO₂e/m³)")
    df_i = pd.DataFrame({
        "Index": ["≥ 0.16", "0.13–0.159", "0.10–0.129", "0.07–0.099", "< 0.07"],
        "คะแนน": [20, 16, 12, 8, 4]
    })
    st.dataframe(df_i, hide_index=True, use_container_width=False)

    st.write("#### หมวด 4: Workability (10 คะแนน) — Slump")
    df_w = pd.DataFrame({
        "Slump": ["7–20 cm (70–200 mm)", "5–6 cm หรือ 21–22 cm", "3–4 cm หรือ 23–25 cm", "< 3 cm หรือ > 25 cm"],
        "คะแนน": [10, 6, 3, 0]
    })
    st.dataframe(df_w, hide_index=True, use_container_width=False)
