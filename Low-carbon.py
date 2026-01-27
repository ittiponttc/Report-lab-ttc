import streamlit as st
import pandas as pd
import numpy as np

st.title("Low Carbon Concrete – Full Score Simulator")

# Emission factor (ตามกติกา)
EF = {
    "Cement": 0.912,
    "GGBFS": 0.0416,
    "Fly ash": 0.004,
    "Aggregates": 0.00747,
    "Water": 0.000541,
    "Superplasticizer": 1.88
}

st.write("### 1️⃣ กรอกส่วนผสมต่อ 1 m³ (kg)")

mix = {}
for mat in EF:
    mix[mat] = st.number_input(mat, min_value=0.0, value=0.0)

# คำนวณ Embodied Carbon
carbon = sum(mix[m] * EF[m] for m in mix)

st.info(f"🌱 Embodied Carbon = **{carbon:.2f} kgCO₂/m³**")

st.write("### 2️⃣ ใส่ผลกำลังอัดเฉลี่ย (24 ชม.)")
strength = st.number_input("Avg Strength (ksc)", value=240.0)

if strength < 240:
    st.error("❌ ไม่ผ่านเกณฑ์ขั้นต่ำ 240 ksc")
else:
    st.success("✅ ผ่านเกณฑ์กำลังอัด")

# --- ส่วนเปรียบเทียบหลายทีม (ตัวอย่าง) ---
st.write("### 3️⃣ ตัวอย่างเปรียบเทียบหลายทีม")

df = pd.DataFrame({
    "Team": ["Your Team", "Team A", "Team B"],
    "Strength": [strength, 245, 255],
    "Carbon": [carbon, 300, 270]
})

passed = df["Strength"] >= 240

S_max = df.loc[passed, "Strength"].max()
C_max = df.loc[passed, "Carbon"].max()
C_min = df.loc[passed, "Carbon"].min()

df["Strength_Score"] = np.where(
    passed,
    50 - abs(df["Strength"] - 240) / (S_max - 240) * 25,
    0
)

df["Carbon_Score"] = np.where(
    passed,
    50 - (df["Carbon"] - C_min) / (C_max - C_min) * 50,
    0
)

df["Total"] = df["Strength_Score"] + df["Carbon_Score"]

st.dataframe(df.round(2))
