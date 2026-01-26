import streamlit as st
import pandas as pd
import numpy as np

st.title("INSEE Low Carbon Concrete – Group Score System")

st.write("## 1️⃣ กำหนดจำนวนกลุ่มแข่งขัน")
n_group = st.number_input("จำนวนกลุ่มแข่งขัน", min_value=1, value=1)

groups = [f"Group {i+1}" for i in range(n_group)]

st.write("## 2️⃣ กรอกข้อมูลทีม")

df = st.data_editor(
    pd.DataFrame({
        "Group": [groups[0]],
        "Team": ["Team A"],
        "Strength_ksc": [240],
        "Cement_OPC": [300],
        "GGBFS": [0],
        "Fly_ash": [0],
        "Limestone": [0],
        "Limestone_fines": [0],
        "Aggregates": [1800],
        "Water": [180],
        "Superplasticizer": [5]
    }),
    num_rows="dynamic"
)

# Emission Factor
EF = {
    "Cement_OPC": 0.912,
    "GGBFS": 0.0416,
    "Fly_ash": 0.004,
    "Limestone": 0.01577,
    "Limestone_fines": 0.01577,
    "Aggregates": 0.00747,
    "Water": 0.000541,
    "Superplasticizer": 1.88
}

if st.button("คำนวณ Embodied Carbon & คะแนน"):
    # --- คำนวณ Embodied Carbon ---
    df["Embodied_Carbon"] = 0.0
    for mat, ef in EF.items():
        df["Embodied_Carbon"] += df[mat] * ef

    df["Pass"] = df["Strength_ksc"] >= 240

    result = []

    for g in groups:
        gdf = df[df["Group"] == g].copy()
        passed = gdf["Pass"]

        if passed.sum() == 0:
            continue

        S_max = gdf.loc[passed, "Strength_ksc"].max()
        C_max = gdf.loc[passed, "Embodied_Carbon"].max()
        C_min = gdf.loc[passed, "Embodied_Carbon"].min()

        gdf["Strength_Score"] = np.where(
            passed,
            50 - abs(gdf["Strength_ksc"] - 240) / (S_max - 240) * 25,
            0
        )

        gdf["Carbon_Score"] = np.where(
            passed,
            50 - (gdf["Embodied_Carbon"] - C_min) / (C_max - C_min) * 50,
            0
        )

        gdf["Total_Score"] = gdf["Strength_Score"] + gdf["Carbon_Score"]
        gdf = gdf.sort_values("Total_Score", ascending=False)

        result.append(gdf)

    final = pd.concat(result)
    st.write("## 📊 ผลคะแนนแยกตามกลุ่ม")
    st.dataframe(final.round(2))
