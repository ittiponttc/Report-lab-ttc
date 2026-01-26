import streamlit as st
import pandas as pd
import numpy as np

st.title("INSEE Low Carbon Concrete – Score Simulator")

st.write("### ใส่ข้อมูลผลการทดสอบ")

df = st.data_editor(
    pd.DataFrame({
        "Team": ["A", "B", "C"],
        "Avg_Strength_ksc": [250, 242, 260],
        "Embodied_Carbon": [280, 300, 260]
    }),
    num_rows="dynamic"
)

if st.button("คำนวณคะแนน"):
    # เงื่อนไขผ่าน
    df["Pass"] = df["Avg_Strength_ksc"] >= 240

    S_max = df.loc[df["Pass"], "Avg_Strength_ksc"].max()
    C_max = df.loc[df["Pass"], "Embodied_Carbon"].max()
    C_min = df.loc[df["Pass"], "Embodied_Carbon"].min()

    def strength_score(row):
        if not row["Pass"]:
            return 0
        return 50 - abs(row["Avg_Strength_ksc"] - 240) / (S_max - 240) * 25

    def carbon_score(row):
        if not row["Pass"]:
            return 0
        return 50 - (row["Embodied_Carbon"] - C_min) / (C_max - C_min) * 50

    df["Strength_Score"] = df.apply(strength_score, axis=1)
    df["Carbon_Score"] = df.apply(carbon_score, axis=1)
    df["Total_Score"] = df["Strength_Score"] + df["Carbon_Score"]

    df = df.sort_values("Total_Score", ascending=False)

    st.write("### 📊 ผลคะแนน")
    st.dataframe(df.round(2))

