import streamlit as st
import pandas as pd
import numpy as np

# =========================================================
# Page setup
# =========================================================
st.set_page_config(
    page_title="INSEE Low Carbon Concrete – Score System",
    layout="wide"
)

st.title("🏗️ INSEE Low Carbon Concrete Competition – Score Simulator")
st.caption("Mix → Embodied Carbon → Score (แยกตามกลุ่มแข่งขัน)")

# =========================================================
# Emission Factor Library (LOCKED – ตามกติกา INSEE)
# =========================================================
EF = {
    "Cement_OPC": 0.912,
    "GGBFS": 0.0416,
    "Fly_ash": 0.004,
    "Limestone": 0.01577,        # ✅ เพิ่ม Limestone
    "Aggregates": 0.00747,
    "Water": 0.000541,
    "Superplasticizer": 1.88
}

with st.expander("📚 Emission Factor Library (kgCO₂/kg)", expanded=False):
    st.dataframe(
        pd.DataFrame.from_dict(EF, orient="index", columns=["kgCO₂/kg"])
    )

# =========================================================
# Input table (รูปแบบเดิม + Group + Limestone)
# =========================================================
st.subheader("📝 กรอกข้อมูลผลการทดสอบและส่วนผสม (ต่อ 1 m³)")

default_df = pd.DataFrame({
    "Group": ["A"],
    "Team": ["Team 1"],
    "Strength_ksc": [240],
    "Cement_OPC": [300],
    "GGBFS": [0],
    "Fly_ash": [0],
    "Limestone": [0],
    "Aggregates": [1800],
    "Water": [180],
    "Superplasticizer": [5]
})

df = st.data_editor(
    default_df,
    num_rows="dynamic",
    use_container_width=True
)

# =========================================================
# Calculation
# =========================================================
if st.button("🧮 คำนวณ Embodied Carbon และคะแนน"):
    calc_df = df.copy()

    # ---- Embodied Carbon calculation ----
    calc_df["Embodied_Carbon"] = 0.0
    for mat, ef in EF.items():
        calc_df["Embodied_Carbon"] += calc_df[mat] * ef

    # ---- Pass / Fail ----
    calc_df["Pass"] = calc_df["Strength_ksc"] >= 240

    results = []

    # ---- Calculate score by Group ----
    for g in calc_df["Group"].unique():
        gdf = calc_df[calc_df["Group"] == g].copy()
        passed = gdf["Pass"]

        if passed.sum() == 0:
            continue

        S_max = gdf.loc[passed, "Strength_ksc"].max()
        C_max = gdf.loc[passed, "Embodied_Carbon"].max()
        C_min = gdf.loc[passed, "Embodied_Carbon"].min()

        # ---- Strength score (50) ----
        gdf["Strength_Score"] = np.where(
            passed,
            50 - abs(gdf["Strength_ksc"] - 240) / (S_max - 240) * 25,
            0
        )

        # ---- Carbon score (50) ----
        gdf["Carbon_Score"] = np.where(
            passed,
            50 - (gdf["Embodied_Carbon"] - C_min) / (C_max - C_min) * 50,
            0
        )

        # ---- Total score ----
        gdf["Total_Score"] = gdf["Strength_Score"] + gdf["Carbon_Score"]

        # ---- Ranking in group ----
        gdf["Rank_in_Group"] = gdf["Total_Score"].rank(
            ascending=False, method="min"
        )

        results.append(gdf)

    if len(results) == 0:
        st.error("❌ ไม่มีทีมใดผ่านเกณฑ์กำลังอัดขั้นต่ำ 240 ksc")
    else:
        final = pd.concat(results).sort_values(
            ["Group", "Rank_in_Group"]
        )

        st.subheader("📊 ผลคะแนน (แยกตามกลุ่มแข่งขัน)")
        st.dataframe(
            final[[
                "Group", "Team",
                "Strength_ksc",
                "Embodied_Carbon",
                "Strength_Score",
                "Carbon_Score",
                "Total_Score",
                "Rank_in_Group"
            ]].round(2),
            use_container_width=True
        )

        # ---- Summary by group ----
        st.subheader("🏆 Top 3 ของแต่ละกลุ่ม")
        for g in final["Group"].unique():
            st.markdown(f"### กลุ่ม {g}")
            st.dataframe(
                final[final["Group"] == g]
                .sort_values("Rank_in_Group")
                .head(3)[[
                    "Rank_in_Group", "Team", "Total_Score"
                ]].round(2),
                use_container_width=True
            )
