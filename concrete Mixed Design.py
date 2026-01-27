import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

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
    # ---- Water content & coarse aggregate volume (ACI typical) ----
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

    # ---- Volume calculations ----
    vol_water = water / 1000
    vol_cement = cement / (sg_cement * 1000)
    vol_coarse_abs = weight_coarse / (sg_coarse * 1000)

    vol_fine = 1 - (
        vol_water +
        vol_cement +
        vol_coarse_abs +
        air_content
    )

    weight_fine = vol_fine * sg_fine * 1000

    return {
        "Water": water,
        "Cement": cement,
        "Fine Aggregate": weight_fine,
        "Coarse Aggregate": weight_coarse
    }


def moisture_correction(weight_ssd, mc, absorption):
    """
    weight_ssd : SSD weight (kg/m3)
    mc         : moisture content (%)
    absorption : absorption (%)
    """
    delta_water = weight_ssd * (mc - absorption) / 100
    batch_weight = weight_ssd * (1 + (mc - absorption) / 100)
    return delta_water, batch_weight


# =========================================================
# Streamlit UI
# =========================================================
st.set_page_config(page_title="Concrete Mix Design", layout="centered")

st.title("🧱 Concrete Mix Design with Moisture Correction")
st.caption("ACI Method | Suitable for Teaching & Practical Use")

# ---------------- Sidebar Inputs ----------------
st.sidebar.header("📥 Mix Design Input")

wc_ratio = st.sidebar.number_input(
    "Water–Cement Ratio (w/c)",
    min_value=0.35, max_value=0.70, value=0.50, step=0.01
)

max_agg_mm = st.sidebar.selectbox(
    "Maximum Aggregate Size (mm)",
    [20, 25, 40]
)

sg_cement = st.sidebar.number_input(
    "Specific Gravity of Cement",
    value=3.15
)

sg_fine = st.sidebar.number_input(
    "Specific Gravity of Fine Aggregate",
    value=2.65
)

sg_coarse = st.sidebar.number_input(
    "Specific Gravity of Coarse Aggregate",
    value=2.70
)

air_content = st.sidebar.slider(
    "Air Content (%)",
    min_value=1.0, max_value=6.0, value=2.0
) / 100

unit_weight_coarse = st.sidebar.number_input(
    "Unit Weight of Coarse Aggregate (kg/m³)",
    value=1600
)

# ---------------- Moisture Input ----------------
st.sidebar.header("💧 Moisture Data")

mc_fine = st.sidebar.number_input(
    "Fine Aggregate MC (%)",
    value=5.0
)

abs_fine = st.sidebar.number_input(
    "Fine Aggregate Absorption (%)",
    value=2.0
)

mc_coarse = st.sidebar.number_input(
    "Coarse Aggregate MC (%)",
    value=1.0
)

abs_coarse = st.sidebar.number_input(
    "Coarse Aggregate Absorption (%)",
    value=0.5
)

# =========================================================
# Calculation
# =========================================================
if st.button("🧮 Calculate Mix Design"):

    # ---- Mix design ----
    mix = concrete_mix_design(
        wc_ratio,
        max_agg_mm,
        sg_cement,
        sg_fine,
        sg_coarse,
        air_content,
        unit_weight_coarse
    )

    df_mix = pd.DataFrame({
        "Material": mix.keys(),
        "SSD Quantity (kg/m³)": [round(v, 1) for v in mix.values()]
    })

    st.subheader("📊 Mix Design (SSD Condition)")
    st.dataframe(df_mix, use_container_width=True)

    # ---- Moisture correction ----
    dw_fine, batch_fine = moisture_correction(
        mix["Fine Aggregate"], mc_fine, abs_fine
    )

    dw_coarse, batch_coarse = moisture_correction(
        mix["Coarse Aggregate"], mc_coarse, abs_coarse
    )

    total_delta_water = dw_fine + dw_coarse
    corrected_water = mix["Water"] - total_delta_water

    df_mc = pd.DataFrame({
        "Item": ["Fine Aggregate", "Coarse Aggregate"],
        "SSD Weight (kg/m³)": [
            round(mix["Fine Aggregate"], 1),
            round(mix["Coarse Aggregate"], 1)
        ],
        "Batch Weight (kg/m³)": [
            round(batch_fine, 1),
            round(batch_coarse, 1)
        ],
        "Δ Water (kg/m³)": [
            round(dw_fine, 1),
            round(dw_coarse, 1)
        ]
    })

    st.subheader("💧 Moisture Correction")
    st.dataframe(df_mc, use_container_width=True)

    st.info(
        f"💧 Original Water = {round(mix['Water'],1)} kg/m³\n\n"
        f"💧 Corrected Mixing Water = {round(corrected_water,1)} kg/m³"
    )

    # ---- Pie Chart ----
    st.subheader("📈 Material Proportion (SSD)")
    fig, ax = plt.subplots()
    ax.pie(
        df_mix["SSD Quantity (kg/m³)"],
        labels=df_mix["Material"],
        autopct="%1.1f%%",
        startangle=90
    )
    ax.axis("equal")
    st.pyplot(fig)

    st.success("คำนวณเรียบร้อย ✔")

# ---------------- Footer ----------------
st.markdown("---")
st.caption(
    "Concrete Technology | ACI-based Mix Design | "
    "รองรับการสอน ป.ตรี–โท และงานหน้างานเบื้องต้น"
)
