"""
Hoek-Brown to Mohr-Coulomb Parameter Conversion
แอปพลิเคชันสำหรับแปลงค่าพารามิเตอร์จาก Hoek-Brown criterion เป็น Mohr-Coulomb

อ้างอิง: Hoek, E., Carranza-Torres, C., and Corkum, B. (2002)
"Hoek-Brown failure criterion – 2002 Edition"

พัฒนาโดย: ภาควิชาครุศาสตร์โยธา คณะครุศาสตร์อุตสาหกรรม
มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ
"""

import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import rcParams
import json
from io import BytesIO

# ตั้งค่า font สำหรับภาษาไทย
plt.rcParams['font.family'] = 'DejaVu Sans'

# ===== Page Configuration =====
st.set_page_config(
    page_title="Hoek-Brown to Mohr-Coulomb Conversion",
    page_icon="🏔️",
    layout="wide"
)

# ===== Custom CSS =====
st.markdown("""
<style>
    .main-header {
        font-size: 28px;
        font-weight: bold;
        color: #1E3A5F;
        text-align: center;
        padding: 10px;
        background: linear-gradient(90deg, #E8F4F8 0%, #D5E8F0 100%);
        border-radius: 10px;
        margin-bottom: 20px;
    }
    .sub-header {
        font-size: 18px;
        color: #2C5F7C;
        border-bottom: 2px solid #2C5F7C;
        padding-bottom: 5px;
        margin-top: 20px;
    }
    .result-box {
        background-color: #F0F8FF;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #2C5F7C;
        margin: 10px 0;
    }
    .formula-box {
        background-color: #FFF8E7;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #D4A574;
        margin: 10px 0;
        font-family: 'Times New Roman', serif;
    }
    .warning-box {
        background-color: #FFF3CD;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #FFC107;
        margin: 10px 0;
    }
    .info-box {
        background-color: #E7F3FF;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #0066CC;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# ===== Header =====
st.markdown('<div class="main-header">🏔️ Hoek-Brown to Mohr-Coulomb Conversion<br>การแปลงค่าพารามิเตอร์จาก Hoek-Brown criterion เป็น Mohr-Coulomb</div>', unsafe_allow_html=True)

# ===== Functions =====
def calculate_hoek_brown_parameters(GSI, mi, D):
    """คำนวณพารามิเตอร์ Hoek-Brown (mb, s, a)"""
    mb = mi * np.exp((GSI - 100) / (28 - 14 * D))
    s = np.exp((GSI - 100) / (9 - 3 * D))
    a = 0.5 + (1/6) * (np.exp(-GSI/15) - np.exp(-20/3))
    return mb, s, a

def calculate_sigma3max(H, gamma, factor=0.95):
    """
    คำนวณค่า σ3max สำหรับ slope
    
    Parameters:
    - H: ความสูงลาดดิน (m)
    - gamma: หน่วยน้ำหนัก (kN/m³)
    - factor: ค่าสัมประสิทธิ์ (default=0.95 สำหรับ slope, 0.72 สำหรับ tunnel)
    
    σ3max = factor × γ × H (หน่วย: kPa → MPa ต้องหาร 1000)
    """
    sigma3max = factor * gamma * H / 1000  # แปลงเป็น MPa
    return sigma3max

def calculate_mohr_coulomb_fit(UCS, mb, s, a, sigma3n):
    """
    คำนวณค่า phi และ c จาก Hoek-Brown criterion
    ใช้สมการจาก Hoek et al. (2002)
    """
    # คำนวณ phi (มุมเสียดทานภายใน)
    numerator = 6 * a * mb * (s + mb * sigma3n)**(a - 1)
    denominator = 2 * (1 + a) * (2 + a) + 6 * a * mb * (s + mb * sigma3n)**(a - 1)
    sin_phi = numerator / denominator
    phi = np.degrees(np.arcsin(sin_phi))
    
    # คำนวณ c (cohesion)
    term1 = UCS * ((1 + 2*a) * s + (1 - a) * mb * sigma3n) * (s + mb * sigma3n)**(a - 1)
    term2 = (1 + a) * (2 + a) * np.sqrt(1 + (6 * a * mb * (s + mb * sigma3n)**(a - 1)) / ((1 + a) * (2 + a)))
    c = term1 / term2
    
    return phi, c

def hoek_brown_criterion(sigma3, UCS, mb, s, a):
    """คำนวณ σ1 จาก Hoek-Brown criterion"""
    sigma1 = sigma3 + UCS * (mb * sigma3 / UCS + s)**a
    return sigma1

def mohr_coulomb_criterion(sigma3, c, phi_deg):
    """คำนวณ σ1 จาก Mohr-Coulomb criterion"""
    phi_rad = np.radians(phi_deg)
    sigma1 = sigma3 * (1 + np.sin(phi_rad)) / (1 - np.sin(phi_rad)) + 2 * c * np.cos(phi_rad) / (1 - np.sin(phi_rad))
    return sigma1

def plot_failure_envelopes(UCS, mb, s, a, c, phi, sigma3max, sigma3n):
    """สร้างกราฟเปรียบเทียบ failure envelopes"""
    fig, axes = plt.subplots(1, 2, figsize=(14, 5))
    
    # === Plot 1: σ1 vs σ3 ===
    ax1 = axes[0]
    
    # สร้างช่วงค่า σ3
    sigma3_range = np.linspace(0, sigma3max * 1.5, 100)
    
    # Hoek-Brown envelope
    sigma1_hb = hoek_brown_criterion(sigma3_range, UCS, mb, s, a)
    ax1.plot(sigma3_range, sigma1_hb, 'b-', linewidth=2, label='Hoek-Brown')
    
    # Mohr-Coulomb envelope
    sigma1_mc = mohr_coulomb_criterion(sigma3_range, c, phi)
    ax1.plot(sigma3_range, sigma1_mc, 'r--', linewidth=2, label='Mohr-Coulomb (Fit)')
    
    # Mark σ3max
    ax1.axvline(x=sigma3max, color='green', linestyle=':', linewidth=1.5, label=f'σ3max = {sigma3max:.4f} MPa')
    
    ax1.set_xlabel('σ₃ (MPa)', fontsize=12)
    ax1.set_ylabel('σ₁ (MPa)', fontsize=12)
    ax1.set_title('Failure Envelope: σ₁ vs σ₃', fontsize=14, fontweight='bold')
    ax1.legend(loc='upper left')
    ax1.grid(True, alpha=0.3)
    ax1.set_xlim(0, None)
    ax1.set_ylim(0, None)
    
    # === Plot 2: Mohr Circle (τ vs σ) ===
    ax2 = axes[1]
    
    # วาด Mohr-Coulomb failure line
    sigma_range = np.linspace(-c/np.tan(np.radians(phi)) * 0.5, sigma3max * 3, 100)
    tau_mc = c + sigma_range * np.tan(np.radians(phi))
    ax2.plot(sigma_range, tau_mc, 'r-', linewidth=2, label=f'Mohr-Coulomb: τ = {c:.3f} + σ·tan({phi:.1f}°)')
    
    # วาด Mohr circles ตัวอย่าง
    sigma3_samples = [0, sigma3max/2, sigma3max]
    colors = ['blue', 'green', 'orange']
    
    for sig3, color in zip(sigma3_samples, colors):
        sig1_hb = hoek_brown_criterion(sig3, UCS, mb, s, a)
        center = (sig1_hb + sig3) / 2
        radius = (sig1_hb - sig3) / 2
        
        theta = np.linspace(0, np.pi, 100)
        x_circle = center + radius * np.cos(theta)
        y_circle = radius * np.sin(theta)
        
        ax2.plot(x_circle, y_circle, color=color, linewidth=1.5, 
                label=f'σ₃ = {sig3:.3f} MPa')
    
    ax2.set_xlabel('σ (MPa)', fontsize=12)
    ax2.set_ylabel('τ (MPa)', fontsize=12)
    ax2.set_title('Mohr Circles and Failure Envelope', fontsize=14, fontweight='bold')
    ax2.legend(loc='upper left', fontsize=9)
    ax2.grid(True, alpha=0.3)
    ax2.set_ylim(0, None)
    ax2.axhline(y=0, color='black', linewidth=0.5)
    ax2.axvline(x=0, color='black', linewidth=0.5)
    
    plt.tight_layout()
    return fig

def save_to_json(data):
    """แปลงข้อมูลเป็น JSON"""
    return json.dumps(data, indent=2, ensure_ascii=False)

def load_from_json(json_str):
    """โหลดข้อมูลจาก JSON"""
    return json.loads(json_str)

# ===== Sidebar - Input Parameters =====
with st.sidebar:
    st.markdown("### 📥 ข้อมูลนำเข้า (Input Parameters)")
    
    # JSON Upload
    st.markdown("---")
    st.markdown("#### 📂 โหลดข้อมูลจาก JSON")
    uploaded_file = st.file_uploader("เลือกไฟล์ JSON", type=['json'], key="json_upload")
    
    if uploaded_file is not None:
        try:
            json_data = json.load(uploaded_file)
            st.success("✅ โหลดไฟล์สำเร็จ!")
            
            # ใช้ค่าจาก JSON
            default_UCS = json_data.get('UCS', 50.0)
            default_GSI = json_data.get('GSI', 54)
            default_mi = json_data.get('mi', 7)
            default_D = json_data.get('D', 1.0)
            default_H = json_data.get('H', 6.0)
            default_gamma = json_data.get('gamma', 26.0)
        except Exception as e:
            st.error(f"❌ ไม่สามารถอ่านไฟล์: {e}")
            default_UCS, default_GSI, default_mi, default_D = 50.0, 54, 7, 1.0
            default_H, default_gamma = 6.0, 26.0
    else:
        default_UCS, default_GSI, default_mi, default_D = 50.0, 54, 7, 1.0
        default_H, default_gamma = 6.0, 26.0
    
    st.markdown("---")
    st.markdown("#### 🪨 Hoek-Brown Classification")
    
    UCS = st.number_input(
        "UCS - Uniaxial Compressive Strength (MPa)",
        min_value=1.0,
        max_value=500.0,
        value=default_UCS,
        step=1.0,
        help="กำลังรับแรงอัดแกนเดียวของหินสมบูรณ์"
    )
    
    GSI = st.number_input(
        "GSI - Geological Strength Index",
        min_value=5,
        max_value=100,
        value=default_GSI,
        step=1,
        help="ดัชนีความแข็งแรงทางธรณีวิทยา (5-100)"
    )
    
    mi = st.number_input(
        "mᵢ - Material constant for intact rock",
        min_value=1,
        max_value=40,
        value=default_mi,
        step=1,
        help="ค่าคงที่ของวัสดุสำหรับหินสมบูรณ์"
    )
    
    D = st.number_input(
        "D - Disturbance factor",
        min_value=0.0,
        max_value=1.0,
        value=default_D,
        step=0.1,
        help="ปัจจัยการรบกวน (0 = undisturbed, 1 = very disturbed)"
    )
    
    st.markdown("---")
    st.markdown("#### 📐 Application Parameters (Slope)")
    
    H = st.number_input(
        "H - Slope height (m)",
        min_value=1.0,
        max_value=500.0,
        value=default_H,
        step=1.0,
        help="ความสูงของลาดดิน"
    )
    
    gamma = st.number_input(
        "γ - Unit weight (kN/m³)",
        min_value=10.0,
        max_value=35.0,
        value=default_gamma,
        step=0.5,
        help="หน่วยน้ำหนักของมวลหิน"
    )
    
    st.markdown("---")
    st.markdown("#### ⚙️ σ3max Calculation Method")
    
    sigma3_method = st.radio(
        "วิธีคำนวณ σ3max",
        ["Slope (factor = 0.95)", "Tunnel (factor = 0.72)", "Custom factor"],
        help="เลือกวิธีคำนวณตามประเภทงาน"
    )
    
    if sigma3_method == "Slope (factor = 0.95)":
        sigma3_factor = 0.95
    elif sigma3_method == "Tunnel (factor = 0.72)":
        sigma3_factor = 0.72
    else:
        sigma3_factor = st.number_input(
            "Custom factor",
            min_value=0.1,
            max_value=2.0,
            value=0.95,
            step=0.05
        )

# ===== Main Calculation =====
# คำนวณพารามิเตอร์ Hoek-Brown
mb, s, a = calculate_hoek_brown_parameters(GSI, mi, D)

# คำนวณ σ3max และ σ3n
sigma3max = calculate_sigma3max(H, gamma, sigma3_factor)
sigma3n = sigma3max / UCS

# คำนวณ Mohr-Coulomb parameters
phi, c = calculate_mohr_coulomb_fit(UCS, mb, s, a, sigma3n)

# ===== Display Results =====
col1, col2 = st.columns(2)

with col1:
    st.markdown('<div class="sub-header">📊 Hoek-Brown Classification</div>', unsafe_allow_html=True)
    
    st.markdown(f"""
    <div class="result-box">
    <table style="width:100%">
        <tr><td><b>UCS</b></td><td>=</td><td><b>{UCS} MPa</b></td></tr>
        <tr><td><b>GSI</b></td><td>=</td><td><b>{GSI}</b></td></tr>
        <tr><td><b>mᵢ</b></td><td>=</td><td><b>{mi}</b></td></tr>
        <tr><td><b>D</b></td><td>=</td><td><b>{D}</b></td></tr>
    </table>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown('<div class="sub-header">🔢 Hoek-Brown Criterion Parameters</div>', unsafe_allow_html=True)
    
    st.markdown(f"""
    <div class="formula-box">
    <p><b>สูตรการคำนวณ:</b></p>
    <p>m<sub>b</sub> = m<sub>i</sub> · exp[(GSI - 100)/(28 - 14D)] = <b>{mb:.4f}</b></p>
    <p>s = exp[(GSI - 100)/(9 - 3D)] = <b>{s:.4f}</b></p>
    <p>a = ½ + ⅙(e<sup>-GSI/15</sup> - e<sup>-20/3</sup>) = <b>{a:.3f}</b></p>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown('<div class="sub-header">📐 Failure Envelope Range</div>', unsafe_allow_html=True)
    
    st.markdown(f"""
    <div class="info-box">
    <p><b>Application: Slope (H = {H} m, Unit Weight = {gamma} kN/m³)</b></p>
    <table style="width:100%">
        <tr><td>σ<sub>3max</sub></td><td>=</td><td>{sigma3_factor} × γ × H / 1000</td><td>=</td><td><b>{sigma3max:.4f} MPa</b></td></tr>
        <tr><td>σ<sub>3n</sub></td><td>=</td><td>σ<sub>3max</sub> / σ<sub>ci</sub></td><td>=</td><td><b>{sigma3n:.6f}</b></td></tr>
    </table>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown('<div class="sub-header">⚡ Mohr-Coulomb Fit</div>', unsafe_allow_html=True)
    
    st.markdown(f"""
    <div class="result-box" style="background-color: #E8F5E9; border-left: 5px solid #4CAF50;">
    <h3 style="color: #2E7D32; text-align: center;">ผลลัพธ์การแปลงค่า</h3>
    <table style="width:100%; font-size: 18px;">
        <tr>
            <td><b>φ (Friction Angle)</b></td>
            <td>=</td>
            <td style="color: #C62828; font-size: 24px;"><b>{phi:.2f}°</b></td>
        </tr>
        <tr>
            <td><b>c (Cohesion)</b></td>
            <td>=</td>
            <td style="color: #C62828; font-size: 24px;"><b>{c:.3f} MPa</b></td>
        </tr>
    </table>
    </div>
    """, unsafe_allow_html=True)

# ===== Detailed Formulas =====
with st.expander("📝 แสดงรายละเอียดสูตรคำนวณ (Detailed Formulas)", expanded=False):
    st.markdown("""
    ### สูตรการแปลง Hoek-Brown เป็น Mohr-Coulomb
    
    **Hoek-Brown Criterion:**
    """)
    st.latex(r"\sigma_1 = \sigma_3 + \sigma_{ci}\left(m_b\frac{\sigma_3}{\sigma_{ci}} + s\right)^a")
    
    st.markdown("**Friction Angle (φ):**")
    st.latex(r"\phi = \sin^{-1}\left[\frac{6am_b(s + m_b\sigma_{3n})^{a-1}}{2(1+a)(2+a) + 6am_b(s + m_b\sigma_{3n})^{a-1}}\right]")
    
    st.markdown("**Cohesion (c):**")
    st.latex(r"c = \frac{\sigma_{ci}[(1+2a)s + (1-a)m_b\sigma_{3n}](s + m_b\sigma_{3n})^{a-1}}{(1+a)(2+a)\sqrt{1 + \frac{6am_b(s+m_b\sigma_{3n})^{a-1}}{(1+a)(2+a)}}}")
    
    st.markdown("""
    **โดยที่:**
    - σ₃ₙ = σ₃ₘₐₓ / σcᵢ (normalized confining stress)
    - σ₃ₘₐₓ = 0.72 × γ × H / 1000 (สำหรับงานลาดดิน)
    
    **Reference:** Hoek, E., Carranza-Torres, C., and Corkum, B. (2002)
    """)

# ===== Plot =====
st.markdown('<div class="sub-header">📈 Failure Envelope Comparison</div>', unsafe_allow_html=True)

fig = plot_failure_envelopes(UCS, mb, s, a, c, phi, sigma3max, sigma3n)
st.pyplot(fig)

# ===== Sensitivity Analysis =====
st.markdown('<div class="sub-header">🔬 Sensitivity Analysis</div>', unsafe_allow_html=True)

with st.expander("วิเคราะห์ความไวของพารามิเตอร์ (Parameter Sensitivity)", expanded=False):
    
    param_choice = st.selectbox(
        "เลือกพารามิเตอร์ที่ต้องการวิเคราะห์",
        ["GSI", "D (Disturbance)", "H (Slope Height)", "mi"]
    )
    
    fig_sens, axes_sens = plt.subplots(1, 2, figsize=(12, 4))
    
    if param_choice == "GSI":
        gsi_range = np.arange(20, 95, 5)
        phi_values = []
        c_values = []
        
        for gsi_val in gsi_range:
            mb_temp, s_temp, a_temp = calculate_hoek_brown_parameters(gsi_val, mi, D)
            phi_temp, c_temp = calculate_mohr_coulomb_fit(UCS, mb_temp, s_temp, a_temp, sigma3n)
            phi_values.append(phi_temp)
            c_values.append(c_temp)
        
        axes_sens[0].plot(gsi_range, phi_values, 'b-o', linewidth=2)
        axes_sens[0].axvline(x=GSI, color='red', linestyle='--', label=f'Current GSI = {GSI}')
        axes_sens[0].set_xlabel('GSI')
        axes_sens[0].set_ylabel('φ (degrees)')
        axes_sens[0].set_title('Friction Angle vs GSI')
        axes_sens[0].legend()
        axes_sens[0].grid(True, alpha=0.3)
        
        axes_sens[1].plot(gsi_range, c_values, 'r-o', linewidth=2)
        axes_sens[1].axvline(x=GSI, color='blue', linestyle='--', label=f'Current GSI = {GSI}')
        axes_sens[1].set_xlabel('GSI')
        axes_sens[1].set_ylabel('c (MPa)')
        axes_sens[1].set_title('Cohesion vs GSI')
        axes_sens[1].legend()
        axes_sens[1].grid(True, alpha=0.3)
        
    elif param_choice == "D (Disturbance)":
        d_range = np.arange(0, 1.05, 0.1)
        phi_values = []
        c_values = []
        
        for d_val in d_range:
            mb_temp, s_temp, a_temp = calculate_hoek_brown_parameters(GSI, mi, d_val)
            phi_temp, c_temp = calculate_mohr_coulomb_fit(UCS, mb_temp, s_temp, a_temp, sigma3n)
            phi_values.append(phi_temp)
            c_values.append(c_temp)
        
        axes_sens[0].plot(d_range, phi_values, 'b-o', linewidth=2)
        axes_sens[0].axvline(x=D, color='red', linestyle='--', label=f'Current D = {D}')
        axes_sens[0].set_xlabel('D')
        axes_sens[0].set_ylabel('φ (degrees)')
        axes_sens[0].set_title('Friction Angle vs Disturbance Factor')
        axes_sens[0].legend()
        axes_sens[0].grid(True, alpha=0.3)
        
        axes_sens[1].plot(d_range, c_values, 'r-o', linewidth=2)
        axes_sens[1].axvline(x=D, color='blue', linestyle='--', label=f'Current D = {D}')
        axes_sens[1].set_xlabel('D')
        axes_sens[1].set_ylabel('c (MPa)')
        axes_sens[1].set_title('Cohesion vs Disturbance Factor')
        axes_sens[1].legend()
        axes_sens[1].grid(True, alpha=0.3)
        
    elif param_choice == "H (Slope Height)":
        h_range = np.arange(5, 105, 5)
        phi_values = []
        c_values = []
        
        for h_val in h_range:
            sigma3max_temp = calculate_sigma3max(h_val, gamma, sigma3_factor)
            sigma3n_temp = sigma3max_temp / UCS
            phi_temp, c_temp = calculate_mohr_coulomb_fit(UCS, mb, s, a, sigma3n_temp)
            phi_values.append(phi_temp)
            c_values.append(c_temp)
        
        axes_sens[0].plot(h_range, phi_values, 'b-o', linewidth=2)
        axes_sens[0].axvline(x=H, color='red', linestyle='--', label=f'Current H = {H} m')
        axes_sens[0].set_xlabel('H (m)')
        axes_sens[0].set_ylabel('φ (degrees)')
        axes_sens[0].set_title('Friction Angle vs Slope Height')
        axes_sens[0].legend()
        axes_sens[0].grid(True, alpha=0.3)
        
        axes_sens[1].plot(h_range, c_values, 'r-o', linewidth=2)
        axes_sens[1].axvline(x=H, color='blue', linestyle='--', label=f'Current H = {H} m')
        axes_sens[1].set_xlabel('H (m)')
        axes_sens[1].set_ylabel('c (MPa)')
        axes_sens[1].set_title('Cohesion vs Slope Height')
        axes_sens[1].legend()
        axes_sens[1].grid(True, alpha=0.3)
        
    else:  # mi
        mi_range = np.arange(4, 35, 2)
        phi_values = []
        c_values = []
        
        for mi_val in mi_range:
            mb_temp, s_temp, a_temp = calculate_hoek_brown_parameters(GSI, mi_val, D)
            phi_temp, c_temp = calculate_mohr_coulomb_fit(UCS, mb_temp, s_temp, a_temp, sigma3n)
            phi_values.append(phi_temp)
            c_values.append(c_temp)
        
        axes_sens[0].plot(mi_range, phi_values, 'b-o', linewidth=2)
        axes_sens[0].axvline(x=mi, color='red', linestyle='--', label=f'Current mi = {mi}')
        axes_sens[0].set_xlabel('mi')
        axes_sens[0].set_ylabel('φ (degrees)')
        axes_sens[0].set_title('Friction Angle vs mi')
        axes_sens[0].legend()
        axes_sens[0].grid(True, alpha=0.3)
        
        axes_sens[1].plot(mi_range, c_values, 'r-o', linewidth=2)
        axes_sens[1].axvline(x=mi, color='blue', linestyle='--', label=f'Current mi = {mi}')
        axes_sens[1].set_xlabel('mi')
        axes_sens[1].set_ylabel('c (MPa)')
        axes_sens[1].set_title('Cohesion vs mi')
        axes_sens[1].legend()
        axes_sens[1].grid(True, alpha=0.3)
    
    plt.tight_layout()
    st.pyplot(fig_sens)

# ===== Summary Table =====
st.markdown('<div class="sub-header">📋 Summary Table</div>', unsafe_allow_html=True)

summary_data = {
    "Parameter": [
        "UCS (σci)", "GSI", "mi", "D",
        "mb", "s", "a",
        "H", "γ", "σ3max", "σ3n",
        "φ (Friction Angle)", "c (Cohesion)"
    ],
    "Value": [
        f"{UCS}", f"{GSI}", f"{mi}", f"{D}",
        f"{mb:.4f}", f"{s:.4f}", f"{a:.3f}",
        f"{H}", f"{gamma}", f"{sigma3max:.4f}", f"{sigma3n:.6f}",
        f"{phi:.2f}", f"{c:.3f}"
    ],
    "Unit": [
        "MPa", "-", "-", "-",
        "-", "-", "-",
        "m", "kN/m³", "MPa", "-",
        "degrees", "MPa"
    ],
    "Description": [
        "Uniaxial Compressive Strength", "Geological Strength Index", 
        "Material constant (intact rock)", "Disturbance factor",
        "Reduced material constant", "Rock mass constant", "Exponent",
        "Slope height", "Unit weight", "Maximum confining stress", "Normalized σ3",
        "Equivalent friction angle", "Equivalent cohesion"
    ]
}

st.dataframe(summary_data, use_container_width=True, hide_index=True)

# ===== Export Data =====
st.markdown('<div class="sub-header">💾 บันทึกข้อมูล (Save Data)</div>', unsafe_allow_html=True)

col_export1, col_export2 = st.columns(2)

with col_export1:
    # สร้างข้อมูลสำหรับ JSON
    export_data = {
        "project_info": {
            "title": "Hoek-Brown to Mohr-Coulomb Conversion",
            "reference": "Hoek et al. (2002)"
        },
        "input_parameters": {
            "UCS": UCS,
            "GSI": GSI,
            "mi": mi,
            "D": D,
            "H": H,
            "gamma": gamma,
            "sigma3_factor": sigma3_factor
        },
        "hoek_brown_parameters": {
            "mb": round(mb, 6),
            "s": round(s, 6),
            "a": round(a, 4)
        },
        "stress_range": {
            "sigma3max_MPa": round(sigma3max, 6),
            "sigma3n": round(sigma3n, 8)
        },
        "mohr_coulomb_parameters": {
            "phi_degrees": round(phi, 2),
            "c_MPa": round(c, 4)
        }
    }
    
    json_str = json.dumps(export_data, indent=2, ensure_ascii=False)
    
    st.download_button(
        label="📥 Download JSON",
        data=json_str,
        file_name="hoek_brown_mohr_coulomb_results.json",
        mime="application/json"
    )

with col_export2:
    # สร้างไฟล์ CSV
    csv_content = """Parameter,Value,Unit,Description
UCS (σci),{},{},Uniaxial Compressive Strength
GSI,{},{},Geological Strength Index
mi,{},{},Material constant (intact rock)
D,{},{},Disturbance factor
mb,{},{},Reduced material constant
s,{},{},Rock mass constant
a,{},{},Exponent
H,{},{},Slope height
γ,{},{},Unit weight
σ3max,{},{},Maximum confining stress
σ3n,{},{},Normalized σ3
φ,{},{},Equivalent friction angle
c,{},{},Equivalent cohesion""".format(
        UCS, "MPa",
        GSI, "-",
        mi, "-",
        D, "-",
        round(mb, 4), "-",
        round(s, 4), "-",
        round(a, 3), "-",
        H, "m",
        gamma, "kN/m³",
        round(sigma3max, 4), "MPa",
        round(sigma3n, 6), "-",
        round(phi, 2), "degrees",
        round(c, 3), "MPa"
    )
    
    st.download_button(
        label="📥 Download CSV",
        data=csv_content,
        file_name="hoek_brown_mohr_coulomb_results.csv",
        mime="text/csv"
    )

# ===== mi Reference Table =====
with st.expander("📚 ตารางค่า mi สำหรับหินประเภทต่างๆ (mi Reference Table)", expanded=False):
    st.markdown("""
    | ประเภทหิน | mi | ตัวอย่าง |
    |-----------|---:|---------|
    | **Carbonate rocks** | | |
    | Crystalline | 12 | Dolomite, Limestone |
    | Clastic | 7 | Chalk, Marl |
    | **Clastic rocks** | | |
    | Coarse | 17 | Conglomerate, Breccia |
    | Medium | 15 | Sandstone |
    | Fine | 7 | Siltstone |
    | Very fine | 4 | Claystone, Shale |
    | **Igneous rocks** | | |
    | Felsic plutonic | 25-32 | Granite, Granodiorite |
    | Mafic plutonic | 25 | Gabbro, Diorite |
    | Volcanic | 13-25 | Basalt, Rhyolite |
    | **Metamorphic rocks** | | |
    | Non-foliated | 19 | Marble, Hornfels |
    | Slightly foliated | 10 | Migmatite |
    | Foliated | 7-9 | Gneiss, Schist |
    | Highly foliated | 4-7 | Phyllite, Slate |
    
    *Reference: Hoek (2007) - Practical Rock Engineering*
    """)

# ===== Footer =====
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666; font-size: 12px;">
    <p>📚 Reference: Hoek, E., Carranza-Torres, C., and Corkum, B. (2002). 
    "Hoek-Brown failure criterion – 2002 Edition"</p>
    <p>🏛️ พัฒนาโดย: ภาควิชาครุศาสตร์โยธา คณะครุศาสตร์อุตสาหกรรม มจพ.</p>
</div>
""", unsafe_allow_html=True)
