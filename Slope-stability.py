import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
from dataclasses import dataclass
from typing import List, Tuple, Optional
import math
from io import BytesIO
from docx import Document

# ============================================================
# Data Classes
# ============================================================

@dataclass
class SoilLayer:
    name: str
    thickness: float
    gamma: float
    gamma_sat: float
    cohesion: float
    phi: float

@dataclass
class SlipCircle:
    x_center: float
    y_center: float
    radius: float

@dataclass
class AnalysisResults:
    method: str
    fs: float
    slices_data: List[dict]
    critical_circle: SlipCircle


# ============================================================
# Geometry Functions
# ============================================================

def get_slope_surface_y(x, geo):
    H = geo["height"]
    m = geo["slope_ratio"]
    crest = geo["crest_width"]
    toe = geo["toe_x"]
    toe_elev = geo["toe_elevation"]

    slope_w = H * m

    if x < toe:
        return toe_elev
    elif x < toe + slope_w:
        return toe_elev + (x - toe) / m
    else:
        return toe_elev + H


# ============================================================
# Slip Circle Generation (Improved)
# ============================================================

def generate_slip_circle_from_points(x1, y1, x2, y2, depth_factor=1.2):
    xm = (x1 + x2) / 2
    ym = (y1 + y2) / 2

    dx = x2 - x1
    dy = y2 - y1
    L = np.sqrt(dx**2 + dy**2)

    if L < 0.1:
        return None

    nx = -dy / L
    ny = dx / L

    depth = L * depth_factor
    xc = xm + nx * depth
    yc = ym + ny * depth

    R = np.sqrt((xc - x1)**2 + (yc - y1)**2)

    return SlipCircle(xc, yc, R)


# ============================================================
# Slice Geometry
# ============================================================

def slice_geometry(circle, geo, n_slices):
    slices = []
    xc, yc, R = circle.x_center, circle.y_center, circle.radius

    x_left = xc - R
    x_right = xc + R
    dx = (x_right - x_left) / n_slices

    for i in range(n_slices):
        x_mid = x_left + (i + 0.5) * dx
        y_surface = get_slope_surface_y(x_mid, geo)

        term = R**2 - (x_mid - xc)**2
        if term <= 0:
            continue

        y_base = yc - np.sqrt(term)

        if y_base >= y_surface:
            continue

        height = y_surface - y_base
        alpha = np.arctan2(x_mid - xc, yc - y_base)

        slices.append({
            "index": i,
            "x_mid": x_mid,
            "height": height,
            "width": dx,
            "alpha": alpha
        })

    return slices


# ============================================================
# Swedish Method with Seismic
# ============================================================

def swedish_method(slices, soil, seismic):
    kh = seismic["kh"]
    kv = seismic["kv"]

    sum_resist = 0
    sum_drive = 0

    for s in slices:
        gamma = soil.gamma
        W = gamma * s["height"] * s["width"]

        alpha = s["alpha"]

        T = W * (np.sin(alpha) + kh * np.cos(alpha))
        N = W * (np.cos(alpha) - kv * np.sin(alpha))

        l = s["width"] / np.cos(alpha)

        resisting = soil.cohesion * l + N * np.tan(np.radians(soil.phi))
        driving = T

        sum_resist += resisting
        sum_drive += driving

    fs = sum_resist / sum_drive if sum_drive != 0 else 999

    return fs


# ============================================================
# Search Critical Circle
# ============================================================

def search_critical_circle(geo, soil, seismic, n_slices):
    H = geo["height"]
    toe = geo["toe_x"]
    slope_w = H * geo["slope_ratio"]
    crest = geo["crest_width"]

    x_vals = np.linspace(toe, toe + slope_w + crest, 20)

    min_fs = 999
    best_circle = None
    best_slices = None

    for i in range(len(x_vals)-5):
        for j in range(i+5, len(x_vals)):
            x1 = x_vals[i]
            x2 = x_vals[j]

            y1 = get_slope_surface_y(x1, geo)
            y2 = get_slope_surface_y(x2, geo)

            circle = generate_slip_circle_from_points(x1, y1, x2, y2)

            if circle is None:
                continue

            slices = slice_geometry(circle, geo, n_slices)
            if len(slices) < 5:
                continue

            fs = swedish_method(slices, soil, seismic)

            if fs < min_fs:
                min_fs = fs
                best_circle = circle
                best_slices = slices

    return AnalysisResults(
        method="Swedish with Seismic",
        fs=min_fs,
        slices_data=best_slices,
        critical_circle=best_circle
    )


# ============================================================
# Plotting
# ============================================================

def plot_slope_and_circle(geo, soil, result):
    fig, ax = plt.subplots(figsize=(10,6))

    H = geo["height"]
    m = geo["slope_ratio"]
    toe = geo["toe_x"]
    crest = geo["crest_width"]
    toe_elev = geo["toe_elevation"]

    slope_w = H * m
    crest_elev = toe_elev + H

    x = [toe-5, toe, toe+slope_w, toe+slope_w+crest]
    y = [toe_elev, toe_elev, crest_elev, crest_elev]

    ax.plot(x, y, 'k-', linewidth=2)

    circle = result.critical_circle
    theta = np.linspace(0, 2*np.pi, 300)
    xc = circle.x_center
    yc = circle.y_center
    R = circle.radius

    xs = xc + R*np.cos(theta)
    ys = yc + R*np.sin(theta)

    x_plot, y_plot = [], []

    for xx, yy in zip(xs, ys):
        if yy <= get_slope_surface_y(xx, geo):
            x_plot.append(xx)
            y_plot.append(yy)

    ax.plot(x_plot, y_plot, 'r-', linewidth=3)

    ax.set_aspect('equal')
    ax.set_title(f"FS = {result.fs:.3f}")
    ax.grid(True)

    return fig


# ============================================================
# Export Word
# ============================================================

def export_word_report(result, geo, soil):
    doc = Document()
    doc.add_heading("Slope Stability Report", level=1)

    doc.add_paragraph(f"Method: {result.method}")
    doc.add_paragraph(f"FS = {result.fs:.3f}")

    c = result.critical_circle
    doc.add_paragraph(f"Center: ({c.x_center:.2f}, {c.y_center:.2f})")
    doc.add_paragraph(f"Radius: {c.radius:.2f} m")

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ============================================================
# Streamlit UI
# ============================================================

def main():
    st.title("Slope Stability with Seismic Analysis")

    st.sidebar.header("Slope Geometry")
    H = st.sidebar.number_input("Height (m)", 2.0, 30.0, 8.0)
    m = st.sidebar.number_input("Slope ratio H:V", 1.0, 4.0, 1.5)
    crest = st.sidebar.number_input("Crest width", 2.0, 20.0, 10.0)

    st.sidebar.header("Soil")
    gamma = st.sidebar.number_input("γ (kN/m³)", 15.0, 22.0, 18.0)
    c = st.sidebar.number_input("c (kPa)", 0.0, 50.0, 20.0)
    phi = st.sidebar.number_input("φ (deg)", 0.0, 40.0, 20.0)

    st.sidebar.header("Seismic")
    kh = st.sidebar.number_input("kh", 0.0, 0.5, 0.0, step=0.01)
    kv = st.sidebar.number_input("kv", -0.5, 0.5, 0.0, step=0.01)

    st.sidebar.header("Slices")
    n_slices = st.sidebar.slider("Number of slices", 5, 50, 20)

    geo = {
        "height": H,
        "slope_ratio": m,
        "crest_width": crest,
        "toe_x": 5,
        "toe_elevation": 0
    }

    soil = SoilLayer("Soil", 5, gamma, gamma, c, phi)
    seismic = {"kh": kh, "kv": kv}

    if st.button("Run Analysis"):
        result = search_critical_circle(geo, soil, seismic, n_slices)

        st.write(f"Factor of Safety = {result.fs:.3f}")

        fig = plot_slope_and_circle(geo, soil, result)
        st.pyplot(fig)

        if st.button("Export Word Report"):
            file = export_word_report(result, geo, soil)
            st.download_button(
                "Download Report",
                file,
                file_name="slope_report.docx"
            )


if __name__ == "__main__":
    main()
