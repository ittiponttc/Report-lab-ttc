def create_excel_report():
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter
    import io

    wb = Workbook()
    ws = wb.active
    ws.title = "UCS Test Report"

    # =========================
    # Styles
    # =========================
    header_font = Font(bold=True, size=12)
    title_font = Font(bold=True, size=14)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    header_fill = PatternFill(
        start_color="CCE5FF",
        end_color="CCE5FF",
        fill_type="solid"
    )

    # =========================
    # Title
    # =========================
    ws.merge_cells('A1:H1')
    ws['A1'] = "KING MONGKUT'S UNIVERSITY OF TECHNOLOGY NORTH BANGKOK"
    ws['A1'].font = title_font
    ws['A1'].alignment = Alignment(horizontal='center')

    ws.merge_cells('A2:H2')
    ws['A2'] = "Unconfined Compression Test (ASTM D2166)"
    ws['A2'].font = header_font
    ws['A2'].alignment = Alignment(horizontal='center')

    # =========================
    # Project Information
    # =========================
    row = 4
    info_data = [
        ("Project Name:", project_name),
        ("Specimen From:", specimen_from),
        ("Location:", location),
        ("Column No.:", column_no),
        ("Depth:", depth),
        ("Sample Number:", sample_number),
        ("Date of Jetting:", str(date_of_jetting)),
        ("Date of Testing:", str(date_of_testing)),
        ("Curing Time (days):", curing_time),
        ("Tested by:", tested_by),
    ]

    for label, value in info_data:
        ws[f'A{row}'] = label
        ws[f'A{row}'].font = Font(bold=True)
        ws[f'B{row}'] = value
        row += 1

    # =========================
    # Sample Properties
    # =========================
    row += 1
    ws[f'A{row}'] = "Sample Properties"
    ws[f'A{row}'].font = title_font
    row += 1

    sample_data = [
        ("Diameter, D (cm)", diameter),
        ("Height, H (cm)", height),
        ("Area, A (cm²)", area),
        ("Volume, V (cm³)", volume),
        ("Weight of Sample, W (g)", weight),
        ("Wet Unit Weight (g/cm³)", wet_unit_weight),
        ("Dry Unit Weight (g/cm³)", dry_unit_weight),
        ("Water Content, w (%)", water_content),
    ]

    for label, value in sample_data:
        ws[f'A{row}'] = label
        ws[f'B{row}'] = f"{value:.2f}" if isinstance(value, float) else value
        row += 1

    # =========================
    # Test Results
    # =========================
    row += 1
    ws[f'A{row}'] = "Test Results"
    ws[f'A{row}'].font = title_font
    row += 1

    results_data = [
        ("Unconfined Compressive Strength, qᵤ (ksc)", ucs),
        ("Undrained Shear Strength, sᵤ (ksc)", undrained_shear_strength),
        ("Failure Strain, εf (%)", failure_strain),
        ("Modulus of Elasticity, E₅₀ (ksc)", modulus_e50),
    ]

    for label, value in results_data:
        ws[f'A{row}'] = label
        ws[f'B{row}'] = f"{value:.2f}"
        row += 1

    # =========================
    # Test Data Sheet
    # =========================
    ws2 = wb.create_sheet("Test Data")

    headers = ['R (mm)', 'P (kg)', 'ξ', 'Ac (cm²)', 'ε (%)', 'σ (ksc)']
    for col, header in enumerate(headers, 1):
        cell = ws2.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(horizontal='center')

    for r_idx, row_data in enumerate(
        df[['Deformation, R (mm)',
            'Axial Load, P (kg)',
            'Axial Strain, ξ (R/10H)',
            'Corrected Area, Ac (cm²)',
            'Axial Strain (%)',
            'Axial Stress, σ (ksc)']].values,
        start=2
    ):
        for c_idx, value in enumerate(row_data, start=1):
            cell = ws2.cell(row=r_idx, column=c_idx, value=round(value, 2))
            cell.border = border
            cell.alignment = Alignment(horizontal='center')

    # =========================
    # ✅ Adjust column widths (FIXED)
    # =========================
    for ws_sheet in [ws, ws2]:
        for column_cells in ws_sheet.columns:
            length = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
            col_letter = get_column_letter(column_cells[0].column)
            ws_sheet.column_dimensions[col_letter].width = max(length + 2, 12)

    # =========================
    # Save to buffer
    # =========================
    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    return excel_buffer
