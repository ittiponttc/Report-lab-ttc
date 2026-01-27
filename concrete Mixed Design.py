import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import json
from io import BytesIO

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
    """
    ACI 211.1 Concrete Mix Design Method
    Returns: dict with Water, Cement, Fine Aggregate, Coarse Aggregate (kg/m³)
    """
    # ---- Water content & coarse aggregate volume (ACI typical) ----
    if max_agg_mm == 20:
        water = 185
        vol_coarse = 0.62
    elif max_agg_mm == 25:
        water = 175
        vol_coarse = 0.64
    else:  # 40 mm
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
        "Coarse Aggregate": weight_coarse,
        # เก็บค่ากลางสำหรับรายงาน
        "vol_water": vol_water,
        "vol_cement": vol_cement,
        "vol_coarse": vol_coarse_abs,
        "vol_fine": vol_fine,
        "vol_air": air_content
    }


def moisture_correction(weight_ssd, mc, absorption):
    """
    Moisture correction for aggregates
    weight_ssd : SSD weight (kg/m³)
    mc         : moisture content (%)
    absorption : absorption (%)
    Returns: (delta_water, batch_weight)
    """
    delta_water = weight_ssd * (mc - absorption) / 100
    batch_weight = weight_ssd * (1 + mc / 100)
    return delta_water, batch_weight


def create_word_report(input_data, mix_result, moisture_result):
    """
    สร้างรายงาน Word แบบละเอียดเป็นขั้นเป็นตอน
    ใช้ Node.js + docx-js
    """
    import subprocess
    import os
    
    # สร้างไฟล์ JS สำหรับสร้าง Word
    js_code = f"""
const {{ Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        HeadingLevel, AlignmentType, WidthType, BorderStyle, ShadingType }} = require('docx');
const fs = require('fs');

// ข้อมูลจาก Python
const input = {json.dumps(input_data, ensure_ascii=False)};
const mix = {json.dumps(mix_result, ensure_ascii=False)};
const moisture = {json.dumps(moisture_result, ensure_ascii=False)};

// Border สำหรับตาราง
const border = {{ style: BorderStyle.SINGLE, size: 1, color: "000000" }};
const borders = {{ top: border, bottom: border, left: border, right: border }};

const doc = new Document({{
  styles: {{
    default: {{
      document: {{
        run: {{ font: "TH SarabunPSK", size: 30 }}  // 15pt
      }}
    }},
    paragraphStyles: [
      {{
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: {{ size: 36, bold: true, font: "TH SarabunPSK" }},
        paragraph: {{ 
          spacing: {{ before: 240, after: 120 }},
          alignment: AlignmentType.CENTER,
          outlineLevel: 0
        }}
      }},
      {{
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: {{ size: 32, bold: true, font: "TH SarabunPSK" }},
        paragraph: {{ 
          spacing: {{ before: 180, after: 120 }},
          outlineLevel: 1
        }}
      }}
    ]
  }},
  
  sections: [{{
    properties: {{
      page: {{
        size: {{ width: 11906, height: 16838 }},  // A4
        margin: {{ top: 1440, right: 1440, bottom: 1440, left: 1440 }}
      }}
    }},
    
    children: [
      // หัวเรื่อง
      new Paragraph({{
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun("รายงานการออกแบบส่วนผสมคอนกรีต")]
      }}),
      
      new Paragraph({{
        alignment: AlignmentType.CENTER,
        spacing: {{ after: 240 }},
        children: [new TextRun({{
          text: "ตามมาตรฐาน ACI 211.1",
          size: 28,
          italics: true
        }})]
      }}),
      
      // ส่วนที่ 1: ข้อมูลนำเข้า
      new Paragraph({{
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("1. ข้อมูลนำเข้าในการออกแบบ")]
      }}),
      
      new Table({{
        width: {{ size: 100, type: WidthType.PERCENTAGE }},
        columnWidths: [4680, 4680],
        rows: [
          new TableRow({{
            children: [
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                shading: {{ fill: "D9D9D9", type: ShadingType.CLEAR }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{
                  children: [new TextRun({{ text: "รายการ", bold: true }})]
                }})]
              }}),
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                shading: {{ fill: "D9D9D9", type: ShadingType.CLEAR }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{
                  children: [new TextRun({{ text: "ค่าที่ใช้", bold: true }})]
                }})]
              }})]
          }}),
          
          new TableRow({{
            children: [
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun("อัตราส่วน น้ำ/ปูนซีเมนต์ (w/c)")] }})]
              }}),
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun(input.wc_ratio.toString())] }})]
              }})]
          }}),
          
          new TableRow({{
            children: [
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun("ขนาดมวลรวมหยาบสูงสุด (mm)")] }})]
              }}),
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun(input.max_agg_mm.toString())] }})]
              }})]
          }}),
          
          new TableRow({{
            children: [
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun("ค่าความถ่วงจำเพาะของปูนซีเมนต์")] }})]
              }}),
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun(input.sg_cement.toFixed(2))] }})]
              }})]
          }}),
          
          new TableRow({{
            children: [
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun("ค่าความถ่วงจำเพาะของมวลรวมละเอียด")] }})]
              }}),
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun(input.sg_fine.toFixed(2))] }})]
              }})]
          }}),
          
          new TableRow({{
            children: [
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun("ค่าความถ่วงจำเพาะของมวลรวมหยาบ")] }})]
              }}),
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun(input.sg_coarse.toFixed(2))] }})]
              }})]
          }}),
          
          new TableRow({{
            children: [
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun("ปริมาณอากาศ (%)")] }})]
              }}),
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun((input.air_content * 100).toFixed(1))] }})]
              }})]
          }}),
          
          new TableRow({{
            children: [
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun("น้ำหนักหน่วยของมวลรวมหยาบ (kg/m³)")] }})]
              }}),
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun(input.unit_weight_coarse.toFixed(0))] }})]
              }})]
          }})]
      }}),
      
      // ส่วนที่ 2: ขั้นตอนการคำนวณ
      new Paragraph({{
        heading: HeadingLevel.HEADING_2,
        spacing: {{ before: 360, after: 120 }},
        children: [new TextRun("2. ขั้นตอนการคำนวณตามวิธี ACI 211.1")]
      }}),
      
      new Paragraph({{
        spacing: {{ after: 120 }},
        children: [new TextRun({{
          text: "ขั้นตอนที่ 1: กำหนดปริมาณน้ำและปริมาณมวลรวมหยาบ",
          bold: true
        }})]
      }}),
      
      new Paragraph({{
        children: [new TextRun(
          `จากตาราง ACI สำหรับขนาดมวลรวมหยาบสูงสุด ${{input.max_agg_mm}} mm:\\n` +
          `  - ปริมาณน้ำ = ${{mix.Water.toFixed(1)}} kg/m³\\n` +
          `  - สัดส่วนปริมาตรมวลรวมหยาบ = ${{input.max_agg_mm === 20 ? '0.62' : input.max_agg_mm === 25 ? '0.64' : '0.68'}}`
        )]
      }}),
      
      new Paragraph({{
        spacing: {{ before: 120, after: 120 }},
        children: [new TextRun({{
          text: "ขั้นตอนที่ 2: คำนวณปริมาณปูนซีเมนต์",
          bold: true
        }})]
      }}),
      
      new Paragraph({{
        children: [new TextRun(
          `ปริมาณปูนซีเมนต์ = น้ำ / (w/c) = ${{mix.Water.toFixed(1)}} / ${{input.wc_ratio}} = ${{mix.Cement.toFixed(1)}} kg/m³`
        )]
      }}),
      
      new Paragraph({{
        spacing: {{ before: 120, after: 120 }},
        children: [new TextRun({{
          text: "ขั้นตอนที่ 3: คำนวณน้ำหนักมวลรวมหยาบ",
          bold: true
        }})]
      }}),
      
      new Paragraph({{
        children: [new TextRun(
          `น้ำหนักมวลรวมหยาบ = สัดส่วนปริมาตร × น้ำหนักหน่วย\\n` +
          `  = ${{input.max_agg_mm === 20 ? '0.62' : input.max_agg_mm === 25 ? '0.64' : '0.68'}} × ${{input.unit_weight_coarse}} = ${{mix["Coarse Aggregate"].toFixed(1)}} kg/m³`
        )]
      }}),
      
      new Paragraph({{
        spacing: {{ before: 120, after: 120 }},
        children: [new TextRun({{
          text: "ขั้นตอนที่ 4: คำนวณปริมาตรสัมบูรณ์ของแต่ละวัสดุ",
          bold: true
        }})]
      }}),
      
      new Paragraph({{
        children: [new TextRun(
          `ปริมาตรน้ำ = ${{mix.Water.toFixed(1)}} / 1000 = ${{mix.vol_water.toFixed(4)}} m³\\n` +
          `ปริมาตรปูนซีเมนต์ = ${{mix.Cement.toFixed(1)}} / (${{input.sg_cement}} × 1000) = ${{mix.vol_cement.toFixed(4)}} m³\\n` +
          `ปริมาตรมวลรวมหยาบ = ${{mix["Coarse Aggregate"].toFixed(1)}} / (${{input.sg_coarse}} × 1000) = ${{mix.vol_coarse.toFixed(4)}} m³\\n` +
          `ปริมาตรอากาศ = ${{(input.air_content * 100).toFixed(1)}}% = ${{mix.vol_air.toFixed(4)}} m³`
        )]
      }}),
      
      new Paragraph({{
        spacing: {{ before: 120, after: 120 }},
        children: [new TextRun({{
          text: "ขั้นตอนที่ 5: คำนวณปริมาตรมวลรวมละเอียด",
          bold: true
        }})]
      }}),
      
      new Paragraph({{
        children: [new TextRun(
          `ปริมาตรมวลรวมละเอียด = 1 - (น้ำ + ปูนซีเมนต์ + มวลรวมหยาบ + อากาศ)\\n` +
          `  = 1 - (${{mix.vol_water.toFixed(4)}} + ${{mix.vol_cement.toFixed(4)}} + ${{mix.vol_coarse.toFixed(4)}} + ${{mix.vol_air.toFixed(4)}})\\n` +
          `  = ${{mix.vol_fine.toFixed(4)}} m³`
        )]
      }}),
      
      new Paragraph({{
        spacing: {{ before: 120, after: 120 }},
        children: [new TextRun({{
          text: "ขั้นตอนที่ 6: คำนวณน้ำหนักมวลรวมละเอียด",
          bold: true
        }})]
      }}),
      
      new Paragraph({{
        children: [new TextRun(
          `น้ำหนักมวลรวมละเอียด = ปริมาตร × ความถ่วงจำเพาะ × 1000\\n` +
          `  = ${{mix.vol_fine.toFixed(4)}} × ${{input.sg_fine}} × 1000\\n` +
          `  = ${{mix["Fine Aggregate"].toFixed(1)}} kg/m³`
        )]
      }}),
      
      // ส่วนที่ 3: ผลลัพธ์ (SSD)
      new Paragraph({{
        heading: HeadingLevel.HEADING_2,
        spacing: {{ before: 360, after: 120 }},
        children: [new TextRun("3. ผลการออกแบบส่วนผสมคอนกรีต (สภาพ SSD)")]
      }}),
      
      new Table({{
        width: {{ size: 100, type: WidthType.PERCENTAGE }},
        columnWidths: [4680, 4680],
        rows: [
          new TableRow({{
            children: [
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                shading: {{ fill: "D9D9D9", type: ShadingType.CLEAR }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{
                  children: [new TextRun({{ text: "วัสดุ", bold: true }})]
                }})]
              }}),
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                shading: {{ fill: "D9D9D9", type: ShadingType.CLEAR }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{
                  children: [new TextRun({{ text: "ปริมาณ (kg/m³)", bold: true }})]
                }})]
              }})]
          }}),
          
          new TableRow({{
            children: [
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun("น้ำ")] }})]
              }}),
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun(mix.Water.toFixed(1))] }})]
              }})]
          }}),
          
          new TableRow({{
            children: [
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun("ปูนซีเมนต์")] }})]
              }}),
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun(mix.Cement.toFixed(1))] }})]
              }})]
          }}),
          
          new TableRow({{
            children: [
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun("มวลรวมละเอียด")] }})]
              }}),
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun(mix["Fine Aggregate"].toFixed(1))] }})]
              }})]
          }}),
          
          new TableRow({{
            children: [
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun("มวลรวมหยาบ")] }})]
              }}),
              new TableCell({{
                borders,
                width: {{ size: 4680, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun(mix["Coarse Aggregate"].toFixed(1))] }})]
              }})]
          }})]
      }}),
      
      // ส่วนที่ 4: การปรับแก้ความชื้น
      new Paragraph({{
        heading: HeadingLevel.HEADING_2,
        spacing: {{ before: 360, after: 120 }},
        children: [new TextRun("4. การปรับแก้เนื่องจากความชื้นในมวลรวม")]
      }}),
      
      new Paragraph({{
        spacing: {{ after: 120 }},
        children: [new TextRun({{
          text: "4.1 การคำนวณสำหรับมวลรวมละเอียด",
          bold: true
        }})]
      }}),
      
      new Paragraph({{
        children: [new TextRun(
          `ความชื้น (MC) = ${{input.mc_fine.toFixed(1)}}%\\n` +
          `การดูดซับน้ำ (Absorption) = ${{input.abs_fine.toFixed(1)}}%\\n` +
          `การเปลี่ยนแปลงน้ำหนักน้ำ = น้ำหนัก SSD × (MC - Absorption) / 100\\n` +
          `  = ${{mix["Fine Aggregate"].toFixed(1)}} × (${{input.mc_fine.toFixed(1)}} - ${{input.abs_fine.toFixed(1)}}) / 100\\n` +
          `  = ${{moisture.dw_fine.toFixed(1)}} kg/m³\\n\\n` +
          `น้ำหนักมวลรวมละเอียดสำหรับผสม = น้ำหนัก SSD × (1 + MC/100)\\n` +
          `  = ${{mix["Fine Aggregate"].toFixed(1)}} × (1 + ${{input.mc_fine.toFixed(1)}}/100)\\n` +
          `  = ${{moisture.batch_fine.toFixed(1)}} kg/m³`
        )]
      }}),
      
      new Paragraph({{
        spacing: {{ before: 120, after: 120 }},
        children: [new TextRun({{
          text: "4.2 การคำนวณสำหรับมวลรวมหยาบ",
          bold: true
        }})]
      }}),
      
      new Paragraph({{
        children: [new TextRun(
          `ความชื้น (MC) = ${{input.mc_coarse.toFixed(1)}}%\\n` +
          `การดูดซับน้ำ (Absorption) = ${{input.abs_coarse.toFixed(1)}}%\\n` +
          `การเปลี่ยนแปลงน้ำหนักน้ำ = น้ำหนัก SSD × (MC - Absorption) / 100\\n` +
          `  = ${{mix["Coarse Aggregate"].toFixed(1)}} × (${{input.mc_coarse.toFixed(1)}} - ${{input.abs_coarse.toFixed(1)}}) / 100\\n` +
          `  = ${{moisture.dw_coarse.toFixed(1)}} kg/m³\\n\\n` +
          `น้ำหนักมวลรวมหยาบสำหรับผสม = น้ำหนัก SSD × (1 + MC/100)\\n` +
          `  = ${{mix["Coarse Aggregate"].toFixed(1)}} × (1 + ${{input.mc_coarse.toFixed(1)}}/100)\\n` +
          `  = ${{moisture.batch_coarse.toFixed(1)}} kg/m³`
        )]
      }}),
      
      new Paragraph({{
        spacing: {{ before: 120, after: 120 }},
        children: [new TextRun({{
          text: "4.3 การปรับแก้ปริมาณน้ำผสม",
          bold: true
        }})]
      }}),
      
      new Paragraph({{
        children: [new TextRun(
          `น้ำที่มาจากมวลรวมทั้งหมด = ${{moisture.dw_fine.toFixed(1)}} + ${{moisture.dw_coarse.toFixed(1)}} = ${{moisture.total_delta_water.toFixed(1)}} kg/m³\\n` +
          `ปริมาณน้ำผสมที่ต้องเติม = ${{mix.Water.toFixed(1)}} - ${{moisture.total_delta_water.toFixed(1)}} = ${{moisture.corrected_water.toFixed(1)}} kg/m³`
        )]
      }}),
      
      // ส่วนที่ 5: สรุปส่วนผสมสำหรับใช้งาน
      new Paragraph({{
        heading: HeadingLevel.HEADING_2,
        spacing: {{ before: 360, after: 120 }},
        children: [new TextRun("5. สรุปส่วนผสมคอนกรีตสำหรับการผสม")]
      }}),
      
      new Table({{
        width: {{ size: 100, type: WidthType.PERCENTAGE }},
        columnWidths: [3120, 3120, 3120],
        rows: [
          new TableRow({{
            children: [
              new TableCell({{
                borders,
                width: {{ size: 3120, type: WidthType.DXA }},
                shading: {{ fill: "D9D9D9", type: ShadingType.CLEAR }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{
                  children: [new TextRun({{ text: "วัสดุ", bold: true }})]
                }})]
              }}),
              new TableCell({{
                borders,
                width: {{ size: 3120, type: WidthType.DXA }},
                shading: {{ fill: "D9D9D9", type: ShadingType.CLEAR }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{
                  children: [new TextRun({{ text: "SSD (kg/m³)", bold: true }})]
                }})]
              }}),
              new TableCell({{
                borders,
                width: {{ size: 3120, type: WidthType.DXA }},
                shading: {{ fill: "D9D9D9", type: ShadingType.CLEAR }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{
                  children: [new TextRun({{ text: "สำหรับผสม (kg/m³)", bold: true }})]
                }})]
              }})]
          }}),
          
          new TableRow({{
            children: [
              new TableCell({{
                borders,
                width: {{ size: 3120, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun("น้ำผสม")] }})]
              }}),
              new TableCell({{
                borders,
                width: {{ size: 3120, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun(mix.Water.toFixed(1))] }})]
              }}),
              new TableCell({{
                borders,
                width: {{ size: 3120, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                shading: {{ fill: "FFFF99", type: ShadingType.CLEAR }},
                children: [new Paragraph({{ children: [new TextRun({{ text: moisture.corrected_water.toFixed(1), bold: true }})] }})]
              }})]
          }}),
          
          new TableRow({{
            children: [
              new TableCell({{
                borders,
                width: {{ size: 3120, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun("ปูนซีเมนต์")] }})]
              }}),
              new TableCell({{
                borders,
                width: {{ size: 3120, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun(mix.Cement.toFixed(1))] }})]
              }}),
              new TableCell({{
                borders,
                width: {{ size: 3120, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                shading: {{ fill: "FFFF99", type: ShadingType.CLEAR }},
                children: [new Paragraph({{ children: [new TextRun({{ text: mix.Cement.toFixed(1), bold: true }})] }})]
              }})]
          }}),
          
          new TableRow({{
            children: [
              new TableCell({{
                borders,
                width: {{ size: 3120, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun("มวลรวมละเอียด")] }})]
              }}),
              new TableCell({{
                borders,
                width: {{ size: 3120, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun(mix["Fine Aggregate"].toFixed(1))] }})]
              }}),
              new TableCell({{
                borders,
                width: {{ size: 3120, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                shading: {{ fill: "FFFF99", type: ShadingType.CLEAR }},
                children: [new Paragraph({{ children: [new TextRun({{ text: moisture.batch_fine.toFixed(1), bold: true }})] }})]
              }})]
          }}),
          
          new TableRow({{
            children: [
              new TableCell({{
                borders,
                width: {{ size: 3120, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun("มวลรวมหยาบ")] }})]
              }}),
              new TableCell({{
                borders,
                width: {{ size: 3120, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                children: [new Paragraph({{ children: [new TextRun(mix["Coarse Aggregate"].toFixed(1))] }})]
              }}),
              new TableCell({{
                borders,
                width: {{ size: 3120, type: WidthType.DXA }},
                margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
                shading: {{ fill: "FFFF99", type: ShadingType.CLEAR }},
                children: [new Paragraph({{ children: [new TextRun({{ text: moisture.batch_coarse.toFixed(1), bold: true }})] }})]
              }})]
          }})]
      }}),
      
      new Paragraph({{
        spacing: {{ before: 240 }},
        children: [new TextRun({{
          text: "หมายเหตุ: ค่าที่ไฮไลต์เป็นสีเหลืองคือส่วนผสมที่ต้องใช้ในการผสมคอนกรีตจริง",
          italics: true
        }})]
      }})
    ]
  }}]
}});

Packer.toBuffer(doc).then(buffer => {{
  fs.writeFileSync('/home/claude/concrete_mix_report.docx', buffer);
  console.log('Word document created successfully!');
}});
"""
    
    # เขียนไฟล์ JS
    with open('/home/claude/create_report.js', 'w', encoding='utf-8') as f:
        f.write(js_code)
    
    # ติดตั้ง docx ถ้ายังไม่มี
    subprocess.run(['npm', 'install', '-g', 'docx'], 
                   capture_output=True, cwd='/home/claude')
    
    # รัน Node.js
    result = subprocess.run(['node', '/home/claude/create_report.js'],
                          capture_output=True, text=True, cwd='/home/claude')
    
    if result.returncode == 0:
        return '/home/claude/concrete_mix_report.docx'
    else:
        raise Exception(f"Error creating Word report: {result.stderr}")


# =========================================================
# Streamlit UI
# =========================================================
st.set_page_config(page_title="Concrete Mix Design - ACI 211", layout="centered")

st.title("🏗️ โปรแกรมออกแบบส่วนผสมคอนกรีต")
st.caption("ตามมาตรฐาน ACI 211.1 | สำหรับการสอนและงานปฏิบัติ")

# ---------------- Upload JSON ----------------
st.sidebar.header("📂 โหลดข้อมูลจากไฟล์")

uploaded_json = st.sidebar.file_uploader("อัพโหลดไฟล์ JSON", type=['json'])

if uploaded_json is not None:
    try:
        loaded_data = json.load(uploaded_json)
        
        # ตรวจสอบว่าเป็นไฟล์ใหม่
        file_id = f"{uploaded_json.name}_{uploaded_json.size}"
        if st.session_state.get('last_uploaded_file') != file_id:
            st.session_state['last_uploaded_file'] = file_id
            
            # อัพเดท session_state
            st.session_state['input_wc_ratio'] = loaded_data.get('wc_ratio', 0.50)
            st.session_state['input_max_agg_mm'] = loaded_data.get('max_agg_mm', 25)
            st.session_state['input_sg_cement'] = loaded_data.get('sg_cement', 3.15)
            st.session_state['input_sg_fine'] = loaded_data.get('sg_fine', 2.65)
            st.session_state['input_sg_coarse'] = loaded_data.get('sg_coarse', 2.70)
            st.session_state['input_air_content'] = loaded_data.get('air_content', 2.0)
            st.session_state['input_unit_weight_coarse'] = loaded_data.get('unit_weight_coarse', 1600)
            st.session_state['input_mc_fine'] = loaded_data.get('mc_fine', 5.0)
            st.session_state['input_abs_fine'] = loaded_data.get('abs_fine', 2.0)
            st.session_state['input_mc_coarse'] = loaded_data.get('mc_coarse', 1.0)
            st.session_state['input_abs_coarse'] = loaded_data.get('abs_coarse', 0.5)
            
            st.sidebar.success("✅ โหลดข้อมูลสำเร็จ!")
            st.rerun()
            
    except Exception as e:
        st.sidebar.error(f"❌ ไม่สามารถอ่านไฟล์ได้: {e}")

# ---------------- Sidebar Inputs ----------------
st.sidebar.header("📥 ข้อมูลสำหรับออกแบบส่วนผสม")

wc_ratio = st.sidebar.number_input(
    "อัตราส่วน น้ำ/ปูนซีเมนต์ (w/c)",
    min_value=0.35, max_value=0.70, step=0.01,
    value=st.session_state.get('input_wc_ratio', 0.50),
    key="input_wc_ratio"
)

max_agg_options = [20, 25, 40]
current_max_agg = st.session_state.get('input_max_agg_mm', 25)
max_agg_idx = max_agg_options.index(current_max_agg) if current_max_agg in max_agg_options else 1

max_agg_mm = st.sidebar.selectbox(
    "ขนาดมวลรวมหยาบสูงสุด (mm)",
    options=max_agg_options,
    index=max_agg_idx,
    key="input_max_agg_mm"
)

sg_cement = st.sidebar.number_input(
    "ค่าความถ่วงจำเพาะของปูนซีเมนต์",
    value=st.session_state.get('input_sg_cement', 3.15),
    key="input_sg_cement"
)

sg_fine = st.sidebar.number_input(
    "ค่าความถ่วงจำเพาะของมวลรวมละเอียด",
    value=st.session_state.get('input_sg_fine', 2.65),
    key="input_sg_fine"
)

sg_coarse = st.sidebar.number_input(
    "ค่าความถ่วงจำเพาะของมวลรวมหยาบ",
    value=st.session_state.get('input_sg_coarse', 2.70),
    key="input_sg_coarse"
)

air_content = st.sidebar.slider(
    "ปริมาณอากาศ (%)",
    min_value=1.0, max_value=6.0,
    value=st.session_state.get('input_air_content', 2.0),
    key="input_air_content"
) / 100

unit_weight_coarse = st.sidebar.number_input(
    "น้ำหนักหน่วยของมวลรวมหยาบ (kg/m³)",
    value=st.session_state.get('input_unit_weight_coarse', 1600),
    key="input_unit_weight_coarse"
)

# ---------------- Moisture Input ----------------
st.sidebar.header("💧 ข้อมูลความชื้นของมวลรวม")

mc_fine = st.sidebar.number_input(
    "ความชื้นมวลรวมละเอียด (%)",
    value=st.session_state.get('input_mc_fine', 5.0),
    key="input_mc_fine"
)

abs_fine = st.sidebar.number_input(
    "การดูดซับน้ำมวลรวมละเอียด (%)",
    value=st.session_state.get('input_abs_fine', 2.0),
    key="input_abs_fine"
)

mc_coarse = st.sidebar.number_input(
    "ความชื้นมวลรวมหยาบ (%)",
    value=st.session_state.get('input_mc_coarse', 1.0),
    key="input_mc_coarse"
)

abs_coarse = st.sidebar.number_input(
    "การดูดซับน้ำมวลรวมหยาบ (%)",
    value=st.session_state.get('input_abs_coarse', 0.5),
    key="input_abs_coarse"
)

# ---------------- Download JSON Button ----------------
export_data = {
    'wc_ratio': wc_ratio,
    'max_agg_mm': max_agg_mm,
    'sg_cement': sg_cement,
    'sg_fine': sg_fine,
    'sg_coarse': sg_coarse,
    'air_content': air_content * 100,
    'unit_weight_coarse': unit_weight_coarse,
    'mc_fine': mc_fine,
    'abs_fine': abs_fine,
    'mc_coarse': mc_coarse,
    'abs_coarse': abs_coarse
}

json_str = json.dumps(export_data, ensure_ascii=False, indent=2)

st.sidebar.download_button(
    label="💾 ดาวน์โหลดข้อมูล (JSON)",
    data=json_str,
    file_name="concrete_mix_input.json",
    mime="application/json"
)

# =========================================================
# Calculation
# =========================================================
if st.button("🧮 คำนวณส่วนผสมคอนกรีต", type="primary"):

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
        "วัสดุ": ["น้ำ", "ปูนซีเมนต์", "มวลรวมละเอียด", "มวลรวมหยาบ"],
        "ปริมาณ (kg/m³)": [
            round(mix["Water"], 1),
            round(mix["Cement"], 1),
            round(mix["Fine Aggregate"], 1),
            round(mix["Coarse Aggregate"], 1)
        ]
    })

    st.subheader("📊 ผลการออกแบบส่วนผสม (สภาพ SSD)")
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
        "รายการ": ["มวลรวมละเอียด", "มวลรวมหยาบ"],
        "น้ำหนัก SSD (kg/m³)": [
            round(mix["Fine Aggregate"], 1),
            round(mix["Coarse Aggregate"], 1)
        ],
        "น้ำหนักสำหรับผสม (kg/m³)": [
            round(batch_fine, 1),
            round(batch_coarse, 1)
        ],
        "Δ น้ำ (kg/m³)": [
            round(dw_fine, 1),
            round(dw_coarse, 1)
        ]
    })

    st.subheader("💧 การปรับแก้เนื่องจากความชื้น")
    st.dataframe(df_mc, use_container_width=True)

    st.info(
        f"💧 **ปริมาณน้ำเดิม** = {round(mix['Water'],1)} kg/m³\n\n"
        f"💧 **ปริมาณน้ำผสมที่ปรับแก้** = {round(corrected_water,1)} kg/m³"
    )

    # ---- Summary table ----
    st.subheader("✅ สรุปส่วนผสมคอนกรีตสำหรับการผสม")
    
    df_final = pd.DataFrame({
        "วัสดุ": ["น้ำผสม", "ปูนซีเมนต์", "มวลรวมละเอียด", "มวลรวมหยาบ"],
        "SSD (kg/m³)": [
            round(mix["Water"], 1),
            round(mix["Cement"], 1),
            round(mix["Fine Aggregate"], 1),
            round(mix["Coarse Aggregate"], 1)
        ],
        "สำหรับผสม (kg/m³)": [
            round(corrected_water, 1),
            round(mix["Cement"], 1),
            round(batch_fine, 1),
            round(batch_coarse, 1)
        ]
    })
    
    st.dataframe(df_final, use_container_width=True)

    # ---- Pie Chart ----
    st.subheader("📈 สัดส่วนวัสดุ (สภาพ SSD)")
    fig, ax = plt.subplots(figsize=(8, 6))
    ax.pie(
        df_mix["ปริมาณ (kg/m³)"],
        labels=df_mix["วัสดุ"],
        autopct="%1.1f%%",
        startangle=90
    )
    ax.axis("equal")
    st.pyplot(fig)

    # ---- Create Word Report Button ----
    st.subheader("📄 สร้างรายงาน Word แบบละเอียด")
    
    if st.button("📝 สร้างรายงาน Word"):
        with st.spinner("กำลังสร้างรายงาน..."):
            try:
                # เตรียมข้อมูลสำหรับรายงาน
                input_dict = {
                    'wc_ratio': wc_ratio,
                    'max_agg_mm': max_agg_mm,
                    'sg_cement': sg_cement,
                    'sg_fine': sg_fine,
                    'sg_coarse': sg_coarse,
                    'air_content': air_content,
                    'unit_weight_coarse': unit_weight_coarse,
                    'mc_fine': mc_fine,
                    'abs_fine': abs_fine,
                    'mc_coarse': mc_coarse,
                    'abs_coarse': abs_coarse
                }
                
                moisture_dict = {
                    'dw_fine': dw_fine,
                    'batch_fine': batch_fine,
                    'dw_coarse': dw_coarse,
                    'batch_coarse': batch_coarse,
                    'total_delta_water': total_delta_water,
                    'corrected_water': corrected_water
                }
                
                # สร้างรายงาน
                report_path = create_word_report(input_dict, mix, moisture_dict)
                
                # อ่านไฟล์เพื่อดาวน์โหลด
                with open(report_path, 'rb') as f:
                    report_bytes = f.read()
                
                st.success("✅ สร้างรายงานสำเร็จ!")
                
                st.download_button(
                    label="📥 ดาวน์โหลดรายงาน Word",
                    data=report_bytes,
                    file_name="รายงานออกแบบส่วนผสมคอนกรีต_ACI211.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            except Exception as e:
                st.error(f"❌ เกิดข้อผิดพลาด: {e}")

    st.success("คำนวณเรียบร้อย ✔")

# ---------------- Footer ----------------
st.markdown("---")
st.caption(
    "🎓 เทคโนโลยีคอนกรีต | การออกแบบส่วนผสมตามมาตรฐาน ACI 211.1 | "
    "รองรับการสอน ป.ตรี–โท และงานหน้างานเบื้องต้น"
)
