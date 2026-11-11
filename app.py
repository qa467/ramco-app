import streamlit as st
from datetime import datetime
import openpyxl
import pandas as pd
from io import BytesIO

import arabic_reshaper
from bidi.algorithm import get_display

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, PageBreak, Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

st.set_page_config(page_title="أذونات تحميل رامكو ستيل", page_icon="🚛", layout="centered")

@st.cache_resource
def load_font():
    # استخدام الخط الموجود بجانب الملف
    try:
        pdfmetrics.registerFont(TTFont('ArabicFont', 'DejaVuSans.ttf'))
    except:
        pass

load_font()

def ar(text):
    if text is None or pd.isna(text):
        return ""
    text_str = str(text).strip()
    if not text_str:
        return ""
    reshaped = arabic_reshaper.reshape(text_str)
    return get_display(reshaped)

ARABIC_DAYS = {
    'Monday': 'الاثنين', 'Tuesday': 'الثلاثاء', 'Wednesday': 'الأربعاء',
    'Thursday': 'الخميس', 'Friday': 'الجمعة', 'Saturday': 'السبت', 'Sunday': 'الأحد'
}

DIAMETER_COLORS = {
    10: colors.black, 12: colors.black,
    14: colors.HexColor('#0E6635'), 16: colors.HexColor('#B7950B'),
    18: colors.HexColor('#A93226'), 20: colors.HexColor('#6E2C00'),
    22: colors.HexColor('#1B4F72'), 25: colors.HexColor('#512E5F'),
    28: colors.black, 32: colors.black,
}

def parse_excel_schedule(file_bytes, selected_day, selected_date):
    wb = openpyxl.load_workbook(file_bytes, data_only=False)
    ws = wb['برنامج تحميل']
    trucks_data = []
    current_client = ""
    current_truck = None
    
    for r in range(1, ws.max_row + 1):
        row_dim = ws.row_dimensions.get(r)
        if row_dim and row_dim.hidden:
            continue
            
        c1 = ws.cell(row=r, column=2).value
        if isinstance(c1, str) and '-' in c1 and ('DETAIL' in c1 or 'TON' in c1):
            parts = c1.split('-')
            current_client = f"{parts[0]} {parts[1]}" if len(parts) >= 2 else parts[0]
            current_truck = None
            continue
            
        tn_cell = ws.cell(row=r, column=12)
        tn = tn_cell.value
        if str(tn).strip().upper() in ['T N', 'TN']:
            continue
            
        items_in_row = []
        for ton_col, l_col, d_col in [(9, 10, 11), (6, 7, 8), (3, 4, 5)]:
            cell_ton = ws.cell(row=r, column=ton_col)
            cell_l = ws.cell(row=r, column=l_col)
            cell_d = ws.cell(row=r, column=d_col)
            ton = cell_ton.value
            l = cell_l.value
            d = cell_d.value
            
            if d is not None and str(d).strip() not in ['', 'D', '0', 'None']:
                try:
                    d_int = int(float(d))
                    l_val = round(float(l), 2) if l is not None else ""
                    ton_val = round(float(ton), 2) if ton is not None else ""
                    
                    has_underline = False
                    for c_obj in [cell_d, cell_l, cell_ton]:
                        if c_obj.font and c_obj.font.underline in ['single', 'double', 'singleAccounting', 'doubleAccounting', True]:
                            has_underline = True
                            break
                        if c_obj.font and c_obj.font.u not in [None, 'none']:
                            has_underline = True
                            break
                            
                    items_in_row.append({'d': d_int, 'l': l_val, 'ton': ton_val, 'is_ameed': has_underline})
                except:
                    pass
        
        has_truck_no = (tn is not None and str(tn).strip() not in ['', '0', 'None'])
        if has_truck_no:
            current_truck = {
                'serial': "", 'client': current_client,
                'truck_no': str(tn).strip(), 'day': selected_day,
                'date': selected_date, 'items': items_in_row
            }
            trucks_data.append(current_truck)
        elif current_truck is not None and len(items_in_row) > 0:
            current_truck['items'].extend(items_in_row)
            
    return trucks_data

def build_single_ticket(truck_info):
    elements = []
    day_date_str = f"اليوم: {truck_info['day']}\nالتاريخ: {truck_info['date']}"
    hdr_data = [
        [ar("F-02-SA-01  |  Rev. 1.0"), "", ar(truck_info['serial']), ar("حديد رامكو\nramco steel")],
        [ar(f"اسم العميل: {truck_info['client']}"), ar(f"رقم الشاحنة: {truck_info['truck_no']}"), ar("اسم السائق: ..............."), ar(day_date_str)]
    ]
    t_hdr = Table(hdr_data, colWidths=[125, 125, 120, 175], rowHeights=[24, 26])
    t_hdr.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), 'ArabicFont'), ('FONTSIZE', (0,0), (-1,-1), 8.5),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOX', (0,0), (-1,-1), 1.2, colors.black), ('INNERGRID', (0,0), (-1,-1), 0.6, colors.grey),
        ('RIGHTPADDING', (0,0), (-1,-1), 3), ('LEFTPADDING', (0,0), (-1,-1), 3),
    ]))
    elements.append(t_hdr)
    elements.append(Spacer(1, 8))
    
    table_data = [[ar("ملاحظات"), ar("عدد الربط"), ar("وزن الربط"), ar("الطول"), ar("القطر (سم / ملم)")]]
    row_styles = []
    items = truck_info['items']
    p_style_center = ParagraphStyle('CellCenter', fontName='ArabicFont', fontSize=10, alignment=1, leading=12)
    
    for idx, it in enumerate(items):
        d_val = it['d']
        l_val = f"{it['l']} متر" if str(it['l']).strip() != '' else ""
        ton_val = f"{it['ton']} طن" if str(it['ton']).strip() != '' else ""
        d_raw_text = f"{d_val} مم (عميد)" if it.get('is_ameed', False) else f"{d_val} مم"
        txt_color = DIAMETER_COLORS.get(d_val, colors.black)
        r_idx = idx + 1
        
        if d_val == 12:
            d_cell = Paragraph(f'<font color="white" backcolor="black">&nbsp;<b>{ar(d_raw_text)}</b>&nbsp;</font>', p_style_center)
        else:
            d_cell = ar(d_raw_text)
            
        row_styles.append(('TEXTCOLOR', (0, r_idx), (-1, r_idx), txt_color))
        row_styles.append(('BACKGROUND', (0, r_idx), (-1, r_idx), colors.white))
        table_data.append(["", "", ar(ton_val), ar(l_val), d_cell])
        
    empty_rows_needed = max(0, 10 - len(items))
    for _ in range(empty_rows_needed):
        table_data.append(["", "", "", "", ""])
        
    t_body = Table(table_data, colWidths=[125, 95, 115, 105, 105], rowHeights=[24]*len(table_data))
    base_style = [
        ('FONTNAME', (0,0), (-1,-1), 'ArabicFont'), ('FONTSIZE', (0,0), (-1,-1), 10),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#ECEFF1')),
        ('BOX', (0,0), (-1,-1), 1.2, colors.black), ('INNERGRID', (0,0), (-1,-1), 0.6, colors.black),
    ] + row_styles
    t_body.setStyle(TableStyle(base_style))
    elements.append(t_body)
    elements.append(Spacer(1, 8))
    
    combined_footer = [
        [ar("ملاحظات:"), ar("S500 ملون على خشب")], ["", ar("أرسل شقف اللوابل")],
        ["", ar("فحص مختبر")],
        [ar("توقيع المستلم: ............................"), ar("مسؤول التحميل: ............................")]
    ]
    t_notes = Table(combined_footer, colWidths=[200, 345], rowHeights=[16, 16, 16, 25])
    t_notes.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), 'ArabicFont'), ('FONTSIZE', (0,0), (-1,-1), 9.5),
        ('ALIGN', (0,0), (-1,2), 'RIGHT'), ('ALIGN', (0,3), (-1,3), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('TEXTCOLOR', (0,0), (-1,2), colors.HexColor('#1B2631')),
        ('LINEBELOW', (0,2), (-1,2), 0.5, colors.grey), ('BOX', (0,0), (-1,-1), 0.8, colors.black),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    elements.append(t_notes)
    return elements

def generate_pdf_bytes(trucks):
    pdf_buffer = BytesIO()
    doc = SimpleDocTemplate(pdf_buffer, pagesize=A4, leftMargin=20, rightMargin=20, topMargin=25, bottomMargin=20)
    story = []
    for i, truck in enumerate(trucks):
        story.extend(build_single_ticket(truck))
        story.append(PageBreak())
        story.extend(build_single_ticket(truck))
        if i < len(trucks) - 1:
            story.append(PageBreak())
    doc.build(story)
    pdf_buffer.seek(0)
    return pdf_buffer

# --- الواجهة ---
st.title("🚛 نظام أذونات التحميل الداخلي")
st.markdown("**شركة حديد رامكو - Ramco Steel**")

selected_date_input = st.date_input("📅 حدد تاريخ التحميل:", datetime.now())
day_en = selected_date_input.strftime('%A')
day_ar = ARABIC_DAYS.get(day_en, '')
date_formatted = selected_date_input.strftime("%Y / %m / %d")

st.info(f"اليوم المعتمد: **{day_ar}** | التاريخ: **{date_formatted}**")

uploaded_file = st.file_uploader("📂 ارفع ملف الإكسيل (production sechdule.xlsx):", type=["xlsx", "xls"])

if uploaded_file and st.button("🚀 إنشاء أذونات التحميل (PDF)", type="primary"):
    with st.spinner("جاري معالجة البيانات وبناء ملف الـ PDF..."):
        trucks = parse_excel_schedule(uploaded_file, day_ar, date_formatted)
        if trucks:
            pdf_data = generate_pdf_bytes(trucks)
            st.success(f"✔️ تم استخراج {len(trucks)} شاحنة بنجاح!")
            st.download_button(
                label="📥 اضغط هنا لتنزيل ملف الـ PDF",
                data=pdf_data,
                file_name=f"اذونات_تحميل_{date_formatted.replace(' / ', '_')}.pdf",
                mime="application/pdf"
            )
        else:
            st.error("لم يتم العثور على شاحنات صالحة في الملف.")