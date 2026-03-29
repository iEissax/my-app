import streamlit as st
import zipfile
import pandas as pd
from lxml import etree
import re
import io

# دالة الترتيب الطبيعي لضمان تسلسل (1, 2, 3 ... 10)
def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', str(s))]

st.set_page_config(page_title="نظام معالجة المحطات", layout="wide")
st.title("📂 مستخرج البيانات الاحترافي (إصدار المطابقة الكاملة)")

uploaded_files = st.file_uploader("اختر ملفات KMZ", type=['kmz'], accept_multiple_files=True)

def extract_data_from_text(full_text, station_code):
    """دالة ذكية لتفكيك النصوص مثل 12/2/1 أو 8/1/1"""
    # تنظيف النص من اسم المحطة
    clean_text = full_text.replace(station_code, "").strip()
    
    # البحث عن نمط الأرقام المفصولة بمائلات (طول/ذراع/عمود) أو (طول/ذراع)
    parts = re.findall(r'\d+', clean_text)
    
    height, arms, col_num = "", "", ""
    
    if len(parts) >= 3:
        height = parts[0]   # الرقم الأول: الطول
        arms = parts[1]     # الرقم الثاني: الذراع
        col_num = parts[2]  # الرقم الثالث: رقم العمود
    elif len(parts) == 2:
        height = parts[0]
        arms = parts[1]
    elif len(parts) == 1:
        col_num = parts[0]

    # إذا كان النص يحتوي على كلمات خاصة
    if "هاي" in clean_text or "mast" in clean_text.lower():
        height, arms = "هاي ماست", 6
    elif "جداري" in clean_text:
        height, arms = "جداري", 1
        
    return height, arms, col_num

def process_kmz(file):
    with zipfile.ZipFile(file, 'r') as f:
        kml_filename = [name for name in f.namelist() if name.endswith('.kml')][0]
        kml_content = f.read(kml_filename)

    tree = etree.fromstring(kml_content)
    ns = {"kml": "http://www.opengis.net/kml/2.2"}
    data = []

    for pm in tree.xpath("//kml:Placemark", namespaces=ns):
        name_text = pm.xpath("./kml:name/text()", namespaces=ns)
        full_name = name_text[0].strip() if name_text else ""
        
        # استخراج المحطة (ج557)
        st_match = re.search(r'(\d+[\u0600-\u06FF]+|[\u0600-\u06FF]+\d+)', full_name)
        station_code = st_match.group(1) if st_match else "غير محدد"

        # جلب الوصف والبيانات
        desc = "".join(pm.xpath("./kml:description/text()", namespaces=ns))
        ext_data = " ".join(pm.xpath(".//kml:value/text()", namespaces=ns))
        all_info = (full_name + " " + desc + " " + ext_data)

        # استخراج البيانات بالمنطق الجديد
        h, a, c = extract_data_from_text(all_info, station_code)
        
        # الإحداثيات خماسية
        coords = pm.xpath(".//kml:coordinates/text()", namespaces=ns)
        lat, lon = 0.0, 0.0
        if coords:
            c_split = coords[0].strip().split(',')
            lat, lon = round(float(c_split[1]), 5), round(float(c_split[0]), 5)

        # الحالة
        detail = "مفقود" if "مفقود" in all_info else ("مغروز" if "مغروز" in all_info else "")

        data.append({
            "المحطة": station_code,
            "رقم العمود": c,
            "طول العمود": h,
            "الذراع": a,
            "الاحداثيات x": lon,
            "الاحداثيات y": lat,
            "التفاصيل": detail
        })
    return pd.DataFrame(data)

if uploaded_files:
    all_data = [process_kmz(f) for f in uploaded_files]
    df = pd.concat(all_data, ignore_index=True)

    # تحويل رقم العمود لنوع عددي للفرز الطبيعي
    df['رقم العمود'] = pd.to_numeric(df['رقم العمود'], errors='coerce').fillna(0).astype(int)
    
    # الترتيب: المحطة أولاً ثم رقم العمود تسلسلياً
    df = df.sort_values(by=['المحطة', 'رقم العمود'], key=lambda x: x.map(natural_sort_key) if x.name == 'المحطة' else x)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('Report')
        worksheet.right_to_left()

        # التنسيقات (خط أسود، حدود كاملة)
        fmt_head = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center', 'font_color': 'black'})
        fmt_stat = workbook.add_format({'bg_color': '#7F7F7F', 'border': 1, 'align': 'center', 'font_color': 'black', 'bold': True})
        fmt_red = workbook.add_format({'bg_color': '#FF0000', 'border': 1, 'align': 'center', 'font_color': 'black'})
        fmt_norm = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'black'})
        fmt_coord = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'black', 'num_format': '0.00000'})

        cols = ["المحطة", "رقم العمود", "طول العمود", "الذراع", "الاحداثيات x", "الاحداثيات y", "التفاصيل"]
        for i, col in enumerate(cols):
            worksheet.write(0, i, col, fmt_head)
            worksheet.set_column(i, i, 15)

        curr_row, last_st = 1, None
        for _, row in df.iterrows():
            if last_st and row['المحطة'] != last_st:
                curr_row += 1 # ترك صف فارغ بين المحطات

            is_red = row['التفاصيل'] in ["مفقود", "مغروز"]
            for j, c_name in enumerate(cols):
                val = row[c_name]
                f = fmt_red if is_red else (fmt_stat if c_name == "المحطة" else (fmt_coord if "الاحداثيات" in c_name else fmt_norm))
                worksheet.write(curr_row, j, val, f)
            
            last_st, curr_row = row['المحطة'], curr_row + 1

    st.success("✅ تم استخراج البيانات وتوزيع (الطول/الذراع/العمود) بدقة.")
    st.download_button("📥 تحميل ملف الإكسيل المعتمد", output.getvalue(), "Lighting_Final_Report.xlsx")
