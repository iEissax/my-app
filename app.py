import streamlit as st
import zipfile
import pandas as pd
from lxml import etree
import re
import io

# دالة الترتيب الطبيعي لضمان تسلسل (1, 2, 10...)
def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', str(s))]

st.set_page_config(page_title="مستخرج بيانات المحطات المطور", layout="wide")
st.title("📂 مستخرج KMZ - استخراج الطول من الوصف")

uploaded_files = st.file_uploader("اختر ملفات KMZ", type=['kmz'], accept_multiple_files=True)

def process_kmz(file):
    with zipfile.ZipFile(file, 'r') as f:
        kml_filename = [name for name in f.namelist() if name.endswith('.kml')][0]
        kml_content = f.read(kml_filename)

    tree = etree.fromstring(kml_content)
    ns = {"kml": "http://www.opengis.net/kml/2.2"}
    data = []

    for pm in tree.xpath("//kml:Placemark", namespaces=ns):
        # 1. العنوان (Name) لاستخراج المحطة ورقم العمود
        name_text = pm.xpath("./kml:name/text()", namespaces=ns)
        full_name = name_text[0].strip() if name_text else ""
        
        station_match = re.search(r'(\d+[\u0600-\u06FF]+|[\u0600-\u06FF]+\d+)', full_name)
        station_code = station_match.group(1) if station_match else "غير محدد"

        # استخراج رقم العمود من العنوان فقط
        clean_name = full_name.replace(station_code, "").strip()
        name_nums = re.findall(r'\d+', clean_name)
        column_num = name_nums[0] if name_nums else ""

        # 2. الوصف (Description) لاستخراج طول العمود والأذرعة
        desc = "".join(pm.xpath("./kml:description/text()", namespaces=ns))
        ext_vals = " ".join(pm.xpath(".//kml:value/text()", namespaces=ns))
        # منطقة البحث عن المواصفات الفنية (الوصف + البيانات الإضافية)
        tech_info = (desc + " " + ext_vals).strip()

        val_height, val_arms = "", ""
        
        # البحث عن نمط 12/2/1 في الوصف
        pattern_match = re.findall(r'(\d+)[/-](\d+)', tech_info)
        if pattern_match:
            val_height = pattern_match[0][0]  # الرقم الأول هو الطول
            val_arms = pattern_match[0][1]    # الرقم الثاني هو الأذرعة
        else:
            # بحث بديل في الوصف عن أرقام الأطوال القياسية
            h_match = re.search(r'\b(12|10|8|6|5)\b', tech_info)
            if h_match:
                val_height = h_match.group(1)

        # معالجة الكلمات الخاصة
        tech_lower = tech_info.lower()
        if "هاي" in tech_lower or "mast" in tech_lower:
            val_height, val_arms = "هاي ماست", 6
        elif "جداري" in tech_lower:
            val_height, val_arms = "جداري", 1

        # 3. الحالة (مفقود/مغروز) - يبحث في كل النصوص
        all_text = (full_name + " " + tech_info).lower()
        details = "مفقود" if "مفقود" in all_text else ("مغروز" if "مغروز" in all_text else "")

        # 4. الإحداثيات (خماسية)
        coords = pm.xpath(".//kml:coordinates/text()", namespaces=ns)
        lat_val, lon_val = 0.0, 0.0
        if coords:
            c_split = coords[0].strip().split(',')
            lat_val, lon_val = round(float(c_split[1]), 5), round(float(c_split[0]), 5)

        data.append({
            "المحطة": station_code,
            "رقم العمود": column_num,
            "طول العمود": val_height,
            "الذراع": val_arms,
            "الاحداثيات x": lon_val,
            "الاحداثيات y": lat_val,
            "التفاصيل": details
        })
    return pd.DataFrame(data)

if uploaded_files:
    all_dfs = [process_kmz(f) for f in uploaded_files]
    df = pd.concat(all_dfs, ignore_index=True)
    
    # تحويل رقم العمود إلى نوع عددي للترتيب الصحيح
    df['رقم العمود'] = pd.to_numeric(df['رقم العمود'], errors='coerce').fillna(0).astype(int)
    
    # الترتيب حسب المحطة ثم رقم العمود
    df = df.sort_values(
        by=['المحطة', 'رقم العمود'], 
        key=lambda x: x.map(natural_sort_key) if x.name == 'المحطة' else x
    )
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook, worksheet = writer.book, writer.book.add_worksheet('Data')
        worksheet.right_to_left()

        # التنسيقات
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center', 'font_color': 'black'})
        station_fmt = workbook.add_format({'bg_color': '#7F7F7F', 'border': 1, 'align': 'center', 'font_color': 'black', 'bold': True})
        red_fmt = workbook.add_format({'bg_color': '#FF0000', 'border': 1, 'align': 'center', 'font_color': 'black'})
        normal_fmt = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'black'})
        coord_fmt = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'black', 'num_format': '0.00000'})

        cols = ["المحطة", "رقم العمود", "طول العمود", "الذراع", "الاحداثيات x", "الاحداثيات y", "التفاصيل"]
        for col_num, col_name in enumerate(cols):
            worksheet.write(0, col_num, col_name, header_fmt)
            worksheet.set_column(col_num, col_num, 15)

        curr_row, last_st = 1, None
        for _, row in df.iterrows():
            if last_st and row['المحطة'] != last_st:
                curr_row += 1 

            row_is_red = str(row['التفاصيل']) in ["مفقود", "مغروز"]
            for col_idx, col_name in enumerate(cols):
                val = row[col_name]
                if row_is_red: fmt = red_fmt
                elif col_name == "المحطة": fmt = station_fmt
                elif "الاحداثيات" in col_name: fmt = coord_fmt
                else: fmt = normal_fmt
                worksheet.write(curr_row, col_idx, val, fmt)
            
            last_st, curr_row = row['المحطة'], curr_row + 1

    st.success("✅ تم تحديث الكود: الطول يُستخرج الآن من الوصف حصراً.")
    st.download_button(label="📥 تحميل الملف المنظم", data=output.getvalue(), file_name="Station_Report_Final.xlsx")
