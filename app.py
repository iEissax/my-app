import streamlit as st
import zipfile
import pandas as pd
from lxml import etree
import re
import io

# تحسين دالة الترتيب لتجنب خطأ unhashable type
def natural_sort_key(s):
    if pd.isna(s) or s == "":
        return []
    # تحويل القائمة إلى "Tuple" لأن التوبل قابل للهاش (Hashable) وعملية الفرز تقبله
    return tuple(int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', str(s)))

st.set_page_config(page_title="مستخرج بيانات المحطات المطور", layout="wide")
st.title("📂 مستخرج KMZ - حل مشكلة الترتيب واستخراج الوصف")

uploaded_files = st.file_uploader("اختر ملفات KMZ", type=['kmz'], accept_multiple_files=True)

def process_kmz(file):
    with zipfile.ZipFile(file, 'r') as f:
        kml_filename = [name for name in f.namelist() if name.endswith('.kml')][0]
        kml_content = f.read(kml_filename)

    tree = etree.fromstring(kml_content)
    ns = {"kml": "http://www.opengis.net/kml/2.2"}
    data = []

    for pm in tree.xpath("//kml:Placemark", namespaces=ns):
        # 1. العنوان (Name) -> رقم المحطة ورقم العمود
        name_text = pm.xpath("./kml:name/text()", namespaces=ns)
        full_name = name_text[0].strip() if name_text else ""
        
        station_match = re.search(r'(\d+[\u0600-\u06FF]+|[\u0600-\u06FF]+\d+)', full_name)
        station_code = station_match.group(1) if station_match else "غير محدد"

        # رقم العمود من العنوان فقط
        clean_name = full_name.replace(station_code, "").strip()
        name_nums = re.findall(r'\d+', clean_name)
        column_num = name_nums[0] if name_nums else ""

        # 2. الوصف (Description) -> طول العمود والأذرعة
        desc = "".join(pm.xpath("./kml:description/text()", namespaces=ns))
        ext_vals = " ".join(pm.xpath(".//kml:value/text()", namespaces=ns))
        tech_info = (desc + " " + ext_vals).strip()

        val_height, val_arms = "", ""
        
        # البحث عن النمط المائل 12/2/1 في الوصف فقط
        pattern_match = re.search(r'(\d+)[/-](\d+)', tech_info)
        if pattern_match:
            val_height = pattern_match.group(1)  # الطول
            val_arms = pattern_match.group(2)    # الأذرعة
        else:
            # بحث بديل عن أرقام قياسية في الوصف
            h_match = re.search(r'\b(12|10|8|6|5)\b', tech_info)
            if h_match: val_height = h_match.group(1)

        # الكلمات الخاصة
        tech_lower = tech_info.lower()
        if "هاي" in tech_lower or "mast" in tech_lower:
            val_height, val_arms = "هاي ماست", 6
        elif "جداري" in tech_lower:
            val_height, val_arms = "جداري", 1

        # 3. الحالة والإحداثيات
        all_text = (full_name + " " + tech_info).lower()
        details = "مفقود" if "مفقود" in all_text else ("مغروز" if "مغروز" in all_text else "")

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
    
    # تحويل رقم العمود لنوع عددي للفرز
    df['رقم العمود'] = pd.to_numeric(df['رقم العمود'], errors='coerce').fillna(0).astype(int)
    
    # الفرز باستخدام Tuple بدلاً من List لتفادي TypeError
    df = df.sort_values(
        by=['المحطة', 'رقم العمود'], 
        key=lambda x: x.map(natural_sort_key) if x.name == 'المحطة' else x
    )
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('Data')
        worksheet.right_to_left()

        # التنسيقات
        f_head = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center', 'font_color': 'black'})
        f_stat = workbook.add_format({'bg_color': '#7F7F7F', 'border': 1, 'align': 'center', 'font_color': 'black', 'bold': True})
        f_red = workbook.add_format({'bg_color': '#FF0000', 'border': 1, 'align': 'center', 'font_color': 'black'})
        f_norm = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'black'})
        f_coord = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'black', 'num_format': '0.00000'})

        cols = ["المحطة", "رقم العمود", "طول العمود", "الذراع", "الاحداثيات x", "الاحداثيات y", "التفاصيل"]
        for i, name in enumerate(cols):
            worksheet.write(0, i, name, f_head)
            worksheet.set_column(i, i, 15)

        curr_row, last_st = 1, None
        for _, row in df.iterrows():
            if last_st and row['المحطة'] != last_st:
                curr_row += 1 

            is_red = str(row['التفاصيل']) in ["مفقود", "مغروز"]
            for j, col_name in enumerate(cols):
                val = row[col_name]
                if is_red: fmt = f_red
                elif col_name == "المحطة": fmt = f_stat
                elif "الاحداثيات" in col_name: fmt = f_coord
                else: fmt = f_norm
                worksheet.write(curr_row, j, val, fmt)
            
            last_st, curr_row = row['المحطة'], curr_row + 1

    st.success("✅ تم حل خطأ النوع (TypeError) وضبط استخراج الطول من الوصف.")
    st.download_button("📥 تحميل الملف النهائي", output.getvalue(), "Lighting_Standard_Final.xlsx")
