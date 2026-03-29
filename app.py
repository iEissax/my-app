import streamlit as st
import zipfile
import pandas as pd
from lxml import etree
import re
import io

# دالة الترتيب الطبيعي (1, 2, 10)
def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', str(s))]

st.set_page_config(page_title="مستخرج بيانات المحطات", layout="wide")
st.title("📂 مستخرج KMZ - فصل رقم العمود عن نمط الطول")

uploaded_files = st.file_uploader("اختر ملفات KMZ", type=['kmz'], accept_multiple_files=True)

def process_kmz(file):
    with zipfile.ZipFile(file, 'r') as f:
        kml_filename = [name for name in f.namelist() if name.endswith('.kml')][0]
        kml_content = f.read(kml_filename)

    tree = etree.fromstring(kml_content)
    ns = {"kml": "http://www.opengis.net/kml/2.2"}
    data = []

    for pm in tree.xpath("//kml:Placemark", namespaces=ns):
        # --- 1. استخراج العنوان (Name) ---
        name_text = pm.xpath("./kml:name/text()", namespaces=ns)
        full_name = name_text[0].strip() if name_text else ""
        
        # --- 2. استخراج رقم المحطة ورقم العمود من العنوان ---
        # استخراج المحطة (مثل ج557)
        st_match = re.search(r'(\d+[\u0600-\u06FF]+|[\u0600-\u06FF]+\d+)', full_name)
        station_code = st_match.group(1) if st_match else "غير محدد"
        
        # استخراج رقم العمود من العنوان (الرقم الذي يتبقى في العنوان بعد حذف المحطة)
        clean_name = full_name.replace(station_code, "").strip()
        name_nums = re.findall(r'\d+', clean_name)
        column_num = name_nums[0] if name_nums else ""

        # --- 3. استخراج الطول والأذرعة من الوصف (النمط 12/2/1) ---
        desc = "".join(pm.xpath("./kml:description/text()", namespaces=ns))
        ext_data = " ".join(pm.xpath(".//kml:value/text()", namespaces=ns))
        
        # البحث عن نمط الأرقام المائلة في الوصف أو البيانات الإضافية
        # نبحث عن نمط X/Y/Z
        pattern_match = re.findall(r'(\d+)[/-](\d+)', desc + " " + ext_data)
        
        val_height, val_arms = "", ""
        if pattern_match:
            val_height = pattern_match[0][0] # الرقم الأول هو الطول (12)
            val_arms = pattern_match[0][1]   # الرقم الثاني هو الذراع (2)
        else:
            # محاولة أخيرة إذا لم يجد النمط المائل، يبحث عن أرقام مفردة
            h_search = re.search(r'\b(12|10|8|6|5)\b', desc + " " + ext_data)
            if h_search: val_height = h_search.group(1)

        # الحالات الخاصة
        all_info_lower = (full_name + " " + desc + " " + ext_data).lower()
        if "هاي" in all_info_lower or "mast" in all_info_lower:
            val_height, val_arms = "هاي ماست", 6
        elif "جداري" in all_info_lower:
            val_height, val_arms = "جداري", 1

        # --- 4. الإحداثيات (خماسية) ---
        coords = pm.xpath(".//kml:coordinates/text()", namespaces=ns)
        lat, lon = 0.0, 0.0
        if coords:
            c_split = coords[0].strip().split(',')
            lat, lon = round(float(c_split[1]), 5), round(float(c_split[0]), 5)

        # --- 5. الحالة ---
        detail = "مفقود" if "مفقود" in all_info_lower else ("مغروز" if "مغروز" in all_info_lower else "")

        data.append({
            "المحطة": station_code,
            "رقم العمود": column_num,
            "طول العمود": val_height,
            "الذراع": val_arms,
            "الاحداثيات x": lon,
            "الاحداثيات y": lat,
            "التفاصيل": detail
        })
    return pd.DataFrame(data)

if uploaded_files:
    all_data = [process_kmz(f) for f in uploaded_files]
    df = pd.concat(all_data, ignore_index=True)

    # الترتيب الطبيعي حسب المحطة ثم رقم العمود
    df['رقم العمود'] = pd.to_numeric(df['رقم العمود'], errors='coerce').fillna(0).astype(int)
    df = df.sort_values(by=['المحطة', 'رقم العمود'], key=lambda x: x.map(natural_sort_key) if x.name == 'المحطة' else x)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('Data')
        worksheet.right_to_left()

        # التنسيقات (خط أسود)
        f_head = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center', 'font_color': 'black'})
        f_stat = workbook.add_format({'bg_color': '#7F7F7F', 'border': 1, 'align': 'center', 'font_color': 'black', 'bold': True})
        f_red = workbook.add_format({'bg_color': '#FF0000', 'border': 1, 'align': 'center', 'font_color': 'black'})
        f_norm = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'black'})
        f_coord = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'black', 'num_format': '0.00000'})

        cols = ["المحطة", "رقم العمود", "طول العمود", "الذراع", "الاحداثيات x", "الاحداثيات y", "التفاصيل"]
        for i, col in enumerate(cols):
            worksheet.write(0, i, col, f_head)
            worksheet.set_column(i, i, 15)

        curr_row, last_st = 1, None
        for _, row in df.iterrows():
            if last_st and row['المحطة'] != last_st:
                curr_row += 1 

            is_red = row['التفاصيل'] in ["مفقود", "مغروز"]
            for j, c_name in enumerate(cols):
                val = row[c_name]
                if is_red: f = f_red
                elif c_name == "المحطة": f = f_stat
                elif "الاحداثيات" in c_name: f = f_coord
                else: f = f_norm
                worksheet.write(curr_row, j, val, f)
            last_st, curr_row = row['المحطة'], curr_row + 1

    st.success("✅ تم الفصل بنجاح: رقم العمود من العنوان، والطول/الذراع من الوصف.")
    st.download_button("📥 تحميل الملف النهائي", output.getvalue(), "Lighting_Standard_Report.xlsx")
