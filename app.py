import streamlit as st
import zipfile
import pandas as pd
from lxml import etree
import re
import io

# دالة الترتيب الطبيعي
def natural_sort_key(s):
    if pd.isna(s) or s == "":
        return tuple()
    return tuple(int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', str(s)))

st.set_page_config(page_title="مستخرج بيانات المحطات", layout="wide")
st.title("📂 معالجة KMZ - تظليل المكرر (أزرق) والمفقود (أحمر)")

uploaded_files = st.file_uploader("اختر ملفات KMZ", type=['kmz'], accept_multiple_files=True)

def process_kmz(file):
    with zipfile.ZipFile(file, 'r') as f:
        kml_filename = [name for name in f.namelist() if name.endswith('.kml')][0]
        kml_content = f.read(kml_filename)

    tree = etree.fromstring(kml_content)
    ns = {"kml": "http://www.opengis.net/kml/2.2"}
    data = []

    for pm in tree.xpath("//kml:Placemark", namespaces=ns):
        # 1. العنوان (Name) -> المحطة، رقم العمود، رقم الفيدر
        name_text = pm.xpath("./kml:name/text()", namespaces=ns)
        full_name = name_text[0].strip() if name_text else ""
        
        st_match = re.search(r'(\d+[\u0600-\u06FF]+|[\u0600-\u06FF]+\d+)', full_name)
        station_code = st_match.group(1) if st_match else "غير محدد"

        clean_name = full_name.replace(station_code, "").strip()
        name_nums = re.findall(r'\d+', clean_name)
        
        column_num = name_nums[0] if len(name_nums) >= 1 else ""
        feeder_num = name_nums[1] if len(name_nums) >= 2 else ""

        # 2. الوصف (Description) -> طول العمود والأذرعة
        desc = "".join(pm.xpath("./kml:description/text()", namespaces=ns))
        ext_vals = " ".join(pm.xpath(".//kml:value/text()", namespaces=ns))
        tech_info = (desc + " " + ext_vals).strip()

        val_height, val_arms = "", ""
        pattern_match = re.search(r'(\d+)[/-](\d+)', tech_info)
        if pattern_match:
            val_height = pattern_match.group(1)
            val_arms = pattern_match.group(2)
        else:
            h_search = re.search(r'\b(12|10|8|6|5)\b', tech_info)
            if h_search: val_height = h_search.group(1)

        tech_lower = tech_info.lower()
        if "هاي" in tech_lower or "mast" in tech_lower:
            val_height, val_arms = "هاي ماست", 6
        elif "جداري" in tech_lower:
            val_height, val_arms = "جداري", 1

        # 3. الإحداثيات والحالة
        coords = pm.xpath(".//kml:coordinates/text()", namespaces=ns)
        lat_v, lon_v = 0.0, 0.0
        if coords:
            c_split = coords[0].strip().split(',')
            lat_v, lon_v = round(float(c_split[1]), 5), round(float(c_split[0]), 5)

        all_txt = (full_name + " " + tech_info).lower()
        detail = "مفقود" if "مفقود" in all_txt else ("مغروز" if "مغروز" in all_txt else "")

        data.append({
            "المحطة": station_code,
            "رقم العمود": column_num,
            "رقم الفيدر": feeder_num,
            "طول العمود": val_height,
            "الذراع": val_arms,
            "الاحداثيات x": lon_v,
            "الاحداثيات y": lat_v,
            "التفاصيل": detail
        })
    return pd.DataFrame(data)

if uploaded_files:
    all_data = [process_kmz(f) for f in uploaded_files]
    df = pd.concat(all_data, ignore_index=True)
    
    # تحويل الأرقام للفرز وتحديد المكرر
    df['رقم العمود'] = pd.to_numeric(df['رقم العمود'], errors='coerce').fillna(0).astype(int)
    df['رقم الفيدر'] = pd.to_numeric(df['رقم الفيدر'], errors='coerce').fillna(0).astype(int)
    
    # الترتيب
    df = df.sort_values(by=['المحطة', 'رقم الفيدر', 'رقم العمود'],
                        key=lambda x: x.map(natural_sort_key) if x.name == 'المحطة' else x)

    # تحديد الصفوف المكررة (نفس المحطة + فيدر + رقم عمود)
    is_duplicate = df.duplicated(subset=['المحطة', 'رقم الفيدر', 'رقم العمود'], keep=False)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook, worksheet = writer.book, writer.book.add_worksheet('Report')
        worksheet.right_to_left()

        # التنسيقات
        fmt_h = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center', 'font_color': 'black'})
        fmt_s = workbook.add_format({'bg_color': '#7F7F7F', 'border': 1, 'align': 'center', 'font_color': 'black', 'bold': True})
        fmt_r = workbook.add_format({'bg_color': '#FF0000', 'border': 1, 'align': 'center', 'font_color': 'black'}) # أحمر للمفقود
        fmt_b = workbook.add_format({'bg_color': '#DDEBF7', 'border': 1, 'align': 'center', 'font_color': 'black'}) # أزرق للمكرر
        fmt_n = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'black'})
        fmt_c = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'black', 'num_format': '0.00000'})

        cols = ["المحطة", "رقم العمود", "رقم الفيدر", "طول العمود", "الذراع", "الاحداثيات x", "الاحداثيات y", "التفاصيل"]
        for i, c_name in enumerate(cols):
            worksheet.write(0, i, c_name, fmt_h)
            worksheet.set_column(i, i, 15)

        curr_row, last_st = 1, None
        for idx, row in df.iterrows():
            if last_st and row['المحطة'] != last_st:
                curr_row += 1 

            row_is_red = str(row['التفاصيل']) in ["مفقود", "مغروز"]
            row_is_blue = is_duplicate.loc[idx]

            for j, c_name in enumerate(cols):
                val = row[c_name]
                # تحديد التنسيق حسب الأولوية: مفقود (أحمر) > مكرر (أزرق)
                if row_is_red: fmt = fmt_r
                elif row_is_blue: fmt = fmt_b
                elif c_name == "المحطة": fmt = fmt_s
                elif "الاحداثيات" in c_name: fmt = fmt_c
                else: fmt = fmt_n
                worksheet.write(curr_row, j, val, fmt)
            
            last_st, curr_row = row['المحطة'], curr_row + 1

    st.success("✅ تم التحديث: تظليل المكرر بالأزرق والمفقود بالأحمر مع فصل الفيدر.")
    st.download_button("📥 تحميل التقرير النهائي", output.getvalue(), "Lighting_Report_Styled.xlsx")
