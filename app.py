import streamlit as st
import zipfile
import pandas as pd
from lxml import etree
import re
import io

# دالة للترتيب الطبيعي (تضمن أن 2 تأتي قبل 10)
def natural_key(string_):
    if isinstance(string_, int): return string_
    return [int(s) if s.isdigit() else s.lower() for s in re.split(r'(\d+)', str(string_))]

st.set_page_config(page_title="مستخرج بيانات المحطات", layout="wide")
st.title("📂 نظام تنظيم محطات الإنارة - الترتيب الطبيعي")

uploaded_files = st.file_uploader("اختر ملفات KMZ", type=['kmz'], accept_multiple_files=True)

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

        # استخراج المحطة
        station_match = re.search(r'(\d+[\u0600-\u06FF]+|[\u0600-\u06FF]+\d+)', full_name)
        station_code = station_match.group(1) if station_match else "غير محدد"

        # استخراج الأرقام (عمود وفيدر)
        clean_name = full_name.replace(station_code, "")
        nums = re.findall(r'\d+', clean_name)
        # تحويلها لأرقام حقيقية للفرز الصحيح
        column_num = int(nums[0]) if len(nums) >= 1 else 0
        feeder_num = int(nums[1]) if len(nums) >= 2 else 0

        # جلب النصوص
        desc = "".join(pm.xpath("./kml:description/text()", namespaces=ns))
        ext_vals = " ".join(pm.xpath(".//kml:value/text()", namespaces=ns))
        search_area = (full_name + " " + desc + " " + ext_vals).strip()

        val_height, val_arms = "", ""
        # معالجة نمط 9/2/2
        pattern_match = re.search(r'(\d+)[/-](\d+)[/-](\d+)', search_area)
        if pattern_match:
            val_height = pattern_match.group(1)
            val_arms = pattern_match.group(2)
        else:
            if "هاي" in search_area.lower(): val_height, val_arms = "هاي ماست", 6
            elif "جداري" in search_area: val_height, val_arms = "جداري", 1
            else:
                h_match = re.search(r'\b(12|10|9|8|6|5)\b', search_area)
                val_height = h_match.group(1) if h_match else ""

        details = ""
        if "مفقود" in search_area: details = "مفقود"
        elif "مغروز" in search_area: details = "مغروز"

        coords = pm.xpath(".//kml:coordinates/text()", namespaces=ns)
        lat_val, lon_val = 0.0, 0.0
        if coords:
            c_split = coords[0].strip().split(',')
            lat_val = round(float(c_split[1]), 5)
            lon_val = round(float(c_split[0]), 5)

        data.append({
            "المحطة": station_code,
            "رقم الفيدر": feeder_num,
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
    
    # تحويل رقم العمود والفيدر لنوع عددي لضمان الترتيب الصحيح (1, 2, 3...)
    df['رقم العمود'] = pd.to_numeric(df['رقم العمود'], errors='coerce').fillna(0).astype(int)
    df['رقم الفيدر'] = pd.to_numeric(df['رقم الفيدر'], errors='coerce').fillna(0).astype(int)

    # الترتيب: المحطة (أبجدي طبيعي)، ثم الفيدر، ثم العمود
    df = df.sort_values(
        by=['المحطة', 'رقم الفيدر', 'رقم العمود'], 
        key=lambda x: x.map(natural_key) if x.name == 'المحطة' else x
    )
    
    is_duplicate = df.duplicated(subset=['المحطة', 'رقم الفيدر', 'رقم العمود'], keep=False)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook, worksheet = writer.book, workbook.add_worksheet('Data')
        worksheet.right_to_left()

        # التنسيقات
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center', 'font_color': 'black'})
        station_fmt = workbook.add_format({'bg_color': '#7F7F7F', 'border': 1, 'align': 'center', 'font_color': 'black', 'bold': True})
        dup_fmt = workbook.add_format({'bg_color': '#DDEBF7', 'border': 1, 'align': 'center', 'font_color': 'black'})
        red_fmt = workbook.add_format({'bg_color': '#FF0000', 'border': 1, 'align': 'center', 'font_color': 'black'})
        normal_fmt = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'black'})
        coord_fmt = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'black', 'num_format': '0.00000'})

        cols = ["المحطة", "رقم الفيدر", "رقم العمود", "طول العمود", "الذراع", "الاحداثيات x", "الاحداثيات y", "التفاصيل"]
        for c, name in enumerate(cols):
            worksheet.write(0, c, name, header_fmt)
            worksheet.set_column(c, c, 15)

        curr_row, last_st = 1, None
        for idx, row in df.iterrows():
            if last_st is not None and row['المحطة'] != last_st:
                curr_row += 1 # الفاصل المطلوب

            is_dup, is_red = is_duplicate.loc[idx], str(row['التفاصيل']) in ["مفقود", "مغروز"]
            for col_idx, col_name in enumerate(cols):
                val = row[col_name]
                if is_red: fmt = red_fmt
                elif is_dup: fmt = dup_fmt
                elif col_name == "المحطة": fmt = station_fmt
                elif "الاحداثيات" in col_name: fmt = coord_fmt
                else: fmt = normal_fmt
                worksheet.write(curr_row, col_idx, val, fmt)
            last_st, curr_row = row['المحطة'], curr_row + 1

    st.success("✅ تم حل مشكلة الترتيب المتسلسل وتفعيل الفرز الهندسي.")
    st.download_button(label="📥 تحميل التقرير النهائي", data=output.getvalue(), file_name="Final_Sorted_Report.xlsx")


