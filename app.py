import streamlit as st
import zipfile
import pandas as pd
from lxml import etree
import re
import io

st.set_page_config(page_title="مستخرج بيانات المحطات", layout="wide")
st.title("📂 مستخرج بيانات KMZ - التنسيق النهائي المعتمد")

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

        # 1. التعرف على رقم المحطة (مثل 904ج أو ج557)
        # تم تبسيط النمط لتفادي أخطاء الـ PatternError
        station_match = re.search(r'(\d+[\u0600-\u06FF]+|[\u0600-\u06FF]+\d+)', full_name)
        station_code = station_match.group(1) if station_match else "غير محدد"

        # 2. استخراج أرقام العمود والفيدر
        clean_name = full_name.replace(station_code, "")
        nums = re.findall(r'\d+', clean_name)
        column_num = int(nums[0]) if len(nums) >= 1 else 0
        feeder_num = int(nums[1]) if len(nums) >= 2 else 0

        # 3. جلب النصوص للبحث عن الطول والأذرعة
        desc = "".join(pm.xpath("./kml:description/text()", namespaces=ns))
        # جلب البيانات من حقول القيمة سواء كانت Data أو SimpleData
        ext_vals = " ".join(pm.xpath(".//kml:value/text()", namespaces=ns))
        search_area = (full_name + " " + desc + " " + ext_vals).strip().lower()

        # --- استخراج الطول (تم إصلاح المنطق لضمان الظهور) ---
        val_height = ""
        if "هاي" in search_area or "mast" in search_area:
            val_height = "هاي ماست"
        elif "جداري" in search_area:
            val_height = "جداري"
        else:
            # البحث عن أرقام الأطوال القياسية
            h_match = re.search(r'\b(12|10|8|6|5)\b', search_area)
            if h_match:
                val_height = h_match.group(1)

        # --- استخراج الأذرعة ---
        lamps = ""
        if val_height == "هاي ماست":
            lamps = 6
        elif any(kw in search_area for kw in ["دبل", "double", "2/2", "ثنائي"]):
            lamps = 2
        elif any(kw in search_area for kw in ["مفرد", "single", "1/1"]):
            lamps = 1

        # 4. تحديد التفاصيل (مفقود/مغروز)
        details = ""
        if "مفقود" in search_area: details = "مفقود"
        elif "مغروز" in search_area: details = "مغروز"

        # 5. الإحداثيات (خماسية)
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
            "الذراع": lamps,
            "الاحداثيات x": lon_val,
            "الاحداثيات y": lat_val,
            "التفاصيل": details
        })
    return pd.DataFrame(data)

if uploaded_files:
    all_dfs = [process_kmz(f) for f in uploaded_files]
    df = pd.concat(all_dfs, ignore_index=True)
    df = df.sort_values(by=['المحطة', 'رقم الفيدر', 'رقم العمود'])
    
    is_duplicate = df.duplicated(subset=['المحطة', 'رقم الفيدر', 'رقم العمود'], keep=False)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('Data')
        worksheet.right_to_left()

        # تنسيقات الخط الأسود
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center', 'font_color': 'black'})
        station_fmt = workbook.add_format({'bg_color': '#7F7F7F', 'border': 1, 'align': 'center', 'font_color': 'black', 'bold': True})
        dup_fmt = workbook.add_format({'bg_color': '#DDEBF7', 'border': 1, 'align': 'center', 'font_color': 'black'})
        red_fmt = workbook.add_format({'bg_color': '#FF0000', 'border': 1, 'align': 'center', 'font_color': 'black'})
        normal_fmt = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'black'})
        coord_fmt = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'black', 'num_format': '0.00000'})

        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, header_fmt)
            worksheet.set_column(col_num, col_num, 15)

        curr_row = 1
        last_st = None
        for idx, row in df.iterrows():
            if last_st is not None and row['المحطة'] != last_st:
                curr_row += 1 

            row_is_dup = is_duplicate.loc[idx]
            row_is_red = str(row['التفاصيل']) in ["مفقود", "مغروز"]

            for col_idx, col_name in enumerate(df.columns):
                val = row[col_name]
                if row_is_red: fmt = red_fmt
                elif row_is_dup: fmt = dup_fmt
                elif col_name == "المحطة": fmt = station_fmt
                elif "الاحداثيات" in col_name: fmt = coord_fmt
                else: fmt = normal_fmt
                worksheet.write(curr_row, col_idx, val, fmt)
            
            last_st = row['المحطة']
            curr_row += 1

    st.success("✅ تم إصلاح الخطأ البرمجي وضمان ظهور الأطوال.")
    st.download_button(label="📥 تحميل الملف النهائي", data=output.getvalue(), file_name="Final_Station_Report.xlsx")
    
