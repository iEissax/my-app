import streamlit as st
import zipfile
import pandas as pd
from lxml import etree
import re
import io

st.set_page_config(page_title="مستخرج بيانات المحطات", layout="wide")
st.title("📂 نظام تنظيم محطات الإنارة - الإصدار الاحترافي")

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

        # 1. استخراج رقم المحطة (ج557 أو 904ج)
        station_match = re.search(r'(\d+[\u0600-\u06FF]+|[\u0600-\u06FF]+\d+)', full_name)
        station_code = station_match.group(1) if station_match else "غير محدد"

        # 2. جلب النصوص للبحث عن (الطول / الذراع / الحالة)
        desc = "".join(pm.xpath("./kml:description/text()", namespaces=ns))
        ext_vals = " ".join(pm.xpath(".//kml:value/text()", namespaces=ns))
        search_area = (full_name + " " + desc + " " + ext_vals).strip()

        # 3. استخراج أرقام العمود والفيدر (بعد استبعاد المحطة)
        clean_name = full_name.replace(station_code, "")
        nums = re.findall(r'\d+', clean_name)
        column_num = int(nums[0]) if len(nums) >= 1 else ""
        feeder_num = int(nums[1]) if len(nums) >= 2 else ""

        val_height = ""
        val_arms = ""

        # --- معالجة نمط 9/2/2 أو 10-1-1 ---
        pattern_match = re.search(r'(\d+)[/-](\d+)[/-](\d+)', search_area)
        if pattern_match:
            val_height = pattern_match.group(1)
            val_arms = pattern_match.group(2)
        else:
            if "هاي" in search_area.lower():
                val_height = "هاي ماست"
                val_arms = 6
            elif "جداري" in search_area:
                val_height = "جداري"
                val_arms = 1
            else:
                h_match = re.search(r'\b(12|10|9|8|6|5)\b', search_area)
                val_height = h_match.group(1) if h_match else ""

        # 4. التفاصيل والإحداثيات
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
    
    # أهم خطوة: الترتيب حسب المحطة ثم الفيدر ثم العمود
    df = df.sort_values(by=['المحطة', 'رقم الفيدر', 'رقم العمود'])
    
    # تحديد المكررات
    is_duplicate = df.duplicated(subset=['المحطة', 'رقم الفيدر', 'رقم العمود'], keep=False)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('Data')
        worksheet.right_to_left()

        # التنسيقات (الخط أسود)
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center', 'font_color': 'black'})
        station_fmt = workbook.add_format({'bg_color': '#7F7F7F', 'border': 1, 'align': 'center', 'font_color': 'black', 'bold': True})
        dup_fmt = workbook.add_format({'bg_color': '#DDEBF7', 'border': 1, 'align': 'center', 'font_color': 'black'})
        red_fmt = workbook.add_format({'bg_color': '#FF0000', 'border': 1, 'align': 'center', 'font_color': 'black'})
        normal_fmt = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'black'})
        coord_fmt = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'black', 'num_format': '0.00000'})

        # كتابة العناوين
        cols = ["المحطة", "رقم الفيدر", "رقم العمود", "طول العمود", "الذراع", "الاحداثيات x", "الاحداثيات y", "التفاصيل"]
        for col_num, col_name in enumerate(cols):
            worksheet.write(0, col_num, col_name, header_fmt)
            worksheet.set_column(col_num, col_num, 15)

        curr_excel_row = 1
        last_station = None

        # كتابة البيانات مع الفصل الإجباري
        for idx, row in df.iterrows():
            # إذا تغيرت المحطة، اترك صفاً فارغاً (هذا ما سيمنع التسلسل المستمر)
            if last_station is not None and row['المحطة'] != last_station:
                curr_excel_row += 1 

            row_is_dup = is_duplicate.loc[idx]
            row_is_red = str(row['التفاصيل']) in ["مفقود", "مغروز"]

            for col_idx, col_name in enumerate(cols):
                val = row[col_name]
                
                if row_is_red: fmt = red_fmt
                elif row_is_dup: fmt = dup_fmt
                elif col_name == "المحطة": fmt = station_fmt
                elif "الاحداثيات" in col_name: fmt = coord_fmt
                else: fmt = normal_fmt
                
                worksheet.write(curr_excel_row, col_idx, val, fmt)
            
            last_station = row['المحطة']
            curr_excel_row += 1

    st.success("✅ تم الفصل بين المحطات بنجاح وتنسيق البيانات.")
    st.download_button(label="📥 تحميل التقرير المنظم", data=output.getvalue(), file_name="Organized_Station_Report.xlsx")
