import streamlit as st
import zipfile
import pandas as pd
from lxml import etree
import re
import io

st.set_page_config(page_title="مستخرج بيانات المحطات", layout="wide")
st.title("📂 نظام تصنيف شبكة الإنارة - إحداثيات خماسية")

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

        # استخراج رقم المحطة (904ج أو ج904)
        station_match = re.search(r'(\d+[\u0600-\u06FF]+|[\u0600-\u06FF]+\d+)', full_name)
        station_code = station_match.group(1) if station_match else "غير محدد"

        # تنظيف النص لاستخراج أرقام العمود والفيدر
        clean_name = full_name.replace(station_code, "")
        nums = re.findall(r'\d+', clean_name)
        
        column_num = int(nums[0]) if len(nums) >= 1 else 0
        feeder_num = int(nums[1]) if len(nums) >= 2 else 0

        # استخراج الطول
        desc = "".join(pm.xpath("./kml:description/text()", namespaces=ns))
        search_area = (full_name + " " + desc).lower()
        val_height = ""
        h_match = re.search(r'(\d{1,2})\s*(?:متر|م|m)\b', search_area)
        if h_match: val_height = h_match.group(1)
        elif "هاي ماست" in search_area: val_height = "هاي ماست"

        # الإحداثيات
        coords = pm.xpath(".//kml:coordinates/text()", namespaces=ns)
        lat_val, lon_val = 0.0, 0.0
        if coords:
            c_split = coords[0].strip().split(',')
            # تقريب الإحداثيات لـ 5 خانات عشرية (خماسيات)
            lat_val = round(float(c_split[1]), 5)
            lon_val = round(float(c_split[0]), 5)

        data.append({
            "المحطة": station_code,
            "رقم الفيدر": feeder_num,
            "رقم العمود": column_num,
            "طول العمود": val_height,
            "الاحداثيات x": lon_val,
            "الاحداثيات y": lat_val
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

        # التنسيقات مع توحيد لون الخط للأسود (font_color: black)
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#BFBFBF', 'border': 1, 'align': 'center', 'font_color': 'black'})
        station_col_fmt = workbook.add_format({'bg_color': '#7F7F7F', 'font_color': 'black', 'bold': True, 'border': 1, 'align': 'center'})
        dup_row_fmt = workbook.add_format({'bg_color': '#DDEBF7', 'border': 1, 'align': 'center', 'font_color': 'black'})
        normal_fmt = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'black'})
        
        # تنسيق خاص للأرقام العشرية (الخماسيات)
        coord_fmt = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'black', 'num_format': '0.00000'})

        # كتابة العناوين
        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, header_fmt)
            worksheet.set_column(col_num, col_num, 18)

        curr_excel_row = 1
        last_station = None

        for idx, row in df.iterrows():
            if last_station is not None and row['المحطة'] != last_station:
                curr_excel_row += 1 

            row_is_dup = is_duplicate.loc[idx]

            for col_idx, col_name in enumerate(df.columns):
                val = row[col_name]
                
                # اختيار التنسيق المناسب
                if row_is_dup:
                    fmt = dup_row_fmt
                elif col_name == "المحطة":
                    fmt = station_col_fmt
                elif "الاحداثيات" in col_name:
                    fmt = coord_fmt
                else:
                    fmt = normal_fmt
                
                worksheet.write(curr_excel_row, col_idx, val, fmt)
            
            last_station = row['المحطة']
            curr_excel_row += 1

    st.success("✅ تم التنسيق: خطوط سوداء، إحداثيات خماسية، وفواصل بين المحطات.")
    st.download_button(
        label="📥 تحميل التقرير النهائي",
        data=output.getvalue(),
        file_name="Lighting_Report_Final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
