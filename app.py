import streamlit as st
import zipfile
import pandas as pd
from lxml import etree
import re
import io

st.set_page_config(page_title="مستخرج بيانات الإنارة", layout="wide")
st.title("📂 نظام استخراج وتصنيف محطات الإنارة")

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

        # استخراج رقم المحطة (مثل ج557)
        station_match = re.search(r'([\u0600-\u06FFa-zA-Z]+\d+)', full_name)
        station_code = station_match.group(1) if station_match else "غير محدد"

        # استخراج الأرقام (عمود، فيدر)
        all_nums = re.findall(r'\d+', full_name)
        numbers = [n for n in all_nums if n not in station_code]
        
        column_num = int(numbers[0]) if len(numbers) >= 1 else 0
        feeder_num = int(numbers[1]) if len(numbers) >= 2 else 0

        # استخراج طول العمود
        desc = pm.xpath("./kml:description/text()", namespaces=ns)
        desc_text = desc[0] if desc else ""
        ext_vals = " ".join(pm.xpath(".//kml:Data/kml:value/text()", namespaces=ns))
        search_area = (full_name + " " + desc_text + " " + ext_vals).strip().lower()

        val_height = ""
        h_match = re.search(r'(\d{1,2})\s*(?:متر|م|m|meter)\b', search_area)
        if h_match: val_height = h_match.group(1)
        elif "هاي ماست" in search_area: val_height = "هاي ماست"

        # الإحداثيات
        coords = pm.xpath(".//kml:coordinates/text()", namespaces=ns)
        lat_val, lon_val = 0.0, 0.0
        if coords:
            c_split = coords[0].strip().split(',')
            lat_val = float(c_split[1])
            lon_val = float(c_split[0])

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
    result_df = pd.concat(all_dfs, ignore_index=True)

    # الترتيب لضمان تجميع المحطات (مثل الصورة)
    result_df = result_df.sort_values(by=['المحطة', 'رقم الفيدر', 'رقم العمود'])

    # تحديد المكررات قبل إضافة الصفوف الفارغة
    duplicates = result_df.duplicated(subset=['المحطة', 'رقم الفيدر', 'رقم العمود'], keep=False)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('MainData')
        worksheet.right_to_left()

        # التنسيقات
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center'})
        station_fmt = workbook.add_format({'bg_color': '#7F7F7F', 'font_color': 'white', 'border': 1, 'align': 'center'})
        dup_fmt = workbook.add_format({'bg_color': '#BDD7EE', 'border': 1, 'align': 'center'}) # أزرق للمكرر
        cell_fmt = workbook.add_format({'border': 1, 'align': 'center'})

        # كتابة العناوين
        columns = result_df.columns.tolist()
        for col_num, col_name in enumerate(columns):
            worksheet.write(0, col_num, col_name, header_fmt)
            worksheet.set_column(col_num, col_num, 15)

        current_row = 1
        last_station = None

        for i, row in result_df.iterrows():
            # إذا تغيرت المحطة، اترك صفاً فارغاً (مثل الصورة)
            if last_station is not None and row['المحطة'] != last_station:
                current_row += 1 

            is_dup = duplicates.loc[i]
            
            for col_num, col_name in enumerate(columns):
                val = row[col_name]
                
                # تطبيق التنسيق: أزرق للمكرر، رمادي لعمود المحطة
                if is_dup:
                    fmt = dup_fmt
                elif col_name == "المحطة":
                    fmt = station_fmt
                else:
                    fmt = cell_fmt
                
                worksheet.write(current_row, col_num, val, fmt)
            
            last_station = row['المحطة']
            current_row += 1

    st.success("✅ تم استخراج البيانات وتنظيمها حسب المحطة.")
    st.download_button(
        label="📥 تحميل ملف الإكسيل المنظم",
        data=output.getvalue(),
        file_name="Lighting_Grid_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
