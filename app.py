import streamlit as st
import zipfile
import pandas as pd
from lxml import etree
import re
import io

st.set_page_config(page_title="مستخرج بيانات المحطات الذكي", layout="wide")
st.title("📂 نظام تقارير الإنارة - مع الملخص التلقائي")

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

        # 1. استخراج رقم المحطة (904ج)
        station_match = re.search(r'(\d+[\u0600-\u06FF]+|[\u0600-\u06FF]+\d+)', full_name)
        station_code = station_match.group(1) if station_match else "غير محدد"

        # 2. استخراج الأرقام (عمود، فيدر)
        clean_name = full_name.replace(station_code, "")
        nums = re.findall(r'\d+', clean_name)
        column_num = int(nums[0]) if len(nums) >= 1 else 0
        feeder_num = int(nums[1]) if len(nums) >= 2 else 0

        # 3. جلب الوصف للحالة والطول
        desc = "".join(pm.xpath("./kml:description/text()", namespaces=ns))
        ext_vals = " ".join(pm.xpath(".//kml:Data/kml:value/text()", namespaces=ns))
        search_area = (full_name + " " + desc + " " + ext_vals).strip()

        details = ""
        if "مفقود" in search_area: details = "مفقود"
        elif "مغروز" in search_area: details = "مغروز"

        val_height = ""
        h_match = re.search(r'(\d{1,2})\s*(?:متر|م|m)\b', search_area.lower())
        if h_match: val_height = h_match.group(1)
        elif "هاي ماست" in search_area.lower(): val_height = "هاي ماست"

        # 4. الإحداثيات (خماسية)
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
            "الاحداثيات x": lon_val,
            "الاحداثيات y": lat_val,
            "التفاصيل": details
        })
    return pd.DataFrame(data)

if uploaded_files:
    all_dfs = [process_kmz(f) for f in uploaded_files]
    df = pd.concat(all_dfs, ignore_index=True)
    df = df.sort_values(by=['المحطة', 'رقم الفيدر', 'رقم العمود'])
    
    # حساب الإحصائيات لكل محطة
    station_counts = df['المحطة'].value_counts()
    is_duplicate = df.duplicated(subset=['المحطة', 'رقم الفيدر', 'رقم العمود'], keep=False)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('Data')
        worksheet.right_to_left()

        # التنسيقات
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#BFBFBF', 'border': 1, 'align': 'center', 'font_color': 'black'})
        station_col_fmt = workbook.add_format({'bg_color': '#7F7F7F', 'border': 1, 'align': 'center', 'font_color': 'black', 'bold': True})
        summary_fmt = workbook.add_format({'bg_color': '#F2F2F2', 'bold': True, 'italic': True, 'font_color': 'black', 'align': 'right', 'border': 1})
        dup_fmt = workbook.add_format({'bg_color': '#DDEBF7', 'border': 1, 'align': 'center', 'font_color': 'black'})
        red_fmt = workbook.add_format({'bg_color': '#FF0000', 'border': 1, 'align': 'center', 'font_color': 'black', 'bold': True})
        normal_fmt = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'black'})
        coord_fmt = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'black', 'num_format': '0.00000'})

        # العناوين
        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, header_fmt)
            worksheet.set_column(col_num, col_num, 20)

        curr_row = 1
        last_st = None

        for idx, row in df.iterrows():
            # عند الانتقال لمحطة جديدة، أضف ملخص المحطة السابقة
            if last_st is not None and row['المحطة'] != last_st:
                count = station_counts[last_st]
                worksheet.merge_range(curr_row, 0, curr_row, len(df.columns)-1, f"--- إجمالي أعمدة المحطة ({last_st}): {count} أعمدة ---", summary_fmt)
                curr_row += 2 # ترك صف إضافي فارغ

            row_is_dup = is_duplicate.loc[idx]
            row_is_red = row['التفاصيل'] in ["مفقود", "مغروز"]

            for col_idx, col_name in enumerate(df.columns):
                val = row[col_name]
                if row_is_red: fmt = red_fmt
                elif row_is_dup: fmt = dup_fmt
                elif col_name == "المحطة": fmt = station_col_fmt
                elif "الاحداثيات" in col_name: fmt = coord_fmt
                else: fmt = normal_fmt
                
                worksheet.write(curr_row, col_idx, val, fmt)
            
            last_st = row['المحطة']
            curr_row += 1

        # إضافة ملخص لآخر محطة في الملف
        if last_st:
            worksheet.merge_range(curr_row, 0, curr_row, len(df.columns)-1, f"--- إجمالي أعمدة المحطة ({last_st}): {station_counts[last_st]} أعمدة ---", summary_fmt)

    st.success("✅ تم تحديث النظام: التقرير الآن يتضمن ملخصاً لكل محطة بشكل آلي.")
    st.download_button(label="📥 تحميل التقرير النهائي المطور", data=output.getvalue(), file_name="Station_Smart_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
