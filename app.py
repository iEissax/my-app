import streamlit as st
import zipfile
import pandas as pd
from lxml import etree
import re
import io

st.set_page_config(page_title="مستخرج بيانات المحطات", layout="wide")
st.title("📂 نظام استخراج وتصنيف محطات الإنارة (نمط ج)")

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

        # 1. التعرف على المحطة (نمط ج557 أو ج900 أو A12)
        # يبحث عن حرف عربي أو إنجليزي ملتصق به أرقام
        station_match = re.search(r'([\u0600-\u06FFa-zA-Z]+\d+)', full_name)
        station_code = station_match.group(1) if station_match else "غير محدد"

        # 2. استخراج بقية الأرقام (الفيدر والعمود)
        # نقوم بإزالة نص المحطة من الاسم الكامل أولاً لضمان دقة استخراج الأرقام الأخرى
        remaining_text = full_name.replace(station_code, "")
        other_numbers = re.findall(r'\d+', remaining_text)
        
        # الترتيب الافتراضي: أول رقم بعد المحطة هو العمود، والثاني هو الفيدر (أو حسب ملفك)
        column_num = int(other_numbers[0]) if len(other_numbers) >= 1 else 0
        feeder_num = int(other_numbers[1]) if len(other_numbers) >= 2 else 0

        # 3. استخراج طول العمود من الوصف أو الاسم
        desc = pm.xpath("./kml:description/text()", namespaces=ns)
        desc_text = desc[0] if desc else ""
        search_area = (full_name + " " + desc_text).lower()
        
        val_height = ""
        # البحث عن نمط "10م" أو "12 م" أو "8m"
        h_match = re.search(r'(\d{1,2})\s*(?:متر|م|m)\b', search_area)
        if h_match:
            val_height = h_match.group(1)
        elif "هاي ماست" in search_area:
            val_height = "هاي ماست"

        # 4. الإحداثيات
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
    df = pd.concat(all_dfs, ignore_index=True)

    # ترتيب البيانات لتجميع المحطات
    df = df.sort_values(by=['المحطة', 'رقم الفيدر', 'رقم العمود'])
    
    # تحديد المكررات (أزرق)
    is_duplicate = df.duplicated(subset=['المحطة', 'رقم الفيدر', 'رقم العمود'], keep=False)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('MainData')
        worksheet.right_to_left()

        # التنسيقات (مطابقة للمثال)
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#BFBFBF', 'border': 1, 'align': 'center'})
        station_col_fmt = workbook.add_format({'bg_color': '#595959', 'font_color': 'white', 'bold': True, 'border': 1, 'align': 'center'})
        dup_fmt = workbook.add_format({'bg_color': '#DDEBF7', 'border': 1, 'align': 'center'}) # أزرق فاتح للمكرر
        normal_fmt = workbook.add_format({'border': 1, 'align': 'center'})

        # كتابة العناوين
        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, header_fmt)
            worksheet.set_column(col_num, col_num, 15)

        curr_row = 1
        last_st = None

        for idx, row in df.iterrows():
            # إضافة صف فارغ عند تغيير المحطة (مثل الصورة)
            if last_st is not None and row['المحطة'] != last_st:
                curr_row += 1

            row_is_dup = is_duplicate.loc[idx]

            for col_idx, col_name in enumerate(df.columns):
                val = row[col_name]
                
                # اختيار التنسيق
                if row_is_dup:
                    fmt = dup_fmt
                elif col_name == "المحطة":
                    fmt = station_col_fmt
                else:
                    fmt = normal_fmt
                
                worksheet.write(curr_row, col_idx, val, fmt)
            
            last_st = row['المحطة']
            curr_row += 1

    st.success("✅ تم التعرف على المحطات وتنظيم الجدول بنجاح.")
    st.download_button(
        label="📥 تحميل ملف الإكسيل المنسق (ج557 / ج900)",
        data=output.getvalue(),
        file_name="Station_Grid_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
