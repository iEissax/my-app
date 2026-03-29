import streamlit as st
import zipfile
import pandas as pd
from lxml import etree
import re
import io

st.set_page_config(page_title="مستخرج بيانات المحطات", layout="wide")
st.title("📂 مستخرج بيانات شبكة الإنارة - نظام المحطات")

# تحميل الملفات
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

        # 1. استخراج رقم المحطة (مثل ج557 أو A12)
        # يبحث عن حرف أو أكثر متبوعاً بأرقام في بداية الكلمة
        station_match = re.search(r'([\u0600-\u06FFa-zA-Z]+\d+)', full_name)
        station_code = station_match.group(1) if station_match else "غير محدد"

        # 2. استخراج رقم الفيدر ورقم العمود من بقية النص
        numbers = re.findall(r'\d+', full_name)
        # إذا استخرجنا المحطة (ج557)، الرقم الأول 557 سيظهر في findall، لذا نتجاوزه
        if station_match and str(numbers[0]) in station_code:
            column_num = int(numbers[1]) if len(numbers) >= 2 else 0
            feeder_num = int(numbers[2]) if len(numbers) >= 3 else 0
        else:
            column_num = int(numbers[0]) if len(numbers) >= 1 else 0
            feeder_num = int(numbers[1]) if len(numbers) >= 2 else 0

        # 3. جلب الوصف والبيانات الإضافية للبحث عن "طول العمود"
        desc = pm.xpath("./kml:description/text()", namespaces=ns)
        desc_text = desc[0] if desc else ""
        ext_vals = " ".join(pm.xpath(".//kml:Data/kml:value/text()", namespaces=ns))
        search_area = (full_name + " " + desc_text + " " + ext_vals).strip().lower()

        # 4. منطق استخراج طول العمود (تعرف ذكي)
        val_height = ""
        # أولوية للأنواع الخاصة
        if any(kw in search_area for kw in ["هاي ماست", "highmast", "high mast"]):
            val_height = "هاي ماست"
        elif any(kw in search_area for kw in ["جداري", "wall"]):
            val_height = "جداري"
        else:
            # البحث عن "رقم" يتبعه م، متر، m، meter
            height_match = re.search(r'(\d{1,2})\s*(?:متر|م|m|meter)\b', search_area)
            if height_match:
                val_height = height_match.group(1)
            else:
                # خيار احتياطي للأرقام القياسية (12, 10, 8, 6)
                fallback = re.findall(r'\b(12|10|8|6)\b', search_area)
                val_height = fallback[0] if fallback else ""

        # 5. عدد الشمعات (الذراع)
        lamps = ""
        if val_height == "هاي ماست":
            lamps = 6
        elif any(kw in search_area for kw in ["دبل", "double", "2/2"]):
            lamps = 2
        elif any(kw in search_area for kw in ["مفرد", "single", "1/1"]):
            lamps = 1

        # 6. الإحداثيات
        coords = pm.xpath(".//kml:coordinates/text()", namespaces=ns)
        lat_val, lon_val = 0.0, 0.0
        if coords:
            coord_split = coords[0].strip().split(',')
            lat_val = float(coord_split[1])
            lon_val = float(coord_split[0])

        data.append({
            "المحطة": station_code,
            "رقم الفيدر": feeder_num,
            "رقم العمود": column_num,
            "طول العمود": val_height,
            "الذراع": lamps,
            "الاحداثيات x": lon_val,
            "الاحداثيات y": lat_val,
            "اسم الشارع": re.search(r'(?:شارع|Street)\s+([^,\n0-9]+)', search_area).group(1).strip() if re.search(r'(?:شارع|Street)\s+([^,\n0-9]+)', search_area) else ""
        })

    return pd.DataFrame(data)

if uploaded_files:
    all_dfs = [process_kmz(f) for f in uploaded_files]
    result_df = pd.concat(all_dfs, ignore_index=True)
    
    # ترتيب البيانات (المحطة ثم الفيدر ثم العمود)
    result_df = result_df.sort_values(by=['المحطة', 'رقم الفيدر', 'رقم العمود'])

    st.success(f"✅ تم استخراج {len(result_df)} عمود إنارة بنجاح.")
    
    # عرض إحصائية للمحطات
    st.write("### 📊 ملخص المحطات المكتشفة:")
    st.table(result_df['المحطة'].value_counts().reset_index().rename(columns={'index':'المحطة', 'المحطة':'عدد الأعمدة'}))

    st.dataframe(result_df, use_container_width=True)

    # تصدير التقرير
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        result_df.to_excel(writer, index=False, sheet_name='MainData')
        workbook = writer.book
        worksheet = writer.sheets['MainData']
        worksheet.right_to_left()

        # تنسيق المحطة بشكل بارز
        station_fmt = workbook.add_format({'bg_color': '#D7E4BC', 'bold': True, 'border': 1, 'align': 'center'})
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#95B3D7', 'border': 1, 'align': 'center'})
        
        for col_num, value in enumerate(result_df.columns.values):
            worksheet.write(0, col_num, value, header_fmt)
            width = 20 if value == "المحطة" else 15
            worksheet.set_column(col_num, col_num, width)

    st.download_button(
        label="📥 تحميل ملف الإكسيل المنسق",
        data=output.getvalue(),
        file_name="Lighting_Report_By_Station.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
