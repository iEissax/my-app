import streamlit as st
import zipfile
import pandas as pd
from lxml import etree
import re
import io

st.set_page_config(page_title="مستخرج بيانات شبكة الإنارة", layout="centered")
st.title("📂 مستخرج بيانات KMZ")

uploaded_file = st.file_uploader("اختر ملف KMZ", type=['kmz'])

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

        # --- تحليل النمط (مثال: ق 17/5/1) ---
        # استخراج الأرقام: 17 هو الأول (عمود)، 5 هو الثاني (فيدر)
        numbers = re.findall(r'\d+', full_name)
        
        # التخصيص حسب طلبك:
        column_num = int(numbers[0]) if len(numbers) >= 1 else 0  # الرقم الأول: العمود
        feeder_num = int(numbers[1]) if len(numbers) >= 2 else 0  # الرقم الثاني: الفيدر
        extra_num = numbers[2] if len(numbers) >= 3 else ""       # أي رقم ثالث إضافي

        # استخراج اسم المحطة (مثل حرف ق)
        station_part = re.search(r'[a-zA-Z\u0600-\u06FF]+', full_name)
        station_code = station_part.group(0) if station_part else ""

        # تنسيق الاسم النهائي في الإكسل (العمود/الفيدر)
        formatted_name = f"{column_num}/{feeder_num}"
        if extra_num:
            formatted_name += f"/{extra_num}"
        if station_code:
            formatted_name = f"{station_code} {formatted_name}"

        # البحث في الوصف والبيانات الممتدة
        desc = pm.xpath("./kml:description/text()", namespaces=ns)
        desc_text = desc[0] if desc else ""
        ext_vals = " ".join(pm.xpath(".//kml:Data/kml:value/text()", namespaces=ns))
        search_area = (desc_text + " " + ext_vals).strip()

        # استخراج الحالة والأطوال والشمعات
        status = "مغروز" if "مغروز" in search_area else ("مفقود" if "مفقود" in search_area else "طبيعي")
        height_match = re.search(r'\b(12|10|9|8|6)\b', search_area)
        val_height = height_match.group(1) if height_match else "غير مسجل"
        lamps = 2 if "دبل" in search_area else (1 if "مفرد" in search_area else 0)

        # الإحداثيات
        coords = pm.xpath(".//kml:coordinates/text()", namespaces=ns)
        lat, lon = (coords[0].split(',')[1], coords[0].split(',')[0]) if coords else (0,0)

        data.append({
            "الاسم المنسق": formatted_name,
            "المحطة": station_code,
            "رقم الفيدر": feeder_num,
            "رقم العمود": column_num,
            "الحالة": status,
            "طول العمود": val_height,
            "عدد الشمعات": lamps,
            "Lat": lat,
            "Long": lon
        })

    df = pd.DataFrame(data)
    
    # الترتيب الصحيح: المحطة أولاً، ثم رقم الفيدر، ثم تسلسل الأعمدة داخل الفيدر
    df = df.sort_values(by=['المحطة', 'رقم الفيدر', 'رقم العمود'], ascending=[True, True, True])
    
    return df.drop(columns=['رقم الفيدر', 'رقم العمود'])

if uploaded_file:
    result_df = process_kmz(uploaded_file)
    st.write("### معاينة البيانات المرتّبة (فيدر ثم عمود):")
    st.dataframe(result_df)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        result_df.to_excel(writer, index=False)
    
    st.download_button("📥 تحميل ملف Excel المنسق", output.getvalue(), "Lighting_Report.xlsx")
