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

        # تحليل النمط (مثال: ق 17/5/1)
        numbers = re.findall(r'\d+', full_name)
        column_num = int(numbers[0]) if len(numbers) >= 1 else 0
        feeder_num = int(numbers[1]) if len(numbers) >= 2 else 0
        extra_num = numbers[2] if len(numbers) >= 3 else ""

        station_part = re.search(r'[a-zA-Z\u0600-\u06FF]+', full_name)
        station_code = station_part.group(0) if station_part else ""

        formatted_name = f"{column_num}/{feeder_num}"
        if extra_num:
            formatted_name += f"/{extra_num}"
        if station_code:
            formatted_name = f"{station_code} {formatted_name}"

        # الوصف والبيانات الممتدة
        desc = pm.xpath("./kml:description/text()", namespaces=ns)
        desc_text = desc[0] if desc else ""
        ext_vals = " ".join(pm.xpath(".//kml:Data/kml:value/text()", namespaces=ns))
        search_area = (desc_text + " " + ext_vals).strip()

        status = "مغروز" if "مغروز" in search_area else ("مفقود" if "مفقود" in search_area else "طبيعي")
        height_match = re.search(r'\b(12|10|9|8|6)\b', search_area)
        val_height = height_match.group(1) if height_match else "غير مسجل"
        lamps = 2 if "دبل" in search_area else (1 if "مفرد" in search_area else 0)

        # --- معالجة الإحداثيات لتطابق Map Marker ---
        coords = pm.xpath(".//kml:coordinates/text()", namespaces=ns)
        lat, lon = 0.0, 0.0
        if coords:
            coord_split = coords[0].strip().split(',')
            # في KML: الترتيب هو [Longitude, Latitude]
            # نحن نعكسهم ليظهر Lat أولاً كما في Map Marker
            lon_val = float(coord_split[0])
            lat_val = float(coord_split[1])
            
            # تقريب لـ 5 خانات عشرية كما في مثالك (24.69230)
            lat = "{:.5f}".format(lat_val)
            lon = "{:.5f}".format(lon_val)

        data.append({
            "الاسم المنسق": formatted_name,
            "المحطة": station_code,
            "الحالة": status,
            "طول العمود": val_height,
            "عدد الشمعات": lamps,
            "الإحداثيات (Lat, Long)": f"{lat},{lon}",
            "رقم الفيدر": feeder_num, # للترتيب
            "رقم العمود": column_num   # للترتيب
        })

    df = pd.DataFrame(data)
    df = df.sort_values(by=['المحطة', 'رقم الفيدر', 'رقم العمود'], ascending=[True, True, True])
    
    # إزالة أعمدة الترتيب قبل العرض والتصدير
    return df.drop(columns=['رقم الفيدر', 'رقم العمود'])

if uploaded_file:
    result_df = process_kmz(uploaded_file)
    st.write("### المعاينة النهائية:")
    st.dataframe(result_df)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        result_df.to_excel(writer, index=False)
    
    st.download_button("📥 تحميل التقرير النهائي", output.getvalue(), "Lighting_Final_Report.xlsx")
