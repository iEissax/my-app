import streamlit as st
import zipfile
import pandas as pd
from lxml import etree
import re
import io

# إعدادات واجهة الموقع
st.set_page_config(page_title="مستخرج بيانات شبكة الإنارة", layout="centered")

st.title("📂 مستخرج بيانات KMZ لشبكات الإنارة")
st.write("قم برفع ملف الـ KMZ المستخرج من Map Marker لتحويله إلى ملف Excel منسق.")

# خانة رفع الملف
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

        # استخراج المحطة
        station_match = re.search(r'\((.*?)\)', full_name)
        station_code = station_match.group(1) if station_match else ""
        
        # تفكيك الأرقام
        clean_name = re.sub(r'\(.*?\)', '', full_name)
        parts = re.split(r'[/|-]', clean_name)
        
        try:
            raw_f = int(re.search(r'\d+', parts[0]).group()) if len(parts) > 0 else 0
            raw_c = int(re.search(r'\d+', parts[1]).group()) if len(parts) > 1 else 0
        except:
            raw_f, raw_c = 0, 0

        # التنسيق المطلوب (العمود/الفيدر)
        formatted_name = f"{raw_c}/{raw_f}"

        # البحث في الوصف والبيانات الممتدة
        desc = pm.xpath("./kml:description/text()", namespaces=ns)
        desc_text = desc[0] if desc else ""
        extended_values = pm.xpath(".//kml:Data/kml:value/text()", namespaces=ns)
        full_search_text = (desc_text + " " + " ".join(extended_values)).strip()

        # استخراج الطول (6, 8, 9, 10, 12)
        height_match = re.search(r'\b(12|10|9|8|6)\b', full_search_text)
        val_height = height_match.group(1) if height_match else "غير مسجل"

        # عدد الشمعات
        val_lamps = 2 if "دبل" in full_search_text else (1 if "مفرد" in full_search_text else 0)
        if val_lamps == 0:
            lamp_num = re.search(r'(\d+)\s*(?:شمعة|كشاف)', full_search_text)
            val_lamps = int(lamp_num.group(1)) if lamp_num else 0

        # الحالة
        status = "مغروز" if "مغروز" in full_search_text else ("مفقود" if "مفقود" in full_search_text else "طبيعي")

        # الإحداثيات
        coords = pm.xpath(".//kml:coordinates/text()", namespaces=ns)
        lat, lon = 0, 0
        if coords:
            c = coords[0].strip().split(',')
            lon, lat = round(float(c[0]), 6), round(float(c[1]), 6)

        data.append({
            "تنسيق (العمود/الفيدر)": formatted_name,
            "المحطة": station_code,
            "رقم الفيدر": raw_f,
            "رقم العمود": raw_c,
            "الحالة": status,
            "طول العمود": val_height,
            "عدد الشمعات": val_lamps,
            "خط العرض": lat,
            "خط الطول": lon
        })

    df = pd.DataFrame(data)
    df = df.sort_values(by=['المحطة', 'رقم الفيدر', 'رقم العمود'])
    df_final = df.drop(columns=['رقم الفيدر', 'رقم العمود'])
    return df_final

if uploaded_file is not None:
    with st.spinner('جاري معالجة الملف...'):
        result_df = process_kmz(uploaded_file)
        
        st.success(f"تمت المعالجة بنجاح! تم العثور على {len(result_df)} نقطة.")
        
        # عرض عينة من البيانات
        st.dataframe(result_df.head(10))

        # تحويل البيانات إلى Excel للتحميل
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            result_df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        st.download_button(
            label="📥 تحميل ملف Excel المنسق",
            data=output.getvalue(),
            file_name="Electrical_Network_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )