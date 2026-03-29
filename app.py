import streamlit as st
import zipfile
import pandas as pd
from lxml import etree
import re
import io

st.set_page_config(page_title="مستخرج بيانات شبكة الإنارة المطور", layout="wide")
st.title("📂 مستخرج بيانات KMZ الاحترافي")
st.info("سيقوم هذا الإصدار بدمج الملفات، منع التكرار، وتصنيف البيانات حسب المحطات.")

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

        # استخراج الأرقام (رقم العمود ورقم الفيدر)
        numbers = re.findall(r'\d+', full_name)
        column_num = int(numbers[0]) if len(numbers) >= 1 else 0
        feeder_num = int(numbers[1]) if len(numbers) >= 2 else 0
        
        # تحسين استخراج كود المحطة (يدعم الحروف العربية والإنجليزية والأرقام المدمجة)
        station_match = re.search(r'([a-zA-Z\u0600-\u06FF]+[\d]*)', full_name)
        station_code = station_match.group(0) if station_match else "غير محدد"

        desc = pm.xpath("./kml:description/text()", namespaces=ns)
        desc_text = desc[0] if desc else ""
        ext_vals = " ".join(pm.xpath(".//kml:Data/kml:value/text()", namespaces=ns))
        search_area = (full_name + " " + desc_text + " " + ext_vals).strip().lower()

        # استخراج اسم الشارع
        street_match = re.search(r'(?:شارع|Street)\s+([^,\n0-9]+)', search_area)
        street_name = street_match.group(1).strip() if street_match else ""

        # تحديد الحالة (مغروز، مفقود، طبيعي)
        observation = "طبيعي"
        details = ""
        if any(kw in search_area for kw in ["مغروز", "magrooz"]):
            observation = "مغروز"
            details = "مغروز"
        elif any(kw in search_area for kw in ["مفقود", "missing"]):
            observation = "مفقود"
            details = "مفقود"

        # تحديد الطول والنوع
        is_highmast = any(kw in search_area for kw in ["هاي ماست", "هايماست", "highmast"])
        is_wall = any(kw in search_area for kw in ["جداري", "wall"])

        if is_highmast:
            val_height, lamps = "هاي ماست", 6
        elif is_wall:
            val_height, lamps = "جداري", 1
        else:
            height_match = re.search(r'(\d{1,2})\s*(?:متر|م|m|meter)\b', search_area)
            val_height = height_match.group(1) if height_match else ""
            
            if any(kw in search_area for kw in ["2/2", "دبل", "double"]):
                lamps = 2
            elif any(kw in search_area for kw in ["1/1", "مفرد", "single"]):
                lamps = 1
            else:
                lamps = ""

        # الإحداثيات
        coords = pm.xpath(".//kml:coordinates/text()", namespaces=ns)
        lat_val, lon_val = 0.0, 0.0
        if coords:
            coord_split = coords[0].strip().split(',')
            lat_val = float(coord_split[1])
            lon_val = float(coord_split[0])

        data.append({
            "المحطة": station_code,
            "رقم العمود": column_num,
            "رقم الفيدر": feeder_num,
            "طول العمود": val_height,
            "الذراع": lamps,
            "الاحداثيات x": lon_val,
            "الاحداثيات y": lat_val,
            "اسم الشارع": street_name,
            "التفاصيل": details,
            "ملاحظة_داخلية": observation 
        })

    return pd.DataFrame(data)

if uploaded_files:
    all_dfs = [process_kmz(f) for f in uploaded_files]
    result_df = pd.concat(all_dfs, ignore_index=True)

    # --- معالجة التكرار وفلترة البيانات ---
    # 1. حذف التكرار التام (نفس المحطة، الفيدر، والعمود)
    initial_count = len(result_df)
    result_df = result_df.drop_duplicates(subset=['المحطة', 'رقم الفيدر', 'رقم العمود'], keep='first')
    removed_count = initial_count - len(result_df)

    # 2. الترتيب المنطقي
    result_df = result_df.sort_values(by=['المحطة', 'رقم الفيدر', 'رقم العمود'])

    # عرض ملخص المحطات في Streamlit
    st.success(f"✅ تم بنجاح! تم العثور على {len(result_df['المحطة'].unique())} محطة مختلفة.")
    if removed_count > 0:
        st.warning(f"⚠️ تم حذف {removed_count} سجل مكرر تلقائياً.")

    # عرض إحصائيات سريعة
    with st.expander("📊 عرض إحصائيات المحطات"):
        stats = result_df.groupby('المحطة').size().reset_index(name='عدد الأعمدة')
        st.table(stats)

    st.dataframe(result_df.drop(columns=['ملاحظة_داخلية']), use_container_width=True)

    # تجهيز ملف الإكسيل مع التنسيق
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        export_df = result_df.drop(columns=['ملاحظة_داخلية'])
        export_df.to_excel(writer, index=False, sheet_name='Data')
        
        workbook = writer.book
        worksheet = writer.sheets['Data']
        worksheet.right_to_left()

        # تعريف التنسيقات
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center'})
        station_fmt = workbook.add_format({'bg_color': '#F2F2F2', 'border': 1, 'align': 'center'})
        red_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1, 'align': 'center'})
        
        # تطبيق التنسيق على العناوين
        for col_num, value in enumerate(export_df.columns.values):
            worksheet.write(0, col_num, value, header_fmt)
            worksheet.set_column(col_num, col_num, 15)

        # تطبيق التنسيق الشرطي والصفوف
        for row_idx in range(len(result_df)):
            row_num = row_idx + 1
            obs = result_df.iloc[row_idx]['ملاحظة_داخلية']
            
            # إذا كان العمود مفقود أو مغروز، يلون الصف باللون الأحمر
            current_fmt = red_fmt if obs in ["مغروز", "مفقود"] else None
            
            for col_idx in range(len(export_df.columns)):
                val = export_df.iloc[row_idx, col_idx]
                # تلوين عمود المحطة بلون مميز دائماً لسهولة القراءة
                if col_idx == 0 and not current_fmt:
                    worksheet.write(row_num, col_idx, val, station_fmt)
                else:
                    worksheet.write(row_num, col_idx, val, current_fmt)

    st.download_button(
        label="📥 تحميل التقرير النهائي (Excel)",
        data=output.getvalue(),
        file_name="Lighting_Network_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
