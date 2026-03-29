
import streamlit as st
import zipfile
import pandas as pd
from lxml import etree
import re
import io

st.set_page_config(page_title="مستخرج بيانات المحطات", layout="wide")
st.title("📂 نظام استخراج وتصنيف محطات الإنارة")

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

        # 1. استخراج رقم المحطة (مثل ج557)
        station_match = re.search(r'([\u0600-\u06FFa-zA-Z]+\d+)', full_name)
        station_code = station_match.group(1) if station_match else "غير محدد"

        # 2. استخراج الأرقام الأخرى (عمود، فيدر)
        all_nums = re.findall(r'\d+', full_name)
        # استبعاد رقم المحطة من قائمة الأرقام إذا وجد
        numbers = [n for n in all_nums if n not in station_code]
        
        column_num = int(numbers[0]) if len(numbers) >= 1 else 0
        feeder_num = int(numbers[1]) if len(numbers) >= 2 else 0

        # 3. جلب النصوص للبحث عن الطول
        desc = pm.xpath("./kml:description/text()", namespaces=ns)
        desc_text = desc[0] if desc else ""
        ext_vals = " ".join(pm.xpath(".//kml:Data/kml:value/text()", namespaces=ns))
        search_area = (full_name + " " + desc_text + " " + ext_vals).strip().lower()

        # 4. استخراج طول العمود
        val_height = ""
        if any(kw in search_area for kw in ["هاي ماست", "highmast"]):
            val_height = "هاي ماست"
        elif any(kw in search_area for kw in ["جداري", "wall"]):
            val_height = "جداري"
        else:
            h_match = re.search(r'(\d{1,2})\s*(?:متر|م|m|meter)\b', search_area)
            val_height = h_match.group(1) if h_match else ""

        # 5. الإحداثيات
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
            "الاحداثيات y": lat_val,
            "اسم الملف": file.name
        })
    return pd.DataFrame(data)

if uploaded_files:
    all_dfs = [process_kmz(f) for f in uploaded_files]
    result_df = pd.concat(all_dfs, ignore_index=True)

    # ترتيب البيانات: المحطة أولاً، ثم الفيدر، ثم رقم العمود
    # هذا يضمن أن كل محطة تكون مجمعة مع أعمدتها بالترتيب
    result_df = result_df.sort_values(by=['المحطة', 'رقم الفيدر', 'رقم العمود'])

    # تحديد الصفوف المكررة (بناءً على المحطة، الفيدر، والعمود)
    is_duplicate = result_df.duplicated(subset=['المحطة', 'رقم الفيدر', 'رقم العمود'], keep=False)

    st.write("### معاينة البيانات المجمعة:")
    st.dataframe(result_df, use_container_width=True)

    # إنشاء ملف الإكسيل مع التنسيق المطلوب
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        result_df.to_excel(writer, index=False, sheet_name='MainData')
        workbook = writer.book
        worksheet = writer.sheets['MainData']
        worksheet.right_to_left()

        # التنسيقات
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center'})
        station_fmt = workbook.add_format({'bg_color': '#EBF1DE', 'border': 1, 'align': 'center'}) # لون للمحطة
        dup_fmt = workbook.add_format({'bg_color': '#BDD7EE', 'border': 1, 'align': 'center'}) # لون أزرق للمكرر
        default_fmt = workbook.add_format({'border': 1, 'align': 'center'})

        # كتابة العناوين
        for col_num, value in enumerate(result_df.columns.values):
            worksheet.write(0, col_num, value, header_fmt)
            worksheet.set_column(col_num, col_num, 15)

        # تلوين الصفوف بناءً على التكرار
        for row_idx in range(len(result_df)):
            row_num = row_idx + 1
            # إذا كان الصف مكرر، نستخدم التنسيق الأزرق
            current_fmt = dup_fmt if is_duplicate.iloc[row_idx] else default_fmt
            
            for col_idx in range(len(result_df.columns)):
                val = result_df.iloc[row_idx, col_idx]
                # تلوين خاص لعمود المحطة إذا لم يكن مكرراً لزيادة الوضوح
                if col_idx == 0 and not is_duplicate.iloc[row_idx]:
                    worksheet.write(row_num, col_idx, val, station_fmt)
                else:
                    worksheet.write(row_num, col_idx, val, current_fmt)

    st.success(f"✅ تم تجهيز التقرير. إجمالي السجلات: {len(result_df)}")
    st.download_button(
        label="📥 تحميل التقرير (المكرر باللون الأزرق)",
        data=output.getvalue(),
        file_name="Station_Report_Formatted.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
