import streamlit as st
import zipfile
import pandas as pd
from lxml import etree
import re
import io

# دالة لضمان الترتيب الطبيعي (1, 2, 10 بدلاً من 1, 10, 2)
def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', str(s))]

st.set_page_config(page_title="مستخرج بيانات المحطات المطور", layout="wide")
st.title("📂 مستخرج بيانات KMZ - الإصدار المهني المعتمد")

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

        # 1. التعرف على رقم المحطة (مثل 904ج أو ج557)
        station_match = re.search(r'(\d+[\u0600-\u06FF]+|[\u0600-\u06FF]+\d+)', full_name)
        station_code = station_match.group(1) if station_match else "غير محدد"

        # 2. جلب النصوص للبحث عن البيانات والحالة
        desc = "".join(pm.xpath("./kml:description/text()", namespaces=ns))
        ext_vals = " ".join(pm.xpath(".//kml:value/text()", namespaces=ns))
        search_area = (full_name + " " + desc + " " + ext_vals).strip()

        # 3. منطق استخراج (الطول / الذراع / رقم العمود) من النمط 12/2/1
        # تنظيف النص من اسم المحطة لاستخراج الأرقام فقط
        clean_text = full_name.replace(station_code, "").strip()
        nums = re.findall(r'\d+', clean_text)
        
        val_height, val_arms, column_num = "", "", ""
        
        if len(nums) >= 3:
            val_height = nums[0]  # الرقم الأول: طول العمود (مثلاً 12)
            val_arms = nums[1]    # الرقم الثاني: الذراع (مثلاً 2)
            column_num = nums[2]  # الرقم الثالث: رقم العمود (مثلاً 1)
        elif len(nums) == 2:
            val_height = nums[0]
            column_num = nums[1]
        elif len(nums) == 1:
            column_num = nums[0]

        # معالجة الكلمات الخاصة إذا لم تكن الأرقام موجودة
        if not val_height:
            if "هاي" in search_area.lower() or "mast" in search_area.lower():
                val_height, val_arms = "هاي ماست", 6
            elif "جداري" in search_area.lower():
                val_height, val_arms = "جداري", 1
            else:
                h_match = re.search(r'\b(12|10|8|6|5)\b', search_area)
                if h_match: val_height = h_match.group(1)

        # 4. تحديد التفاصيل (مفقود/مغروز) لغرض التلوين
        details = ""
        if "مفقود" in search_area: details = "مفقود"
        elif "مغروز" in search_area: details = "مغروز"

        # 5. الإحداثيات (دقة خماسية)
        coords = pm.xpath(".//kml:coordinates/text()", namespaces=ns)
        lat_val, lon_val = 0.0, 0.0
        if coords:
            c_split = coords[0].strip().split(',')
            lat_val = round(float(c_split[1]), 5)
            lon_val = round(float(c_split[0]), 5)

        data.append({
            "المحطة": station_code,
            "رقم العمود": column_num,
            "طول العمود": val_height,
            "الذراع": val_arms,
            "الاحداثيات x": lon_val,
            "الاحداثيات y": lat_val,
            "التفاصيل": details
        })
    return pd.DataFrame(data)

if uploaded_files:
    all_dfs = [process_kmz(f) for f in uploaded_files]
    df = pd.concat(all_dfs, ignore_index=True)
    
    # تحويل رقم العمود إلى أرقام للفرز الصحيح
    df['رقم العمود'] = pd.to_numeric(df['رقم العمود'], errors='coerce').fillna(0).astype(int)
    
    # الترتيب حسب المحطة (طبيعي) ثم رقم العمود
    df = df.sort_values(
        by=['المحطة', 'رقم العمود'], 
        key=lambda x: x.map(natural_sort_key) if x.name == 'المحطة' else x
    )
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('Data')
        worksheet.right_to_left()

        # تعريف التنسيقات (الخط أسود دائماً)
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center', 'font_color': 'black'})
        station_fmt = workbook.add_format({'bg_color': '#7F7F7F', 'border': 1, 'align': 'center', 'font_color': 'black', 'bold': True})
        red_fmt = workbook.add_format({'bg_color': '#FF0000', 'border': 1, 'align': 'center', 'font_color': 'black'})
        normal_fmt = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'black'})
        coord_fmt = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'black', 'num_format': '0.00000'})

        # كتابة العناوين
        cols = ["المحطة", "رقم العمود", "طول العمود", "الذراع", "الاحداثيات x", "الاحداثيات y", "التفاصيل"]
        for col_num, col_name in enumerate(cols):
            worksheet.write(0, col_num, col_name, header_fmt)
            worksheet.set_column(col_num, col_num, 15)

        curr_row = 1
        last_st = None
        
        for idx, row in df.iterrows():
            # إضافة صف فارغ عند تغيير المحطة
            if last_st is not None and row['المحطة'] != last_st:
                curr_row += 1 

            row_is_red = str(row['التفاصيل']) in ["مفقود", "مغروز"]

            for col_idx, col_name in enumerate(cols):
                val = row[col_name]
                
                if row_is_red: fmt = red_fmt
                elif col_name == "المحطة": fmt = station_fmt
                elif "الاحداثيات" in col_name: fmt = coord_fmt
                else: fmt = normal_fmt
                
                worksheet.write(curr_row, col_idx, val, fmt)
            
            last_st = row['المحطة']
            curr_row += 1

    st.success("✅ تم استخراج البيانات وتنسيقها وفقاً للنمط المطلوب.")
    st.download_button(label="📥 تحميل الملف النهائي", data=output.getvalue(), file_name="Station_Standard_Report.xlsx")
