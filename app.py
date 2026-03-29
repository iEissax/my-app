import streamlit as st
import zipfile
import pandas as pd
from lxml import etree
import re
import io

def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', str(s))]

st.set_page_config(page_title="نظام معالجة المحطات", layout="wide")
st.title("📂 مستخرج البيانات - تحديث منطق (الطول والأذرعة فقط)")

uploaded_files = st.file_uploader("اختر ملفات KMZ", type=['kmz'], accept_multiple_files=True)

def parse_pole_data(text, station_code):
    """تفكيك النمط 12/2/1: الأول طول، الثاني أذرعة، والثالث يتجاهل"""
    clean_text = text.replace(station_code, "").strip()
    nums = re.findall(r'\d+', clean_text)
    
    height, arms, col_num = "", "", ""
    
    # تنفيذ قاعدتك: 12/2/1 -> 12 طول، 2 ذراع
    if len(nums) >= 2:
        height = nums[0]   # الرقم الأول = الطول
        arms = nums[1]     # الرقم الثاني = الأذرعة
    
    # البحث عن رقم العمود بعيداً عن النمط (إذا وجد رقم وحيد في مكان آخر)
    # أو سيبقى فارغاً كما في طلبك الأخير
    if "هاي" in text or "mast" in text.lower():
        height, arms = "هاي ماست", 6
    elif "جداري" in text:
        height, arms = "جداري", 1
        
    return height, arms

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
        
        st_match = re.search(r'(\d+[\u0600-\u06FF]+|[\u0600-\u06FF]+\d+)', full_name)
        station_code = st_match.group(1) if st_match else "غير محدد"

        desc = "".join(pm.xpath("./kml:description/text()", namespaces=ns))
        ext_data = " ".join(pm.xpath(".//kml:value/text()", namespaces=ns))
        all_info = (full_name + " " + desc + " " + ext_data)

        # تطبيق المنطق: استخراج طول وذراع فقط
        h, a = parse_pole_data(full_name, station_code)
        
        # محاولة استخراج رقم العمود إذا كان موجوداً كرقم مستقل (خارج النمط المائل)
        # إذا كان طلبك أن لا نسجل أي رقم عمود من النمط، سيبقى هذا الحقل فارغاً
        col_num = "" 

        coords = pm.xpath(".//kml:coordinates/text()", namespaces=ns)
        lat, lon = 0.0, 0.0
        if coords:
            c_split = coords[0].strip().split(',')
            lat, lon = round(float(c_split[1]), 5), round(float(c_split[0]), 5)

        detail = "مفقود" if "مفقود" in all_info else ("مغروز" if "مغروز" in all_info else "")

        data.append({
            "المحطة": station_code,
            "رقم العمود": col_num,
            "طول العمود": h,
            "الذراع": a,
            "الاحداثيات x": lon,
            "الاحداثيات y": lat,
            "التفاصيل": detail
        })
    return pd.DataFrame(data)

if uploaded_files:
    all_dfs = [process_kmz(f) for f in uploaded_files]
    df = pd.concat(all_dfs, ignore_index=True)

    # الترتيب حسب المحطة
    df = df.sort_values(by=['المحطة'], key=lambda x: x.map(natural_sort_key))

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('Report')
        worksheet.right_to_left()

        # التنسيقات (خط أسود)
        f_head = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center', 'font_color': 'black'})
        f_stat = workbook.add_format({'bg_color': '#7F7F7F', 'border': 1, 'align': 'center', 'font_color': 'black', 'bold': True})
        f_red = workbook.add_format({'bg_color': '#FF0000', 'border': 1, 'align': 'center', 'font_color': 'black'})
        f_norm = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'black'})
        f_coord = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'black', 'num_format': '0.00000'})

        cols = ["المحطة", "رقم العمود", "طول العمود", "الذراع", "الاحداثيات x", "الاحداثيات y", "التفاصيل"]
        for i, col in enumerate(cols):
            worksheet.write(0, i, col, f_head)
            worksheet.set_column(i, i, 15)

        curr_row, last_st = 1, None
        for _, row in df.iterrows():
            if last_st and row['المحطة'] != last_st:
                curr_row += 1 

            is_red = row['التفاصيل'] in ["مفقود", "مغروز"]
            for j, c_name in enumerate(cols):
                val = row[c_name]
                if is_red: f = f_red
                elif c_name == "المحطة": f = f_stat
                elif "الاحداثيات" in c_name: f = f_coord
                else: f = f_norm
                worksheet.write(curr_row, j, val, f)
            last_st, curr_row = row['المحطة'], curr_row + 1

    st.success("✅ تم التعديل: لن يتم تسجيل الرقم الثالث كـ 'رقم عمود'.")
    st.download_button("📥 تحميل الملف المعدل", output.getvalue(), "Lighting_Report_NoColNum.xlsx")
