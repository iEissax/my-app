import streamlit as st
import easyocr
import pandas as pd
import numpy as np
from PIL import Image
from pdf2image import convert_from_bytes

st.set_page_config(page_title="مستخرج بيانات الكهرباء", layout="wide")

st.title("⚡ مستخرج بيانات محطات الكهرباء")

@st.cache_resource
def load_reader():
    return easyocr.Reader(['ar', 'en'], gpu=False)

reader = load_reader()

uploaded_files = st.file_uploader("اختر الصور أو ملفات PDF", type=['png', 'jpg', 'jpeg', 'pdf'], accept_multiple_files=True)

def process_image(img_input):
    """دالة لمعالجة الصورة واستخراج النصوص"""
    img_np = np.array(img_input)
    text_results = reader.readtext(img_np, detail=0)
    
    keys = {
        "رقم العداد": "العداد",
        "رقم الاشتراك": "الاشتراك",
        "رقم المحطة": "محطة",
        "جهد التغذية": "التغذية",
        "سعة القاطع": "القاطع",
        "سعة المحطة": "سعة المحطة",
        "قراءة الشركة": "قراءة شركة"
    }
    
    extracted = {}
    for label, search_key in keys.items():
        for i, word in enumerate(text_results):
            if search_key in word:
                if i + 1 < len(text_results):
                    extracted[label] = text_results[i+1]
                break
        if label not in extracted:
            extracted[label] = "غير موجود"
    return extracted

if uploaded_files:
    all_results = []
    for uploaded_file in uploaded_files:
        with st.spinner(f'جاري معالجة {uploaded_file.name}...'):
            # إذا كان الملف PDF
            if uploaded_file.type == "application/pdf":
                images = convert_from_bytes(uploaded_file.read())
                for i, page_img in enumerate(images):
                    data = process_image(page_img)
                    data['اسم الملف'] = f"{uploaded_file.name} (صفحة {i+1})"
                    all_results.append(data)
            # إذا كان ملف صورة
            else:
                img = Image.open(uploaded_file)
                data = process_image(img)
                data['اسم الملف'] = uploaded_file.name
                all_results.append(data)

    if all_results:
        df = pd.DataFrame(all_results)
        st.dataframe(df)
        csv = df.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 تحميل النتائج Excel", csv, "data.csv", "text/csv")
