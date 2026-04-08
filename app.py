import streamlit as st
import easyocr
import pandas as pd
import numpy as np
from PIL import Image
import re

# إعداد الصفحة
st.set_page_config(page_title="مستخرج بيانات الكهرباء", layout="wide")

st.markdown("""
    <style>
    .main { text-align: right; direction: rtl; }
    div.stButton > button { width: 100%; }
    </style>
    """, unsafe_allow_html=True)

st.title("⚡ مستخرج بيانات محطات الكهرباء")
st.info("ارفع صور التقارير أو ملفات PDF لاستخراج البيانات تلقائياً")

# تحميل محرك القراءة (العربية والإنجليزية)
@st.cache_resource
def load_reader():
    return easyocr.Reader(['ar', 'en'], gpu=False)

reader = load_reader()

uploaded_files = st.file_uploader("اختر الصور/الملفات", type=['png', 'jpg', 'jpeg', 'pdf'], accept_multiple_files=True)

def extract_data(text_list):
    full_text = " ".join(text_list)
    # قاموس لتخزين البيانات المستخرجة
    extracted = {}
    
    # تعريف الكلمات المفتاحية والبحث عن القيم التي تليها
    keys = {
        "رقم العداد": "العداد",
        "رقم الاشتراك": "الاشتراك",
        "رقم المحطة": "محطة",
        "جهد التغذية": "التغذية",
        "سعة القاطع": "القاطع",
        "سعة المحطة": "سعة المحطة",
        "قراءة الشركة": "قراءة شركة"
    }

    for label, search_key in keys.items():
        for i, word in enumerate(text_list):
            if search_key in word:
                # محاولة جلب الكلمة التالية (القيمة)
                if i + 1 < len(text_list):
                    extracted[label] = text_list[i+1]
                else:
                    extracted[label] = "غير موجود"
                break
        if label not in extracted:
            extracted[label] = "لم يتم العثور"
            
    return extracted

if uploaded_files:
    all_results = []
    for uploaded_file in uploaded_files:
        img = Image.open(uploaded_file)
        img_np = np.array(img)
        
        with st.spinner(f'جاري قراءة {uploaded_file.name}...'):
            # التعرف على النص
            text_results = reader.readtext(img_np, detail=0)
            data = extract_data(text_results)
            data['اسم الملف'] = uploaded_file.name
            all_results.append(data)

    # عرض النتائج في جدول
    df = pd.DataFrame(all_results)
    # إعادة ترتيب الأعمدة لتبدأ باسم الملف
    cols = ['اسم الملف'] + [c for c in df.columns if c != 'اسم الملف']
    df = df[cols]
    
    st.subheader("📊 البيانات المستخرجة")
    st.dataframe(df)

    # تصدير إلى Excel
    csv = df.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        label="📥 تحميل النتائج كملف Excel",
        data=csv,
        file_name="extracted_electricity_data.csv",
        mime="text/csv",
    )
