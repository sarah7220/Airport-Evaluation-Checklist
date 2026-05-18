import streamlit as st
import base64
import os

from type1_RUH import run_type_1_RUH
from type1_JED import run_type_1_JED
from type2 import run_type_2
from type3 import run_type_3
from type4 import run_type_4  
from type5 import run_type_5

# ─────────────────────────────────────────────
# إعدادات الصفحة
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="تحليل بيانات المطارات",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─────────────────────────────────────────────
# تهيئة الذاكرة (Session State)
# ─────────────────────────────────────────────
if 'city' not in st.session_state: st.session_state.city = None
if 'airport' not in st.session_state: st.session_state.airport = None

# ─────────────────────────────────────────────
# شعار 
# ─────────────────────────────────────────────
def get_base64_image(name):
    for ext in ['.jpg', '.jpeg', '.png']:
        path = name + ext
        if os.path.exists(path):
            try:
                with open(path, "rb") as f:
                    return base64.b64encode(f.read()).decode()
            except:
                continue
    return None

logo_b64 = get_base64_image("logo")

# ─────────────────────────────────────────────
# CSS (التصميم المعدل لحل مشكلة تداخل النصوص والأيقونات)
# ─────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Arabic:wght@300;400;600;700&display=swap');

    /* 1. تطبيق الخط العربي بشكل عام لكن بدون فرضه على الأيقونات */
    body, p, h1, h2, h3, h4, h5, h6, div, span, label, button, input, select {
        font-family: 'IBM Plex Sans Arabic', sans-serif;
    }

    /* 2. حماية أيقونات Streamlit (مثل الأسهم) لكي لا تتحول إلى نص */
    span[class*="material-symbols"], span[class*="icon"], i {
        font-family: 'Material Symbols Rounded', 'Material Icons', sans-serif !important;
        direction: ltr !important;
    }

    /* 3. تطبيق الاتجاه العربي (RTL) بشكل آمن على الحاويات الرئيسية فقط */
    [data-testid="stAppViewContainer"], [data-testid="stSidebar"] {
        direction: rtl;
    }

    /* 4. محاذاة النصوص لليمين */
    p, h1, h2, h3, h4, h5, h6, label {
        text-align: right !important;
    }

    /* --- باقي تنسيقاتك الجميلة بدون تغيير --- */
    .stApp, .main {
        background-color: #EEF2F7;
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #071120 0%, #0d1f3c 100%) !important;
    }

    [data-testid="stSidebar"] > div:first-child {
        padding-top: 0 !important;
    }

    .logo-wrap {
        width: 100%;
        background: #fff;
        overflow: hidden;
        line-height: 0;
    }

    .logo-wrap img {
        width: 100%;
        display: block;
        object-fit: cover;
    }

    .sidebar-header {
        padding: 1.2rem 1rem;
        border-bottom: 1px solid rgba(255,255,255,0.07);
        text-align: center;
    }

    .sidebar-header h2 {
        margin: 0 0 3px !important;
        color: #ffffff !important;
        font-size: 1.15rem;
        font-weight: 700;
    }

    .sidebar-header p {
        margin: 0 !important;
        color: #6b8cba !important;
        font-size: 0.78rem;
        font-weight: 300;
    }

    .nav-label {
        color: #4f6f9a;
        font-size: 0.68rem;
        font-weight: 600;
        letter-spacing: 1.8px;
        text-transform: uppercase;
        padding: 1.1rem 1rem 0.4rem;
        display: block;
    }

    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] p {
        color: #c2d0e3 !important;
    }

    [data-testid="stSidebar"] [data-baseweb="radio"] {
        background: transparent;
        padding: 9px 12px;
        border-radius: 10px;
        margin-bottom: 3px;
        transition: background 0.2s;
    }

    [data-testid="stSidebar"] [data-baseweb="radio"]:hover {
        background: rgba(255,255,255,0.06);
    }

    .page-header {
        background: linear-gradient(130deg, #071120 0%, #122a56 55%, #183470 100%);
        border-radius: 20px;
        padding: 40px 48px;
        margin-bottom: 28px;
        position: relative;
        overflow: hidden;
        display: flex;
        align-items: center;
        justify-content: space-between;
        box-shadow: 0 10px 36px rgba(7, 17, 32, 0.2);
    }

    .header-text h1 {
        color: #ffffff !important;
        font-size: 1.9rem !important;
        font-weight: 700 !important;
        margin: 0 0 6px !important;
        text-align: right !important;
    }

    .header-text p {
        color: #7ca3cc;
        font-size: 0.9rem;
        margin: 0;
        font-weight: 300;
    }

    .header-plane {
        font-size: 5rem;
        opacity: 0.12;
        position: relative;
        z-index: 1;
    }

    .section-bar {
        display: flex;
        align-items: center;
        gap: 12px;
        margin-bottom: 16px;
        direction: rtl !important;
    }

    .section-bar h2 {
        color: #071120;
        font-size: 1.15rem;
        font-weight: 700;
        margin: 0;
    }

    .section-badge {
        background: linear-gradient(135deg, #122a56, #1d4ed8);
        color: #fff;
        padding: 3px 13px;
        border-radius: 20px;
        font-size: 0.74rem;
        font-weight: 600;
    }

    hr {
        border: none;
        border-top: 1px solid #d4dce9;
        margin: 0 0 22px;
    }

    .stButton > button {
        background: linear-gradient(135deg, #122a56, #1d4ed8);
        color: white;
        border-radius: 10px;
        padding: 0.5rem 1.8rem;
        font-weight: 600;
    }

    #MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# Sidebar
# ─────────────────────────────────────────────
PAGES = [
    "Type 1 - RUH",
    "Type 1 - JED",
    "Type 2",
    "Type 3",
    "Type 4",
    "Type 5",
]

with st.sidebar:
    if logo_b64:
        st.markdown(f'<div class="logo-wrap"><img src="data:image/jpeg;base64,{logo_b64}"></div>', unsafe_allow_html=True)

    st.markdown("""
    <div class="sidebar-header">
        <h2>تحليل بيانات المطارات</h2>
        <p>لوحة التحكم الرئيسية</p>
    </div>
    <span class="nav-label">اختر فئة المطار</span>
    """, unsafe_allow_html=True)

    selection = st.radio(
        label="",
        options=PAGES,
        index=5,
        label_visibility="collapsed"
    )

# ─────────────────────────────────────────────
# Header
# ─────────────────────────────────────────────
st.markdown("""
<div class="page-header">
    <div class="header-text">
        <h1>نظام أتمتة تقييم جودة المطارات</h1>
        <p>الهيئة العامة للطيران المدني — قطاع الجودة و تجربة العميل</p>
    </div>
    <div class="header-plane">✈</div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# Section title
# ─────────────────────────────────────────────
st.markdown(f"""
<div class="section-bar">
    <h2>📋 {selection}</h2>
    <span class="section-badge">تحليل البيانات</span>
</div>
<hr>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# المحتوى (Logics)
# ─────────────────────────────────────────────
if selection == "Type 1 - RUH":
    run_type_1_RUH()
elif selection == "Type 1 - JED":
    run_type_1_JED()
elif selection == "Type 2":
    run_type_2()
elif selection == "Type 3":
    run_type_3()
elif selection == "Type 4":
    run_type_4()
elif selection == "Type 5":
    run_type_5()