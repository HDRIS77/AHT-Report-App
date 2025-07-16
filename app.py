import pandas as pd
import streamlit as st
from io import BytesIO

# إعداد صفحة Streamlit
st.set_page_config(page_title="نظام تحليل AHT", layout="wide")
st.title("📊 نظام تحليل وقت التعامل (AHT)")

# قسم رفع الملفات - في عمود جانبي
with st.sidebar:
    st.header("📤 رفع الملفات المطلوبة")
    
    # رفع ملفات Raw Data
    raw_overall_file = st.file_uploader("ملف البيانات الرئيسي (Raw Overall)", type=['xlsx', 'csv'])
    raw_urdu_file = st.file_uploader("ملف بيانات الأردية (Raw Urdu)", type=['xlsx', 'csv'])
    
    # رفع ملف HC
    hc_file = st.file_uploader("ملف الموظفين (HC)", type=['xlsx', 'csv'])
    
    # إعدادات AHT
    st.header("⚙️ إعدادات AHT")
    target_aht = st.number_input("هدف AHT (بالثواني)", min_value=0, value=300, step=1)
    min_aht = st.number_input("الحد الأدنى لـ AHT", min_value=0, value=180, step=1)
    max_aht = st.number_input("الحد الأقصى لـ AHT", min_value=0, value=600, step=1)

# قسم التحليل الرئيسي
if raw_overall_file and raw_urdu_file:
    try:
        # قراءة الملفات
        df_overall = pd.read_excel(raw_overall_file) if raw_overall_file.name.endswith('.xlsx') else pd.read_csv(raw_overall_file)
        df_urdu = pd.read_excel(raw_urdu_file) if raw_urdu_file.name.endswith('.xlsx') else pd.read_csv(raw_urdu_file)
        
        # معالجة ملف HC إذا موجود
        df_hc = None
        if hc_file:
            df_hc = pd.read_excel(hc_file) if hc_file.name.endswith('.xlsx') else pd.read_csv(hc_file)
        
        # حساب AHT (مثال - يمكن تعديله حسب هيكل بياناتك)
        if 'Handle_Time' in df_overall.columns:
            # حساب AHT لكل موظف/فريق
            aht_results = df_overall.groupby('Agent_Name')['Handle_Time'].mean().reset_index()
            aht_results = aht_results.rename(columns={'Handle_Time': 'AHT'})
            
            # تصفية حسب الحدود المدخلة
            aht_results = aht_results[(aht_results['AHT'] >= min_aht) & (aht_results['AHT'] <= max_aht)]
            
            # عرض النتائج
            st.header("📈 نتائج تحليل AHT")
            
            # عرض في أعمدة
            col1, col2, col3 = st.columns(3)
            
            with col1:
                avg_aht = aht_results['AHT'].mean()
                st.metric("متوسط AHT", f"{avg_aht:.2f} ثانية", delta=f"{avg_aht-target_aht:.2f} vs الهدف")
            
            with col2:
                min_agent = aht_results.loc[aht_results['AHT'].idxmin()]
                st.metric("أفضل أداء", f"{min_agent['AHT']:.2f} ث", min_agent['Agent_Name'])
            
            with col3:
                max_agent = aht_results.loc[aht_results['AHT'].idxmax()]
                st.metric("أضعف أداء", f"{max_agent['AHT']:.2f} ث", max_agent['Agent_Name'])
            
            # عرض جدول البيانات
            st.dataframe(aht_results.sort_values('AHT'), height=400)
            
            # تحميل النتائج
            st.download_button(
                label="⬇️ تنزيل نتائج AHT",
                data=aht_results.to_csv(index=False),
                file_name="aht_results.csv",
                mime="text/csv"
            )
        else:
            st.error("الملف الرئيسي لا يحتوي على عمود Handle_Time لحساب AHT")
    
    except Exception as e:
        st.error(f"حدث خطأ في معالجة البيانات: {str(e)}")
else:
    st.warning("الرجاء رفع الملفات المطلوبة لبدء التحليل")
    st.info("""
    **تعليمات رفع الملفات:**
    1. اختر ملف البيانات الرئيسي (Raw Overall)
    2. اختر ملف بيانات الأردية (Raw Urdu)
    3. (اختياري) اختر ملف الموظفين (HC)
    4. اضبط إعدادات AHT حسب احتياجاتك
    """)
