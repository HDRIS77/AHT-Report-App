import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
import streamlit as st
from io import BytesIO

# ... (بقية الاستيرادات والثوابت)

def validate_hc_file(df_hc):
    """التحقق من وجود الأعمدة المطلوبة في ملف HC"""
    required_columns = ['Employee_ID', 'Team_Leader']
    missing_columns = [col for col in required_columns if col not in df_hc.columns]
    
    if missing_columns:
        st.error(f"ملف HC ينقصه الأعمدة التالية: {', '.join(missing_columns)}")
        st.info("يجب أن يحتوي ملف HC على الأعمدة التالية على الأقل: Employee_ID و Team_Leader")
        return False
    return True

def main():
    st.title("نظام تحليل أداء خدمة العملاء المتكامل")
    
    # رفع الملفات
    st.subheader("رفع ملفات البيانات المطلوبة")
    
    # إنشاء أعمدة لعرض حقول الرفع بشكل أنيق
    col1, col2, col3 = st.columns(3)
    
    with col1:
        raw_overall_file = st.file_uploader("البيانات الكلية (Raw Overall)", type=['xlsx', 'csv'])
    
    with col2:
        raw_urdu_file = st.file_uploader("بيانات الأردية (Raw Urdu)", type=['xlsx', 'csv'])
    
    with col3:
        hc_file = st.file_uploader("ملف الموظفين (HC)", 
                                 type=['xlsx', 'csv'],
                                 help="يجب أن يحتوي على عمود Employee_ID (رقم الموظف) وعمود Team_Leader (اسم المشرف)")

    target_aht = st.number_input("هدف وقت التعامل (AHT) بالثواني", value=300, min_value=0)

    if st.button("إنشاء التقرير"):
        if not raw_overall_file or not raw_urdu_file:
            st.warning("الرجاء رفع ملفات البيانات الأساسية (البيانات الكلية وبيانات الأردية)")
            return

        try:
            # قراءة الملفات الأساسية
            df_raw_overall = pd.read_excel(raw_overall_file) if raw_overall_file.name.endswith('.xlsx') else pd.read_csv(raw_overall_file)
            df_raw_urdu = pd.read_excel(raw_urdu_file) if raw_urdu_file.name.endswith('.xlsx') else pd.read_csv(raw_urdu_file)

            # معالجة ملف HC إذا تم رفعه
            df_hc = None
            if hc_file:
                try:
                    df_hc = pd.read_excel(hc_file) if hc_file.name.endswith('.xlsx') else pd.read_csv(hc_file)
                    
                    if not validate_hc_file(df_hc):
                        return
                        
                    # تحويل الأعمدة إلى strings لتجنب مشاكل المطابقة
                    df_hc['Employee_ID'] = df_hc['Employee_ID'].astype(str)
                    
                except Exception as e:
                    st.error(f"خطأ في قراءة ملف HC: {str(e)}")
                    return

            # ... (بقية عمليات المعالجة وإنشاء التقرير)

        except Exception as e:
            st.error(f"حدث خطأ غير متوقع: {str(e)}")

if __name__ == "__main__":
    main()
