import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
import streamlit as st
from io import BytesIO

# تعريف الثوابت والألوان
HEADER_FILL = PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid")
LIGHT_PURPLE_FILL = PatternFill(start_color="E4D9F5", end_color="E4D9F5", fill_type="solid")
THIN_BORDER = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
CENTER_ALIGN = Alignment(horizontal='center', vertical='center')
HEADER_FONT = Font(bold=True, color="FFFFFF")

def calculate_all_metrics(df_raw_overall, df_raw_urdu, df_hc, target_aht):
    """
    حساب جميع المقاييس المطلوبة من البيانات الخام
    """
    metrics = {}
    
    try:
        # دمج بيانات HC مع البيانات الأخرى
        if df_hc is not None:
            hc_mapping = df_hc.set_index('Employee_ID')['Team_Leader'].to_dict()
            df_raw_overall['Team_Leader'] = df_raw_overall['Employee_ID'].map(hc_mapping)
            df_raw_urdu['Team_Leader'] = df_raw_urdu['Employee_ID'].map(hc_mapping)
        
        # 1. حساب AHT (متوسط وقت التعامل) لكل TL
        aht_scores = df_raw_overall.groupby('Team_Leader')['Handle_Time'].mean().to_dict()
        
        # ... (بقية الحسابات كما هي مع تعديل groupby لاستخدام Team_Leader بدلاً من AD)
        
    except Exception as e:
        st.error(f"حدث خطأ في حساب المقاييس: {str(e)}")
        return {}
    
    return metrics

def main():
    st.title("نظام تحليل أداء خدمة العملاء المتكامل")
    
    # رفع الملفات
    st.subheader("رفع ملفات البيانات المطلوبة")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        raw_overall_file = st.file_uploader("ملف البيانات الكلية (Raw Overall)", type=['xlsx', 'csv'])
    
    with col2:
        raw_urdu_file = st.file_uploader("ملف بيانات الأردية (Raw Urdu)", type=['xlsx', 'csv'])
    
    with col3:
        hc_file = st.file_uploader("ملف التوظيف (HC)", type=['xlsx', 'csv'], help="يجب أن يحتوي على أعمدة Employee_ID و Team_Leader")
    
    target_aht = st.number_input("هدف AHT المطلوب (بالثواني)", value=300, min_value=0)
    
    if st.button("إنشاء التقرير"):
        if raw_overall_file and raw_urdu_file:
            try:
                # قراءة الملفات
                df_raw_overall = pd.read_excel(raw_overall_file) if raw_overall_file.name.endswith('.xlsx') else pd.read_csv(raw_overall_file)
                df_raw_urdu = pd.read_excel(raw_urdu_file) if raw_urdu_file.name.endswith('.xlsx') else pd.read_csv(raw_urdu_file)
                
                # قراءة ملف HC إذا تم رفعه
                df_hc = None
                if hc_file:
                    df_hc = pd.read_excel(hc_file) if hc_file.name.endswith('.xlsx') else pd.read_csv(hc_file)
                    # التحقق من وجود الأعمدة المطلوبة
                    if not all(col in df_hc.columns for col in ['Employee_ID', 'Team_Leader']):
                        st.warning("ملف HC يجب أن يحتوي على أعمدة Employee_ID و Team_Leader")
                        return
                
                # حساب جميع المقاييس
                with st.spinner('جاري معالجة البيانات...'):
                    metrics = calculate_all_metrics(df_raw_overall, df_raw_urdu, df_hc, target_aht)
                
                # ... (بقية الكود كما هو)

if __name__ == "__main__":
    main()
