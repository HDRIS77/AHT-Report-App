import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
import streamlit as st
from io import BytesIO

# ... (الثوابت والتعريفات الأخرى تبقى كما هي)

def main():
    st.title("نظام تحليل أداء خدمة العملاء المتكامل")
    
    # رفع الملفات
    st.subheader("رفع ملفات البيانات المطلوبة")
    
    try:
        raw_overall_file = st.file_uploader("ملف البيانات الكلية (Raw Overall)", type=['xlsx', 'csv'])
        raw_urdu_file = st.file_uploader("ملف بيانات الأردية (Raw Urdu)", type=['xlsx', 'csv'])
        hc_file = st.file_uploader("ملف التوظيف (HC)", type=['xlsx', 'csv'], 
                                 help="يجب أن يحتوي على أعمدة Employee_ID و Team_Leader")
        
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
                        try:
                            df_hc = pd.read_excel(hc_file) if hc_file.name.endswith('.xlsx') else pd.read_csv(hc_file)
                            if not all(col in df_hc.columns for col in ['Employee_ID', 'Team_Leader']):
                                st.warning("ملف HC يجب أن يحتوي على أعمدة Employee_ID و Team_Leader")
                                return
                        except Exception as hc_error:
                            st.error(f"خطأ في قراءة ملف HC: {str(hc_error)}")
                            return
                    
                    # حساب المقاييس
                    try:
                        with st.spinner('جاري معالجة البيانات...'):
                            metrics = calculate_all_metrics(df_raw_overall, df_raw_urdu, df_hc, target_aht)
                    except Exception as calc_error:
                        st.error(f"خطأ في حساب المقاييس: {str(calc_error)}")
                        return
                    
                    # إنشاء التقرير
                    if metrics:
                        try:
                            with st.spinner('جاري إنشاء التقرير...'):
                                report_data = create_full_report(metrics, target_aht)
                            
                            if report_data:
                                st.success("تم إنشاء التقرير بنجاح!")
                                st.download_button(
                                    label="⬇️ تنزيل التقرير",
                                    data=report_data,
                                    file_name="AHT_Full_Report.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                        except Exception as report_error:
                            st.error(f"خطأ في إنشاء التقرير: {str(report_error)}")
                
                except Exception as main_error:
                    st.error(f"حدث خطأ رئيسي: {str(main_error)}")
            else:
                st.warning("الرجاء رفع ملفات البيانات المطلوبة")
    
    except Exception as ui_error:
        st.error(f"حدث خطأ في واجهة المستخدم: {str(ui_error)}")

if __name__ == "__main__":
    main()
