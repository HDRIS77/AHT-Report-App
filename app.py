import pandas as pd
import streamlit as st

def process_hc_file(hc_file):
    """معالجة ملف HC مع دعم جميع أشكال أسماء الأعمدة"""
    try:
        # قراءة الملف
        df_hc = pd.read_excel(hc_file) if hc_file.name.endswith('.xlsx') else pd.read_csv(hc_file)
        
        # خريطة شاملة لأسماء الأعمدة المقبولة
        ACCEPTED_COLUMNS = {
            'employee_id': [
                'hr id', 'employee id', 'employee_id', 'emp_id', 
                'id', 'employee', 'staff id', 'رقم الموظف',
                'الرقم الوظيفي', 'كود الموظف'
            ],
            'team_leader': [
                'tl', 'team leader', 'team_leader', 'supervisor',
                'manager', 'المشرف', 'المسؤول', 'المشرف المسؤول',
                'رئيس الفريق'
            ]
        }
        
        # البحث عن أعمدة مطابقة
        matched_columns = {}
        for standard_col, aliases in ACCEPTED_COLUMNS.items():
            for alias in aliases:
                # البحث بدون حساسية لحالة الأحرف وبإزالة الفراغات
                normalized_columns = [col.strip().lower() for col in df_hc.columns]
                if alias.lower().strip() in normalized_columns:
                    # الحصول على اسم العمود الأصلي
                    original_col = df_hc.columns[normalized_columns.index(alias.lower().strip())]
                    matched_columns[standard_col] = original_col
                    break
        
        # التحقق من وجود الأعمدة المطلوبة
        if len(matched_columns) < 2:
            missing = [col for col in ACCEPTED_COLUMNS.keys() if col not in matched_columns]
            st.error("⚠️ ملف HC لا يحتوي على الأعمدة المطلوبة")
            
            # رسالة مساعدة موسعة
            st.markdown("""
            ### الأعمدة المطلوبة وأشكالها المقبولة:
            
            **1. عمود الموظف (Employee_ID):**  
            يمكن أن يحمل أي من هذه الأسماء:  
            - `HR ID`، `Employee ID`، `Emp_ID`، `ID`  
            - `رقم الموظف`، `الرقم الوظيفي`، `كود الموظف`
            
            **2. عمود المشرف (Team_Leader):**  
            يمكن أن يحمل أي من هذه الأسماء:  
            - `TL`، `Team Leader`، `Supervisor`  
            - `المشرف`، `المسؤول`، `رئيس الفريق`
            
            **ملاحظة:** الأسماء غير حساسة لحالة الأحرف (كبيرة/صغيرة) أو للفراغات
            """)
            
            # عرض أعمدة الملف المرفوع للمساعدة في تحديد المشكلة
            st.warning("🔍 الأعمدة الموجودة في ملفك الحالي:")
            st.write(list(df_hc.columns))
            
            # زر لتنزيل نموذج باستخدام الأسماء التي يعرفها النظام
            st.markdown("---")
            st.markdown("### لمساعدتك في إنشاء ملف صحيح:")
            sample_data = {
                'HR ID': ['1001', '1002', '1003'],
                'TL': ['أحمد علي', 'محمد حسن', 'فاطمة عبدالله']
            }
            sample_df = pd.DataFrame(sample_data)
            
            st.download_button(
                label="⬇️ تنزيل نموذج ملف HC (جاهز للاستخدام)",
                data=sample_df.to_csv(index=False, encoding='utf-8-sig'),
                file_name="hc_template.csv",
                mime="text/csv",
                help="قم بملء هذا النموذج ببياناتك وحفظه كملف CSV أو Excel"
            )
            return None
        
        # إعادة تسمية الأعمدة للشكل القياسي
        df_hc = df_hc.rename(columns=matched_columns)
        
        # تنظيف وتحويل أنواع البيانات
        df_hc['employee_id'] = df_hc['employee_id'].astype(str).str.strip()
        df_hc['team_leader'] = df_hc['team_leader'].astype(str).str.strip()
        
        # إزالة الصفوف الفارغة
        df_hc = df_hc.dropna(subset=['employee_id', 'team_leader'])
        
        return df_hc[['employee_id', 'team_leader']].drop_duplicates()
    
    except Exception as e:
        st.error(f"❌ حدث خطأ أثناء معالجة الملف: {str(e)}")
        return None

def main():
    st.set_page_config(page_title="نظام إدارة ملف الموظفين", layout="wide")
    
    st.title("📊 نظام إدارة ملف الموظفين (HC)")
    st.markdown("---")
    
    # قسم رفع الملف
    with st.expander("📤 رفع ملف HC", expanded=True):
        hc_file = st.file_uploader(
            "اختر ملف HC (Excel أو CSV)",
            type=['xlsx', 'csv'],
            help="يجب أن يحتوي الملف على عمود للموظفين وعمود للمشرفين بأي من الأسماء المقبولة"
        )
    
    if hc_file:
        with st.spinner("🔍 جاري تحليل الملف..."):
            df_hc = process_hc_file(hc_file)
        
        if df_hc is not None:
            st.success("✅ تم تحميل الملف بنجاح!")
            st.balloons()
            
            # عرض النتائج
            col1, col2 = st.columns([2, 1])
            
            with col1:
                st.subheader("📝 معاينة البيانات")
                st.dataframe(df_hc.head(10), height=300)
            
            with col2:
                st.subheader("📊 إحصاءات")
                st.metric("عدد الموظفين", len(df_hc))
                st.metric("عدد المشرفين", df_hc['team_leader'].nunique())
            
            # تحميل البيانات المعالجة
            st.markdown("---")
            st.subheader("💾 حفظ البيانات المعالجة")
            
            export_format = st.radio("اختر صيغة التصدير:", ('Excel', 'CSV'))
            
            if export_format == 'Excel':
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_hc.to_excel(writer, index=False)
                st.download_button(
                    label="⬇️ تنزيل كملف Excel",
                    data=output.getvalue(),
                    file_name="employees_processed.xlsx",
                    mime="application/vnd.ms-excel"
                )
            else:
                st.download_button(
                    label="⬇️ تنزيل كملف CSV",
                    data=df_hc.to_csv(index=False, encoding='utf-8-sig'),
                    file_name="employees_processed.csv",
                    mime="text/csv"
                )

if __name__ == "__main__":
    main()
