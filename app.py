import pandas as pd
import streamlit as st

def process_hc_file(hc_file):
    """معالجة وتحقق من ملف HC"""
    try:
        # قراءة الملف
        df_hc = pd.read_excel(hc_file) if hc_file.name.endswith('.xlsx') else pd.read_csv(hc_file)
        
        # تنظيف أسماء الأعمدة (إزالة فراغات، تحويل لحروف صغيرة)
        df_hc.columns = df_hc.columns.str.strip().str.lower()
        
        # خريطة لأسماء الأعمدة البديلة المقبولة
        column_aliases = {
            'employee_id': ['id', 'employee id', 'emp_id', 'رقم الموظف'],
            'team_leader': ['tl', 'team leader', 'المشرف', 'المشرف المسؤول']
        }
        
        # البحث عن أعمدة مطابقة
        found_columns = {}
        for standard_col, aliases in column_aliases.items():
            for alias in aliases:
                if alias.lower() in df_hc.columns:
                    found_columns[standard_col] = alias
                    break
        
        # إذا لم نجد جميع الأعمدة المطلوبة
        if len(found_columns) < 2:
            missing = set(column_aliases.keys()) - set(found_columns.keys())
            st.error(f"ملف HC ينقصه الأعمدة التالية: {', '.join(missing)}")
            st.markdown("""
            **ملاحظات مهمة:**
            - يجب أن يحتوي ملف HC على الأعمدة التالية (أو ما يعادلها):
              - `Employee_ID` (أو: ID, Employee ID, Emp_ID, رقم الموظف)
              - `Team_Leader` (أو: TL, Team Leader, المشرف, المشرف المسؤول)
            - يمكنك تنزيل نموذج ملف HC باستخدام الزر أدناه
            """)
            
            # زر لتنزيل نموذج
            sample_hc = pd.DataFrame(columns=['Employee_ID', 'Team_Leader'])
            st.download_button(
                label="⬇️ تنزيل نموذج ملف HC",
                data=sample_hc.to_csv(index=False),
                file_name="hc_template.csv",
                mime="text/csv"
            )
            return None
        
        # إعادة تسمية الأعمدة للشكل القياسي
        df_hc = df_hc.rename(columns=found_columns)
        
        # تنظيف البيانات
        df_hc['employee_id'] = df_hc['employee_id'].astype(str).str.strip()
        df_hc['team_leader'] = df_hc['team_leader'].astype(str).str.strip()
        
        return df_hc[['employee_id', 'team_leader']]
    
    except Exception as e:
        st.error(f"حدث خطأ أثناء قراءة ملف HC: {str(e)}")
        return None

def main():
    st.title("نظام تحليل الأداء - إدارة ملف HC")
    
    # رفع الملف
    hc_file = st.file_uploader("رفع ملف HC (الموظفين والمشرفين)", 
                             type=['xlsx', 'csv'],
                             help="يجب أن يحتوي على عمود للموظفين (Employee_ID) وعمود للمشرفين (Team_Leader)")
    
    if hc_file:
        df_hc = process_hc_file(hc_file)
        
        if df_hc is not None:
            st.success("تم تحميل ملف HC بنجاح!")
            st.dataframe(df_hc.head())
            
            # عرض إحصاءات
            st.subheader("إحصاءات ملف HC")
            col1, col2 = st.columns(2)
            
            with col1:
                st.metric("عدد الموظفين", len(df_hc))
                
            with col2:
                st.metric("عدد المشرفين", df_hc['team_leader'].nunique())
            
            # عرض توزيع المشرفين
            st.bar_chart(df_hc['team_leader'].value_counts())

if __name__ == "__main__":
    main()
