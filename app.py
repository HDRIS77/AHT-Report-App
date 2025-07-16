import pandas as pd
import streamlit as st

def process_hc_file(hc_file):
    """معالجة وتحقق من ملف HC مع أسماء أعمدة بديلة"""
    try:
        # قراءة الملف
        df_hc = pd.read_excel(hc_file) if hc_file.name.endswith('.xlsx') else pd.read_csv(hc_file)
        
        # خريطة لأسماء الأعمدة المقبولة (الأصلية والبديلة)
        column_mapping = {
            'employee_id': ['hr id', 'employee_id', 'emp_id', 'id', 'رقم الموظف'],
            'team_leader': ['tl', 'team_leader', 'team leader', 'المشرف', 'المشرف المسؤول']
        }
        
        # البحث عن أعمدة مطابقة في الملف
        found_columns = {}
        for standard_col, aliases in column_mapping.items():
            for alias in aliases:
                if alias.lower() in [col.lower().strip() for col in df_hc.columns]:
                    # الحفاظ على اسم العمود الأصلي كما في الملف
                    original_col_name = [col for col in df_hc.columns if col.lower().strip() == alias.lower()][0]
                    found_columns[standard_col] = original_col_name
                    break
        
        # التحقق من وجود جميع الأعمدة المطلوبة
        if len(found_columns) < 2:
            missing = [col for col in column_mapping.keys() if col not in found_columns]
            st.error(f"ملف HC ينقصه الأعمدة التالية: {', '.join(missing)}")
            
            st.markdown("""
            **البدائل المقبولة لهذه الأعمدة:**
            - `Employee_ID` (أو: HR ID, Emp_ID, ID, رقم الموظف)
            - `Team_Leader` (أو: TL, Team Leader, المشرف)
            """)
            
            # عرض أعمدة الملف المرفوع للمساعدة في اكتشاف المشكلة
            st.warning(f"أعمدة الملف المرفوع: {list(df_hc.columns)}")
            
            # زر لتنزيل نموذج
            sample_hc = pd.DataFrame(columns=['HR ID', 'TL'])
            st.download_button(
                label="⬇️ تنزيل نموذج ملف HC",
                data=sample_hc.to_csv(index=False),
                file_name="hc_template.csv",
                mime="text/csv"
            )
            return None
        
        # إعادة تسمية الأعمدة للشكل القياسي
        df_hc = df_hc.rename(columns={
            found_columns['employee_id']: 'employee_id',
            found_columns['team_leader']: 'team_leader'
        })
        
        # تنظيف البيانات
        df_hc['employee_id'] = df_hc['employee_id'].astype(str).str.strip()
        df_hc['team_leader'] = df_hc['team_leader'].astype(str).str.strip()
        
        return df_hc[['employee_id', 'team_leader']]
    
    except Exception as e:
        st.error(f"حدث خطأ أثناء قراءة ملف HC: {str(e)}")
        return None

def main():
    st.title("نظام إدارة ملف الموظفين (HC)")
    
    # رفع الملف
    hc_file = st.file_uploader("رفع ملف HC", 
                             type=['xlsx', 'csv'],
                             help="يجب أن يحتوي على عمود للموظفين (HR ID أو Employee_ID) وعمود للمشرفين (TL أو Team_Leader)")
    
    if hc_file:
        df_hc = process_hc_file(hc_file)
        
        if df_hc is not None:
            st.success("تم تحميل ملف HC بنجاح!")
            
            # عرض البيانات
            st.subheader("معاينة البيانات")
            st.dataframe(df_hc.head())
            
            # إحصاءات
            st.subheader("الإحصاءات")
            col1, col2 = st.columns(2)
            with col1:
                st.metric("عدد الموظفين", df_hc.shape[0])
            with col2:
                st.metric("عدد المشرفين", df_hc['team_leader'].nunique())
            
            # تحميل البيانات المعالجة
            st.download_button(
                label="⬇️ تنزيل البيانات المعالجة",
                data=df_hc.to_csv(index=False),
                file_name="processed_hc.csv",
                mime="text/csv"
            )

if __name__ == "__main__":
    main()
