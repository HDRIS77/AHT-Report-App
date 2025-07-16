# --- IMPORTANT FOR DEPLOYMENT ---
# Your requirements.txt file on GitHub must contain the following lines:
# streamlit
# pandas
# openpyxl
# xlsxwriter
# xlrd

import streamlit as st
import pandas as pd
import io
import numpy as np

# --- Main App Logic ---
st.set_page_config(page_title="Final Report Generator", layout="wide")

st.title("📈 مولد تقارير الأداء النهائي")
st.write("ارفع ملفات البيانات الخام، ملف الموظفين، وأدخل الرقم المستهدف لإنشاء التقرير.")

# --- Step 1: File Uploaders and Target Input ---
uploaded_overall = st.file_uploader("1. ارفع ملف البيانات الشامل (Raw Overall)", type=["xlsx", "xls"])
uploaded_urdu = st.file_uploader("2. ارفع ملف بيانات الأردو (Raw Urdu)", type=["xlsx", "xls"])
uploaded_hc = st.file_uploader("3. ارفع ملف بيانات الموظفين (HC File)", type=["xlsx", "xls"])
aht_target = st.number_input("4. أدخل الرقم المستهدف للـ AHT (مثلاً: 450)", value=450)


# --- Step 2: Process files when all inputs are ready ---
if uploaded_overall and uploaded_urdu and uploaded_hc:
    try:
        # Read uploaded files
        df_overall = pd.read_excel(uploaded_overall)
        df_urdu = pd.read_excel(uploaded_urdu)
        df_hc = pd.read_excel(uploaded_hc)
        st.success("تم قراءة جميع الملفات بنجاح. جاري حساب التقرير...")

        # --- Step 3: Prepare and Merge Data ---
        
        # --- A. Prepare HC Data (Rename columns) ---
        df_hc.columns = df_hc.columns.str.strip() # Clean column names
        column_mapping_hc = {
            'Email': 'agent_email',
            'TL': 'Team leader',
            'SPV': 'Supervisor'
        }
        df_hc.rename(columns=column_mapping_hc, inplace=True)
        
        required_hc_columns = ['agent_email', 'Team leader', 'Supervisor']
        if not all(col in df_hc.columns for col in required_hc_columns):
            st.error(f"خطأ: ملف HC يجب أن يحتوي على الأعمدة التالية: 'Email', 'TL', 'SPV'. الأعمدة التي تم العثور عليها: {list(df_hc.columns)}")
        else:
            # --- B. Prepare Raw Data (Combine and calculate Total_Time) ---
            # Assume raw data files have consistent naming
            df_raw_combined = pd.concat([df_overall, df_urdu], ignore_index=True)
            
            # Ensure time columns are numeric, coercing errors
            for col in ['handling_time', 'wrap_up_time']:
                if col in df_raw_combined.columns:
                    df_raw_combined[col] = pd.to_numeric(df_raw_combined[col], errors='coerce').fillna(0)

            # Calculate Total_Time for AHT calculations
            if 'handling_time' in df_raw_combined.columns and 'wrap_up_time' in df_raw_combined.columns:
                df_raw_combined['Total_Time'] = df_raw_combined['handling_time'] + df_raw_combined['wrap_up_time']
            else:
                st.error("خطأ: ملفات البيانات الخام يجب أن تحتوي على عمودي 'handling_time' و 'wrap_up_time'.")


            # --- C. Merge Data ---
            if 'agent_email' not in df_raw_combined.columns:
                 st.error("خطأ: ملفات البيانات الخام يجب أن تحتوي على عمود 'agent_email' لربطها بملف الموظفين.")
            else:
                df_merged = pd.merge(df_raw_combined, df_hc, on='agent_email', how='left')

                # --- Step 4: Perform Calculations ---
                all_team_leaders = df_merged['Team leader'].dropna().unique()
                df_view = pd.DataFrame({'Team leader': all_team_leaders})

                # Calculate metrics
                urdu_data = df_merged[df_merged['language'] == 'Urdu']
                df_view['Urdu'] = df_view['Team leader'].map(urdu_data.groupby('Team leader')['Total_Time'].mean())
                
                arabic_data = df_merged[df_merged['language'] == 'Arabic']
                df_view['Arabic'] = df_view['Team leader'].map(arabic_data.groupby('Team leader')['Total_Time'].mean())
                df_view['Over all AHT Score'] = df_view['Team leader'].map(df_merged.groupby('Team leader')['Total_Time'].mean())
                
                if 'Agent Status' in df_merged.columns:
                    production_data = df_merged[df_merged['Agent Status'] == 'Production']
                    df_view['Tenured AHT'] = df_view['Team leader'].map(production_data.groupby('Team leader')['Total_Time'].mean())
                    nesting_data = df_merged[df_merged['Agent Status'] == 'Nesting']
                    df_view['Nesting AHT'] = df_view['Team leader'].map(nesting_data.groupby('Team leader')['Total_Time'].mean())
                
                df_view['# Chats Arabic'] = df_view['Team leader'].map(arabic_data.groupby('Team leader').size())
                df_view['# Chats Urdu'] = df_view['Team leader'].map(urdu_data.groupby('Team leader').size())
                
                df_view['Var From Target'] = df_view['Over all AHT Score'] - aht_target
                df_view['Var From Target'] = df_view['Var From Target'].apply(lambda x: x if x > 0 else np.nan)
                df_view['Status'] = np.where(df_view['Var From Target'].isna(), 'Achieved', 'Not Achieved')
                
                if 'FRT' in df_merged.columns:
                    df_merged['FRT'] = pd.to_numeric(df_merged['FRT'], errors='coerce')
                    df_view['FRT'] = df_view['Team leader'].map(df_merged.groupby('Team leader')['FRT'].mean())
                
                if 'Country' in df_merged.columns:
                    country_pivot = df_merged.pivot_table(index='Team leader', columns='Country', values='Total_Time', aggfunc='mean')
                    df_view = pd.merge(df_view, country_pivot, on='Team leader', how='left')
                
                # Clean up and finalize the report
                final_columns_order = [
                    'Team leader', 'Urdu', 'Arabic', 'Over all AHT Score', 'Tenured AHT', 'Nesting AHT',
                    '# Chats Arabic', '# Chats Urdu', 'Var From Target', 'Status', 'Readiness', 'FRT',
                    'EG', 'JO', 'BH', 'AE', 'QA', 'KW', 'OM', 'Releasing', 'ZTP', 'Missed', 'Chats', 'Pass', 'Fail'
                ]
                for col in final_columns_order:
                    if col not in df_view.columns:
                        df_view[col] = '-'
                df_view = df_view[final_columns_order]
                df_view.fillna('-', inplace=True)
                df_view = df_view.round(2)

                # --- Step 5: Display results and provide download button ---
                st.subheader("معاينة للنتائج (View)")
                st.dataframe(df_view)

                # --- Create the final Excel file with formatting ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_view.to_excel(writer, sheet_name='View', index=False)
                    
                    workbook = writer.book
                    worksheet = writer.sheets['View']
                    header_format = workbook.add_format({'bold': True, 'valign': 'top', 'fg_color': '#002060', 'font_color': 'white', 'border': 1, 'align': 'center'})
                    cell_format = workbook.add_format({'border': 1, 'align': 'center'})
                    green_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1, 'align': 'center'})
                    red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1, 'align': 'center'})
                    
                    worksheet.set_default_row(18)
                    for col_num, value in enumerate(df_view.columns.values):
                        worksheet.write(0, col_num, value, header_format)
                    worksheet.set_column('A:AZ', 12, cell_format)
                    worksheet.conditional_format('J2:J100', {'type': 'cell', 'criteria': '==', 'value': '"Achieved"', 'format': green_format})
                    worksheet.conditional_format('J2:J100', {'type': 'cell', 'criteria': '==', 'value': '"Not Achieved"', 'format': red_format})

                st.download_button(
                    label="⬇️ تحميل التقرير النهائي",
                    data=output.getvalue(),
                    file_name="Final_Performance_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"حدث خطأ غير متوقع: {e}")
