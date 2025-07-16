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
        # Map your file's column names to the names the script expects
        column_mapping = {
            'Email': 'agent_email',
            'TL': 'Team leader',
            'SPV': 'Supervisor'
        }
        df_hc.rename(columns=column_mapping, inplace=True)
        
        # Check if required columns exist after renaming
        required_hc_columns = ['agent_email', 'Team leader', 'Supervisor']
        if not all(col in df_hc.columns for col in required_hc_columns):
            st.error(f"خطأ: ملف HC يجب أن يحتوي على الأعمدة التالية: 'Email', 'TL', 'SPV'. الأعمدة التي تم العثور عليها: {list(df_hc.columns)}")
        else:
            # --- B. Prepare Raw Data (Merge and add HC info) ---
            # Combine both raw data files
            df_raw_combined = pd.concat([df_overall, df_urdu], ignore_index=True)
            
            # Merge with HC data to add Team leader and Supervisor to each chat
            # We assume the raw data has an 'agent_email' column to link with HC data
            if 'agent_email' not in df_raw_combined.columns:
                st.error("خطأ: ملفات البيانات الخام يجب أن تحتوي على عمود 'agent_email' لربطها بملف الموظفين.")
            else:
                df_merged = pd.merge(df_raw_combined, df_hc, on='agent_email', how='left')

                # --- Step 4: Perform Calculations ---
                # Get a unique list of all team leaders to build the final report
                all_team_leaders = df_merged['Team leader'].dropna().unique()
                df_view = pd.DataFrame({'Team leader': all_team_leaders})

                # Define a function for safe calculation
                def calculate_metric(df, group_col, value_col, agg_func='mean'):
                    if value_col in df.columns:
                        return df.groupby(group_col)[value_col].agg(agg_func)
                    return pd.Series(dtype='float64')

                # Calculate metrics
                df_view['Urdu'] = df_view['Team leader'].map(calculate_metric(df_merged[df_merged['language'] == 'Urdu'], 'Team leader', 'Total_Time'))
                df_view['Arabic'] = df_view['Team leader'].map(calculate_metric(df_merged[df_merged['language'] == 'Arabic'], 'Team leader', 'Total_Time'))
                df_view['Over all AHT Score'] = df_view['Team leader'].map(calculate_metric(df_merged, 'Team leader', 'Total_Time'))
                df_view['Tenured AHT'] = df_view['Team leader'].map(calculate_metric(df_merged[df_merged['Agent Status'] == 'Production'], 'Team leader', 'Total_Time'))
                df_view['Nesting AHT'] = df_view['Team leader'].map(calculate_metric(df_merged[df_merged['Agent Status'] == 'Nesting'], 'Team leader', 'Total_Time'))
                df_view['# Chats Arabic'] = df_view['Team leader'].map(calculate_metric(df_merged[df_merged['language'] == 'Arabic'], 'Team leader', 'language', 'count'))
                df_view['# Chats Urdu'] = df_view['Team leader'].map(calculate_metric(df_merged[df_merged['language'] == 'Urdu'], 'Team leader', 'language', 'count'))
                df_view['FRT'] = df_view['Team leader'].map(calculate_metric(df_merged, 'Team leader', 'FRT'))

                # ... (Add other calculations here in the same way) ...

                # Clean up and finalize the report
                df_view.fillna('-', inplace=True)
                # ... (rest of the finalization and formatting code) ...

                # --- Step 5: Display results and provide download button ---
                st.subheader("معاينة للنتائج (View)")
                st.dataframe(df_view)

                # --- Create the final Excel file with formatting ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_view.to_excel(writer, sheet_name='View', index=False)

                    # ... (rest of the formatting code) ...

                st.download_button(
                    label="⬇️ تحميل التقرير النهائي",
                    data=output.getvalue(),
                    file_name="Final_Performance_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"حدث خطأ غير متوقع: {e}")
