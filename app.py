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
        df_hc_original = df_hc.copy() # Keep original for HC sheet
        df_hc.columns = df_hc.columns.str.strip()
        column_mapping_hc = {'Email': 'agent_email', 'TL': 'Team leader', 'SPV': 'Supervisor'}
        df_hc.rename(columns=column_mapping_hc, inplace=True)
        
        required_hc_columns = ['agent_email', 'Team leader', 'Supervisor']
        if not all(col in df_hc.columns for col in required_hc_columns):
            st.error(f"خطأ: ملف HC يجب أن يحتوي على الأعمدة التالية: 'Email', 'TL', 'SPV'.")
        else:
            # --- B. Prepare Raw Data (Calculate Total_Time) ---
            def prepare_raw_df(df):
                for col in ['handling_time', 'wrap_up_time', 'agent_first_reply_time']:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                if 'handling_time' in df.columns and 'wrap_up_time' in df.columns:
                    df['Total_Time'] = df['handling_time'] + df['wrap_up_time']
                else:
                    df['Total_Time'] = 0
                return df

            df_overall = prepare_raw_df(df_overall)
            df_urdu = prepare_raw_df(df_urdu)

            # --- C. Merge Data ---
            df_merged = pd.merge(df_overall, df_hc, on='agent_email', how='left')
            urdu_data_merged = pd.merge(df_urdu, df_hc, on='agent_email', how='left')

            # --- Step 4: Perform Calculations for BOTH sheets ---
            
            # --- CALCULATIONS FOR "AGENT VIEW" SHEET ---
            # Start with a base aggregation dictionary
            agg_dict = {
                'Total_Time': pd.NamedAgg(column='Total_Time', aggfunc='sum'),
                'agent_email': pd.NamedAgg(column='agent_email', aggfunc='count')
            }
            # Dynamically add Pass/Fail aggregation if the column exists
            if 'Pass/Fail' in df_merged.columns:
                agg_dict['Pass'] = pd.NamedAgg(column='Pass/Fail', aggfunc=lambda x: (x == 'Pass').sum())
                agg_dict['Fail'] = pd.NamedAgg(column='Pass/Fail', aggfunc=lambda x: (x == 'Fail').sum())

            df_agent_view = df_merged.groupby(['agent_email']).agg(**agg_dict).reset_index()
            df_agent_view.rename(columns={'agent_email': 'Email', 'agent_email_count': 'Total_Chats'}, inplace=True)

            # Ensure Pass/Fail columns exist before calculations
            if 'Pass' not in df_agent_view.columns: df_agent_view['Pass'] = 0
            if 'Fail' not in df_agent_view.columns: df_agent_view['Fail'] = 0
            
            df_agent_view.rename(columns={'Email': 'agent_email', 'Total_Chats': '# Chats'}, inplace=True)

            df_agent_view['AHT Score'] = (df_agent_view['Total_Time'] / df_agent_view['# Chats']).round(2)
            df_agent_view['Var from Target'] = df_agent_view['AHT Score'] - aht_target
            df_agent_view['Var from Target'] = df_agent_view['Var from Target'].apply(lambda x: x if x > 0 else np.nan)
            df_agent_view['Status'] = np.where(df_agent_view['Var from Target'].isna(), 'Achieved', 'Not Achieved')
            df_agent_view['Readiness'] = (df_agent_view['Pass'] / (df_agent_view['Pass'] + df_agent_view['Fail'])).fillna(0)

            # Merge with HC data to get agent info
            df_agent_view = pd.merge(df_hc, df_agent_view, on='agent_email', how='left')
            
            # Reorder and select final columns for Agent View
            agent_view_cols = ['HR ID', 'Full Name', 'Email', 'TL', 'SPV', 'AHT Score', '# Chats', 'Var from Target', 'Status', 'Pass', 'Fail', 'Readiness']
            df_agent_view.rename(columns={'agent_email': 'Email', 'Team leader': 'TL', 'Supervisor': 'SPV'}, inplace=True)
            for col in agent_view_cols:
                if col not in df_agent_view.columns:
                    df_agent_view[col] = '-'
            df_agent_view = df_agent_view[agent_view_cols]


            # --- CALCULATIONS FOR "VIEW" (TEAM LEADER) SHEET ---
            all_team_leaders = df_merged['Team leader'].dropna().unique()
            df_view = pd.DataFrame({'Team leader': all_team_leaders})
            
            # ... (Add all calculations for the View sheet here, checking for column existence) ...
            # This part needs to be fully built out to match the image
            df_view['Over all AHT Score'] = df_view['Team leader'].map(df_merged.groupby('Team leader')['Total_Time'].mean())


            # --- Step 5: Display results and provide download button ---
            st.subheader("معاينة للنتائج")
            st.write("**ملخص قادة الفرق (View)**")
            st.dataframe(df_view)
            st.write("**التقرير التفصيلي للموظفين (Agent View)**")
            st.dataframe(df_agent_view)

            # --- Create the final Excel file with formatting ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_view.to_excel(writer, sheet_name='View', index=False)
                df_agent_view.to_excel(writer, sheet_name='Agent View', index=False)
                df_hc_original.to_excel(writer, sheet_name='HC', index=False)
                
                workbook = writer.book
                # --- Define Formats ---
                header_format_view = workbook.add_format({'bold': True, 'valign': 'top', 'fg_color': '#002060', 'font_color': 'white', 'border': 1, 'align': 'center'})
                header_format_agent = workbook.add_format({'bold': True, 'valign': 'top', 'fg_color': '#2F5597', 'font_color': 'white', 'border': 1, 'align': 'center'})
                cell_format = workbook.add_format({'border': 1, 'align': 'center'})
                green_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1, 'align': 'center'})
                red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1, 'align': 'center'})
                yellow_format = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500', 'border': 1, 'align': 'center'})

                # --- Format "View" Sheet ---
                worksheet_view = writer.sheets['View']
                worksheet_view.set_default_row(18)
                for col_num, value in enumerate(df_view.columns.values):
                    worksheet_view.write(0, col_num, value, header_format_view)
                worksheet_view.set_column('A:AZ', 12, cell_format)
                
                # --- Format "Agent View" Sheet ---
                worksheet_agent = writer.sheets['Agent View']
                worksheet_agent.set_default_row(18)
                for col_num, value in enumerate(df_agent_view.columns.values):
                    worksheet_agent.write(0, col_num, value, header_format_agent)
                worksheet_agent.set_column('A:L', 18, cell_format)
                worksheet_agent.conditional_format('I2:I100', {'type': 'cell', 'criteria': '==', 'value': '"Achieved"', 'format': green_format})
                worksheet_agent.conditional_format('I2:I100', {'type': 'cell', 'criteria': '==', 'value': '"Not Achieved"', 'format': red_format})
                worksheet_agent.conditional_format('L2:L100', {'type': '3_color_scale', 'min_value': 0, 'mid_value': 0.8, 'max_value': 1, 'min_color': '#F8696B', 'mid_color': '#FFEB84', 'max_color': '#63BE7B'})

                # --- Format "HC" Sheet ---
                worksheet_hc = writer.sheets['HC']
                for col_num, value in enumerate(df_hc_original.columns.values):
                    worksheet_hc.write(0, col_num, value, header_format_view)
                worksheet_hc.set_column('A:Z', 18, cell_format)

            st.download_button(
                label="⬇️ تحميل التقرير النهائي",
                data=output.getvalue(),
                file_name="Final_Performance_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"حدث خطأ غير متوقع: {e}")
