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
        df_hc_original = df_hc.copy() 
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
            agg_dict = {
                'Total_Time': pd.NamedAgg(column='Total_Time', aggfunc='sum'),
                '# Chats': pd.NamedAgg(column='Total_Time', aggfunc='count')
            }
            if 'Pass/Fail' in df_merged.columns:
                agg_dict['Pass'] = pd.NamedAgg(column='Pass/Fail', aggfunc=lambda x: (x == 'Pass').sum())
                agg_dict['Fail'] = pd.NamedAgg(column='Pass/Fail', aggfunc=lambda x: (x == 'Fail').sum())

            df_agent_view = df_merged.groupby(['agent_email']).agg(**agg_dict).reset_index()

            if 'Pass' not in df_agent_view.columns: df_agent_view['Pass'] = 0
            if 'Fail' not in df_agent_view.columns: df_agent_view['Fail'] = 0
            
            df_agent_view['AHT Score'] = (df_agent_view['Total_Time'] / df_agent_view['# Chats']).round(2)
            df_agent_view['Var from Target'] = df_agent_view['AHT Score'] - aht_target
            df_agent_view['Var from Target'] = df_agent_view['Var from Target'].apply(lambda x: x if x > 0 else np.nan)
            df_agent_view['Status'] = np.where(df_agent_view['Var from Target'].isna(), 'Achieved', 'Not Achieved')
            
            total_pass_fail = df_agent_view['Pass'] + df_agent_view['Fail']
            df_agent_view['Readiness'] = np.where(total_pass_fail > 0, df_agent_view['Pass'] / total_pass_fail, 0)

            df_agent_view = pd.merge(df_hc, df_agent_view, on='agent_email', how='left')
            
            agent_view_cols = ['HR ID', 'Full Name', 'Email', 'TL', 'SPV', 'AHT Score', '# Chats', 'Var from Target', 'Status', 'Pass', 'Fail', 'Readiness']
            df_agent_view.rename(columns={'agent_email': 'Email', 'Team leader': 'TL', 'Supervisor': 'SPV'}, inplace=True)
            for col in agent_view_cols:
                if col not in df_agent_view.columns:
                    df_agent_view[col] = '-'
            df_agent_view = df_agent_view[agent_view_cols]


            # --- CALCULATIONS FOR "VIEW" (TEAM LEADER) SHEET ---
            # This function will create one summary table (like the BPO section)
            def create_summary_table(df, df_urdu_agg, target, lob_name=None):
                if lob_name:
                    df_filtered = df[df['LOB'] == lob_name]
                else: # For overall summary
                    df_filtered = df

                if df_filtered.empty:
                    return pd.DataFrame()

                all_team_leaders = df_filtered['Team leader'].dropna().unique()
                df_summary = pd.DataFrame({'Team leader': all_team_leaders})

                # Calculate all metrics
                df_summary['Urdu'] = df_summary['Team leader'].map(df_urdu_agg.groupby('Team leader')['Total_Time'].mean())
                arabic_data = df_filtered[df_filtered['language'] == 'Arabic']
                df_summary['Arabic'] = df_summary['Team leader'].map(arabic_data.groupby('Team leader')['Total_Time'].mean())
                df_summary['Over all AHT Score'] = df_summary['Team leader'].map(df_filtered.groupby('Team leader')['Total_Time'].mean())
                
                if 'Agent Status' in df_filtered.columns:
                    prod_data = df_filtered[df_filtered['Agent Status'] == 'Production']
                    df_summary['Tenured AHT'] = df_summary['Team leader'].map(prod_data.groupby('Team leader')['Total_Time'].mean())
                    nest_data = df_filtered[df_filtered['Agent Status'] == 'Nesting']
                    df_summary['Nesting AHT'] = df_summary['Team leader'].map(nest_data.groupby('Team leader')['Total_Time'].mean())

                df_summary['# Chats Arabic'] = df_summary['Team leader'].map(arabic_data.groupby('Team leader').size())
                df_summary['# Chats Urdu'] = df_summary['Team leader'].map(df_urdu_agg.groupby('Team leader').size())
                
                df_summary['Var From Target'] = df_summary['Over all AHT Score'] - target
                df_summary['Var From Target'] = df_summary['Var From Target'].apply(lambda x: x if x > 0 else np.nan)
                df_summary['Status'] = np.where(df_summary['Var From Target'].isna(), 'Achieved', 'Not Achieved')

                # ... add other calculations (FRT, Country, etc.) here ...
                
                return df_summary

            # --- Create each section for the "View" sheet ---
            # IMPORTANT: This assumes your HC file has a column named 'LOB'
            if 'LOB' not in df_merged.columns:
                st.error("خطأ: ملف HC يجب أن يحتوي على عمود 'LOB' لتصنيف الأقسام (BPO, Brightscouts, etc.).")
            else:
                bpo_summary = create_summary_table(df_merged, urdu_data_merged, aht_target, "BPO")
                brightscouts_summary = create_summary_table(df_merged, urdu_data_merged, aht_target, "Brightscouts")
                # Add other sections here...

                # --- Step 5: Display results and provide download button ---
                st.subheader("معاينة للنتائج")
                st.write("**ملخص قادة الفرق (View)**")
                if not bpo_summary.empty:
                    st.write("### BPO")
                    st.dataframe(bpo_summary)
                if not brightscouts_summary.empty:
                    st.write("### Brightscouts")
                    st.dataframe(brightscouts_summary)
                
                st.write("**التقرير التفصيلي للموظفين (Agent View)**")
                st.dataframe(df_agent_view)

                # --- Create the final Excel file with formatting ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    # Write each section to the "View" sheet
                    current_row = 0
                    if not bpo_summary.empty:
                        bpo_summary.to_excel(writer, sheet_name='View', startrow=current_row, index=False)
                        current_row += len(bpo_summary) + 2 # Add space
                    if not brightscouts_summary.empty:
                        brightscouts_summary.to_excel(writer, sheet_name='View', startrow=current_row, index=False, header=False) # No header for second table
                        current_row += len(brightscouts_summary) + 2
                    # ... write other sections ...

                    df_agent_view.to_excel(writer, sheet_name='Agent View', index=False)
                    df_hc_original.to_excel(writer, sheet_name='HC', index=False)
                    
                    # ... (Add all the detailed formatting code here as before) ...

                st.download_button(
                    label="⬇️ تحميل التقرير النهائي",
                    data=output.getvalue(),
                    file_name="Final_Performance_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"حدث خطأ غير متوقع: {e}")
