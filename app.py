import streamlit as st
import pandas as pd
import io

# --- Main App Logic ---

st.set_page_config(page_title="AHT Report Generator", layout="wide")

st.title("📈 مولد تقارير الأداء التلقائي")
st.write("ارفع ملفات البيانات الخام وملف الموظفين (HC) لإنشاء التقرير النهائي.")

# --- Step 1: File Uploaders ---
uploaded_raw_data = st.file_uploader(
    "1. ارفع ملف أو أكثر من ملفات بيانات المحادثات الخام",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True
)

uploaded_hc_data = st.file_uploader(
    "2. ارفع ملف بيانات الموظفين (HC)",
    type=["xlsx", "xls", "csv"]
)

# --- Step 2: Process files when both are uploaded ---
if uploaded_raw_data and uploaded_hc_data:
    try:
        # --- Robust file reading function ---
        def read_uploaded_file(file, file_type_name):
            """Tries to read an uploaded file, providing specific error feedback."""
            try:
                # First, try to read as an Excel file
                return pd.read_excel(file)
            except Exception as e1:
                st.warning(f"لم يتمكن من قراءة '{file.name}' كملف إكسل. يحاول الآن قراءته كملف CSV...")
                file.seek(0)
                try:
                    # Try reading as CSV with robust settings
                    return pd.read_csv(file, encoding='windows-1256', sep=',', on_bad_lines='skip', engine='python')
                except Exception as e2:
                    st.error(f"فشل في قراءة الملف '{file.name}' ({file_type_name}). قد يكون الملف تالفًا. جرب فتحه في Excel وإعادة حفظه.")
                    st.error(f"تفاصيل الخطأ: {e2}")
                    return None

        # --- Process multiple raw data files ---
        all_raw_data_dfs = []
        st.write("--- قراءة ملفات البيانات الخام ---")
        all_raw_files_valid = True
        for file in uploaded_raw_data:
            df_single_raw = read_uploaded_file(file, "بيانات خام")
            if df_single_raw is None:
                all_raw_files_valid = False
                break 
            all_raw_data_dfs.append(df_single_raw)
        
        # --- Read the single HC data file ---
        df_hc = read_uploaded_file(uploaded_hc_data, "بيانات الموظفين")

        # --- Proceed only if all files are read successfully ---
        if all_raw_files_valid and df_hc is not None:
            df_raw = pd.concat(all_raw_data_dfs, ignore_index=True)
            st.success("تم دمج جميع ملفات البيانات بنجاح.")
            st.success(f"تم تحميل ملف بيانات الموظفين ({uploaded_hc_data.name}) بنجاح.")

            df_hc_original = df_hc.copy()
            df_hc.columns = df_hc.columns.str.strip()
            column_mapping = {'Email': 'agent_email', 'TL': 'Team leader', 'SPV': 'Supervisor'}
            df_hc.rename(columns=column_mapping, inplace=True)

            required_hc_columns = ['agent_email', 'Team leader', 'Supervisor']
            if not all(col in df_hc.columns for col in required_hc_columns):
                st.error(f"خطأ: ملف HC يجب أن يحتوي على الأعمدة التالية بعد إعادة التسمية: {required_hc_columns}")
            else:
                # --- Step 3: Merge Data and Calculate Metrics ---
                df_merged = pd.merge(df_raw, df_hc, on='agent_email', how='left')

                # --- FIX: Convert time columns to numeric BEFORE aggregation ---
                # This ensures that any non-numeric values (like text) become NaN, then are filled with 0.
                df_merged['handling_time'] = pd.to_numeric(df_merged['handling_time'], errors='coerce').fillna(0)
                df_merged['wrap_up_time'] = pd.to_numeric(df_merged['wrap_up_time'], errors='coerce').fillna(0)
                if 'agent_first_reply_time' in df_merged.columns:
                    df_merged['agent_first_reply_time'] = pd.to_numeric(df_merged['agent_first_reply_time'], errors='coerce').fillna(0)

                count_col = 'chat_id' if 'chat_id' in df_merged.columns else 'handling_time'
                agg_dict = {
                    'handling_time': ('handling_time', 'sum'),
                    'wrap_up_time': ('wrap_up_time', 'sum'),
                    count_col: (count_col, 'count')
                }
                if 'agent_first_reply_time' in df_merged.columns: agg_dict['FRT'] = ('agent_first_reply_time', 'mean')
                
                df_agent_view = df_merged.groupby(['Team leader', 'Supervisor', 'agent_email']).agg(**agg_dict).reset_index()
                df_agent_view.rename(columns={count_col: 'Total Chats'}, inplace=True)
                df_agent_view['AHT'] = ((df_agent_view['handling_time'] + df_agent_view['wrap_up_time']) / df_agent_view['Total Chats']).round(2)

                final_cols = ['Team leader', 'agent_email', 'Total Chats', 'AHT', 'FRT']
                for col in final_cols:
                    if col not in df_agent_view.columns: df_agent_view[col] = 0
                df_agent_view = df_agent_view[final_cols]

                df_view = df_agent_view.groupby('Team leader').agg({
                    'Total Chats': 'sum', 'AHT': 'mean', 'FRT': 'mean'
                }).reset_index()
                df_view['AHT'] = df_view['AHT'].round(2)

                # --- Step 4: Display results and provide download button ---
                st.subheader("معاينة للنتائج")
                st.write("**ملخص قادة الفرق (View)**")
                st.dataframe(df_view)
                st.write("**التقرير التفصيلي للموظفين (Agent View)**")
                st.dataframe(df_agent_view)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_view.to_excel(writer, sheet_name='View', index=False)
                    df_agent_view.to_excel(writer, sheet_name='Agent View', index=False)
                    df_hc_original.to_excel(writer, sheet_name='HC', index=False)

                    workbook = writer.book
                    header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#4472C4', 'font_color': 'white', 'border': 1})
                    
                    worksheet_view = writer.sheets['View']
                    for c, v in enumerate(df_view.columns.values): worksheet_view.write(0, c, v, header_format)
                    worksheet_view.set_column('A:E', 18)

                    worksheet_agent = writer.sheets['Agent View']
                    for c, v in enumerate(df_agent_view.columns.values): worksheet_agent.write(0, c, v, header_format)
                    worksheet_agent.set_column('A:E', 18)
                    
                    worksheet_hc = writer.sheets['HC']
                    for c, v in enumerate(df_hc_original.columns.values): worksheet_hc.write(0, c, v, header_format)
                    worksheet_hc.set_column('A:Z', 18)

                st.download_button(
                    label="⬇️ تحميل التقرير النهائي",
                    data=output.getvalue(),
                    file_name="Final_Performance_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"حدث خطأ غير متوقع بعد قراءة الملفات: {e}")
