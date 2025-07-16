import streamlit as st
import pandas as pd
import io

# --- NEW: Install the required library for writing formatted Excel files ---
# Streamlit handles this via the requirements.txt file, so no !pip install is needed here.

# --------------------------------------------------------------------------
# Main App Logic
# --------------------------------------------------------------------------

st.set_page_config(page_title="AHT Report Generator", layout="wide")

st.title("?? „Ê·œ  ﬁ«—Ì— «·√œ«¡ «· ·ﬁ«∆Ì")
st.write("«—›⁄ „·›«  «·»Ì«‰«  «·Œ«„ Ê„·› «·„ÊŸ›Ì‰ (HC) ·≈‰‘«¡ «· ﬁ—Ì— «·‰Â«∆Ì.")

# --- Step 1: File Uploaders ---
uploaded_raw_data = st.file_uploader(
    "1. «—›⁄ „·› √Ê √ﬂÀ— „‰ „·›«  »Ì«‰«  «·„Õ«œÀ«  «·Œ«„",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

uploaded_hc_data = st.file_uploader(
    "2. «—›⁄ „·› »Ì«‰«  «·„ÊŸ›Ì‰ (HC)",
    type=["xlsx", "xls"]
)

# --- Step 2: Process files when both are uploaded ---
if uploaded_raw_data and uploaded_hc_data:
    try:
        # --- Process multiple raw data files ---
        all_raw_data_dfs = []
        st.write("--- ﬁ—«¡… „·›«  «·»Ì«‰«  «·Œ«„ ---")
        for file in uploaded_raw_data:
            st.write(f"ﬁ—«¡… «·„·›: {file.name}...")
            df_single_raw = pd.read_excel(file)
            all_raw_data_dfs.append(df_single_raw)

        df_raw = pd.concat(all_raw_data_dfs, ignore_index=True)
        st.success(" „ œ„Ã Ã„Ì⁄ „·›«  «·»Ì«‰«  »‰Ã«Õ.")

        # --- Read the single HC data file ---
        df_hc = pd.read_excel(uploaded_hc_data)
        st.success(f" „  Õ„Ì· „·› »Ì«‰«  «·„ÊŸ›Ì‰ ({uploaded_hc_data.name}) »‰Ã«Õ.")

        df_hc_original = df_hc.copy()
        df_hc.columns = df_hc.columns.str.strip()

        column_mapping = {'Email': 'agent_email', 'TL': 'Team leader', 'SPV': 'Supervisor'}
        df_hc.rename(columns=column_mapping, inplace=True)

        required_hc_columns = ['agent_email', 'Team leader', 'Supervisor']
        if not all(col in df_hc.columns for col in required_hc_columns):
            st.error(f"Œÿ√: „·› HC ÌÃ» √‰ ÌÕ ÊÌ ⁄·Ï «·√⁄„œ… «· «·Ì… »⁄œ ≈⁄«œ… «· ”„Ì…: {required_hc_columns}")
        else:
            # --- Step 3: Merge Data and Calculate Metrics ---
            df_merged = pd.merge(df_raw, df_hc, on='agent_email', how='left')

            count_col = 'chat_id' if 'chat_id' in df_merged.columns else 'handling_time'
            agg_dict = {
                'handling_time': ('handling_time', 'sum'),
                'wrap_up_time': ('wrap_up_time', 'sum'),
                count_col: (count_col, 'count')
            }
            if 'agent_first_reply_time' in df_merged.columns: agg_dict['FRT'] = ('agent_first_reply_time', 'mean')
            if 'nps_score' in df_merged.columns: agg_dict['NPS'] = ('nps_score', 'mean')
            if 'csat_score' in df_merged.columns: agg_dict['CSAT'] = ('csat_score', 'mean')
            if 'reopen_count' in df_merged.columns: agg_dict['Re-open'] = ('reopen_count', 'sum')

            df_agent_view = df_merged.groupby(['Team leader', 'Supervisor', 'agent_email']).agg(**agg_dict).reset_index()
            df_agent_view.rename(columns={count_col: 'Total Chats'}, inplace=True)
            df_agent_view['AHT'] = ((df_agent_view['handling_time'] + df_agent_view['wrap_up_time']) / df_agent_view['Total Chats']).round(2)

            final_cols = ['Team leader', 'agent_email', 'Total Chats', 'AHT', 'FRT', 'NPS', 'CSAT', 'Re-open']
            for col in final_cols:
                if col not in df_agent_view.columns: df_agent_view[col] = 0
            df_agent_view = df_agent_view[final_cols]

            df_view = df_agent_view.groupby('Team leader').agg({
                'Total Chats': 'sum', 'AHT': 'mean', 'FRT': 'mean',
                'NPS': 'mean', 'CSAT': 'mean', 'Re-open': 'sum'
            }).reset_index()
            df_view['AHT'] = df_view['AHT'].round(2)

            # --- Step 4: Display results and provide download button ---
            st.subheader("„⁄«Ì‰… ··‰ «∆Ã")
            st.write("**„·Œ’ ﬁ«œ… «·›—ﬁ (View)**")
            st.dataframe(df_view)
            st.write("**«· ﬁ—Ì— «· ›’Ì·Ì ··„ÊŸ›Ì‰ (Agent View)**")
            st.dataframe(df_agent_view)

            # Create the Excel file in memory
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_view.to_excel(writer, sheet_name='View', index=False)
                df_agent_view.to_excel(writer, sheet_name='Agent View', index=False)
                df_hc_original.to_excel(writer, sheet_name='HC', index=False)

                workbook = writer.book
                header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#4472C4', 'font_color': 'white', 'border': 1})
                cell_format = workbook.add_format({'border': 1})
                green_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1})
                yellow_format = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500', 'border': 1})
                red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1})

                # Format sheets
                worksheet_view = writer.sheets['View']
                for c, v in enumerate(df_view.columns.values): worksheet_view.write(0, c, v, header_format)
                worksheet_view.conditional_format('C2:C100', {'type': 'cell', 'criteria': '<=', 'value': 400, 'format': green_format})
                worksheet_view.conditional_format('C2:C100', {'type': 'cell', 'criteria': 'between', 'minimum': 401, 'maximum': 500, 'format': yellow_format})
                worksheet_view.conditional_format('C2:C100', {'type': 'cell', 'criteria': '>', 'value': 500, 'format': red_format})
                worksheet_view.set_column('A:H', 15, cell_format)

                worksheet_agent = writer.sheets['Agent View']
                for c, v in enumerate(df_agent_view.columns.values): worksheet_agent.write(0, c, v, header_format)
                worksheet_agent.conditional_format('D2:D100', {'type': 'cell', 'criteria': '<=', 'value': 400, 'format': green_format})
                worksheet_agent.conditional_format('D2:D100', {'type': 'cell', 'criteria': 'between', 'minimum': 401, 'maximum': 500, 'format': yellow_format})
                worksheet_agent.conditional_format('D2:D100', {'type': 'cell', 'criteria': '>', 'value': 500, 'format': red_format})
                worksheet_agent.set_column('A:H', 18, cell_format)
                
                worksheet_hc = writer.sheets['HC']
                for c, v in enumerate(df_hc_original.columns.values): worksheet_hc.write(0, c, v, header_format)
                worksheet_hc.set_column('A:Z', 18, cell_format)

            st.download_button(
                label="??  Õ„Ì· «· ﬁ—Ì— «·‰Â«∆Ì",
                data=output.getvalue(),
                file_name="Final_Performance_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"ÕœÀ Œÿ√ €Ì— „ Êﬁ⁄: {e}")
