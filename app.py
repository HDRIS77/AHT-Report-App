import streamlit as st
import pandas as pd
import io
import numpy as np

# --- Main App Logic ---
st.set_page_config(page_title="Final Report Generator", layout="wide")

st.title("📈 مولد تقارير الأداء النهائي")
st.write("ارفع ملفات البيانات الخام، ملف الأردو، وأدخل الرقم المستهدف لإنشاء التقرير.")

# --- Step 1: File Uploaders and Target Input ---
uploaded_overall = st.file_uploader("1. ارفع ملف البيانات الشامل (Raw Overall)", type=["xlsx", "xls"])
uploaded_urdu = st.file_uploader("2. ارفع ملف بيانات الأردو (Raw Urdu)", type=["xlsx", "xls"])
aht_target = st.number_input("3. أدخل الرقم المستهدف للـ AHT (مثلاً: 450)", value=450)


# --- Step 2: Process files when all inputs are ready ---
if uploaded_overall is not None and uploaded_urdu is not None:
    try:
        # Read uploaded files
        df_overall = pd.read_excel(uploaded_overall)
        df_urdu = pd.read_excel(uploaded_urdu)
        st.success("تم قراءة الملفات بنجاح. جاري حساب التقرير...")

        # --- Step 3: Define Column Mapping ---
        # This mapping is based on the Excel formulas provided by the user
        column_map = {
            'C': 'Country', 'H': 'FRT', 'AC': 'Pass/Fail', 'AD': 'Team leader',
            'AF': 'Agent Status', 'AH': 'Total_Time', 'AI': 'Releasing',
            'AJ': 'ZTP', 'AK': 'Missed', 'AO': 'language'
        }
        
        def rename_cols_by_position(df, mapping):
            # Creates a dictionary to map Excel column letters to integer indices
            col_indices = {chr(ord('A') + i): i for i in range(26)}
            for i in range(26): col_indices['A' + chr(ord('A') + i)] = 26 + i
            
            rename_dict = {}
            for letter, name in mapping.items():
                if letter in col_indices and col_indices[letter] < len(df.columns):
                    # Map the original column name at the specified index to the new name
                    rename_dict[df.columns[col_indices[letter]]] = name
            df.rename(columns=rename_dict, inplace=True)
            return df

        # Apply the renaming to the dataframes
        df_overall = rename_cols_by_position(df_overall, column_map)
        df_urdu = rename_cols_by_position(df_urdu, {'AD': 'Team leader', 'AH': 'Total_Time'})

        # --- Step 4: Perform Calculations for the "View" Sheet ---
        # Get a unique list of all team leaders to build the final report
        all_team_leaders = pd.concat([df_overall['Team leader'], df_urdu['Team leader']]).dropna().unique()
        df_view = pd.DataFrame({'Team leader': all_team_leaders})

        # --- Calculate each metric based on the provided formulas ---
        
        # Urdu AHT
        df_view['Urdu'] = df_view['Team leader'].map(df_urdu.groupby('Team leader')['Total_Time'].mean())
        
        # Arabic AHT
        arabic_data = df_overall[df_overall['language'] == 'Arabic']
        df_view['Arabic'] = df_view['Team leader'].map(arabic_data.groupby('Team leader')['Total_Time'].mean())
        
        # Overall AHT Score
        df_view['Over all AHT Score'] = df_view['Team leader'].map(df_overall.groupby('Team leader')['Total_Time'].mean())
        
        # Tenured and Nesting AHT
        production_data = df_overall[df_overall['Agent Status'] == 'Production']
        df_view['Tenured AHT'] = df_view['Team leader'].map(production_data.groupby('Team leader')['Total_Time'].mean())
        nesting_data = df_overall[df_overall['Agent Status'] == 'Nesting']
        df_view['Nesting AHT'] = df_view['Team leader'].map(nesting_data.groupby('Team leader')['Total_Time'].mean())
        
        # Chat Counts
        df_view['# Chats Arabic'] = df_view['Team leader'].map(arabic_data.groupby('Team leader').size())
        df_view['# Chats Urdu'] = df_view['Team leader'].map(df_urdu.groupby('Team leader').size())
        
        # Variance and Status
        df_view['Var From Target'] = df_view['Over all AHT Score'] - aht_target
        df_view['Var From Target'] = df_view['Var From Target'].apply(lambda x: x if x > 0 else np.nan)
        df_view['Status'] = np.where(df_view['Var From Target'].isna(), 'Achieved', 'Not Achieved')
        
        # FRT
        df_view['FRT'] = df_view['Team leader'].map(df_overall.groupby('Team leader')['FRT'].mean())
        
        # AHT per Country
        country_pivot = df_overall.pivot_table(index='Team leader', columns='Country', values='Total_Time', aggfunc='mean')
        df_view = pd.merge(df_view, country_pivot, on='Team leader', how='left')
        
        # Other Counts
        if 'Releasing' in df_overall.columns:
            df_view['Releasing'] = df_view['Team leader'].map(df_overall[df_overall['Releasing'] == 'Releasing'].groupby('Team leader').size())
        if 'ZTP' in df_overall.columns:
            df_view['ZTP'] = df_view['Team leader'].map(df_overall[df_overall['ZTP'] == 'ZTP'].groupby('Team leader').size())
        if 'Missed' in df_overall.columns:
            df_view['Missed'] = df_view['Team leader'].map(df_overall[df_overall['Missed'] == 'Missed'].groupby('Team leader').size())
        
        df_view['Chats'] = df_view['Team leader'].map(df_overall.groupby('Team leader').size())
        
        if 'Pass/Fail' in df_overall.columns:
            df_view['Pass'] = df_view['Team leader'].map(df_overall[df_overall['Pass/Fail'] == 'Pass'].groupby('Team leader').size())
            df_view['Fail'] = df_view['Team leader'].map(df_overall[df_overall['Pass/Fail'] == 'Fail'].groupby('Team leader').size())
        
        # Clean up the final dataframe
        df_view.fillna('-', inplace=True)
        # Ensure correct column order as per the image
        final_columns_order = [
            'Team leader', 'Urdu', 'Arabic', 'Over all AHT Score', 'Tenured AHT', 'Nesting AHT',
            '# Chats Arabic', '# Chats Urdu', 'Var From Target', 'Status', 'Readiness', 'FRT',
            'EG', 'JO', 'BH', 'AE', 'QA', 'KW', 'OM', 'Releasing', 'ZTP', 'Missed', 'Chats', 'Pass', 'Fail'
        ]
        # Add missing columns with placeholder
        for col in final_columns_order:
            if col not in df_view.columns:
                df_view[col] = '-'
        df_view = df_view[final_columns_order]
        
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
            
            # Define formats to match the user's image
            header_format = workbook.add_format({'bold': True, 'valign': 'top', 'fg_color': '#002060', 'font_color': 'white', 'border': 1, 'align': 'center'})
            cell_format = workbook.add_format({'border': 1, 'align': 'center'})
            green_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1, 'align': 'center'})
            red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1, 'align': 'center'})
            yellow_format = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500', 'border': 1, 'align': 'center'})
            
            worksheet.set_default_row(18)
            
            # Write header with format
            for col_num, value in enumerate(df_view.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Apply formatting to all cells
            worksheet.set_column('A:AZ', 12, cell_format)
            
            # Apply conditional formatting for 'Status' column
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
