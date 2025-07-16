import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.utils import get_column_letter
import streamlit as st

def create_excel_template(output_path):
    # إنشاء مصنف جديد
    wb = Workbook()
    
    # إزالة الورقة الافتراضية
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']
    
    # إنشاء ورقة View
    view_sheet = wb.create_sheet("View")
    
    # تعريف الأنماط
    header_fill = PatternFill(start_color="FF7030A0", end_color="FF7030A0", fill_type="solid")
    light_purple_fill = PatternFill(start_color="FFE4D9F5", end_color="FFE4D9F5", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFFFF", end_color="FFFFFFFF", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), 
                         right=Side(style='thin'), 
                         top=Side(style='thin'), 
                         bottom=Side(style='thin'))
    center_alignment = Alignment(horizontal='center', vertical='center')
    header_font = Font(bold=True, color="FFFFFF")
    
    # عناوين ورقة View
    view_headers = [
        "TL", "Urdu", "Arabic", "Over all AHT Score", "Tenured AHT", "Nesting AHT", 
        "# Chats Arabic", "# Chats Urdu", "Var From Target", "Status", "Readiness", 
        "FRT", "EG", "JO", "BH", "AE", "QA", "KW", "OM", "Releasing", "ZTP", 
        "Missed", "", "Chats", "Pass", "Fail"
    ]
    
    # إضافة العناوين إلى ورقة View
    for col_num, header in enumerate(view_headers, 1):
        cell = view_sheet.cell(row=1, column=col_num, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
        cell.alignment = center_alignment
    
    # تعيين عرض الأعمدة لورقة View
    view_column_widths = {
        'A': 25, 'B': 8, 'C': 8, 'D': 18, 'E': 12, 'F': 12, 'G': 15, 'H': 12,
        'I': 15, 'J': 10, 'K': 10, 'L': 8, 'M': 5, 'N': 5, 'O': 5, 'P': 5,
        'Q': 5, 'R': 5, 'S': 5, 'T': 10, 'U': 5, 'V': 8, 'W': 5, 'X': 8,
        'Y': 8, 'Z': 8
    }
    
    for col, width in view_column_widths.items():
        view_sheet.column_dimensions[col].width = width
    
    # إضافة أقسام البيانات النموذجية (مشابهة للمثال الخاص بك)
    sections = [
        ("TL", 23),
        ("SPV", 4),
        ("Tenurity", 5),
        ("BPO", 1),
        ("country", 7),
        ("Timing", 5),
        ("CR", 38)
    ]
    
    current_row = 2
    for section, row_count in sections:
        view_sheet.cell(row=current_row, column=1, value=section).fill = light_purple_fill
        for row in range(current_row, current_row + row_count):
            for col in range(1, 27):  # الأعمدة من A إلى Z
                cell = view_sheet.cell(row=row, column=col)
                cell.border = thin_border
                cell.alignment = center_alignment
                if col > 1:
                    cell.value = "-" if col not in [20, 21, 22, 24, 25, 26] else 0
        current_row += row_count + 1
    
    # حالة خاصة لقسم BPO
    view_sheet['K26'].value = "#DIV/0!"
    view_sheet['X26'].value = 32762
    view_sheet['Y26'].value = 32762
    
    # حالة خاصة لقسم country
    for row in range(30, 37):
        view_sheet.cell(row=row, column=11).value = "#DIV/0!"
    
    # إنشاء ورقة Agent View
    agent_sheet = wb.create_sheet("Agent View")
    
    # عناوين ورقة Agent View
    agent_headers = [
        "HR ID", "Full Name", "Email", "TL", "SPV", "AHT Score", "FRT", 
        "# Chats", "Var From Target", "Status", "Pass", "Fail", "Readiness"
    ]
    
    # إضافة العناوين إلى ورقة Agent View
    for col_num, header in enumerate(agent_headers, 1):
        cell = agent_sheet.cell(row=1, column=col_num, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
        cell.alignment = center_alignment
    
    # تعيين عرض الأعمدة لورقة Agent View
    agent_column_widths = {
        'A': 8, 'B': 45, 'C': 45, 'D': 15, 'E': 15, 'F': 12, 'G': 8,
        'H': 10, 'I': 15, 'J': 10, 'K': 8, 'L': 8, 'M': 12
    }
    
    for col, width in agent_column_widths.items():
        agent_sheet.column_dimensions[col].width = width
    
    # إضافة بيانات نموذجية إلى Agent View
    sample_agent_data = [
        [4768, "Ahmed Mohamed Saber Abdelhamid Ahmed", "ahmed.mohamed.2449_bseg.ext@talabat.com", 
         "Michael Fawzy", "Mohamed Hamada", "-", "-", 0, "-", "Achieved", 0, 0, "-"],
        [1301, "Mohamed Mousa Ramdan Hassan", "mohamed.mousa.d12_bseg.ext@talabat.com", 
         "Michael Fawzy", "Mohamed Hamada", "-", "-", 0, "-", "Achieved", 0, 0, "-"],
        [4127, "Mohamed Yehia Youssef Ahmed Nagaty", "mohamed.youssef.2445_bseg.ext@talabat.com", 
         "Michael Fawzy", "Mohamed Hamada", "-", "-", 0, "-", "Achieved", 0, 0, "-"],
        [4761, "Helena Ishak Sabet Ishak", "helena.ishak.2449_bseg.ext@talabat.com", 
         "Michael Fawzy", "Mohamed Hamada", "-", "-", 0, "-", "Achieved", 0, 0, "-"],
        [4060, "Basmala Samer Ibrahim Havez", "basmala.samer.2444_bseg.ext@talabat.com", 
         "Michael Fawzy", "Mohamed Hamada", "-", "-", 0, "-", "Achieved", 0, 0, "-"]
    ]
    
    for row_num, row_data in enumerate(sample_agent_data, 2):
        for col_num, cell_value in enumerate(row_data, 1):
            cell = agent_sheet.cell(row=row_num, column=col_num, value=cell_value)
            cell.border = thin_border
            cell.alignment = center_alignment
    
    # حفظ المصنف
    wb.save(output_path)

def main():
    st.title("منشئ تقارير إكسل")
    st.write("هذه الأداة تنشئ تقارير إكسل بالتنسيق المحدد.")
    
    # أدوات رفع الملفات
    report1 = st.file_uploader("رفع التقرير 1", type=['xlsx', 'csv'])
    report2 = st.file_uploader("رفع التقرير 2", type=['xlsx', 'csv'])
    hc_file = st.file_uploader("رفع ملف HC", type=['xlsx', 'csv'])
    
    output_filename = st.text_input("اسم ملف الإخراج (بدون امتداد)", "output_report")
    
    if st.button("إنشاء التقرير"):
        if not report1 or not report2 or not hc_file:
            st.warning("الرجاء رفع جميع الملفات المطلوبة")
        else:
            # معالجة الملفات وإنشاء المخرجات
            output_path = f"{output_filename}.xlsx"
            create_excel_template(output_path)
            
            # هنا يمكنك إضافة منطق المعالجة الفعلي للبيانات
            # حالياً، نقوم فقط بإنشاء قالب
            
            st.success("تم إنشاء التقرير بنجاح! يمكنك تنزيل الملف أدناه.")
            with open(output_path, "rb") as f:
                st.download_button(
                    label="تنزيل تقرير إكسل",
                    data=f,
                    file_name=output_path,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()
