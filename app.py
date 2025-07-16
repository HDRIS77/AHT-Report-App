import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils.dataframe import dataframe_to_rows
import streamlit as st
import numpy as np

def calculate_metrics(df_report1, df_report2, df_hc):
    """دالة لحساب المقاييس المطلوبة"""
    # هنا نقوم بدمج البيانات وإجراء الحسابات
    # يمكنك تعديل هذه الدالة حسب احتياجاتك الدقيقة
    
    # مثال لحساب بعض القيم (يجب تعديله حسب منطق عملك)
    df_combined = pd.concat([df_report1, df_report2])
    
    # حساب المتوسطات
    aht_score = df_combined['AHT'].mean() if 'AHT' in df_combined.columns else 0
    chats_arabic = df_combined['Arabic_Chats'].sum() if 'Arabic_Chats' in df_combined.columns else 0
    chats_urdu = df_combined['Urdu_Chats'].sum() if 'Urdu_Chats' in df_combined.columns else 0
    
    # حساب النسب المئوية
    pass_rate = (df_combined['Pass'].sum() / df_combined['Chats'].sum()) * 100 if 'Chats' in df_combined.columns else 0
    fail_rate = (df_combined['Fail'].sum() / df_combined['Chats'].sum()) * 100 if 'Chats' in df_combined.columns else 0
    
    return {
        'aht_score': aht_score,
        'chats_arabic': chats_arabic,
        'chats_urdu': chats_urdu,
        'pass_rate': pass_rate,
        'fail_rate': fail_rate,
        'var_from_target': aht_score - 300  # مثال: assuming target is 300
    }

def create_excel_report(data, output_path):
    """إنشاء تقرير إكسل مع القيم المحسوبة"""
    wb = Workbook()
    
    # إزالة الورقة الافتراضية
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']
    
    # إنشاء ورقة View
    view_sheet = wb.create_sheet("View")
    
    # تعريف الأنماط
    header_fill = PatternFill(start_color="FF7030A0", end_color="FF7030A0", fill_type="solid")
    light_purple_fill = PatternFill(start_color="FFE4D9F5", end_color="FFE4D9F5", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))
    center_alignment = Alignment(horizontal='center', vertical='center')
    header_font = Font(bold=True, color="FFFFFF")
    
    # عناوين ورقة View
    view_headers = [
        "TL", "Urdu", "Arabic", "Over all AHT Score", "Tenured AHT", "Nesting AHT", 
        "# Chats Arabic", "# Chats Urdu", "Var From Target", "Status", "Readiness", 
        "FRT", "EG", "JO", "BH", "AE", "QA", "KW", "OM", "Releasing", "ZTP", 
        "Missed", "", "Chats", "Pass", "Fail"
    ]
    
    # إضافة العناوين
    for col_num, header in enumerate(view_headers, 1):
        cell = view_sheet.cell(row=1, column=col_num, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
        cell.alignment = center_alignment
    
    # تعبئة البيانات المحسوبة
    # قسم TL
    view_sheet.cell(row=2, column=4, value=data['aht_score']).number_format = '0.00'
    view_sheet.cell(row=2, column=7, value=data['chats_arabic'])
    view_sheet.cell(row=2, column=8, value=data['chats_urdu'])
    view_sheet.cell(row=2, column=9, value=data['var_from_target']).number_format = '0.00'
    view_sheet.cell(row=2, column=24, value=data['chats_arabic'] + data['chats_urdu'])
    view_sheet.cell(row=2, column=25, value=round((data['pass_rate']/100) * (data['chats_arabic'] + data['chats_urdu'])))
    view_sheet.cell(row=2, column=26, value=round((data['fail_rate']/100) * (data['chats_arabic'] + data['chats_urdu']))
    
    # قسم BPO (مثال)
    view_sheet.cell(row=26, column=4, value=data['aht_score']).number_format = '0.00'
    view_sheet.cell(row=26, column=7, value=data['chats_arabic'])
    view_sheet.cell(row=26, column=8, value=data['chats_urdu'])
    view_sheet.cell(row=26, column=24, value=data['chats_arabic'] + data['chats_urdu'])
    view_sheet.cell(row=26, column=25, value=round((data['pass_rate']/100) * (data['chats_arabic'] + data['chats_urdu'])))
    view_sheet.cell(row=26, column=26, value=round((data['fail_rate']/100) * (data['chats_arabic'] + data['chats_urdu'])))
    
    # حفظ الملف
    wb.save(output_path)

def main():
    st.title("نظام تحليل أداء خدمة العملاء")
    st.write("يرجى رفع الملفات المطلوبة لإنشاء التقرير")
    
    # رفع الملفات
    report1 = st.file_uploader("التقرير الأول (Excel/CSV)", type=['xlsx', 'csv'])
    report2 = st.file_uploader("التقرير الثاني (Excel/CSV)", type=['xlsx', 'csv'])
    hc_file = st.file_uploader("ملف HC (Excel/CSV)", type=['xlsx', 'csv'])
    
    if st.button("إنشاء التقرير"):
        if report1 and report2 and hc_file:
            try:
                # قراءة الملفات
                df1 = pd.read_excel(report1) if report1.name.endswith('.xlsx') else pd.read_csv(report1)
                df2 = pd.read_excel(report2) if report2.name.endswith('.xlsx') else pd.read_csv(report2)
                df_hc = pd.read_excel(hc_file) if hc_file.name.endswith('.xlsx') else pd.read_csv(hc_file)
                
                # حساب المقاييس
                metrics = calculate_metrics(df1, df2, df_hc)
                
                # إنشاء التقرير
                output_path = "customer_performance_report.xlsx"
                create_excel_report(metrics, output_path)
                
                # عرض النتائج
                st.success("تم إنشاء التقرير بنجاح!")
                st.write("ملخص النتائج:")
                st.write(f"متوسط AHT: {metrics['aht_score']:.2f}")
                st.write(f"عدد المحادثات العربية: {metrics['chats_arabic']}")
                st.write(f"عدد المحادثات الأردية: {metrics['chats_urdu']}")
                st.write(f"معدل النجاح: {metrics['pass_rate']:.2f}%")
                
                # زر التنزيل
                with open(output_path, "rb") as f:
                    st.download_button(
                        label="تنزيل التقرير",
                        data=f,
                        file_name=output_path,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"حدث خطأ: {str(e)}")
        else:
            st.warning("الرجاء رفع جميع الملفات المطلوبة")

if __name__ == "__main__":
    main()
