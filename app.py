import pandas as pd
import streamlit as st
from io import BytesIO

# إعداد صفحة Streamlit
st.set_page_config(page_title="نظام تحليل AHT", layout="wide")
st.title("📊 نظام تحليل وقت التعامل (AHT)")

# القاموس الشامل لأسماء أعمدة AHT المقبولة
AHT_COLUMN_NAMES = [
    'handle_time', 'aht', 'duration', 'call_duration',
    'talk_time', 'avg_handle_time', 'time_spent',
    'وقت_التعامل', 'مدة_المكالمة', 'الوقت_المنقضي'
]

# قسم رفع الملفات - في عمود جانبي
with st.sidebar:
    st.header("📤 رفع الملفات المطلوبة")
    raw_file = st.file_uploader("ملف البيانات الرئيسي", type=['xlsx', 'csv'], key='main_file')
    hc_file = st.file_uploader("ملف الموظفين (HC)", type=['xlsx', 'csv'], key='hc_file')
    
    st.header("⚙️ إعدادات التحليل")
    target_aht = st.number_input("هدف AHT (بالثواني)", min_value=0, value=300, step=1)
    selected_aht_column = st.selectbox("اختر عمود وقت التعامل", [])
    date_column = st.text_input("اسم عمود التاريخ (اختياري)", value="Date")

def detect_aht_column(df):
    """اكتشاف عمود AHT تلقائياً"""
    detected_columns = []
    for col in df.columns:
        col_lower = str(col).lower().strip()
        for aht_name in AHT_COLUMN_NAMES:
            if aht_name in col_lower:
                detected_columns.append(col)
                break
    return detected_columns

def analyze_aht(data, aht_column, target):
    """تحليل بيانات AHT"""
    analysis = {}
    
    # التحليل الأساسي
    analysis['avg_aht'] = data[aht_column].mean()
    analysis['median_aht'] = data[aht_column].median()
    analysis['aht_vs_target'] = analysis['avg_aht'] - target
    
    # تحليل حسب الموظف (إذا كان ملف HC موجوداً)
    if 'employee_id' in data.columns and 'team_leader' in data.columns:
        agent_stats = data.groupby(['team_leader', 'employee_id'])[aht_column].agg(['mean', 'count']).reset_index()
        agent_stats.columns = ['المشرف', 'رقم الموظف', 'متوسط AHT', 'عدد المكالمات']
        analysis['agent_stats'] = agent_stats.sort_values('متوسط AHT')
    
    return analysis

if raw_file:
    try:
        # قراءة الملف الرئيسي
        df = pd.read_excel(raw_file) if raw_file.name.endswith('.xlsx') else pd.read_csv(raw_file)
        
        # اكتشاف أعمدة AHT تلقائياً
        aht_columns = detect_aht_column(df)
        
        if not aht_columns:
            st.error("""
            **لم يتم العثور على عمود وقت التعامل (AHT) في الملف**
            
            الأعمدة المقبولة:
            - الإنجليزية: Handle_Time, AHT, Duration, Call_Time
            - العربية: وقت_التعامل، مدة_المكالمة
            
            **الأعمدة الموجودة في ملفك:**
            """ + str(list(df.columns)))
            
            st.download_button(
                label="⬇️ تنزيل نموذج لملف البيانات",
                data=pd.DataFrame(columns=['Employee', 'Handle_Time', 'Date']).to_csv(index=False),
                file_name="data_template.csv",
                mime="text/csv"
            )
            st.stop()
        
        # إذا وجدنا أعمدة AHT
        with st.sidebar:
            selected_aht_column = st.selectbox(
                "اختر عمود وقت التعامل",
                options=aht_columns,
                index=0,
                key='aht_column_selector'
            )
        
        # معالجة ملف HC إذا كان موجوداً
        if hc_file:
            df_hc = pd.read_excel(hc_file) if hc_file.name.endswith('.xlsx') else pd.read_csv(hc_file)
            # دمج البيانات مع ملف HC
            df = pd.merge(df, df_hc, on='employee_id', how='left')
        
        # تحليل البيانات
        analysis = analyze_aht(df, selected_aht_column, target_aht)
        
        # عرض النتائج
        st.header("📊 نتائج تحليل AHT")
        
        # المقاييس الرئيسية
        col1, col2, col3 = st.columns(3)
        col1.metric("متوسط AHT", f"{analysis['avg_aht']:.2f} ثانية")
        col2.metric("الوسيط", f"{analysis['median_aht']:.2f} ثانية")
        col3.metric("الفرق عن الهدف", f"{analysis['aht_vs_target']:.2f} ثانية", 
                   delta_color="inverse")
        
        # توزيع AHT
        st.subheader("توزيع أوقات التعامل")
        st.bar_chart(df[selected_aht_column].value_counts(bins=10))
        
        # نتائج الموظفين (إذا كانت متاحة)
        if 'agent_stats' in analysis:
            st.subheader("أداء الموظفين")
            st.dataframe(analysis['agent_stats'], height=400)
            
            # تصدير النتائج
            output = BytesIO()
            analysis['agent_stats'].to_excel(output, index=False)
            st.download_button(
                label="⬇️ تنزيل نتائج الموظفين",
                data=output.getvalue(),
                file_name="aht_results.xlsx",
                mime="application/vnd.ms-excel"
            )
    
    except Exception as e:
        st.error(f"حدث خطأ في معالجة البيانات: {str(e)}")
else:
    st.warning("الرجاء رفع ملف البيانات الرئيسي لبدء التحليل")
    st.info("""
    **ملاحظات مهمة:**
    1. تأكد أن الملف يحتوي على عمود لوقت التعامل (AHT)
    2. الأعمدة المقبولة: Handle_Time, AHT, Duration، أو ما يعادلها
    3. يمكنك تنزيل نموذج لملف البيانات إذا كنت غير متأكد من الهيكل المطلوب
    """)
