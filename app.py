import pandas as pd
import streamlit as st
from io import BytesIO

# إعداد صفحة Streamlit
st.set_page_config(page_title="نظام تحليل AHT", layout="wide")
st.title("📊 نظام تحليل وقت التعامل (AHT)")

# قاموس موسع لأسماء أعمدة AHT المقبولة (بما في ذلك handling_time)
AHT_COLUMN_NAMES = [
    'handle_time', 'handling_time', 'aht', 'duration', 'call_duration',
    'talk_time', 'avg_handle_time', 'time_spent', 'service_time',
    'وقت_التعامل', 'مدة_المكالمة', 'الوقت_المنقضي', 'زمن_الخدمة'
]

# قسم رفع الملفات
with st.sidebar:
    st.header("📤 رفع الملفات المطلوبة")
    raw_file = st.file_uploader("ملف البيانات الرئيسي", type=['xlsx', 'csv'])
    hc_file = st.file_uploader("ملف الموظفين (HC) - اختياري", type=['xlsx', 'csv'])
    
    st.header("⚙️ إعدادات التحليل")
    target_aht = st.number_input("هدف AHT (بالثواني)", min_value=0, value=300, step=1)

def detect_aht_column(df):
    """اكتشاف عمود AHT تلقائياً مع دعم handling_time"""
    detected = []
    for col in df.columns:
        col_normalized = str(col).lower().strip().replace(' ', '_')
        for aht_name in AHT_COLUMN_NAMES:
            if aht_name in col_normalized:
                detected.append(col)
                break
    return detected

def analyze_data(df, aht_column, target):
    """تحليل بيانات AHT مع معالجة متقدمة"""
    results = {}
    
    # تنظيف البيانات
    df[aht_column] = pd.to_numeric(df[aht_column], errors='coerce')
    df = df.dropna(subset=[aht_column])
    
    # التحليلات الأساسية
    results['avg'] = df[aht_column].mean()
    results['median'] = df[aht_column].median()
    results['target_diff'] = results['avg'] - target
    
    # تحليل حسب المشرف (إذا كان HC موجوداً)
    if hc_file and 'employee_id' in df.columns:
        hc_data = pd.read_excel(hc_file) if hc_file.name.endswith('.xlsx') else pd.read_csv(hc_file)
        df = pd.merge(df, hc_data, on='employee_id')
        
        if 'team_leader' in df.columns:
            leader_stats = df.groupby('team_leader')[aht_column].agg(['mean', 'count', 'std'])
            results['leader_stats'] = leader_stats.sort_values('mean')
    
    return results, df

if raw_file:
    try:
        # قراءة الملف
        df = pd.read_excel(raw_file) if raw_file.name.endswith('.xlsx') else pd.read_csv(raw_file)
        
        # اكتشاف أعمدة AHT
        aht_columns = detect_aht_column(df)
        
        if not aht_columns:
            st.error("""
            **لم يتم العثور على عمود وقت التعامل في الملف**
            
            الأعمدة المقبولة تشمل:
            - handling_time
            - handle_time
            - duration
            - وقت_التعامل
            
            **الأعمدة الموجودة في ملفك:**
            """ + str(list(df.columns)))
            
            # نموذج للتنزيل
            sample_df = pd.DataFrame({
                'employee_id': [1001, 1002],
                'handling_time': [120, 180],
                'date': pd.to_datetime(['2023-01-01', '2023-01-02'])
            })
            
            st.download_button(
                label="⬇️ تنزيل نموذج ملف بيانات",
                data=sample_df.to_csv(index=False),
                file_name="aht_data_template.csv",
                mime="text/csv"
            )
            st.stop()
        
        # إذا وجدنا أعمدة AHT
        selected_aht = st.sidebar.selectbox(
            "اختر عمود وقت التعامل",
            options=aht_columns,
            index=0
        )
        
        # التحليل
        results, processed_df = analyze_data(df, selected_aht, target_aht)
        
        # عرض النتائج
        st.header("📈 نتائج تحليل وقت التعامل")
        
        # المقاييس الرئيسية
        cols = st.columns(3)
        cols[0].metric("المتوسط", f"{results['avg']:.1f} ثانية")
        cols[1].metric("الوسيط", f"{results['median']:.1f} ثانية")
        cols[2].metric("الفرق عن الهدف", 
                      f"{results['target_diff']:.1f} ثانية",
                      delta_color="inverse")
        
        # التوزيع
        st.subheader("توزيع أوقات التعامل")
        st.bar_chart(processed_df[selected_aht].value_counts(bins=10))
        
        # نتائج المشرفين (إذا متاحة)
        if 'leader_stats' in results:
            st.subheader("أداء الفرق حسب المشرفين")
            st.dataframe(results['leader_stats'].style
                        .background_gradient(cmap='Blues')
                        .format("{:.1f}"))
            
            # تصدير النتائج
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                results['leader_stats'].to_excel(writer, sheet_name='Team_Leaders')
                processed_df.to_excel(writer, sheet_name='Raw_Data', index=False)
            
            st.download_button(
                label="⬇️ تنزيل التقرير الكامل",
                data=output.getvalue(),
                file_name="aht_full_report.xlsx",
                mime="application/vnd.ms-excel"
            )
    
    except Exception as e:
        st.error(f"حدث خطأ: {str(e)}")
else:
    st.info("""
    **تعليمات الاستخدام:**
    1. ارفع ملف البيانات الرئيسي (يجب أن يحتوي على عمود وقت التعامل)
    2. (اختياري) ارفع ملف الموظفين (HC) لتحليل أداء الفرق
    3. حدد هدف AHT المطلوب
    4. استعرض النتائج وقم بتنزيل التقارير
    """)
