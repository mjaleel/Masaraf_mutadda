import streamlit as st
import pandas as pd
import re
from rapidfuzz import fuzz
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from io import BytesIO
import random

# ============ دوال المساعدة ============

def generate_color():
    """توليد ألوان مختلفة للأسماء المتشابهة"""
    colors = [
        "FFB3BA", "BAFFC9", "BAE1FF", "FFFFBA", "FFDFBA", "E0BBE4",
        "957DAD", "D4A5A5", "A8E6CF", "DCEDC1", "FFD3B6", "FFAAA5",
        "FF8B94", "A8D8EA", "AA96DA", "FCBAD3", "C9CBA3", "FFE66D",
        "F38181", "95E1D3", "EAFFD0", "FCE38A", "F54748", "7FE7CC"
    ]
    return random.choice(colors)

def normalize_name(name):
    """تطبيع الاسم للمطابقة"""
    if pd.isnull(name):
        return ""
    name = str(name).strip()
    name = name.replace("ه", "ة").replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    name = name.replace("ى", "ي").replace("ئ", "ي")
    name = re.sub(r'(عبد)([^\s])', r'\1 \2', name)
    return " ".join(name.split()).lower()

def get_first_three_words(name):
    """استخراج أول ثلاث كلمات من الاسم"""
    if pd.isnull(name) or name == "":
        return ""
    words = str(name).split()
    return " ".join(words[:3]) if len(words) >= 3 else " ".join(words)

def is_first_three_words_match(name1, name2):
    """التحقق من تطابق أول ثلاث كلمات"""
    words1 = name1.split()
    words2 = name2.split()
    length = min(len(words1), len(words2), 3)
    if length == 0:
        return False
    return all(words1[i] == words2[i] for i in range(length))

# ============ وظائف المطابقة المتقدمة ============

def find_best_match_in_db(normalized_name, database_map, name_column_db, threshold=85):
    """دالة داخلية للبحث عن أفضل مطابقة في قاعدة بيانات محددة"""
    best_match = None
    best_score = 0
    
    # محاولة المطابقة الدقيقة أو بناءً على Score
    for db_name in database_map.keys():
        score = fuzz.ratio(normalized_name, db_name)
        if score > best_score:
            best_score = score
            best_match = db_name

    # التحقق من الشروط الصارمة (نسبة التشابه + أول 3 كلمات)
    if best_score >= threshold and best_match:
        if is_first_three_words_match(normalized_name, best_match) or best_match.startswith(normalized_name):
            return best_match, best_score
            
    # محاولة البحث عن "يبدأ بـ" في حال فشل الـ Score
    for db_name in database_map.keys():
        if db_name.startswith(normalized_name) or normalized_name.startswith(db_name):
            return db_name, fuzz.ratio(normalized_name, db_name)
            
    return None, 0

def match_dual_databases(names_df, db1_df, db2_df, name_col_file, name_col_db1, name_col_db2, selected_cols_db1, selected_cols_db2):
    """منطق المطابقة المتسلسل بين قاعدتين بيانات"""
    names_df = names_df.copy()
    results = []
    
    # تجهيز القواعد (تطبيع وإزالة تكرار)
    db1_df["norm"] = db1_df[name_col_db1].apply(normalize_name)
    db2_df["norm"] = db2_df[name_col_db2].apply(normalize_name)
    
    db1_map = db1_df.drop_duplicates(subset=["norm"]).set_index("norm").to_dict(orient="index")
    db2_map = db2_df.drop_duplicates(subset=["norm"]).set_index("norm").to_dict(orient="index")

    for _, row in names_df.iterrows():
        original_name = row[name_col_file]
        norm_name = normalize_name(original_name)
        
        match_info = {
            "الاسم الأصلي": original_name,
            "الاسم المطابق": "",
            "نسبة التطابق": "",
            "المصدر": "❌ لم يتم العثور",
            "ملاحظة": "❌ غير متوفر"
        }
        # إضافة أعمدة فارغة افتراضياً
        for col in selected_cols_db1: match_info[f"[ق1] {col}"] = ""
        for col in selected_cols_db2: match_info[f"[ق2] {col}"] = ""

        # الخطوة 1: البحث في القاعدة الأولى (المصارف الموطنة)
        best_m1, score1 = find_best_match_in_db(norm_name, db1_map, name_col_db1)
        
        if best_m1:
            data = db1_map[best_m1]
            match_info.update({
                "الاسم المطابق": data[name_col_db1],
                "نسبة التطابق": f"{round(score1)}%",
                "المصدر": "✅ القاعدة الأولى (الموطنة)",
                "ملاحظة": "✅ تم التطابق"
            })
            for col in selected_cols_db1:
                match_info[f"[ق1] {col}"] = data.get(col, "")
        else:
            # الخطوة 2: البحث في القاعدة الثانية (الرافدين) إذا لم يجد في الأولى
            best_m2, score2 = find_best_match_in_db(norm_name, db2_map, name_col_db2)
            if best_m2:
                data = db2_map[best_m2]
                match_info.update({
                    "الاسم المطابق": data[name_col_db2],
                    "نسبة التطابق": f"{round(score2)}%",
                    "المصدر": "✅ القاعدة الثانية (الرافدين)",
                    "ملاحظة": "✅ تم التطابق"
                })
                for col in selected_cols_db2:
                    match_info[f"[ق2] {col}"] = data.get(col, "")
        
        results.append(match_info)
        
    return pd.DataFrame(results)

# ============ دوال التصدير (نفس الدوال السابقة مع تحسين بسيط) ============

def to_excel_basic(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
        wb = writer.book
        ws = wb.active
        red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
        for column in ws.columns:
            max_length = max(len(str(cell.value)) if cell.value else 0 for cell in column)
            ws.column_dimensions[column[0].column_letter].width = max_length + 2
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            if row[4].value and "❌" in str(row[4].value): # عمود الملاحظة
                for cell in row: cell.fill = red_fill
    output.seek(0)
    return output

# ============ واجهة المستخدم ============

st.set_page_config(page_title="نظام الرواتب الذكي", layout="wide")
st.markdown("### 👨‍💻 مبرمج النظام: محمد عبدالجليل")
st.title("🏦 نظام مطابقة بيانات الرواتب المتعدد")

password = st.sidebar.text_input("قفل النظام:", type="password")

if password == "mjaleel":
    tab1, tab2, tab3, tab4 = st.tabs([
        "📋 مطابقة بسيطة", 
        "🏢 مطابقة الأقسام", 
        "🔄 مطابقة ثلاثية",
        "🏦 مطابقة الرواتب (قاعدتين)"
    ])

    # (احتفظنا بالتابات القديمة كما هي في كودك الأصلي...)
    # سأقوم بكتابة محتوى Tab 4 المطلوب فقط للاختصار ولأنه التعديل الأساسي

    with tab4:
        st.subheader("🏦 البحث المتسلسل في قواعد بيانات المصارف")
        st.warning("⚠️ سيتم البحث أولاً في (القاعدة الأولى)، وإذا لم يتم العثور على تطابق، سيتم الانتقال تلقائياً للبحث في (قاعدة الرافدين).")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            f_names = st.file_uploader("📄 ملف الأسماء المطلوب مطابقتها", type="xlsx", key="f_names")
        with col2:
            f_db1 = st.file_uploader("📊 القاعدة 1 (المصارف الموطنة)", type="xlsx", key="f_db1")
        with col3:
            f_db2 = st.file_uploader("🏛️ القاعدة 2 (مصرف الرافدين)", type="xlsx", key="f_db2")

        if f_names and f_db1 and f_db2:
            df_names = pd.read_excel(f_names)
            df_db1 = pd.read_excel(f_db1)
            df_db2 = pd.read_excel(f_db2)

            st.markdown("---")
            c1, c2, c3 = st.columns(3)
            with c1:
                col_name_file = st.selectbox("اسم الموظف (الملف الأصلي):", df_names.columns)
            with c2:
                col_name_db1 = st.selectbox("اسم الموظف (قاعدة 1):", df_db1.columns)
            with c3:
                col_name_db2 = st.selectbox("اسم الموظف (قاعدة 2):", df_db2.columns)

            cc1, cc2 = st.columns(2)
            with cc1:
                sel_db1 = st.multiselect("أعمدة لجلبها من قاعدة 1 (مثلاً IBAN):", [c for c in df_db1.columns if c != col_name_db1])
            with cc2:
                sel_db2 = st.multiselect("أعمدة لجلبها من قاعدة 2 (مثلاً IBAN):", [c for c in df_db2.columns if c != col_name_db2])

            if st.button("🚀 بدء مطابقة الرواتب والآيبان", type="primary"):
                with st.spinner("جاري فحص القواعد..."):
                    results_df = match_dual_databases(
                        df_names, df_db1, df_db2,
                        col_name_file, col_name_db1, col_name_db2,
                        sel_db1, sel_db2
                    )
                
                st.success("✅ اكتملت العملية")
                
                # عرض إحصائيات دقيقة
                m_total = len(results_df)
                m_db1 = len(results_df[results_df["المصدر"].str.contains("الأولى")])
                m_db2 = len(results_df[results_df["المصدر"].str.contains("الثانية")])
                m_fail = m_total - (m_db1 + m_db2)
                
                sc1, sc2, sc3, sc4 = st.columns(4)
                sc1.metric("العدد الكلي", m_total)
                sc2.metric("من القاعدة 1", m_db1)
                sc3.metric("من القاعدة 2", m_db2)
                sc4.metric("لم يتم العثور", m_fail, delta_color="inverse")

                st.dataframe(results_df, use_container_width=True)
                
                excel_out = to_excel_basic(results_df)
                st.download_button(
                    "⬇️ تحميل كشف الرواتب النهائي",
                    excel_out,
                    file_name="مطابقة_الرواتب_الموحد.xlsx"
                )

    # ... يمكنك وضع التابات الأخرى هنا (Tab1, Tab2, Tab3 بنفس الكود السابق)

elif password:
    st.error("❌ كلمة المرور غير صحيحة.")
else:
    st.info("الرجاء إدخال كلمة المرور للوصول إلى صلاحيات المطابقة.")
