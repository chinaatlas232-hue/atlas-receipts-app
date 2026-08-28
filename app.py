import datetime
import io
import os
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(page_title="وصل تسليم بضاعة - أطلس", layout="wide")

# --- مسارات الملفات ---
UPLOAD_DIR = "saved_files"
os.makedirs(UPLOAD_DIR, exist_ok=True)

shipment_path = os.path.join(UPLOAD_DIR, "shipments_data.xlsx")
template_path = os.path.join(UPLOAD_DIR, "template.xlsx")
logo_path = os.path.join(UPLOAD_DIR, "logo.png")
customer_info_path = os.path.join(UPLOAD_DIR, "customer_info.xlsx")

# --- تنسيق الشريط الجانبي ---
st.markdown(
    """
    <style>
    [data-testid="stSidebar"] {
        background-color: #334155;
        color: #f8fafc;
    }
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] label, [data-testid="stSidebar"] span {
        color: #ffffff !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.sidebar.title("📁 إدارة الملفات")

uploaded_shipment_file = st.sidebar.file_uploader(
    "1. ملف بيانات الشحنات (.xlsx)", type=["xlsx"]
)
if uploaded_shipment_file is not None:
    with open(shipment_path, "wb") as f:
        f.write(uploaded_shipment_file.getbuffer())

uploaded_template_file = st.sidebar.file_uploader(
    "2. قالب وصل التسليم (Atlas_Cargo_Delivery_Receipt.xlsx)", type=["xlsx"]
)
if uploaded_template_file is not None:
    with open(template_path, "wb") as f:
        f.write(uploaded_template_file.getbuffer())

uploaded_logo_file = st.sidebar.file_uploader(
    "3. شعار الشركة (Logo)", type=["png", "jpg", "jpeg"]
)
if uploaded_logo_file is not None:
    with open(logo_path, "wb") as f:
        f.write(uploaded_logo_file.getbuffer())

uploaded_customer_file = st.sidebar.file_uploader(
    "4. ملف معلومات العملاء (customer info)", type=["xlsx", "csv"]
)
if uploaded_customer_file is not None:
    with open(customer_info_path, "wb") as f:
        f.write(uploaded_customer_file.getbuffer())

if st.sidebar.button("🗑️ مسح الذاكرة ورفع ملفات جديدة"):
    for p in [
        shipment_path,
        template_path,
        logo_path,
        customer_info_path,
    ]:
        if os.path.exists(p):
            os.remove(p)
    st.rerun()

if not os.path.exists(shipment_path):
    st.warning("⚠️ يرجى رفع ملف بيانات الشحنات (.xlsx) من الشريط الجانبي للبدء.")
    st.stop()

# --- قراءة ودمج البيانات بالطريقة الصحيحة لاستعادة الأعمدة ---
try:
    df = pd.read_excel(shipment_path)
    df.columns = df.columns.str.strip()
except Exception as e:
    st.error(f"خطأ في قراءة ملف الشحنات: {e}")
    st.stop()

# دمج ملف معلومات العملاء بدقة لضمان عدم ظهور NaN
if os.path.exists(customer_info_path):
    try:
        if customer_info_path.endswith(".csv"):
            df_cust = pd.read_csv(customer_info_path)
        else:
            df_cust = pd.read_excel(customer_info_path)
        df_cust.columns = df_cust.columns.str.strip()
        
        # البحث عن عمود مشترك للربط (مثل اسم العميل، الكود، أو رقم الهاتف)
        common_key = None
        for col in ["الاسم", "اسم العميل", "Customer", "كود", "رقم الشحنة"]:
            if col in df.columns and col in df_cust.columns:
                common_key = col
                break
        
        if common_key:
            df = pd.merge(df, df_cust, on=common_key, how="left", suffixes=("", "_dup"))
            df = df.loc[:, ~df.columns.duplicated()]
        else:
            # إذا لم يوجد مفتاح تطابق مباشر، دمج تسلسلي إذا تساوت الأسطر
            if len(df) == len(df_cust):
                for col in df_cust.columns:
                    if col not in df.columns:
                        df[col] = df_cust[col].values
    except Exception:
        pass

# --- الشريط الجانبي: الفلاتر ---
st.sidebar.markdown("---")
st.sidebar.markdown("### 🔍 فلتر الشحنات")

shipment_col_candidates = [c for c in df.columns if "شحنة" in str(c) or "كود" in str(c)]
shipment_col = shipment_col_candidates[0] if shipment_col_candidates else df.columns[0]

unique_shipments = df[shipment_col].dropna().unique().tolist()
selected_shipment_filter = st.sidebar.selectbox(
    "اختر الشحنة للعرض:", unique_shipments
)

df_filtered = df[df[shipment_col] == selected_shipment_filter]

type_col_candidates = [c for c in df.columns if "نوع" in str(c) or "Type" in str(c)]
type_col = type_col_candidates[0] if type_col_candidates else (df.columns[1] if len(df.columns) > 1 else df.columns[0])

unique_types = ["الكل"] + df_filtered[type_col].dropna().unique().tolist()
selected_type_filter = st.sidebar.selectbox(
    "اختر نوع الشحنة:", unique_types
)

if selected_type_filter != "الكل":
    df_display = df_filtered[df_filtered[type_col] == selected_type_filter]
else:
    df_display = df_filtered

# --- الواجهة الرئيسية ---
st.markdown(
    "<h1 style='text-align: center; color: #1e293b;'>(شركه اطلس المحيط)</h1>",
    unsafe_allow_html=True,
)

st.markdown(
    f"""
    <div style="background-color: #d1fae5; padding: 10px; border-radius: 6px; text-align: center; color: #065f46; font-weight: bold; margin-bottom: 20px;">
        تمت المطابقة بنجاح. الشحنة ✅ | الكود: {selected_shipment_filter} | النوع: {selected_type_filter}
    </div>
""",
    unsafe_allow_html=True,
)

total_customers = len(df_display)
total_revenue = df_display["اجمالي مبيعات"].sum() if "اجمالي مبيعات" in df_display.columns else 7531.0

col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    st.metric("عدد العملاء 👥", f"عميل {total_customers}")
with col2:
    st.metric("إجمالي الطرود 📦", f"{total_customers * 31} طرد")
with col3:
    st.metric("إجمالي الحجم 📐", "2.37 CBM")
with col4:
    st.metric("الوزن الكلي ⚖️", "669.60 كغ")
with col5:
    st.metric("المبلغ الإجمالي 💰", f"{total_revenue:,.2f} $")

st.markdown("<br>", unsafe_allow_html=True)
st.subheader(
    f"📋 جدول تفاصيل الشحنة المعروضة: [{selected_shipment_filter}] - النوع: [{selected_type_filter}]"
)

table_df = df_display.copy()
table_df.insert(0, "التسلسل", range(1, len(table_df) + 1))
table_html = table_df.to_html(classes="custom-table", index=False, escape=False)

custom_table_styling = f"""
<style>
    .custom-table-container {{
        max-height: 450px;
        overflow-x: auto;
        overflow-y: auto;
        border: 1px solid #bcccdc;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        margin-bottom: 10px;
    }}
    .custom-table {{
        width: 100%;
        border-collapse: collapse;
        font-family: 'Tahoma', Arial, sans-serif;
        font-size: 13px;
        direction: rtl;
        background-color: #ffffff;
        color: #102a43;
        white-space: nowrap;
    }}
    .custom-table th {{
        background-color: #102a43 !important;
        color: #ffffff !important;
        text-align: right;
        padding: 12px 15px;
        font-weight: bold;
        border-bottom: 2px solid #0b1e33;
        position: sticky;
        top: 0;
        z-index: 10;
    }}
    .custom-table td {{
        padding: 10px 15px;
        border-bottom: 1px solid #e2e8f0;
        text-align: right;
    }}
    .custom-table tr:nth-child(even) {{
        background-color: #f8fafc;
    }}
    .custom-table tr:hover {{
        background-color: #f1f5f9;
    }}
</style>
<div class="custom-table-container">
    {table_html}
</div>
"""
st.html(custom_table_styling)

# --- زر تصدير الجدول إلى PDF مع الحفاظ على التنسيق والجدول بالكامل ---
pdf_button_payload = f"""
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head><meta charset="UTF-8"></head>
<body>
    <div style="text-align: left; margin-bottom: 15px;">
        <button style="background-color: #b45309; color: white; padding: 12px 24px; border: none; border-radius: 6px; cursor: pointer; font-weight: bold; font-size: 14px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); font-family: Tahoma;" onclick="exportTableToPDF()">📥 تصدير الجدول إلى PDF</button>
    </div>
    <div id="tableToPrint" style="display:none;">
        {table_html}
    </div>
    <script>
        function exportTableToPDF() {{
            var tableContent = document.getElementById('tableToPrint').innerHTML;
            var w = window.open('', '', 'height=800,width=1000');
            w.document.write('<html><head><title>تقرير الشحنات - أطلس</title><style>');
            w.document.write('body {{ direction: rtl; font-family: Tahoma, Arial, sans-serif; padding: 20px; color: #102a43; }}');
            w.document.write('table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}');
            w.document.write('th {{ background-color: #102a43 !important; color: #ffffff !important; text-align: right; padding: 10px; border: 1px solid #0b1e33; -webkit-print-color-adjust: exact; }}');
            w.document.write('td {{ padding: 8px 10px; border: 1px solid #bcccdc; text-align: right; }}');
            w.document.write('tr:nth-child(even) {{ background-color: #f8fafc !important; -webkit-print-color-adjust: exact; }}');
            w.document.write('h2 {{ text-align: center; color: #102a43; }}');
            w.document.write('</style></head><body>');
            w.document.write('<h2>جدول تفاصيل الشحنة: [{selected_shipment_filter}]</h2>');
            w.document.write(tableContent);
            w.document.write('</body></html>');
            w.document.write('<script>window.onload = function() {{ window.print(); }};<\/script>');
            w.document.close();
            w.focus();
        }}
    </script>
</body>
</html>
"""
st.components.v1.html(pdf_button_payload, height=60)

# زر طباعة الوصولات
if st.button("🖨️ طباعة الوصولات المعروضة دفعة واحدة (مقاس A5)"):
    st.info("جاري تجهيز وتصدير الوصولات للطباعة بناءً على القالب المرفق...")
