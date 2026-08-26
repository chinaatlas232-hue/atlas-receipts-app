import os
import streamlit as st

# إنشاء مجلد لحفظ الملفات بشكل دائم على السيرفر إذا لم يكن موجوداً
UPLOAD_DIR = "saved_files"
os.makedirs(UPLOAD_DIR, exist_ok=True)

# مسارات الملفات الثابتة
shipment_path = os.path.join(UPLOAD_DIR, "shipments_data.xlsx")
template_path = os.path.join(UPLOAD_DIR, "template.xlsx")
logo_path = os.path.join(UPLOAD_DIR, "logo.png")

st.sidebar.header("⚙️ إدارة الملفات")

# 1. ملف بيانات الشحنات
uploaded_shipment = st.sidebar.file_uploader(
    "1. ملف بيانات الشحنات (تعبئة وصل اطلس.xlsx)", type=["xlsx"]
)
if uploaded_shipment is not None:
  with open(shipment_path, "wb") as f:
    f.write(uploaded_shipment.getbuffer())
  st.sidebar.success("تم حفظ ملف الشحنات بنجاح!")

# التحقق مما إذا كان الملف مخزناً مسبقاً على السيرفر
active_shipment = (
    shipment_path
    if os.path.exists(shipment_path) and uploaded_shipment is None
    else uploaded_shipment
)


# 2. قالب وصل التسليم
uploaded_template = st.sidebar.file_uploader(
    "2. قالب وصل التسليم (Atlas_Cargo_Delivery_Receipt.xlsx)", type=["xlsx"]
)
if uploaded_template is not None:
  with open(template_path, "wb") as f:
    f.write(uploaded_template.getbuffer())
  st.sidebar.success("تم حفظ القالب بنجاح!")

active_template = (
    template_path
    if os.path.exists(template_path) and uploaded_template is None
    else uploaded_template
)


# 3. شعار الشركة
uploaded_logo = st.sidebar.file_uploader(
    "3. شعار الشركة (Logo) - اختيارى", type=["png", "jpg", "jpeg"]
)
if uploaded_logo is not None:
  with open(logo_path, "wb") as f:
    f.write(uploaded_logo.getbuffer())
  st.sidebar.success("تم حفظ الشعار بنجاح!")

active_logo = (
    logo_path
    if os.path.exists(logo_path) and uploaded_logo is None
    else uploaded_logo
)
