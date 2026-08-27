# إنشاء مجلد لحفظ الملفات بشكل دائم على السيرفر
UPLOAD_DIR = "saved_files"
os.makedirs(UPLOAD_DIR, exist_ok=True)

shipment_path = os.path.join(UPLOAD_DIR, "shipments_data.xlsx")
template_path = os.path.join(UPLOAD_DIR, "template.xlsx")
logo_path = os.path.join(UPLOAD_DIR, "logo.png")
# تحديد مسار حفظ ملف معلومات العملاء الجديد
customer_info_path = os.path.join(UPLOAD_DIR, "customer_info.xlsx")

# شريط جانبي لإدارة الملفات والفلتر
with st.sidebar:
  st.header("⚙️ إدارة الملفات")

  uploaded_data_file = st.file_uploader(
      "1. ملف بيانات الشحنات (تعبئة وصل اطلس.xlsx)", type=["xlsx"]
  )
  if uploaded_data_file is not None:
    with open(shipment_path, "wb") as f:
      f.write(uploaded_data_file.getbuffer())
    st.sidebar.success("تم حفظ ملف الشحنات بنجاح!")

  uploaded_template_file = st.file_uploader(
      "2. قالب وصل التسليم (Atlas_Cargo_Delivery_Receipt.xlsx)", type=["xlsx"]
  )
  if uploaded_template_file is not None:
    with open(template_path, "wb") as f:
      f.write(uploaded_template_file.getbuffer())
    st.sidebar.success("تم حفظ القالب بنجاح!")

  uploaded_logo = st.file_uploader(
      "3. شعار الشركة (Logo) - اختيارى", type=["png", "jpg", "jpeg"]
  )
  if uploaded_logo is not None:
    with open(logo_path, "wb") as f:
      f.write(uploaded_logo.getbuffer())
    st.sidebar.success("تم حفظ الشعار بنجاح!")

  # --- زر رفع معلومات العملاء الجديد ---
  uploaded_customer_file = st.file_uploader("coustemr info", type=["xlsx", "csv"])
  if uploaded_customer_file is not None:
    with open(customer_info_path, "wb") as f:
      f.write(uploaded_customer_file.getbuffer())
    st.sidebar.success("تم حفظ ملف معلومات العملاء (coustemr info) بنجاح!")

  if st.button("🗑️ مسح الذاكرة ورفع ملفات جديدة"):
    for path in [
        shipment_path,
        template_path,
        logo_path,
        customer_info_path,
    ]:
      if os.path.exists(path):
        os.remove(path)
    st.sidebar.warning("تم مسح الملفات المحفوظة بنجاح.")
    st.rerun()
