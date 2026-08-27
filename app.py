import datetime
import io
import os
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(page_title="وصل تسليم بضاعة - أطلس", layout="wide")

# --- تنسيق الألوان العام وتغيير لون الشريط الجانبي وزر المسح ---
st.markdown(
    """
    <style>
    /* لون الشريط الجانبي: رمادي غامق بدرجة متوسطة */
    [data-testid="stSidebar"] {
        background-color: #334155;
        color: #f8fafc;
    }
    /* تغيير لون النصوص والعناوين داخل الشريط الجانبي لتكون واضحة */
    [data-testid="stSidebar"] h1, 
    [data-testid="stSidebar"] h2, 
    [data-testid="stSidebar"] h3, 
    [data-testid="stSidebar"] label, 
    [data-testid="stSidebar"] .stMarkdown {
        color: #f8fafc !important;
    }
    /* تخصيص زر مسح الذاكرة ليكون أحمر غامق */
    [data-testid="stSidebar"] button[kind="secondary"] {
        background-color: #991b1b !important;
        color: white !important;
        border: 1px solid #7f1d1d !important;
        font-weight: bold !important;
    }
    [data-testid="stSidebar"] button[kind="secondary"]:hover {
        background-color: #7f1d1d !important;
        color: white !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.title("📦 النظام المالي والفني - وصل تسليم البضائع")

# إنشاء مجلد لحفظ الملفات بشكل دائم على السيرفر
UPLOAD_DIR = "saved_files"
os.makedirs(UPLOAD_DIR, exist_ok=True)

shipment_path = os.path.join(UPLOAD_DIR, "shipments_data.xlsx")
template_path = os.path.join(UPLOAD_DIR, "template.xlsx")
logo_path = os.path.join(UPLOAD_DIR, "logo.png")
customers_db_path = os.path.join(UPLOAD_DIR, "coustmer_info.xlsx")

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

  uploaded_cust_file = st.file_uploader(
      "2. ملف معلومات العملاء (coustmer info.xlsx)", type=["xlsx"]
  )
  if uploaded_cust_file is not None:
    with open(customers_db_path, "wb") as f:
      f.write(uploaded_cust_file.getbuffer())
    st.sidebar.success("تم حفظ ملف العملاء بنجاح!")

  uploaded_template_file = st.file_uploader(
      "3. قالب وصل التسليم (Atlas_Cargo_Delivery_Receipt.xlsx)", type=["xlsx"]
  )
  if uploaded_template_file is not None:
    with open(template_path, "wb") as f:
      f.write(uploaded_template_file.getbuffer())
    st.sidebar.success("تم حفظ القالب بنجاح!")

  uploaded_logo = st.file_uploader(
      "4. شعار الشركة (Logo) - اختيارى", type=["png", "jpg", "jpeg"]
  )
  if uploaded_logo is not None:
    with open(logo_path, "wb") as f:
      f.write(uploaded_logo.getbuffer())
    st.sidebar.success("تم حفظ الشعار بنجاح!")

  if st.button("🗑️ مسح الذاكرة ورفع ملفات جديدة"):
    for path in [shipment_path, template_path, logo_path, customers_db_path]:
      if os.path.exists(path):
        os.remove(path)
    st.sidebar.warning("تم مسح الملفات المحفوظة بنجاح.")
    st.rerun()

  # --- فلتر الشحنات في القائمة الجانبية ---
  st.markdown("---")
  st.header("🔍 فلتر الشحنات")
  selected_shipment_filter = "الكل"
  selected_code_filter = "الكل"
  selected_type_filter = "الكل"

  if os.path.exists(shipment_path):
    try:
      temp_df = pd.read_excel(shipment_path)
      temp_df.columns = temp_df.columns.str.strip()
      
      if "الشحنة" in temp_df.columns:
        temp_df["الشحنة"] = temp_df["الشحنة"].fillna("بدون شحنة").astype(str).str.replace(".0", "").str.strip()
        shipment_list = ["الكل"] + sorted(temp_df["الشحنة"].unique().tolist())
        selected_shipment_filter = st.selectbox(
            "اختر الشحنة للعرض:", shipment_list
        )
        
      filtered_temp_df = temp_df.copy()
      if selected_shipment_filter != "الكل" and "الشحنة" in filtered_temp_df.columns:
        filtered_temp_df = filtered_temp_df[
            filtered_temp_df["الشحنة"] == selected_shipment_filter
        ]

      if "الكود" in filtered_temp_df.columns:
        filtered_temp_df["الكود"] = filtered_temp_df["الكود"].fillna("بدون كود").astype(str).str.replace(".0", "").str.strip()
        code_list = ["الكل"] + sorted(filtered_temp_df["الكود"].unique().tolist())
        selected_code_filter = st.selectbox(
            "اختر أو ابحث برقم الكود:", code_list
        )
        
      if selected_code_filter != "الكل" and "الكود" in filtered_temp_df.columns:
        filtered_temp_df = filtered_temp_df[
            filtered_temp_df["الكود"] == selected_code_filter
        ]

      # فلتر نوع الشحنة (بحري / جوي)
      type_col = None
      for col in ["نوع الشحنة", "النوع"]:
        if col in filtered_temp_df.columns:
          type_col = col
          break
          
      if type_col:
        filtered_temp_df[type_col] = filtered_temp_df[type_col].fillna("غير محدد").astype(str).str.strip()
        type_list = ["الكل"] + sorted(filtered_temp_df[type_col].unique().tolist())
        selected_type_filter = st.selectbox(
            "اختر نوع الشحنة:", type_list
        )

    except Exception:
      pass

active_data_file = shipment_path if os.path.exists(shipment_path) else None
active_template_file = template_path if os.path.exists(template_path) else None
active_logo = logo_path if os.path.exists(logo_path) else None
active_customers_db = customers_db_path if os.path.exists(customers_db_path) else None

if active_data_file is not None and active_template_file is not None:
  try:
    df = pd.read_excel(active_data_file)
    df.columns = df.columns.str.strip()

    # --- المعالجة الدقيقة والمطابقة الذكية لبيانات العملاء من coustmer info.xlsx ---
    if active_customers_db is not None:
      try:
        cust_df = pd.read_excel(active_customers_db)
        cust_df.columns = cust_df.columns.str.strip()
        
        cust_code_col = None
        for col in cust_df.columns:
          c_lower = str(col).lower()
          if any(k in c_lower for k in ["new code", "code", "ats", "كود"]):
            cust_code_col = col
            break
        if not cust_code_col and len(cust_df.columns) > 1:
          cust_code_col = cust_df.columns[1]

        cust_name_col = None
        for col in cust_df.columns:
          c_lower = str(col).lower()
          if any(k in c_lower for k in ["name surname", "name", "surname", "الاسم", "اسم"]):
            cust_name_col = col
            break
        if not cust_name_col and len(cust_df.columns) > 2:
          cust_name_col = cust_df.columns[2]

        cust_phone_col = None
        for col in cust_df.columns:
          c_lower = str(col).lower()
          if any(k in c_lower for k in ["phone 1", "phone", "tel", "mobile", "هاتف", "موبايل", "رقم"]):
            cust_phone_col = col
            break

        cust_address_col = None
        for col in cust_df.columns:
          c_lower = str(col).lower()
          if any(k in c_lower for k in ["address", "عنوان", "استلام"]):
            cust_address_col = col
            break

        if cust_code_col:
          cust_df["clean_code"] = cust_df[cust_code_col].fillna("").astype(str).str.strip().str.replace(".0", "").str.lower()
          
          name_dict = {}
          phone_dict = {}
          address_dict = {}
          
          for _, c_row in cust_df.iterrows():
            c_code = c_row["clean_code"]
            if c_code and c_code != "nan" and c_code != "":
              if cust_name_col and pd.notna(c_row.get(cust_name_col)):
                val_n = str(c_row[cust_name_col]).strip()
                if val_n and val_n.lower() != "nan":
                  name_dict[c_code] = val_n
                  
              if cust_phone_col and pd.notna(c_row.get(cust_phone_col)):
                val_p = str(c_row[cust_phone_col]).strip()
                if val_p and val_p.lower() != "nan":
                  phone_dict[c_code] = val_p
                  
              if cust_address_col and pd.notna(c_row.get(cust_address_col)):
                val_a = str(c_row[cust_address_col]).strip()
                if val_a and val_a.lower() != "nan":
                  address_dict[c_code] = val_a

          for col_name in ["الاسم", "رقم الهاتف", "عنوان استلام البظاعة"]:
            if col_name not in df.columns:
              df[col_name] = ""

          for idx, row in df.iterrows():
            raw_code_val = str(row.get("الكود", "")).strip().replace(".0", "").lower()
            if raw_code_val and raw_code_val != "nan" and raw_code_val != "بدون كود" and raw_code_val != "":
              
              curr_name = str(row.get("الاسم", "")).strip()
              if (not curr_name or curr_name.lower() in ["nan", "none", ""]) and raw_code_val in name_dict:
                df.at[idx, "الاسم"] = name_dict[raw_code_val]
                
              curr_phone = str(row.get("رقم الهاتف", "")).strip()
              if (not curr_phone or curr_phone.lower() in ["nan", "none", ""]) and raw_code_val in phone_dict:
                df.at[idx, "رقم الهاتف"] = phone_dict[raw_code_val]

              curr_addr = str(row.get("عنوان استلام البظاعة", "")).strip()
              if (not curr_addr or curr_addr.lower() in ["nan", "none", ""]) and raw_code_val in address_dict:
                df.at[idx, "عنوان استلام البظاعة"] = address_dict[raw_code_val]
      except Exception as ex:
        st.sidebar.warning(f"ملاحظة حول قراءة قاعدة بيانات العملاء: {ex}")

    if "الشحنة" in df.columns:
      df["الشحنة"] = df["الشحنة"].fillna("بدون شحنة").astype(str).str.replace(".0", "").str.strip()
    if "الكود" in df.columns:
      df["الكود"] = df["الكود"].fillna("بدون كود").astype(str).str.replace(".0", "").str.strip()

    type_col_name = None
    for col in ["نوع الشحنة", "النوع"]:
      if col in df.columns:
        type_col_name = col
        break
        
    if type_col_name:
      df[type_col_name] = df[type_col_name].fillna("غير محدد").astype(str).str.strip()

    if selected_shipment_filter != "الكل" and "الشحنة" in df.columns:
      df = df[df["الشحنة"] == selected_shipment_filter]

    if selected_code_filter != "الكل" and "الكود" in df.columns:
      df = df[df["الكود"] == selected_code_filter]

    if selected_type_filter != "الكل" and type_col_name:
      df = df[df[type_col_name] == selected_type_filter]

    if df.empty:
      st.warning("⚠️ لا توجد بيانات مطابقة للفلاتر المحددة.")
      st.stop()

    today_date = datetime.date.today().strftime("%Y-%m-%d")

    import base64
    logo_base64 = ""
    if active_logo and os.path.exists(active_logo):
      with open(active_logo, "rb") as img_file:
        logo_base64 = base64.b64encode(img_file.read()).decode("utf-8")

    st.success(
        f"✅ تم دمج البيانات بنجاح وسحب الحقول المفقودة. الشحنة المعروضة:"
        f" **{selected_shipment_filter}** | الكود: **{selected_code_filter}** | النوع: **{selected_type_filter}**"
    )
    st.markdown("---")

    total_clients_count = len(df)
    total_packages_count = 0
    total_weight_sum = 0.0
    total_cbm_sum = 0.0
    total_sales_sum = 0.0

    receipts_data_list = []
    all_receipts_html_for_print = ""

    for index, row in df.iterrows():
      shipment = str(row.get("الشحنة", "بدون شحنة")).strip()
      code = str(row.get("الكود", "بدون كود")).strip()
      
      display_code = "" if code == "بدون كود" else code
      display_shipment = "" if shipment == "بدون شحنة" else shipment

      name = ""
      for col in ["الاسم", "الاسم "]:
        if col in df.columns and pd.notna(row.get(col)):
          name = str(row.get(col)).strip()
          break
      if not name or name.lower() == "nan":
        name = "عميل غير محدد"

      file_name_id = (
          f"Shipment_{shipment}_Client_{name}"
          if shipment and name
          else (shipment if shipment else f"Receipt_{index}")
      )

      weight = float(row.get("الوزن", 0) or 0)
      total_weight_sum += weight

      cbm_value = 0.0
      for col in df.columns:
        col_lower = str(col).lower()
        if any(k in col_lower for k in ["حجم", "cbm", "dimension", "C.B.M", "الحجم"]):
          val = row.get(col, 0)
          if pd.notna(val):
            try:
              cbm_value = float(val)
            except:
              pass
            break
      total_cbm_sum += cbm_value

      packages = 0
      try:
        packages = int(row.get("عدد الطرود", 0) or 0)
      except Exception:
        packages = 0
      total_packages_count += packages

      price_per_kg = 0
      for col in ["سعر الكيلو", "سعر الكيلو ", "السعر"]:
        if col in df.columns:
          val = row.get(col, 0)
          if pd.notna(val):
            try:
              price_per_kg = float(val)
            except:
              pass
            break

      total_sales = 0
      for col in ["اجمالي مبيعات", "اجمالي مبيعات ", "الاجمالي", "المبلغ"]:
        if col in df.columns:
          val = row.get(col, 0)
          if pd.notna(val):
            try:
              total_sales = float(val)
            except:
              pass
            break
      if total_sales == 0 and price_per_kg > 0 and weight > 0:
        total_sales = weight * price_per_kg

      total_sales_sum += total_sales

      phone_raw = ""
      for col in ["رقم الهاتف", "رقم الهاتف "]:
        if col in df.columns:
          phone_raw = row.get(col, "")
          break
      phone = str(phone_raw).strip()
      if phone.endswith(".0"):
        phone = phone[:-2]
      phone = phone.replace("+", "").strip()
      if phone.startswith("964"):
        phone = phone[3:]
      formatted_phone = f"+964 {phone}" if phone and phone.lower() != "nan" else ""

      address = ""
      for col in ["عنوان استلام البظاعة", "العنوان", "عنوان"]:
        if col in df.columns:
          address = str(row.get(col, "")).strip()
          break
      if address.lower() == "nan":
        address = ""

      shipment_type = ""
      if type_col_name:
        shipment_type = str(row.get(type_col_name, "")).strip()
      if shipment_type.lower() == "nan":
        shipment_type = ""

      wb = openpyxl.load_workbook(active_template_file)
      ws = wb.active

      ws["B4"] = display_code
      ws["D4"] = today_date
      ws["B5"] = name
      ws["B6"] = address
      ws["D5"] = formatted_phone
      ws["B7"] = display_shipment
      ws["D6"] = packages
      ws["B8"] = shipment_type
      ws["D7"] = weight

      output = io.BytesIO()
      wb.save(output)
      output.seek(0)

      logo_img_tag = f'<img src="data:image/png;base64,{logo_base64}" style="max-height: 52px; max-width: 60px; margin-left: 10px; vertical-align: middle; mix-blend-mode: multiply; filter: contrast(120%) brightness(105%);">' if logo_base64 else ''

      single_receipt_html = f"""
            <div class="receipt-page" style="
                padding: 15px; 
                font-family: 'Tahoma', Arial, sans-serif; 
                direction: rtl; 
                border: 2px solid #102a43; 
                width: 100%; 
                max-width: 148mm; 
                margin: auto auto 20px auto; 
                background: #ffffff; 
                color: #102a43;
                box-sizing: border-box;
                page-break-after: always;
                break-after: page;
            ">
                <table style="width: 100%; border-bottom: 2px solid #102a43; padding-bottom: 8px; margin-bottom: 12px; border-collapse: collapse;">
                    <tr>
                        <td style="text-align: right; vertical-align: middle;">
                            <div style="display: flex; align-items: center;">
                                {logo_img_tag}
                                <div>
                                    <h2 style="margin: 0; font-size: 15px; color: #102a43;">أطلس المحيط للتجارة العامة</h2>
                                    <p style="margin: 2px 0 0; font-size: 10px; color: #627d98;">OCEAN ATLAS GENERAL TRADING</p>
                                </div>
                            </div>
                        </td>
                        <td style="text-align: left; vertical-align: middle;">
                            <h3 style="margin: 0; font-size: 13px; color: #b45309;">وصل تسليم بضاعة</h3>
                            <p style="margin: 2px 0 0; font-size: 10px; color: #334e68;">Cargo Delivery Receipt</p>
                        </td>
                    </tr>
                </table>

                <table style="width: 100%; font-size: 11px; border-collapse: collapse; margin-bottom: 10px;">
                    <tr style="background-color: #f0f4f8;">
                        <td style="padding: 5px; border: 1px solid #bcccdc; width: 50%;"><strong>كود العميل:</strong> <span style="color: #b45309; font-weight: bold;">{display_code}</span></td>
                        <td style="padding: 5px; border: 1px solid #bcccdc; width: 50%;"><strong>رقم الشحنة:</strong> <span style="color: #b45309; font-weight: bold;">{display_shipment}</span></td>
                    </tr>
                    <tr>
                        <td style="padding: 5px; border: 1px solid #bcccdc;"><strong>اسم العميل:</strong> <span style="font-weight: bold;">{name}</span></td>
                        <td style="padding: 5px; border: 1px solid #bcccdc;"><strong>رقم الهاتف:</strong> <span style="direction: ltr; display: inline-block; font-weight: bold;">{formatted_phone}</span></td>
                    </tr>
                    <tr style="background-color: #f0f4f8;">
                        <td style="padding: 5px; border: 1px solid #bcccdc;"><strong>عنوان الاستلام:</strong> <span style="color: #486581; font-weight: bold;">{address}</span></td>
                        <td style="padding: 5px; border: 1px solid #bcccdc;"><strong>عدد الطرود:</strong> 📦 {packages} طرد</td>
                    </tr>
                    <tr>
                        <td style="padding: 5px; border: 1px solid #bcccdc;"><strong>تاريخ الإصدار:</strong> <span style="color: #b45309; font-weight: bold;">{today_date}</span></td>
                        <td style="padding: 5px; border: 1px solid #bcccdc;"><strong>الوزن الإجمالي:</strong> <span style="color: #102a43; font-weight: bold;">{weight} كغ</span></td>
                    </tr>
                    <tr style="background-color: #f0f4f8;">
                        <td style="padding: 5px; border: 1px solid #bcccdc;"><strong>نوع الشحنة:</strong> {shipment_type}</td>
                        <td style="padding: 5px; border: 1px solid #bcccdc;"><strong>حجم الشحنة (CBM):</strong> <span style="color: #b45309; font-weight: bold;">{cbm_value}</span></td>
                    </tr>
                    <tr>
                        <td style="padding: 5px; border: 1px solid #bcccdc;" colspan="2"><strong>سعر الكيلو:</strong> {price_per_kg:,.2f} $</td>
                    </tr>
                    <tr style="background-color: #fef3c7;">
                        <td style="padding: 5px; border: 1px solid #f59e0b;" colspan="2"><strong>إجمالي المبيعات:</strong> <span style="color: #b45309; font-weight: bold; font-size: 12px;">{total_sales:,.2f} $</span> &nbsp;&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;&nbsp; <strong>طريقة الدفع:</strong> [ &nbsp; ] نقداً &nbsp;&nbsp; [ &nbsp; ] أجل</td>
                    </tr>
                </table>

                <div style="background-color: #fffbeb; border: 1px solid #fde68a; padding: 6px; border-radius: 4px; margin-bottom: 10px;">
                    <p style="margin: 0; font-size: 10px; color: #92400e; line-height: 1.3;">
                        <strong>إقرار الاستلام:</strong><br>
                        أقر أنا الموقع أدناه، بأنني استلمت البضاعة والشحنة المذكورة أعلاه كاملة، وبحالة سليمة وممتازة، ومطابقة لكافة الأوزان والأوصاف المدونة.
                    </p>
                </div>

                <table style="width: 100%; font-size: 11px; margin-top: 5px; margin-bottom: 10px;">
                    <tr>
                        <td style="width: 50%; padding: 2px;">
                            <strong>اسم المستلم:</strong><br><br>
                            ............................................
                        </td>
                        <td style="width: 50%; padding: 2px; text-align: left;">
                            <strong>توقيع وختم المستلم:</strong><br><br>
                            ............................................
                        </td>
                    </tr>
                </table>

                <div style="border-top: 1px dashed #bcccdc; margin-top: 10px; padding-top: 6px; text-align: center; font-size: 9.5px; color: #334e68;">
                    <span>📍 العنوان: بغداد - المنصور - تقاطع الواد</span>
                    <span style="margin: 0 10px;">|</span>
                    <span style="direction: ltr; display: inline-block;">📞 هاتف: 07858588899 / 07814518989</span>
                </div>
            </div>
            """

      all_receipts_html_for_print += single_receipt_html
      receipts_data_list.append({
          "index": index,
          "name": name,
          "code": display_code,
          "shipment": display_shipment,
          "total_sales": total_sales,
          "output": output,
          "file_name_id": file_name_id,
          "single_html": single_receipt_html,
      })

    st.markdown(
        """
        <style>
        .metric-card-1 { background-color: #eff6ff; border: 1px solid #bfdbfe; padding: 15px; border-radius: 10px; text-align: center; }
        .metric-card-2 { background-color: #f0fdf4; border: 1px solid #bbf7d0; padding: 15px; border-radius: 10px; text-align: center; }
        .metric-card-3 { background-color: #f5f3ff; border: 1px solid #ddd6fe; padding: 15px; border-radius: 10px; text-align: center; }
        .metric-card-4 { background-color: #fffbeb; border: 1px solid #fde68a; padding: 15px; border-radius: 10px; text-align: center; }
        .metric-card-5 { background-color: #fdf2f8; border: 1px solid #fbcfe8; padding: 15px; border-radius: 10px; text-align: center; }
        </style>
    """,
        unsafe_allow_html=True,
    )

    m1, m2, m3, m4, m5 = st.columns(5)
    with m1:
      st.markdown(
          f"""
            <div class="metric-card-1">
                <p style="margin: 0; color: #1e40af; font-weight: bold; font-size: 14px;">👥 عدد العملاء</p>
                <h3 style="margin: 5px 0 0; color: #1e3a8a; font-size: 20px;">{total_clients_count} عميل</h3>
            </div>
        """,
          unsafe_allow_html=True,
      )
    with m2:
      st.markdown(
          f"""
            <div class="metric-card-2">
                <p style="margin: 0; color: #166534; font-weight: bold; font-size: 14px;">📦 إجمالي الطرود</p>
                <h3 style="margin: 5px 0 0; color: #14532d; font-size: 20px;">{total_packages_count} طرد</h3>
            </div>
        """,
          unsafe_allow_html=True,
      )
    with m3:
      st.markdown(
          f"""
            <div class="metric-card-3">
                <p style="margin: 0; color: #5b21b6; font-weight: bold; font-size: 14px;">📐 إجمالي الحجم</p>
                <h3 style="margin: 5px 0 0; color: #4c1d95; font-size: 20px;">{total_cbm_sum:,.2f} CBM</h3>
            </div>
        """,
          unsafe_allow_html=True,
      )
    with m4:
      st.markdown(
          f"""
            <div class="metric-card-4">
                <p style="margin: 0; color: #92400e; font-weight: bold; font-size: 14px;">⚖️ الوزن الكلي</p>
                <h3 style="margin: 5px 0 0; color: #78350f; font-size: 20px;">{total_weight_sum:,.2f} كغ</h3>
            </div>
        """,
          unsafe_allow_html=True,
      )
    with m5:
      st.markdown(
          f"""
            <div class="metric-card-5">
                <p style="margin: 0; color: #9d174d; font-weight: bold; font-size: 14px;">💰 المبلغ الإجمالي</p>
                <h3 style="margin: 5px 0 0; color: #831843; font-size: 20px;">{total_sales_sum:,.2f} $</h3>
            </div>
        """,
          unsafe_allow_html=True,
      )

    st.markdown("<br>", unsafe_allow_html=True)

    # --- جدول تفاصيل الشحنة ---
    st.subheader(f"📋 جدول تفاصيل الشحنة المعروضة: [{selected_shipment_filter}] - النوع: [{selected_type_filter}]")

    display_table_df = df.copy()
    display_table_df.insert(0, "التسلسل", range(1, len(display_table_df) + 1))

    table_html = display_table_df.to_html(
        classes="custom-table", index=False, escape=False
    )

    custom_table_styling = f"""
    <style>
        .custom-table-container {{
            max-height: 650px;
            overflow-x: auto;
            overflow-y: auto;
            border: 1px solid #bcccdc;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            margin-bottom: 20px;
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
    st.markdown("---")

    master_payload = f"""
        <!DOCTYPE html>
        <html lang="ar" dir="rtl">
        <head>
            <meta charset="UTF-8">
            <style>
                body {{
                    font-family: Tahoma, sans-serif;
                    background-color: #ffffff;
                    margin: 0;
                    padding: 10px;
                    direction: rtl;
                }}
                .master-btn {{
                    background-color: #047857;
                    color: white;
                    padding: 14px 28px;
                    border: none;
                    border-radius: 8px;
                    cursor: pointer;
                    font-weight: bold;
                    font-size: 16px;
                    width: 100%;
                    max-width: 500px;
                    display: block;
                    margin: 0 auto;
                    box-shadow: 0 4px 6px rgba(0,0,0,0.1);
                    text-align: center;
                }}
                .master-btn:hover {{
                    background-color: #065f46;
                }}
            </style>
        </head>
        <body>
            <button class="master-btn" onclick="printAllReceipts()">
                🖨️ طباعة الوصولات المعروضة دفعة واحدة (مقاس A5)
            </button>

            <script>
                const allReceiptsContent = `{all_receipts_html_for_print.replace('`', '\\`').replace('$', '\\$')}`;

                function printAllReceipts() {{
                    var printWin = window.open('', '', 'height=900,width=800');
                    printWin.document.write('<html><head><title>طباعة الوصولات</title><style>@page {{ size: A5; margin: 5mm; }} body {{ direction: rtl; font-family: Tahoma, sans-serif; background: #fff; margin: 0; padding: 0; }} .receipt-page {{ page-break-after: always; break-after: page; margin-bottom: 20px; }}</style></head><body>');
                    printWin.document.write(allReceiptsContent);
                    printWin.document.write('</body></html>');
                    printWin.document.close();
                    printWin.focus();
                    setTimeout(function(){{ printWin.print(); printWin.close(); }}, 600);
                }}
            </script>
        </body>
        </html>
        """

    st.components.v1.html(master_payload, height=75)
    st.markdown("---")

    for item in receipts_data_list:
      index = item["index"]
      shipment = item["shipment"]
      file_name_id = item["file_name_id"]
      single_html_content = item["single_html"]

      expander_label = f"📄 وصل العميل: {item['name']}  |  كود: {item['code'] if item['code'] else 'بدون'}  |  الشحنة: {shipment}  |  الإجمالي: {item['total_sales']:,.2f} دولار"

      with st.expander(expander_label, expanded=False):
        st.download_button(
            label=f"📥 تنزيل إكسل الوصل (الشحنة: {shipment})",
            data=item["output"],
            file_name=f"Delivery_Receipt_{file_name_id}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            key=f"download_excel_{index}",
        )

        st.markdown("<br>", unsafe_allow_html=True)

        safe_html_payload = f"""
            <!DOCTYPE html>
            <html lang="ar" dir="rtl">
            <head>
                <meta charset="UTF-8">
                <style>
                    body {{
                        font-family: Tahoma, sans-serif;
                        background-color: #ffffff;
                        margin: 0;
                        padding: 5px;
                        direction: rtl;
                    }}
                    .btn-container {{
                        display: flex;
                        gap: 12px;
                        margin-top: 15px;
                        margin-bottom: 10px;
                    }}
                    .action-btn {{
                        background-color: #102a43;
                        color: white;
                        padding: 12px 20px;
                        border: none;
                        border-radius: 6px;
                        cursor: pointer;
                        font-weight: bold;
                        font-size: 14px;
                        flex: 1;
                        text-align: center;
                    }}
                    .pdf-btn {{
                        background-color: #b45309;
                    }}
                </style>
            </head>
            <body>
                {single_html_content}
                
                <div class="btn-container">
                    <button class="action-btn" onclick="printReceipt()">
                        🖨️ طباعة الوصل (ورقيّاً)
                    </button>
                    <button class="action-btn pdf-btn" onclick="savePdfReceipt()">
                        📑 حفظ PDF (مقاس A5)
                    </button>
                </div>

                <script>
                    const receiptContent = `{single_html_content.replace('`', '\\`').replace('$', '\\$')}`;
                    const fileNameId = '{file_name_id}';

                    function printReceipt() {{
                        var printWin = window.open('', '', 'height=800,width=800');
                        printWin.document.write('<html><head><title>طباعة الشحنة</title><style>@page {{ size: A5; margin: 5mm; }} body {{ direction: rtl; font-family: Tahoma, sans-serif; background: #fff; margin: 0; padding: 0; }}</style></head><body>');
                        printWin.document.write(receiptContent);
                        printWin.document.write('</body></html>');
                        printWin.document.close();
                        printWin.focus();
                        setTimeout(function(){{ printWin.print(); printWin.close(); }}, 500);
                    }}

                    function savePdfReceipt() {{
                        var printWin = window.open('', '', 'height=800,width=800');
                        printWin.document.write('<html><head><title>' + fileNameId + '</title><style>@page {{ size: A5; margin: 5mm; }} body {{ direction: rtl; font-family: Tahoma, sans-serif; background: #fff; margin: 0; padding: 0; }}</style></head><body>');
                        printWin.document.write(receiptContent);
                        printWin.document.write('</body></html>');
                        printWin.document.close();
                        printWin.focus();
                        setTimeout(function(){{ 
                            printWin.document.title = fileNameId;
                            printWin.print(); 
                        }}, 600);
                    }}
                </script>
            </body>
            </html>
            """

        st.components.v1.html(
            safe_html_payload, height=750, scrolling=True
        )

      st.markdown("---")

  except Exception as e:
    st.error(f"حدث خطأ أثناء قراءة أو معالجة الملفات: {e}")
else:
  st.info(
      "الرجاء رفع ملف بيانات الشحنات وقالب الوصل وملف معلومات العملاء من الشريط الجانبي لتظهر المعاينة والطباعة."
  )
