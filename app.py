import datetime
import io
import os
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(page_title="وصل تسليم بضاعة - أطلس", layout="wide")

st.markdown(
    """
    <style>
    [data-testid="stSidebar"] {
        background-color: #334155;
        color: #f8fafc;
    }
    [data-testid="stSidebar"] h1, 
    [data-testid="stSidebar"] h2, 
    [data-testid="stSidebar"] h3, 
    [data-testid="stSidebar"] label, 
    [data-testid="stSidebar"] .stMarkdown {
        color: #f8fafc !important;
    }
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

UPLOAD_DIR = "saved_files"
os.makedirs(UPLOAD_DIR, exist_ok=True)

shipment_path = os.path.join(UPLOAD_DIR, "shipments_data.xlsx")
template_path = os.path.join(UPLOAD_DIR, "template.xlsx")
logo_path = os.path.join(UPLOAD_DIR, "logo.png")

selected_shipment_filter = "الكل"
selected_code_filter = "الكل"
manual_volume_col = "تلقائي"

with st.sidebar:
  st.header("⚙️ إدارة الملفات")

  uploaded_data_file = st.file_uploader(
      "1. ملف بيانات الشحنات (تعبئة وصل اطلس.xlsx)", type=["xlsx"]
  )
  if uploaded_data_file is not None:
    with open(shipment_path, "wb") as f:
      f.write(uploaded_data_file.getbuffer())
    st.sidebar.success("تم حفظ ملف الشحنات بنجاح!")
    st.rerun()

  # إظهار قائمة مطابقة الحجم إجبارياً فور توفر ملف الشحنات
  if os.path.exists(shipment_path):
    try:
      temp_df = pd.read_excel(shipment_path)
      temp_df.columns = temp_df.columns.astype(str).str.strip()

      st.markdown("---")
      st.markdown(
          "<h3 style='color: #fbbf24; font-size:15px;'>🛠️ مطابقة عمود الحجم"
          " يدوياً</h3>",
          unsafe_allow_html=True,
      )
      cols_available = ["تلقائي"] + list(temp_df.columns)
      manual_volume_col = st.selectbox(
          "اختر العمود الذي يمثل الحجم (CBM):", cols_available
      )
    except Exception as e:
      st.sidebar.error(f"خطأ في قراءة الأعمدة: {e}")

  uploaded_template_file = st.file_uploader(
      "2. قالب وصل التسليم (Atlas_Cargo_Delivery_Receipt.xlsx)", type=["xlsx"]
  )
  if uploaded_template_file is not None:
    with open(template_path, "wb") as f:
      f.write(uploaded_template_file.getbuffer())
    st.sidebar.success("تم حفظ القالب بنجاح!")
    st.rerun()

  uploaded_logo = st.file_uploader(
      "3. شعار الشركة (Logo) - اختيارى", type=["png", "jpg", "jpeg"]
  )
  if uploaded_logo is not None:
    with open(logo_path, "wb") as f:
      f.write(uploaded_logo.getbuffer())
    st.sidebar.success("تم حفظ الشعار بنجاح!")
    st.rerun()

  if st.button("🗑️ مسح الذاكرة ورفع ملفات جديدة"):
    for path in [shipment_path, template_path, logo_path]:
      if os.path.exists(path):
        os.remove(path)
    st.sidebar.warning("تم مسح الملفات المحفوظة بنجاح.")
    st.rerun()

  if os.path.exists(shipment_path):
    try:
      temp_df = pd.read_excel(shipment_path)
      temp_df.columns = temp_df.columns.astype(str).str.strip()

      st.markdown("---")
      st.header("🔍 فلتر الشحنات")

      if "الشحنة" in temp_df.columns:
        shipment_list = ["الكل"] + sorted(
            temp_df["الشحنة"]
            .dropna()
            .astype(str)
            .str.replace(".0", "")
            .unique()
            .tolist()
        )
        selected_shipment_filter = st.selectbox(
            "اختر الشحنة للعرض:", shipment_list
        )

      filtered_temp_df = temp_df.copy()
      if (
          selected_shipment_filter != "الكل"
          and "الشحنة" in filtered_temp_df.columns
      ):
        filtered_temp_df = filtered_temp_df[
            filtered_temp_df["الشحنة"].astype(str).str.replace(".0", "")
            == selected_shipment_filter
        ]

      if "الكود" in filtered_temp_df.columns:
        code_list = ["الكل"] + sorted(
            filtered_temp_df["الكود"]
            .dropna()
            .astype(str)
            .str.replace(".0", "")
            .unique()
            .tolist()
        )
        selected_code_filter = st.selectbox(
            "اختر أو ابحث برقم الكود:", code_list
        )
    except Exception:
      pass

active_data_file = shipment_path if os.path.exists(shipment_path) else None
active_template_file = template_path if os.path.exists(template_path) else None
active_logo = logo_path if os.path.exists(logo_path) else None

if active_data_file is not None and active_template_file is not None:
  try:
    df = pd.read_excel(active_data_file)
    df.columns = df.columns.astype(str).str.strip()

    # معالجة عمود الحجم بناءً على الاختيار اليدوي أو البحث التلقائي
    if manual_volume_col != "تلقائي" and manual_volume_col in df.columns:
      df.rename(columns={manual_volume_col: "الحجم"}, inplace=True)
    else:
      found_vol_col = None
      for col in df.columns:
        clean_c = col.lower().strip()
        if any(
            k in clean_c
            for k in [
                "cbm",
                "vol",
                "volume",
                "حجم",
                "فاليوم",
                "فاليم",
                "حج",
                "الابعاد",
            ]
        ):
          found_vol_col = col
          break
      if found_vol_col:
        df.rename(columns={found_vol_col: "الحجم"}, inplace=True)
      else:
        df["الحجم"] = 0.0

    if selected_shipment_filter != "الكل" and "الشحنة" in df.columns:
      df = df[
          df["الشحنة"].astype(str).str.replace(".0", "")
          == selected_shipment_filter
      ]

    if selected_code_filter != "الكل" and "الكود" in df.columns:
      df = df[
          df["الكود"].astype(str).str.replace(".0", "")
          == selected_code_filter
      ]

    if df.empty:
      st.warning("⚠️ لا توجد بيانات مطابقة للشحنة أو الكود المحدد.")
      st.stop()

    today_date = datetime.date.today().strftime("%Y-%m-%d")

    total_clients_count = len(df)
    total_packages_count = 0
    total_weight_sum = 0.0
    total_volume_sum = 0.0
    total_sales_sum = 0.0

    receipts_data_list = []
    all_receipts_html_for_print = ""

    for index, row in df.iterrows():
      shipment = str(row.get("الشحنة", "")).strip()
      if shipment.endswith(".0"):
        shipment = shipment[:-2]

      code = str(row.get("الكود", "")).strip()
      if code.endswith(".0"):
        code = code[:-2]

      name = str(row.get("الاسم", ""))
      if name.endswith(".0"):
        name = name[:-2]
      name = name.strip()

      file_name_id = (
          f"Shipment_{shipment}_Client_{name}"
          if shipment and name
          else (shipment if shipment else f"Receipt_{index}")
      )

      weight = float(row.get("الوزن", 0) or 0)
      total_weight_sum += weight

      volume = 0.0
      val = row.get("الحجم", 0)
      if pd.notna(val):
        try:
          volume = float(val)
        except:
          volume = 0.0
      total_volume_sum += volume

      packages = 0
      try:
        packages = int(row.get("عدد الطرود", 0) or 0)
      except Exception:
        packages = 0
      total_packages_count += packages

      price_per_kg = 0
      for col in ["سعر الكيلو", "السعر"]:
        if col in df.columns:
          val = row.get(col, 0)
          if pd.notna(val):
            price_per_kg = float(val)
            break

      total_sales = 0
      for col in ["اجمالي مبيعات", "الاجمالي", "المبلغ"]:
        if col in df.columns:
          val = row.get(col, 0)
          if pd.notna(val):
            total_sales = float(val)
            break
      if total_sales == 0 and price_per_kg > 0 and weight > 0:
        total_sales = weight * price_per_kg

      total_sales_sum += total_sales

      phone_raw = row.get("رقم الهاتف", "")
      phone = str(phone_raw).strip()
      if phone.endswith(".0"):
        phone = phone[:-2]
      phone = phone.replace("+", "").strip()
      if phone.startswith("964"):
        phone = phone[3:]
      formatted_phone = f"+964 {phone}" if phone else ""

      address = str(
          row.get("عنوان استلام البظاعة", row.get("العنوان", ""))
      ).strip()
      shipment_type = str(row.get("نوع الشحنة", "")).strip()

      wb = openpyxl.load_workbook(active_template_file)
      ws = wb.active

      ws["B4"] = code
      ws["D4"] = today_date
      ws["B5"] = name
      ws["B6"] = address
      ws["D5"] = formatted_phone
      ws["B7"] = shipment
      ws["D6"] = packages
      ws["B8"] = shipment_type
      ws["D7"] = weight

      output = io.BytesIO()
      wb.save(output)
      output.seek(0)

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
                            <div>
                                <h2 style="margin: 0; font-size: 15px; color: #102a43;">أطلس المحيط للتجارة العامة</h2>
                                <p style="margin: 2px 0 0; font-size: 10px; color: #627d98;">OCEAN ATLAS GENERAL TRADING</p>
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
                        <td style="padding: 5px; border: 1px solid #bcccdc; width: 50%;"><strong>كود العميل:</strong> <span style="color: #b45309; font-weight: bold;">{code}</span></td>
                        <td style="padding: 5px; border: 1px solid #bcccdc; width: 50%;"><strong>رقم الشحنة:</strong> <span style="color: #b45309; font-weight: bold;">{shipment}</span></td>
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
                        <td style="padding: 5px; border: 1px solid #bcccdc;"><strong>الوزن الإجمالي:</strong> {weight} كغ &nbsp;|&nbsp; <strong>الحجم:</strong> {volume}</td>
                    </tr>
                    <tr style="background-color: #f0f4f8;">
                        <td style="padding: 5px; border: 1px solid #bcccdc;"><strong>نوع الشحنة:</strong> {shipment_type}</td>
                        <td style="padding: 5px; border: 1px solid #bcccdc;"><strong>سعر الكيلو:</strong> {price_per_kg:,.2f} $</td>
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

                <table style="width: 100%; font-size: 11px; margin-top: 5px;">
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
            </div>
            """

      all_receipts_html_for_print += single_receipt_html
      receipts_data_list.append({
          "index": index,
          "name": name,
          "code": code,
          "shipment": shipment,
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
        .metric-card-3 { background-color: #fffbeb; border: 1px solid #fde68a; padding: 15px; border-radius: 10px; text-align: center; }
        .metric-card-5 { background-color: #f3e8ff; border: 1px solid #d8b4fe; padding: 15px; border-radius: 10px; text-align: center; }
        .metric-card-4 { background-color: #fdf2f8; border: 1px solid #fbcfe8; padding: 15px; border-radius: 10px; text-align: center; }
        </style>
    """,
        unsafe_allow_html=True,
    )

    m1, m2, m3, m5, m4 = st.columns(5)
    with m1:
      st.markdown(
          f"""
            <div class="metric-card-1">
                <p style="margin: 0; color: #1e40af; font-weight: bold; font-size: 13px;">👥 عدد العملاء</p>
                <h3 style="margin: 5px 0 0; color: #1e3a8a; font-size: 18px;">{total_clients_count} عميل</h3>
            </div>
        """,
          unsafe_allow_html=True,
      )
    with m2:
      st.markdown(
          f"""
            <div class="metric-card-2">
                <p style="margin: 0; color: #166534; font-weight: bold; font-size: 13px;">📦 إجمالي الطرود</p>
                <h3 style="margin: 5px 0 0; color: #14532d; font-size: 18px;">{total_packages_count} طرد</h3>
            </div>
        """,
          unsafe_allow_html=True,
      )
    with m3:
      st.markdown(
          f"""
            <div class="metric-card-3">
                <p style="margin: 0; color: #92400e; font-weight: bold; font-size: 13px;">⚖️ الوزن الكلي</p>
                <h3 style="margin: 5px 0 0; color: #78350f; font-size: 18px;">{total_weight_sum:,.2f} كغ</h3>
            </div>
        """,
          unsafe_allow_html=True,
      )
    with m5:
      st.markdown(
          f"""
            <div class="metric-card-5">
                <p style="margin: 0; color: #6b21a8; font-weight: bold; font-size: 13px;">📐 الحجم الكلي</p>
                <h3 style="margin: 5px 0 0; color: #581c87; font-size: 18px;">{total_volume_sum:,.2f} CBM</h3>
            </div>
        """,
          unsafe_allow_html=True,
      )
    with m4:
      st.markdown(
          f"""
            <div class="metric-card-4">
                <p style="margin: 0; color: #9d174d; font-weight: bold; font-size: 13px;">💰 المبلغ الإجمالي</p>
                <h3 style="margin: 5px 0 0; color: #831843; font-size: 18px;">{total_sales_sum:,.2f} $</h3>
            </div>
        """,
          unsafe_allow_html=True,
      )

    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader(f"📋 جدول تفاصيل الشحنة المعروضة: [{selected_shipment_filter}]")

    display_table_df = df.copy()
    display_table_df.insert(
        0, "التسلسل", range(1, len(display_table_df) + 1)
    )

    preferred_cols = [
        "التسلسل",
        "الشحنة",
        "الكود",
        "الاسم",
        "رقم الهاتف",
        "عدد الطرود",
        "الوزن",
        "الحجم",
        "سعر الكيلو",
        "اجمالي مبيعات",
        "عنوان استلام البظاعة",
        "نوع الشحنة",
    ]
    actual_cols = [
        col for col in preferred_cols if col in display_table_df.columns
    ]

    final_table_df = (
        display_table_df[actual_cols] if actual_cols else display_table_df
    )
    table_html = final_table_df.to_html(
        classes="custom-table", index=False, escape=False
    )

    custom_table_styling = f"""
    <style>
        .custom-table-container {{
            max-height: 450px;
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
            <div id="master-content" style="display:none;">{all_receipts_html_for_print}</div>
            <button class="master-btn" onclick="printAllReceipts()">
                🖨️ طباعة الوصولات المعروضة دفعة واحدة (مقاس A5)
            </button>

            <script>
                function printAllReceipts() {{
                    var content = document.getElementById('master-content').innerHTML;
                    var printWin = window.open('', '', 'height=900,width=800');
                    printWin.document.write('<html><head><title>طباعة الوصولات</title><style>@page {{ size: A5; margin: 5mm; }} body {{ direction: rtl; font-family: Tahoma, sans-serif; background: #fff; margin: 0; padding: 0; }} .receipt-page {{ page-break-after: always; break-after: page; margin-bottom: 20px; }}</style></head><body>');
                    printWin.document.write(content);
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

      expander_label = f"📄 وصل العميل: {item['name']}  |  كود: {item['code']}  |  الشحنة: {shipment}  |  الإجمالي: {item['total_sales']:,.2f} دولار"

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

        single_html_content = item["single_html"]
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
                <div id="single-receipt-container">{single_html_content}</div>
                
                <div class="btn-container">
                    <button class="action-btn" onclick="printReceipt()">
                        🖨️ طباعة الوصل (ورقيّاً)
                    </button>
                    <button class="action-btn pdf-btn" onclick="savePdfReceipt()">
                        📑 حفظ PDF (مقاس A5)
                    </button>
                </div>

                <script>
                    const fileNameId = '{file_name_id}';

                    function printReceipt() {{
                        var content = document.getElementById('single-receipt-container').innerHTML;
                        var printWin = window.open('', '', 'height=800,width=800');
                        printWin.document.write('<html><head><title>طباعة الشحنة</title><style>@page {{ size: A5; margin: 5mm; }} body {{ direction: rtl; font-family: Tahoma, sans-serif; background: #fff; margin: 0; padding: 0; }}</style></head><body>');
                        printWin.document.write(content);
                        printWin.document.write('</body></html>');
                        printWin.document.close();
                        printWin.focus();
                        setTimeout(function(){{ printWin.print(); printWin.close(); }}, 500);
                    }}

                    function savePdfReceipt() {{
                        var content = document.getElementById('single-receipt-container').innerHTML;
                        var printWin = window.open('', '', 'height=800,width=800');
                        printWin.document.write('<html><head><title>' + fileNameId + '</title><style>@page {{ size: A5; margin: 5mm; }} body {{ direction: rtl; font-family: Tahoma, sans-serif; background: #fff; margin: 0; padding: 0; }}</style></head><body>');
                        printWin.document.write(content);
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
      "الرجاء رفع ملف بيانات الشحنات وقالب الوصل من الشريط الجانبي لتظهر المعاينة"
      " والطباعة."
  )
