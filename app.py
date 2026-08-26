import datetime
import io
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(page_title="وصل تسليم بضاعة - أطلس", layout="wide")

st.title("📦 النظام المالي والفني - وصل تسليم البضائع")

# تهيئة Session State لحفظ الملفات والبيانات حتى لا تضيع عند التحديث
if "df_saved" not in st.session_state:
  st.session_state.df_saved = None
if "template_saved" not in st.session_state:
  st.session_state.template_saved = None
if "logo_saved" not in st.session_state:
  st.session_state.logo_saved = None

# شريط جانبي لإدارة الملفات المرفوعة وحفظها
with st.sidebar:
  st.header("⚙️ إدارة الملفات")

  uploaded_data_file = st.file_uploader(
      "1. ملف بيانات الشحنات (تعبئة وصل اطلس.xlsx)", type=["xlsx"]
  )
  if uploaded_data_file is not None:
    st.session_state.df_saved = uploaded_data_file

  uploaded_template_file = st.file_uploader(
      "2. قالب وصل التسليم (Atlas_Cargo_Delivery_Receipt.xlsx)", type=["xlsx"]
  )
  if uploaded_template_file is not None:
    st.session_state.template_saved = uploaded_template_file

  uploaded_logo = st.file_uploader(
      "3. شعار الشركة (Logo) - اختيارى", type=["png", "jpg", "jpeg"]
  )
  if uploaded_logo is not None:
    st.session_state.logo_saved = uploaded_logo

  if st.button("🗑️ مسح الذاكرة ورفع ملفات جديدة"):
    st.session_state.df_saved = None
    st.session_state.template_saved = None
    st.session_state.logo_saved = None
    st.rerun()

# استخدام الملفات المحفوظة في الذاكرة المؤقتة للجلسة
active_data_file = st.session_state.df_saved
active_template_file = st.session_state.template_saved
active_logo = st.session_state.logo_saved

if active_data_file is not None and active_template_file is not None:
  try:
    # قراءة بيانات الشحنات
    df = pd.read_excel(active_data_file)
    df.columns = df.columns.str.strip()

    # الحصول على تاريخ اليوم تلقائياً
    today_date = datetime.date.today().strftime("%Y-%m-%d")

    st.success(
        f"✅ الملفات محفوظة في الذاكرة ومحمية ضد التحديث. تاريخ الإصدار:"
        f" {today_date}"
    )
    st.markdown("---")

    # معالجة الشعار بدون مربعات أو حدود خلفية
    logo_html = ""
    if active_logo is not None:
      import base64

      bytes_data = active_logo.getvalue()
      base64_logo = base64.b64encode(bytes_data).decode("utf-8")
      logo_html = f'<img src="data:image/png;base64,{base64_logo}" style="height: 45px; width: 45px; object-fit: cover; border-radius: 50%; border: none; outline: none; background: transparent; margin-left: 10px; vertical-align: middle;" />'

    all_receipts_html_for_print = ""
    receipts_data_list = []

    # حلقة تكرارية لكل صف في ملف البيانات
    for index, row in df.iterrows():
      # استخراج البيانات الأساسية
      shipment = str(row.get("الشحنة", "")).strip()
      if shipment.endswith(".0"):
        shipment = shipment[:-2]

      code = str(row.get("الكود", "")).strip()
      if code.endswith(".0"):
        code = code[:-2]

      name = str(row.get("الاسم", row.get("الاسم ", ""))).strip()

      # معرف فريد للملف
      file_name_id = f"Shipment_{shipment}_Client_{name}" if shipment and name else (shipment if shipment else f"Receipt_{index}")

      weight = float(row.get("الوزن", 0) or 0)
      packages = row.get("عدد الطرود", 0)

      # استخراج سعر الكيلو وحسابه بأمان
      price_per_kg = 0
      for col in ["سعر الكيلو", "سعر الكيلو ", "السعر"]:
        if col in df.columns:
          val = row.get(col, 0)
          if pd.notna(val):
            price_per_kg = float(val)
            break

      # استخراج إجمالي المبيعات أو حسابه تلقائياً
      total_sales = 0
      for col in ["اجمالي مبيعات", "اجمالي مبيعات ", "الاجمالي", "المبلغ"]:
        if col in df.columns:
          val = row.get(col, 0)
          if pd.notna(val):
            total_sales = float(val)
            break
      if total_sales == 0 and price_per_kg > 0 and weight > 0:
        total_sales = weight * price_per_kg

      # تنظيف رقم الهاتف وإضافة مسافة بين +964 وباقي الرقم
      phone_raw = row.get("رقم الهاتف", row.get("رقم الهاتف ", ""))
      phone = str(phone_raw).strip()
      if phone.endswith(".0"):
        phone = phone[:-2]

      phone = phone.replace("+", "").strip()
      if phone.startswith("964"):
        phone = phone[3:]

      formatted_phone = f"+964 {phone}" if phone else ""

      # استخراج العنوان ونوع الشحنة مع التحقق من وجود الأعمدة
      address = ""
      for col in ["عنوان استلام البظاعة", "العنوان", "عنوان"]:
        if col in df.columns:
          address = str(row.get(col, "")).strip()
          break

      shipment_type = ""
      for col in ["نوع الشحنة", "النوع"]:
        if col in df.columns:
          shipment_type = str(row.get(col, "")).strip()
          break

      # تعبئة خلايا الإكسل للقالب الرسمي
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

      # تصميم HTML للوصل الواحد (مع فاصل صفحة مخصص للطباعة الكلية)
      single_receipt_html = f"""
            <div class="receipt-page" style="
                padding: 15px; 
                font-family: 'Tahoma', Arial, sans-serif; 
                direction: rtl; 
                border: 2px solid #102a43; 
                width: 100%; 
                max-width: 148mm; 
                margin: auto auto 25px auto; 
                background: #ffffff; 
                color: #102a43;
                box-sizing: border-box;
                page-break-after: always;
                break-after: page;
            ">
                <!-- رأس الوصل مع شعار بدون مربعات -->
                <table style="width: 100%; border-bottom: 2px solid #102a43; padding-bottom: 8px; margin-bottom: 12px;">
                    <tr>
                        <td style="text-align: right; vertical-align: middle;">
                            <div style="display: flex; align-items: center;">
                                {logo_html}
                                <div>
                                    <h2 style="margin: 0; font-size: 16px; color: #102a43;">أطلس المحيط للتجارة العامة</h2>
                                    <p style="margin: 2px 0 0; font-size: 10px; color: #627d98;">OCEAN ATLAS GENERAL TRADING</p>
                                </div>
                            </div>
                        </td>
                        <td style="text-align: left; vertical-align: middle;">
                            <h3 style="margin: 0; font-size: 14px; color: #b45309;">وصل تسليم بضاعة</h3>
                            <p style="margin: 2px 0 0; font-size: 11px; color: #334e68;">Cargo Delivery Receipt</p>
                        </td>
                    </tr>
                </table>

                <!-- تفاصيل الوصل الرئيسية -->
                <table style="width: 100%; font-size: 12px; border-collapse: collapse; margin-bottom: 12px;">
                    <tr style="background-color: #f0f4f8;">
                        <td style="padding: 6px; border: 1px solid #bcccdc; width: 50%;"><strong>كود العميل:</strong> <span style="color: #b45309; font-weight: bold;">{code}</span></td>
                        <td style="padding: 6px; border: 1px solid #bcccdc; width: 50%;"><strong>رقم الشحنة:</strong> <span style="color: #b45309; font-weight: bold;">{shipment}</span></td>
                    </tr>
                    <tr>
                        <td style="padding: 6px; border: 1px solid #bcccdc;"><strong>اسم العميل:</strong> <span style="font-weight: bold;">{name}</span></td>
                        <td style="padding: 6px; border: 1px solid #bcccdc;"><strong>رقم الهاتف:</strong> <span style="direction: ltr; display: inline-block; font-weight: bold;">{formatted_phone}</span></td>
                    </tr>
                    <tr style="background-color: #f0f4f8;">
                        <td style="padding: 6px; border: 1px solid #bcccdc;"><strong>عنوان الاستلام:</strong> <span style="color: #486581; font-weight: bold;">{address}</span></td>
                        <td style="padding: 6px; border: 1px solid #bcccdc;"><strong>عدد الطرود:</strong> 📦 {packages} طرد</td>
                    </tr>
                    <tr>
                        <td style="padding: 6px; border: 1px solid #bcccdc;"><strong>تاريخ الإصدار:</strong> <span style="color: #b45309; font-weight: bold;">{today_date}</span></td>
                        <td style="padding: 6px; border: 1px solid #bcccdc;"><strong>الوزن الإجمالي:</strong> {weight} كغ</td>
                    </tr>
                    <tr style="background-color: #f0f4f8;">
                        <td style="padding: 6px; border: 1px solid #bcccdc;"><strong>نوع الشحنة:</strong> {shipment_type}</td>
                        <td style="padding: 6px; border: 1px solid #bcccdc;"><strong>سعر الكيلو:</strong> {price_per_kg:,.2f} $</td>
                    </tr>
                    <tr style="background-color: #fef3c7;">
                        <td style="padding: 6px; border: 1px solid #f59e0b;" colspan="2"><strong>إجمالي المبيعات:</strong> <span style="color: #b45309; font-weight: bold; font-size: 13px;">{total_sales:,.2f} $</span> &nbsp;&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;&nbsp; <strong>طريقة الدفع:</strong> [ &nbsp; ] نقداً &nbsp;&nbsp; [ &nbsp; ] أجل</td>
                    </tr>
                </table>

                <!-- إقرار استلام البضاعة -->
                <div style="background-color: #fffbeb; border: 1px solid #fde68a; padding: 8px; border-radius: 4px; margin-bottom: 12px;">
                    <p style="margin: 0; font-size: 11px; color: #92400e; line-height: 1.4;">
                        <strong>إقرار الاستلام:</strong><br>
                        أقر أنا الموقع أدناه، بأنني استلمت البضاعة والشحنة المذكورة أعلاه كاملة، وبحالة سليمة وممتازة، ومطابقة لكافة الأوزان والأوصاف المدونة.
                    </p>
                </div>

                <!-- تواقيع الاستلام -->
                <table style="width: 100%; font-size: 12px; margin-top: 5px;">
                    <tr>
                        <td style="width: 50%; padding: 4px;">
                            <strong>اسم المستلم:</strong><br><br>
                            ............................................
                        </td>
                        <td style="width: 50%; padding: 4px; text-align: left;">
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

    # --- زر الطباعة الكلية في أعلى الصفحة باستخدام Streamlit الأصلي ---
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
      master_btn = st.button(
          "🖨️ طباعة جميع الوصولات دفعة واحدة (A5)",
          use_container_width=True,
          type="primary",
      )

    if master_btn:
      # حقن كود JavaScript لتنفيذ الطباعة الشاملة في نافذة جديدة نظيفة
      escaped_html = (
          all_receipts_html_for_print.replace("`", "\\`")
          .replace("$", "\\$")
          .replace('"', '\\"')
      )
      master_js = f"""
            <script>
                var printWin = window.open('', '', 'height=900,width=800');
                printWin.document.write('<html><head><title>طباعة جميع الوصولات - A5</title><style>@page {{ size: A5; margin: 5mm; }} body {{ direction: rtl; font-family: Tahoma, sans-serif; background: #fff; margin: 0; padding: 0; }} .receipt-page {{ page-break-after: always; break-after: page; margin-bottom: 20px; }}</style></head><body>');
                printWin.document.write("{escaped_html}");
                printWin.document.write('</body></html>');
                printWin.document.close();
                printWin.focus();
                setTimeout(function(){{ printWin.print(); printWin.close(); }}, 600);
            </script>
            """
      st.components.v1.html(master_js, height=0)

    st.markdown("---")

    # حلقة تكرارية لعرض الوصولات تفصيلياً مع أزرارها الخاصة بالتنزيل والطباعة الفردية
    for item in receipts_data_list:
      index = item["index"]
      shipment = item["shipment"]
      file_name_id = item["file_name_id"]

      with st.expander(
          f"📄 وصل العميل: {item['name']} | كود العميل: {item['code']} | الشحنة:"
          f" {shipment} | الإجمالي: {item['total_sales']:,.0f} $",
          expanded=False,
      ):
        st.download_button(
            label=f"📥 تنزيل إكسل الوصل (الشحنة: {shipment})",
            data=item["output"],
            file_name=f"Delivery_Receipt_{file_name_id}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            key=f"download_excel_{index}",
        )

        single_ui_html = (
            f"""
                {item['single_html']}
                <br>
                <div style="display: flex; gap: 10px; margin-bottom: 15px; direction: rtl;">
                    <button onclick="
                        var printWin = window.open('', '', 'height=800,width=800');
                        printWin.document.write('<html><head><title>طباعة الشحنة {shipment}</title><style>@page {{ size: A5; margin: 5mm; }} body {{ direction: rtl; font-family: Tahoma, sans-serif; background: #fff; margin: 0; padding: 0; }}</style></head><body>');
                        printWin.document.write(`{item['single_html']}`);
                        printWin.document.write('</body></html>');
                        printWin.document.close();
                        printWin.focus();
                        setTimeout(function(){{ printWin.print(); printWin.close(); }}, 500);
                    " style="
                        background-color: #102a43; color: white; padding: 10px 16px; border: none; border-radius: 6px; cursor: pointer; font-weight: bold; font-size: 14px; flex: 1; text-align: center;
                    ">
                        🖨️ طباعة الوصل (ورقيّاً)
                    </button>

                    <button onclick="
                        var printWin = window.open('', '', 'height=800,width=800');
                        printWin.document.write('<html><head><title>{file_name_id}</title><style>@page {{ size: A5; margin: 5mm; }} body {{ direction: rtl; font-family: Tahoma, sans-serif; background: #fff; margin: 0; padding: 0; }}</style></head><body>');
                        printWin.document.write(`{item['single_html']}`);
                        printWin.document.write('</body></html>');
                        printWin.document.close();
                        printWin.focus();
                        setTimeout(function(){{ 
                            printWin.document.title = '{file_name_id}';
                            printWin.print(); 
                        }}, 600);
                    " style="
                        background-color: #b45309; color: white; padding: 10px 16px; border: none; border-radius: 6px; cursor: pointer; font-weight: bold; font-size: 14px; flex: 1; text-align: center;
                    ">
                        📑 حفظ PDF (مقاس A5)
                    </button>
                </div>
                """
        )
        st.components.v1.html(single_ui_html, height=650)

      st.markdown("---")

  except Exception as e:
    st.error(f"حدث خطأ أثناء قراءة أو معالجة الملفات: {e}")
else:
  st.info(
      "الرجاء رفع ملف بيانات الشحنات وقالب الوصل من الشريط الجانبي لتظهر المعاينة"
      " والطباعة."
  )
