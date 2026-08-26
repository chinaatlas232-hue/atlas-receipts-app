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
      logo_html = f'<img src="data:image/png;base64,{base64_logo}" style="height: 55px; width: 55px; object-fit: cover; border-radius: 50%; border: none; outline: none; background: transparent; margin-left: 12px; vertical-align: middle;" />'

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

      # حل مشكلة تشابه أرقام الشحنات بدمج رقم الشحنة مع اسم العميل أو كود العميل ليكون الملف فريداً
      file_name_id = f"{shipment}_{name}" if shipment and name else (shipment if shipment else f"Receipt_{index}")

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

      # عرض الوصل داخل تطبيق Streamlit مع إظهار تفاصيل واضحة لكل عميل
      with st.expander(
          f"📄 وصل العميل: {name} | الشحنة: {shipment} | الإجمالي: {total_sales:,.0f} $",
          expanded=True,
      ):
        # زر تنزيل الإكسل
        st.download_button(
            label=f"📥 تنزيل إكسل الوصل (الشحنة: {shipment} - العميل: {name})",
            data=output,
            file_name=f"Delivery_Receipt_{file_name_id}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            key=f"download_excel_{index}",
        )

        # تصميم HTML للوصل المخصص للطباعة وحفظه بصيغة PDF
        clean_receipt_html = f"""
                <div id="receipt-print-{index}" style="
                    padding: 40px; 
                    font-family: 'Tahoma', Arial, sans-serif; 
                    direction: rtl; 
                    border: 3px solid #102a43; 
                    width: 100%; 
                    max-width: 700px; 
                    margin: auto; 
                    background: #ffffff; 
                    color: #102a43;
                ">
                    <!-- رأس الوصل مع شعار بدون مربعات -->
                    <table style="width: 100%; border-bottom: 2px solid #102a43; padding-bottom: 10px; margin-bottom: 25px;">
                        <tr>
                            <td style="text-align: right; vertical-align: middle;">
                                <div style="display: flex; align-items: center;">
                                    {logo_html}
                                    <div>
                                        <h2 style="margin: 0; font-size: 22px; color: #102a43;">أطلس المحيط للتجارة العامة</h2>
                                        <p style="margin: 3px 0 0; font-size: 12px; color: #627d98;">OCEAN ATLAS GENERAL TRADING</p>
                                    </div>
                                </div>
                            </td>
                            <td style="text-align: left; vertical-align: middle;">
                                <h3 style="margin: 0; font-size: 18px; color: #b45309;">وصل تسليم بضاعة</h3>
                                <p style="margin: 3px 0 0; font-size: 13px; color: #334e68;">Cargo Delivery Receipt</p>
                            </td>
                        </tr>
                    </table>

                    <!-- تفاصيل الوصل الرئيسية -->
                    <table style="width: 100%; font-size: 14px; border-collapse: collapse; margin-bottom: 25px;">
                        <tr style="background-color: #f0f4f8;">
                            <td style="padding: 10px; border: 1px solid #bcccdc; width: 50%;"><strong>رقم الشحنة:</strong> <span style="color: #b45309; font-weight: bold;">{shipment}</span></td>
                            <td style="padding: 10px; border: 1px solid #bcccdc; width: 50%;"><strong>كود العميل:</strong> <span style="font-weight: bold;">{code}</span></td>
                        </tr>
                        <tr>
                            <td style="padding: 10px; border: 1px solid #bcccdc;"><strong>اسم العميل:</strong> <span style="font-weight: bold;">{name}</span></td>
                            <td style="padding: 10px; border: 1px solid #bcccdc;"><strong>رقم الهاتف:</strong> <span style="direction: ltr; display: inline-block; font-weight: bold;">{formatted_phone}</span></td>
                        </tr>
                        <tr style="background-color: #f0f4f8;">
                            <td style="padding: 10px; border: 1px solid #bcccdc;"><strong>عنوان الاستلام:</strong> <span style="color: #486581; font-weight: bold;">{address}</span></td>
                            <td style="padding: 10px; border: 1px solid #bcccdc;"><strong>عدد الطرود:</strong> 📦 {packages} طرد</td>
                        </tr>
                        <tr>
                            <td style="padding: 10px; border: 1px solid #bcccdc;"><strong>تاريخ الإصدار:</strong> <span style="color: #b45309; font-weight: bold;">{today_date}</span></td>
                            <td style="padding: 10px; border: 1px solid #bcccdc;"><strong>الوزن الإجمالي:</strong> {weight} كغ</td>
                        </tr>
                        <tr style="background-color: #f0f4f8;">
                            <td style="padding: 10px; border: 1px solid #bcccdc;"><strong>نوع الشحنة:</strong> {shipment_type}</td>
                            <td style="padding: 10px; border: 1px solid #bcccdc;"><strong>سعر الكيلو:</strong> {price_per_kg:,.2f} $</td>
                        </tr>
                        <tr style="background-color: #fef3c7;">
                            <td style="padding: 10px; border: 1px solid #f59e0b;" colspan="2"><strong>إجمالي المبيعات:</strong> <span style="color: #b45309; font-weight: bold; font-size: 15px;">{total_sales:,.2f} $</span> &nbsp;&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;&nbsp; <strong>طريقة الدفع:</strong> [ &nbsp; ] نقداً &nbsp;&nbsp; [ &nbsp; ] أجل</td>
                        </tr>
                    </table>

                    <!-- إقرار استلام البضاعة -->
                    <div style="background-color: #fffbeb; border: 1px solid #fde68a; padding: 15px; border-radius: 4px; margin-bottom: 30px;">
                        <p style="margin: 0; font-size: 13px; color: #92400e; line-height: 1.7;">
                            <strong>إقرار الاستلام:</strong><br>
                            أقر أنا الموقع أدناه، بأنني استلمت البضاعة والشحنة المذكورة أعلاه كاملة، وبحالة سليمة وممتازة، ومطابقة لكافة الأوزان والأوصاف المدونة.
                        </p>
                    </div>

                    <!-- تواقيع الاستلام -->
                    <table style="width: 100%; font-size: 14px; margin-top: 20px;">
                        <tr>
                            <td style="width: 50%; padding: 10px;">
                                <strong>اسم المستلم:</strong><br><br>
                                ....................................................
                            </td>
                            <td style="width: 50%; padding: 10px; text-align: left;">
                                <strong>توقيع وختم المستلم:</strong><br><br>
                                ....................................................
                            </td>
                        </tr>
                    </table>
                </div>
                <br>
                <div style="display: flex; gap: 10px;">
                    <!-- زر الطباعة الورقية -->
                    <button onclick="
                        var printWin = window.open('', '', 'height=800,width=1000');
                        printWin.document.write('<html><head><title>طباعة الشحنة {shipment}</title></head><body style=\\'direction: rtl; font-family: Tahoma; background: #fff;\\'>');
                        printWin.document.write(document.getElementById('receipt-print-{index}').innerHTML);
                        printWin.document.write('</body></html>');
                        printWin.document.close();
                        printWin.focus();
                        setTimeout(function(){{ printWin.print(); printWin.close(); }}, 500);
                    " style="
                        background-color: #102a43; 
                        color: white; 
                        padding: 12px 20px; 
                        border: none; 
                        border-radius: 6px; 
                        cursor: pointer; 
                        font-weight: bold; 
                        font-size: 15px;
                        flex: 1;
                    ">
                        🖨️ طباعة الوصل ورقيّاً
                    </button>

                    <!-- زر حفظ كـ PDF بالاسم الفريد للشحنة والعميل -->
                    <button onclick="
                        var printWin = window.open('', '', 'height=800,width=1000');
                        printWin.document.write('<html><head><title>{file_name_id}</title></head><body style=\\'direction: rtl; font-family: Tahoma; background: #fff;\\'>');
                        printWin.document.write(document.getElementById('receipt-print-{index}').innerHTML);
                        printWin.document.write('</body></html>');
                        printWin.document.close();
                        printWin.focus();
                        setTimeout(function(){{ 
                            printWin.document.title = '{file_name_id}';
                            printWin.print(); 
                        }}, 600);
                    " style="
                        background-color: #b45309; 
                        color: white; 
                        padding: 12px 20px; 
                        border: none; 
                        border-radius: 6px; 
                        cursor: pointer; 
                        font-weight: bold; 
                        font-size: 15px;
                        flex: 1;
                    ">
                        📑 حفظ بصيغة PDF ({file_name_id})
                    </button>
                </div>
                """
        st.components.v1.html(clean_receipt_html, height=620)

      st.markdown("---")

  except Exception as e:
    st.error(f"حدث خطأ أثناء قراءة أو معالجة الملفات: {e}")
else:
  st.info(
      "الرجاء رفع ملف بيانات الشحنات وقالب الوصل من الشريط الجانبي لتظهر المعاينة"
      " والطباعة."
  )
