import datetime
import io
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(page_title="وصل تسليم بضاعة - أطلس", layout="wide")

st.title("📦 النظام المالي والفني - وصل تسليم البضائع")

# 1. رفع ملف البيانات الرئيسي
uploaded_data_file = st.file_uploader(
    "1. الرجاء رفع ملف بيانات الشحنات (تعبئة وصل اطلس.xlsx)", type=["xlsx"]
)

# 2. رفع قالب الوصل
uploaded_template_file = st.file_uploader(
    "2. الرجاء رفع قالب وصل التسليم (Atlas_Cargo_Delivery_Receipt.xlsx)",
    type=["xlsx"],
)

if uploaded_data_file is not None and uploaded_template_file is not None:
  try:
    # قراءة بيانات الشحنات
    df = pd.read_excel(uploaded_data_file)
    df.columns = df.columns.str.strip()

    # الحصول على تاريخ اليوم تلقائياً
    today_date = datetime.date.today().strftime("%Y-%m-%d")

    st.success(
        f"تم تحميل الملفات بنجاح. تاريخ الإصدار التلقائي هو: {today_date}"
    )
    st.markdown("---")

    # حلقة تكرارية لكل صف في ملف البيانات
    for index, row in df.iterrows():
      # استخراج البيانات وتنظيفها
      shipment = str(row.get("الشحنة", "")).strip()
      code = str(row.get("الكود", "")).strip()
      weight = row.get("الوزن", 0)
      packages = row.get("عدد الطرود", 0)
      volume = row.get("الحجم", 0)

      price_per_kg = row.get("سعر الكيلو", row.get("سعر الكيلو ", 0))
      total_sales = row.get("اجمالي مبيعات", row.get("اجمالي مبيعات ", 0))
      name = str(row.get("الاسم", row.get("الاسم ", ""))).strip()

      phone_raw = row.get("رقم الهاتف", row.get("رقم الهاتف ", ""))
      phone = str(phone_raw).strip()
      if phone.endswith(".0"):
        phone = phone[:-2]

      address = str(
          row.get("عنوان استلام البظاعة", row.get("عنوان استلام البظاعة ", ""))
      ).strip()
      shipment_type = str(
          row.get("نوع الشحنة", row.get("نوع الشحنة ", ""))
      ).strip()

      # تعبئة خلايا الأكسل للقالب الرسمي
      wb = openpyxl.load_workbook(uploaded_template_file)
      ws = wb.active

      ws["B4"] = code
      ws["D4"] = today_date
      ws["B5"] = name
      ws["B6"] = address  # حقل العنوان المنفصل تحت اسم العميل
      ws["D5"] = phone
      ws["B7"] = shipment
      ws["D6"] = packages
      ws["B8"] = shipment_type
      ws["D7"] = weight

      output = io.BytesIO()
      wb.save(output)
      output.seek(0)

      # عرض الوصل داخل تطبيق Streamlit
      with st.expander(
          f"📄 وصل تسليم رقم: {code} | العميل: {name} | التاريخ: {today_date}",
          expanded=True,
      ):
        st.download_button(
            label=f"📥 تنزيل إكسل الوصل الرسمي ({name})",
            data=output,
            file_name=f"Delivery_Receipt_{code}_{name}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            key=f"download_{index}",
        )

        # تصميم HTML الفخم مع حقل العنوان المنفصل تحت اسم العميل مباشرة
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
                    <!-- رأس الوصل -->
                    <table style="width: 100%; border-bottom: 2px solid #102a43; padding-bottom: 10px; margin-bottom: 25px;">
                        <tr>
                            <td style="text-align: right;">
                                <h2 style="margin: 0; font-size: 22px; color: #102a43;">أطلس المحيط للتجارة العامة</h2>
                                <p style="margin: 3px 0 0; font-size: 12px; color: #627d98;">OCEAN ATLAS GENERAL TRADING</p>
                            </td>
                            <td style="text-align: left;">
                                <h3 style="margin: 0; font-size: 18px; color: #b45309;">وصل تسليم بضاعة</h3>
                                <p style="margin: 3px 0 0; font-size: 13px; color: #334e68;">Cargo Delivery Receipt</p>
                            </td>
                        </tr>
                    </table>

                    <!-- تفاصيل الوصل الرئيسية مع العنوان المستقل تحت اسم العميل -->
                    <table style="width: 100%; font-size: 14px; border-collapse: collapse; margin-bottom: 25px;">
                        <tr>
                            <td style="padding: 10px; border: 1px solid #bcccdc; width: 50%;"><strong>رقم الوصل (Code):</strong> <span style="color: #0066cc; font-weight: bold;">{code}</span></td>
                            <td style="padding: 10px; border: 1px solid #bcccdc; width: 50%;"><strong>تاريخ الإصدار:</strong> <span style="color: #b45309; font-weight: bold;">{today_date}</span></td>
                        </tr>
                        <tr style="background-color: #f0f4f8;">
                            <td style="padding: 10px; border: 1px solid #bcccdc;"><strong>اسم العميل:</strong> <span style="font-weight: bold;">{name}</span></td>
                            <td style="padding: 10px; border: 1px solid #bcccdc;"><strong>رقم الهاتف:</strong> {phone}</td>
                        </tr>
                        <tr>
                            <td style="padding: 10px; border: 1px solid #bcccdc;"><strong>عنوان الاستلام:</strong> <span style="color: #486581; font-weight: bold;">{address}</span></td>
                            <td style="padding: 10px; border: 1px solid #bcccdc;"><strong>عدد الطرود:</strong> {packages} طرد</td>
                        </tr>
                        <tr style="background-color: #f0f4f8;">
                            <td style="padding: 10px; border: 1px solid #bcccdc;"><strong>تفاصيل الشحنة:</strong> {shipment}</td>
                            <td style="padding: 10px; border: 1px solid #bcccdc;"><strong>الوزن الإجمالي:</strong> {weight} كغ</td>
                        </tr>
                        <tr>
                            <td style="padding: 10px; border: 1px solid #bcccdc;"><strong>نوع الشحنة:</strong> {shipment_type}</td>
                            <td style="padding: 10px; border: 1px solid #bcccdc;">
                                <strong>طريقة الدفع:</strong> 
                                &nbsp;&nbsp; [ &nbsp; ] نقداً &nbsp;&nbsp; [ &nbsp; ] أجل
                            </td>
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
                <button onclick="
                    var printWin = window.open('', '', 'height=800,width=1000');
                    printWin.document.write('<html><head><title>طباعة وصل تسليم بضاعة - أطلس</title></head><body style=\\'direction: rtl; font-family: Tahoma; background: #fff;\\'>');
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
                    width: 100%;
                ">
                    🖨️ طباعة وصل تسليم البضاعة للعميل: {name}
                </button>
                """
        st.components.v1.html(clean_receipt_html, height=540)

      st.markdown("---")

  except Exception as e:
    st.error(f"حدث خطأ أثناء قراءة أو معالجة الملفات: {e}")
else:
  st.info("الرجاء رفع ملف بيانات الشحنات وقالب الوصل لتظهر المعاينة والطباعة.")
