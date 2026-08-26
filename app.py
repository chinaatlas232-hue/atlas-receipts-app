import io
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(page_title="تعبئة وتوليد وصل أطلس", layout="wide")

st.title("📦 نظام تعبئة وطباعة وصل أطلس للشحن")

# 1. رفع ملف البيانات الرئيسي
uploaded_data_file = st.file_uploader(
    "1.الرجاء رفع ملف بيانات الشحنات (تعبئة وصل اطلس.xlsx)", type=["xlsx"]
)

# 2. رفع قالب الوصل
uploaded_template_file = st.file_uploader(
    "2.الرجاء رفع قالب وصل التسليم (Atlas_Cargo_Delivery_Receipt.xlsx)",
    type=["xlsx"],
)

if uploaded_data_file is not None and uploaded_template_file is not None:
  try:
    # قراءة بيانات الشحنات
    df = pd.read_excel(uploaded_data_file)
    df.columns = df.columns.str.strip()

    st.success("تم قراءة ملف البيانات وقالب الوصل بنجاح!")
    st.markdown("---")

    # حلقة تكرارية لكل صف في ملف البيانات لتعبئة الوصل الخاص به
    for index, row in df.iterrows():
      # استخراج وتنقيب البيانات بدقة من الأعمدة
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

      # فتح قالب الوصل باستخدام openpyxl وتعبئته بالبيانات المطابقة تماماً للخلايا
      wb = openpyxl.load_workbook(uploaded_template_file)
      ws = wb.active

      # تعبئة خلايا القالب بناءً على تصميم ملف الـ Excel الخاص بالوصل
      ws["B4"] = code  # رقم الوصل / Receipt No (نضع الكود هنا)
      ws["B5"] = name  # اسم العميل / Client Name
      ws["D5"] = phone  # رقم الهاتف / Phone
      ws["B6"] = (
          f"{shipment} - {address}"  # اسم الشحنة / Cargo Name (مع العنوان)
      )
      ws["D6"] = packages  # عدد الطرود / Pieces
      ws["B7"] = shipment_type  # نوع الشحنة / Cargo Type
      ws["D7"] = weight  # الوزن الفعلي / Actual Weight

      # حفظ الملف في الذاكرة لتسهيل التحميل المباشر
      output = io.BytesIO()
      wb.save(output)
      output.seek(0)

      # عرض تفاصيل الوصل في الواجهة
      with st.expander(
          f"📁 الوصل رقم {index + 1} | العميل: {name} | الكود: {code}",
          expanded=True,
      ):
        col1, col2 = st.columns(2)
        with col1:
          st.write(f"- **رقم الشحنة:** `{shipment}`")
          st.write(f"- **الكود (رقم الوصل):** `{code}`")
          st.write(f"- **اسم العميل:** {name}")
          st.write(f"- **رقم الهاتف:** `{phone}`")
          st.write(f"- **نوع الشحنة:** {shipment_type}")
        with col2:
          st.write(f"- **عنوان الاستلام:** {address}")
          st.write(f"- **الوزن الفعلي:** {weight} كغ")
          st.write(f"- **عدد الطرود:** {packages}")
          st.write(f"- **سعر الكيلو:** {price_per_kg}")
          st.write(f"- **إجمالي المبيعات:** {total_sales}")

        # زر تحميل ملف الإكسل المعبأ لهذا الوصل
        st.download_button(
            label=f"📥 تحميل وصل العميل {name} (Excel)",
            data=output,
            file_name=f"Atlas_Receipt_{code}_{name}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            key=f"download_{index}",
        )

        # كود HTML + جافاسكريبت لطباعة الوصل بتنسيق نظيف ومنسق
        receipt_html = f"""
                <div id="print-area-{index}" style="padding: 20px; font-family: Tahoma, sans-serif; direction: rtl; border: 2px solid #333; width: 100%; max-width: 600px; margin: auto; background: #fff;">
                    <h2 style="text-align: center; margin-bottom: 5px;">أطلس المحيط للتجارة العامة</h2>
                    <h4 style="text-align: center; margin-top: 0; color: #555;">وصل تسليم الشحنات (CARGO DELIVERY RECEIPT)</h4>
                    <hr/>
                    <table style="width: 100%; font-size: 14px; line-height: 1.8;">
                        <tr>
                            <td><strong>رقم الوصل:</strong> {code}</td>
                            <td><strong>التاريخ:</strong> 2026 / / </td>
                        </tr>
                        <tr>
                            <td><strong>اسم العميل:</strong> {name}</td>
                            <td><strong>رقم الهاتف:</strong> {phone}</td>
                        </tr>
                        <tr>
                            <td><strong>اسم الشحنة:</strong> {shipment} ({address})</td>
                            <td><strong>عدد الطرود:</strong> {packages}</td>
                        </tr>
                        <tr>
                            <td><strong>نوع الشحنة:</strong> {shipment_type}</td>
                            <td><strong>الوزن الفعلي:</strong> {weight} كغ</td>
                        </tr>
                        <tr>
                            <td colspan="2"><strong>طريقة الدفع:</strong> [  ] نقداً    [  ] أجل</td>
                        </tr>
                    </table>
                    <br>
                    <p style="font-size: 12px; border-top: 1px dashed #ccc; padding-top: 10px;">
                        <strong>إقرار الاستلام والتسليم:</strong><br>
                        أقر أنا الموقع أدناه، بأنني استلمت الشحنة المذكورة أعلاه سليمة وبحالة جيدة ومطابقة لكافة الأوزان والأوصاف المذكورة.
                    </p>
                    <br>
                    <table style="width: 100%; font-size: 14px;">
                        <tr>
                            <td><strong>اسم المستلم:</strong> ........................................</td>
                            <td><strong>التوقيع:</strong> ..........................</td>
                        </tr>
                    </table>
                </div>
                <br>
                <button onclick="
                    var printWin = window.open('', '', 'height=700,width=900');
                    printWin.document.write('<html><head><title>طباعة الوصل</title></head><body style=\\'direction: rtl; font-family: Tahoma;\\'>');
                    printWin.document.write(document.getElementById('print-area-{index}').innerHTML);
                    printWin.document.write('</body></html>');
                    printWin.document.close();
                    printWin.focus();
                    setTimeout(function(){{ printWin.print(); printWin.close(); }}, 500);
                " style="background-color: #2e7d32; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer; font-weight: bold; width: 100%;">
                    🖨️ طباعة وصل العميل {name}
                </button>
                """
        st.components.v1.html(receipt_html, height=340)

      st.markdown("---")

  except Exception as e:
    st.error(f"حدث خطأ أثناء معالجة الملفات: {e}")
else:
  st.info("الرجاء رفع كلا الملفين (ملف بيانات الشحنات وقالب الوصل) لتبدأ العملية.")
