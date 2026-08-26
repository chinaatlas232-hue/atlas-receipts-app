import pandas as pd
import streamlit as st

st.title("تعبئة وصل أطلس")

# إضافة أداة لرفع ملف الأكسل مباشرة من المتصفح
uploaded_file = st.file_uploader(
    "الرجاء رفع ملف الأكسل (تعبئة وصل أطلس.xlsx)", type=["xlsx"]
)

if uploaded_file is not None:
  try:
    # قراءة الملف المرفوع
    df = pd.read_excel(uploaded_file)

    # تنظيف أسماء الأعمدة من المسافات الزائدة لضمان مطابقتها تماماً
    df.columns = df.columns.str.strip()

    st.success("تم قراءة الملف بنجاح!")

    # استعراض وتعبئة البيانات لكل صف مع زر طباعة خاص لكل وصل
    for index, row in df.iterrows():
      # تنظيف وقراءة البيانات بدقة
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

      # استخدام حاوية (Container) لكل وصل لتنظيم العرض والطباعة
      with st.container():
        st.markdown(f"### وصل رقم {index + 1} (الكود: {code})")
        st.write(f"- **الشحنة:** `{shipment}`")
        st.write(f"- **الكود:** `{code}`")
        st.write(f"- **اسم العميل:** {name}")
        st.write(f"- **رقم الهاتف:** `{phone}`")
        st.write(f"- **نوع الشحنة:** {shipment_type}")
        st.write(f"- **عنوان الاستلام:** {address}")
        st.write(
            f"- **الوزن:** {weight} كغ | **عدد الطرود:** {packages} | **الحجم:**"
            f" {volume}"
        )
        st.write(f"- **سعر الكيلو:** {price_per_kg}")
        st.write(f"- **إجمالي المبيعات:** {total_sales}")

        # كود جافاسكريبت لفتح نافذة طباعة خاصة بهذا الوصل عند الضغط على الزر
        print_html = f"""
                <div id="receipt-{index}" style="padding: 20px; font-family: Tahoma, sans-serif; direction: rtl;">
                    <h2>وصل أطلس للشحن</h2>
                    <hr/>
                    <p><strong>رقم الشحنة:</strong> {shipment}</p>
                    <p><strong>الكود:</strong> {code}</p>
                    <p><strong>اسم العميل:</strong> {name}</p>
                    <p><strong>رقم الهاتف:</strong> {phone}</p>
                    <p><strong>نوع الشحنة:</strong> {shipment_type}</p>
                    <p><strong>عنوان استلام البضاعة:</strong> {address}</p>
                    <p><strong>الوزن:</strong> {weight} كغ</p>
                    <p><strong>عدد الطرود:</strong> {packages}</p>
                    <p><strong>الحجم:</strong> {volume}</p>
                    <p><strong>سعر الكيلو:</strong> {price_per_kg}</p>
                    <p><strong>إجمالي المبيعات:</strong> {total_sales}</p>
                </div>
                <button onclick="
                    var printWindow = window.open('', '', 'height=600,width=800');
                    printWindow.document.write('<html><head><title>طباعة الوصل</title>');
                    printWindow.document.write('</head><body style=\\"direction: rtl; font-family: Tahoma;\\">');
                    printWindow.document.write(document.getElementById('receipt-{index}').innerHTML);
                    printWindow.document.write('</body></html>');
                    printWindow.document.close();
                    printWindow.focus();
                    setTimeout(function(){{ printWindow.print(); printWindow.close(); }}, 500);
                " style="background-color: #ff4b4b; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer; font-weight: bold;">
                    🖨️ طباعة هذا الوصل
                </button>
                <br><br>
                """
        st.components.v1.html(print_html, height=120)
        st.markdown("---")

  except Exception as e:
    st.error(f"حدث خطأ أثناء قراءة الملف: {e}")
else:
  st.info("الرجاء رفع الملف المذكور أعلاه ليبدأ الكود بالعمل.")
