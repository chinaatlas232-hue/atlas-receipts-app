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

    # استعراض وتعبئة البيانات لكل صف
    for index, row in df.iterrows():
      # تنظيف وقراءة البيانات مع مراعاة المسافات وأسماء الأعمدة الدقيقة
      shipment = str(row.get("الشحنة", "")).strip()
      code = str(row.get("الكود", "")).strip()
      weight = row.get("الوزن", 0)
      packages = row.get("عدد الطرود", 0)
      volume = row.get("الحجم", 0)

      # قراءة سعر الكيلو والإجمالي مع معالجة المسافات في أسماء الأعمدة الأصلية
      price_per_kg = row.get("سعر الكيلو", row.get("سعر الكيلو ", 0))
      total_sales = row.get("اجمالي مبيعات", row.get("اجمالي مبيعات ", 0))

      name = str(row.get("الاسم", row.get("الاسم ", ""))).strip()

      # معالجة رقم الهاتف لضمان ظهوره بشكل نصي صحيح
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

      # عرض النتائج بشكل كامل وواضح
      st.markdown(f"### وصل رقم {index + 1}")
      st.write(f"- **الكود:** `{code}`")
      st.write(f"- **اسم العميل:** {name}")
      st.write(f"- **رقم الهاتف:** `{phone}`")
      st.write(f"- **نوع الشحنة:** {shipment_type}")
      st.write(f"- **عنوان الاستلام:** {address}")
      st.write(f"- **الوزن:** {weight} كغ | **عدد الطرود:** {packages}")
      st.write(f"- **سعر الكيلو:** {price_per_kg}")
      st.write(f"- **إجمالي المبيعات:** {total_sales}")
      st.markdown("---")

  except Exception as e:
    st.error(f"حدث خطأ أثناء قراءة الملف: {e}")
else:
  st.info("الرجاء رفع الملف المذكور أعلاه ليبدأ الكود بالعمل.")
