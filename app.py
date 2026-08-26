import pandas as pd
import streamlit as st

st.title("تعبئة وصل أطلس")

# إضافة أداة لرفع ملف الأكسل مباشرة من المتصفح لمنع مشاكل مسار الملف
uploaded_file = st.file_uploader("الرجاء رفع ملف الأكسل (تعبئة وصل أطلس.xlsx)", type=["xlsx"])

if uploaded_file is not None:
  try:
    # قراءة الملف المرفوع
    df = pd.read_excel(uploaded_file)

    # تنظيف أسماء الأعمدة من المسافات الزائدة
    df.columns = df.columns.str.strip()

    st.success("تم قراءة الملف بنجاح!")
    st.write("معاينة البيانات:")
    st.dataframe(df)

    # استعراض وتعبئة البيانات لكل صف
    for index, row in df.iterrows():
      # تنظيف البيانات وقراءتها بشكل صحيح
      shipment = str(row.get("الشحنة", "")).strip()
      code = str(row.get("الكود", "")).strip()
      weight = row.get("الوزن", 0)
      packages = row.get("عدد الطرود", 0)
      volume = row.get("الحجم", 0)
      price_per_kg = row.get("سعر الكيلو", 0)
      total_sales = row.get("اجمالي مبيعات", 0)
      name = str(row.get("الاسم", "")).strip()

      # معالجة رقم الهاتف لضمان ظهوره بشكل صحيح
      phone = str(row.get("رقم الهاتف", "")).strip()
      if phone.endswith(".0"):
        phone = phone[:-2]  # إزالة الفاصلة العشرية لو ظهرت من القراءة العددية

      address = str(row.get("عنوان استلام البظاعة", "")).strip()
      shipment_type = str(row.get("نوع الشحنة", "")).strip()

      # عرض النتائج لكل وصل على حدة للتأكد من مطابقتها
      st.markdown(f"### وصل رقم {index + 1}")
      st.write(f"- **الكود:** `{code}`")
      st.write(f"- **اسم العميل:** {name}")
      st.write(f"- **رقم الهاتف:** `{phone}`")
      st.write(f"- **نوع الشحنة:** {shipment_type}")
      st.write(f"- **عنوان الاستلام:** {address}")
      st.write(f"- **الوزن:** {weight} | **عدد الطرود:** {packages}")
      st.markdown("---")

  except Exception as e:
    st.error(f"حدث خطأ أثناء قراءة الملف: {e}")
else:
  st.info("الرجاء رفع الملف المذكور أعلاه ليبدأ الكود بالعمل.")
