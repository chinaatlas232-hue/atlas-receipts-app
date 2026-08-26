import pandas as pd

# قراءة ملف الأكسل
file_path = "تعبئة  وصل اطلس.xlsx"
df = pd.read_excel(file_path)

# تنظيف أسماء الأعمدة من المسافات الزائدة لضمان مطابقتها تماماً
df.columns = df.columns.str.strip()

# استعراض البيانات أو تعبئتها بالشكل المطلوب
for index, row in df.iterrows():
  # تنظيف القيم النصية من المسافات البيضاء الفائضة
  shipment = str(row["الشحنة"]).strip()
  code = str(row["الكود"]).strip()
  weight = row["الوزن"]
  packages = row["عدد الطرود"]
  volume = row["الحجم"]
  price_per_kg = row["سعر الكيلو"]
  total_sales = row["اجمالي مبيعات"]
  name = str(row["الاسم"]).strip()

  # التأكد من التعامل مع رقم الهاتف كنص لكي لا يفقد الصفر في البداية أو يظهر بشكل خاطئ
  phone = str(row["رقم الهاتف"]).strip()
  if not phone.startswith("0") and not phone.startswith("964"):
    phone = "0" + phone  # تعديل تنسيق الهاتف إذا لزم الأمر

  address = str(row["عنوان استلام البظاعة"]).strip()
  shipment_type = str(row["نوع الشحنة"]).strip()

  # طباعة البيانات للتأكد من مطابقتها الصحيحة
  print(f"--- سجل رقم {index + 1} ---")
  print(f"الكود: {code}")
  print(f"اسم العميل: {name}")
  print(f"رقم الهاتف: {phone}")
  print(f"نوع الشحنة: {shipment_type}")
  print(f"العنوان: {address}")
  print(f"الوزن: {weight} | عدد الطرود: {packages} | الإجمالي: {total_sales}\n")
