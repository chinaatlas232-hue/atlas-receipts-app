import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

st.set_page_config(page_title="نظام وصولات أطلس المحيط المتطور", page_icon="🧾", layout="wide")

st.title("🧾 نظام توليد وإصدار وصولات الشحنات - نسخة مطورة ومتعددة")
st.markdown("هذه الصفحة مخصصة لقراءة جدول الشحنات وإصدار وصولات التسليم (CARGO DELIVERY RECEIPT) بشكل فردي أو جماعي (دفعة واحدة لجميع شحنات العميل).")

uploaded_file = st.file_uploader("قم برفع ملف الشحنات الخاص بك (Excel أو CSV):", type=["xlsx", "csv"])

df = None
if uploaded_file is not None:
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
    st.success("تم رفع الملف بنجاح!")

if df is not None:
    with st.expander("🔍 معاينة جدول الشحنات المرفق"):
        st.dataframe(df.head(15))
    
    cols = df.columns.tolist()
    st.markdown("---")
    st.subheader("⚙️ إعدادات اختيار العميل والشحنة")
    
    col_a, col_b = st.columns(2)
    with col_a:
        client_col = st.selectbox("اختر عمود 'اسم العميل' أو 'الكود':", cols, index=0 if len(cols)>0 else 0)
    with col_b:
        mark_col = st.selectbox("اختر عمود 'علامة الشحن' (Shipping Mark):", cols, index=1 if len(cols)>1 else 0)

    if client_col and mark_col:
        client_list = df[client_col].dropna().unique().tolist()
        selected_client = st.selectbox("اختر العميل المطلوبة شحنته:", client_list)
        
        filtered_df = df[df[client_col] == selected_client]
        st.write(f"إجمالي الشحنات المتاحة للعميل (عددها: {len(filtered_df)}):")
        st.dataframe(filtered_df)
        
        selected_mark = st.selectbox("اختر الشحنة / العلامة المراد إصدار وصل لها:", filtered_df[mark_col].dropna().unique().tolist())
        row = filtered_df[filtered_df[mark_col] == selected_mark].iloc[0]
        
        st.markdown("### 📝 تفاصيل بيانات الوصل للتعديل أو المراجعة:")
        col1, col2 = st.columns(2)
        with col1:
            receipt_no = st.text_input("رقم الوصل الأساسي (Receipt No):", value=f"ATLAS-{str(row.get('No.', '2026-01'))}")
            receipt_date = st.text_input("التاريخ (Date):", value="2026/08/26")
            client_name_val = st.text_input("اسم العميل / الكود (Client Name):", value=str(selected_client))
            phone_val = st.text_input("رقم الهاتف (Phone):", value="")
        with col2:
            cargo_name_val = st.text_input("اسم الشحنة / علامة الشحن (Cargo Name):", value=str(selected_mark))
            cargo_type_val = st.text_input("نوع الشحنة (Cargo Type):", value=str(row.get('نوع البضاعة', '')))
            pieces_val = st.text_input("عدد الطرود / الكارتون (Pieces):", value=str(row.get('عدد الكارتون', '1')))
            weight_val = st.text_input("الوزن الفعلي (kg):", value=str(row.get('الوزن', '0')))
            
        receiver_name_val = st.text_input("اسم المستلم (Receiver Name):", value="")

        # دالة لبناء صفحة وصل واحدة داخل الـ PDF
        def build_single_receipt_elements(data, styles):
            elements = []
            
            # الهوية والترويسة (Header)
            title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], alignment=1, fontSize=16, textColor=colors.HexColor("#1F4E78"), spaceAfter=2)
            sub_style = ParagraphStyle('SubStyle', parent=styles['Heading2'], alignment=1, fontSize=11, textColor=colors.HexColor("#2F5597"), spaceAfter=10)
            
            elements.append(Paragraph("<b>أطلس المحيط للتجارة العامة والشحن الدولي</b>", title_style))
            elements.append(Paragraph("وصل تسليم الشحنات الرسمية (CARGO DELIVERY RECEIPT)", sub_style))
            elements.append(Spacer(1, 5))
            
            table_data = [
                [Paragraph("<b>رقم الوصل / Receipt No:</b>", styles['Normal']), Paragraph(str(data['receipt_no']), styles['Normal']),
                 Paragraph("<b>التاريخ / Date:</b>", styles['Normal']), Paragraph(str(data['date']), styles['Normal'])],
                [Paragraph("<b>اسم العميل / Client Name:</b>", styles['Normal']), Paragraph(str(data['client_name']), styles['Normal']),
                 Paragraph("<b>رقم الهاتف / Phone:</b>", styles['Normal']), Paragraph(str(data['phone']), styles['Normal'])],
                [Paragraph("<b>اسم الشحنة / Cargo Name:</b>", styles['Normal']), Paragraph(str(data['cargo_name']), styles['Normal']),
                 Paragraph("<b>عدد الطرود / Pieces:</b>", styles['Normal']), Paragraph(str(data['pieces']), styles['Normal'])],
                [Paragraph("<b>نوع الشحنة / Cargo Type:</b>", styles['Normal']), Paragraph(str(data['cargo_type']), styles['Normal']),
                 Paragraph("<b>الوزن الفعلي (kg):</b>", styles['Normal']), Paragraph(str(data['weight']), styles['Normal'])],
                [Paragraph("<b>اسم المستلم / Receiver Name:</b>", styles['Normal']), Paragraph(str(data['receiver_name']), styles['Normal']),
                 Paragraph("<b>التوقيع / Signature:</b>", styles['Normal']), Paragraph("____________________", styles['Normal'])]
            ]
            
            t = Table(table_data, colWidths=[120, 150, 120, 150])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F8F9FA")),
                ('GRID', (0,0), (-1,-1), 1, colors.HexColor("#1F4E78")),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('TOPPADDING', (0,0), (-1,-1), 8),
                ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ]))
            
            elements.append(t)
            elements.append(Spacer(1, 10))
            
            dec_style = ParagraphStyle('DecStyle', parent=styles['Normal'], alignment=1, fontSize=8.5, textColor=colors.HexColor("#444444"))
            elements.append(Paragraph("<b>إقرار الاستلام والتسليم:</b> أقر أنا المستلم بأن الشحنة المذكورة أعلاه قد استلمتها بحالة سليمة وكاملة، ولا يحق لي المطالبة بأي نقص بعد التوقيع والاستلام.", dec_style))
            elements.append(Spacer(1, 15))
            
            return elements

        # توليد وصل فردي
        def generate_receipt_pdf(data):
            buffer = BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
            styles = getSampleStyleSheet()
            elements = build_single_receipt_elements(data, styles)
            doc.build(elements)
            buffer.seek(0)
            return buffer

        # توليد جميع وصولات العميل دفعة واحدة (Batch PDF)
        def generate_batch_receipts_pdf(f_df, base_data):
            buffer = BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
            styles = getSampleStyleSheet()
            elements = []
            
            for index, r in f_df.iterrows():
                single_data = {
                    "receipt_no": f"{base_data['receipt_no']}-{index+1}",
                    "date": base_data['date'],
                    "client_name": base_data['client_name'],
                    "phone": base_data['phone'],
                    "cargo_name": str(r.get(mark_col, '')),
                    "cargo_type": str(r.get('نوع البضاعة', '')),
                    "pieces": str(r.get('عدد الكارتون', '1')),
                    "weight": str(r.get('الوزن', '0')),
                    "receiver_name": base_data['receiver_name']
                }
                receipt_elems = build_single_receipt_elements(single_data, styles)
                elements.extend(receipt_elems)
                # إضافة فاصل صفحات بين الوصلات إذا كانت متعددة
                elements.append(PageBreak())
                
            doc.build(elements)
            buffer.seek(0)
            return buffer

        st.markdown("---")
        col_btn1, col_btn2 = st.columns(2)
        
        with col_btn1:
            if st.button("🖨️ إصدار وتحميل الوصل المحدد (PDF)"):
                receipt_payload = {
                    "receipt_no": receipt_no, "date": receipt_date, "client_name": client_name_val,
                    "phone": phone_val, "cargo_name": cargo_name_val, "cargo_type": cargo_type_val,
                    "pieces": pieces_val, "weight": weight_val, "receiver_name": receiver_name_val
                }
                pdf_out = generate_receipt_pdf(receipt_payload)
                st.success("تم إصدار الوصل الفردي بنجاح!")
                st.download_button(
                    label="📥 تحميل الوصل الفردي (PDF)", data=pdf_out,
                    file_name=f"Receipt_{selected_client}_{selected_mark}.pdf", mime="application/pdf"
                )
                
        with col_btn2:
            if st.button("🚀 إصدار جميع وصولات هذا العميل بملف واحد (PDF)"):
                base_payload = {
                    "receipt_no": receipt_no, "date": receipt_date, "client_name": client_name_val,
                    "phone": phone_val, "receiver_name": receiver_name_val
                }
                batch_pdf_out = generate_batch_receipts_pdf(filtered_df, base_payload)
                st.success(f"تم إصدار جميع وصولات العميل ({len(filtered_df)} شحنة) في ملف واحد بنجاح!")
                st.download_button(
                    label="📥 تحميل ملف كافة الوصولات دفعة واحدة (PDF)", data=batch_pdf_out,
                    file_name=f"All_Receipts_{selected_client}.pdf", mime="application/pdf"
                )
else:
    st.info("📌 يرجى رفع ملف الشحنات للبدء.")
