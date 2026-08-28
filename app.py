st.markdown("<br>", unsafe_allow_html=True)
    st.subheader(f"📋 جدول تفاصيل الشحنة المعروضة: [{selected_shipment_filter}] - النوع: [{selected_type_filter}]")

    display_table_df = df.copy()
    display_table_df.insert(0, "التسلسل", range(1, len(display_table_df) + 1))
    table_html = display_table_df.to_html(classes="custom-table", index=False, escape=False)

    custom_table_styling = f"""
    <style>
        .custom-table-container {{
            max-height: 450px;
            overflow-x: auto;
            overflow-y: auto;
            border: 1px solid #bcccdc;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            margin-bottom: 10px;
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
    <div id="tableToPrint" class="custom-table-container">
        {table_html}
    </div>
    
    <div style="text-align: left; margin-bottom: 20px;">
        <button style="background-color: #b45309; color: white; padding: 10px 20px; border: none; border-radius: 6px; cursor: pointer; font-weight: bold; font-size: 14px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);" onclick="exportTableToPDF()">📥 تصدير الجدول إلى PDF</button>
    </div>

    <script>
        function exportTableToPDF() {{
            var tableContent = document.getElementById('tableToPrint').innerHTML;
            var w = window.open('', '', 'height=800,width=1000');
            w.document.write('<html><head><title>تقرير الشحنات - أطلس</title><style>');
            w.document.write('body {{ direction: rtl; font-family: Tahoma, Arial, sans-serif; padding: 20px; color: #102a43; }}');
            w.document.write('table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}');
            w.document.write('th {{ background-color: #102a43 !important; color: #ffffff !important; text-align: right; padding: 10px; border: 1px solid #0b1e33; -webkit-print-color-adjust: exact; }}');
            w.document.write('td {{ padding: 8px 10px; border: 1px solid #bcccdc; text-align: right; }}');
            w.document.write('tr:nth-child(even) {{ background-color: #f8fafc !important; -webkit-print-color-adjust: exact; }}');
            w.document.write('h2 {{ text-align: center; color: #102a43; }}');
            w.document.write('</style></head><body>');
            w.document.write('<h2>جدول تفاصيل الشحنة: [{selected_shipment_filter}]</h2>');
            w.document.write(tableContent);
            w.document.write('</body></html>');
            w.document.write('<script>window.onload = function() {{ window.print(); }};<\/script>');
            w.document.close();
            w.focus();
        }}
    </script>
    """
    st.html(custom_table_styling)
