import base64

# --- داخل حلقة التكرار للوصولات ---
logo_html = ""
if os.path.exists(logo_path):
  with open(logo_path, "rb") as f:
    encoded_logo = base64.b64encode(f.read()).decode("utf-8")
    logo_html = f'<img src="data:image/png;base64,{encoded_logo}" style="max-height: 40px; vertical-align: middle; margin-left: 10px;" />'

single_receipt_html = f"""
    <div class="receipt-page" style="
        padding: 15px; 
        font-family: 'Tahoma', Arial, sans-serif; 
        direction: rtl; 
        border: 2px solid #102a43; 
        width: 100%; 
        max-width: 148mm; 
        margin: auto auto 20px auto; 
        background: #ffffff; 
        color: #102a43;
        box-sizing: border-box;
        page-break-after: always;
        break-after: page;
    ">
        <!-- رأس الوصل مع الشعار -->
        <table style="width: 100%; border-bottom: 2px solid #102a43; padding-bottom: 8px; margin-bottom: 12px; border-collapse: collapse;">
            <tr>
                <td style="text-align: right; vertical-align: middle;">
                    <div style="display: flex; align-items: center;">
                        {logo_html}
                        <div>
                            <h2 style="margin: 0; font-size: 15px; color: #102a43;">أطلس المحيط للتجارة العامة</h2>
                            <p style="margin: 2px 0 0; font-size: 10px; color: #627d98;">OCEAN ATLAS GENERAL TRADING</p>
                        </div>
                    </div>
                </td>
                <td style="text-align: left; vertical-align: middle;">
                    <h3 style="margin: 0; font-size: 13px; color: #b45309;">وصل تسليم بضاعة</h3>
                    <p style="margin: 2px 0 0; font-size: 10px; color: #334e68;">Cargo Delivery Receipt</p>
                </td>
            </tr>
        </table>
        <!-- باقي تفاصيل الوصل... -->
    </div>
"""
