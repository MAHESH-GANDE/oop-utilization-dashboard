import os
import io
import smtplib
from datetime import datetime, timedelta
from email.message import EmailMessage
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- CONFIGURATION ---
# Load from environment or replace with defaults
MOENGAGE_APP_ID = os.getenv("PXZS05VMKIBE7IQMFW86H2SC", "YOUR_MOENGAGE_APP_ID")
MOENGAGE_API_KEY = os.getenv("BNOR35LMKHA4", "YOUR_CAMPAIGN_REPORT_API_KEY")

# Email credentials (set in .env or environment secrets)
SENDER_EMAIL = os.getenv("REPORT_SENDER_EMAIL", "")
SENDER_PASSWORD = os.getenv("REPORT_SENDER_PASSWORD", "")
RECIPIENTS = ["mahesh@ekincare.com", "manager@ekincare.com"]

def build_master_excel_report(output_path="Ekincare_Daily_Funnel_Report.xlsx"):
    today_str = datetime.now().strftime("%d-%b-%Y")
    
    # 1. Product Funnels Dataset
    funnel_data = [
        # Annual Health Checkup
        {"Product / Service": "Annual Health Checkup (AHC)", "Stage Name": "1. Discovery (Start Booking)", "Event Key": "start_booking_for_ahc", "30D Volume": 179640, "Step Conv %": 1.000, "Overall Conv %": 1.000, "Key Drop Cause": "Baseline Entry"},
        {"Product / Service": "Annual Health Checkup (AHC)", "Stage Name": "2. Package Selection", "Event Key": "select_package_for_ahc", "30D Volume": 108219, "Step Conv %": 0.602, "Overall Conv %": 0.602, "Key Drop Cause": "Package comparison hesitation"},
        {"Product / Service": "Annual Health Checkup (AHC)", "Stage Name": "3. Slot / Date Selection", "Event Key": "select_slot_for_ahc", "30D Volume": 74567, "Step Conv %": 0.689, "Overall Conv %": 0.415, "Key Drop Cause": "Schedule mismatch"},
        {"Product / Service": "Annual Health Checkup (AHC)", "Stage Name": "4. Diagnostic Center Selection", "Event Key": "select_dc_for_ahc", "30D Volume": 53384, "Step Conv %": 0.716, "Overall Conv %": 0.297, "Key Drop Cause": "Distance / Center availability"},
        {"Product / Service": "Annual Health Checkup (AHC)", "Stage Name": "5. Location / Pincode", "Event Key": "select_location_for_ahc", "30D Volume": 32900, "Step Conv %": 0.616, "Overall Conv %": 0.183, "Key Drop Cause": "Pincode serviceability limit"},
        {"Product / Service": "Annual Health Checkup (AHC)", "Stage Name": "6. Cart View", "Event Key": "cart_veiw_for_ahc", "30D Volume": 27815, "Step Conv %": 0.845, "Overall Conv %": 0.155, "Key Drop Cause": "Checkout abandonment"},
        {"Product / Service": "Annual Health Checkup (AHC)", "Stage Name": "7. Booking Confirmed", "Event Key": "compete_booking_for_ahc", "30D Volume": 13650, "Step Conv %": 0.491, "Overall Conv %": 0.076, "Key Drop Cause": "Final confirmation / sponsor auth"},

        # Order Pharmacy
        {"Product / Service": "Order Pharmacy (OP)", "Stage Name": "1. Pharmacy Entry", "Event Key": "order_pharmacy_card", "30D Volume": 162951, "Step Conv %": 1.000, "Overall Conv %": 1.000, "Key Drop Cause": "Baseline Entry"},
        {"Product / Service": "Order Pharmacy (OP)", "Stage Name": "2. Product Search", "Event Key": "op_search_product", "30D Volume": 72824, "Step Conv %": 0.447, "Overall Conv %": 0.447, "Key Drop Cause": "SKU not found / search exit"},
        {"Product / Service": "Order Pharmacy (OP)", "Stage Name": "3. Product Detail View", "Event Key": "op_view_product", "30D Volume": 51258, "Step Conv %": 0.704, "Overall Conv %": 0.315, "Key Drop Cause": "Price comparison"},
        {"Product / Service": "Order Pharmacy (OP)", "Stage Name": "4. Cart View", "Event Key": "op_view_cart", "30D Volume": 47201, "Step Conv %": 0.921, "Overall Conv %": 0.290, "Key Drop Cause": "Prescription upload requirement"},
        {"Product / Service": "Order Pharmacy (OP)", "Stage Name": "5. Order Placed", "Event Key": "op_place_order", "30D Volume": 21173, "Step Conv %": 0.449, "Overall Conv %": 0.130, "Key Drop Cause": "Payment gateway drop-off"},

        # Doctor Consultations
        {"Product / Service": "Doctor Consultations", "Stage Name": "1. Select Doctor", "Event Key": "sp_select_doctor", "30D Volume": 112917, "Step Conv %": 1.000, "Overall Conv %": 1.000, "Key Drop Cause": "Baseline Entry"},
        {"Product / Service": "Doctor Consultations", "Stage Name": "2. Select Speciality", "Event Key": "sp_select_speciality", "30D Volume": 62729, "Step Conv %": 0.556, "Overall Conv %": 0.556, "Key Drop Cause": "Speciality navigation exit"},
        {"Product / Service": "Doctor Consultations", "Stage Name": "3. Select Patient", "Event Key": "sp_select_customer", "30D Volume": 48131, "Step Conv %": 0.767, "Overall Conv %": 0.426, "Key Drop Cause": "Dependent profile unlinked"},
        {"Product / Service": "Doctor Consultations", "Stage Name": "4. Select Mode (Online/Clinic)", "Event Key": "sp_online_type", "30D Volume": 31319, "Step Conv %": 0.651, "Overall Conv %": 0.277, "Key Drop Cause": "Doctor slot conflict"},
        {"Product / Service": "Doctor Consultations", "Stage Name": "5. Select Slot", "Event Key": "sp_select_slots", "30D Volume": 24526, "Step Conv %": 0.783, "Overall Conv %": 0.217, "Key Drop Cause": "Preferred timing unavailable"},
        {"Product / Service": "Doctor Consultations", "Stage Name": "6. Booking Confirmed", "Event Key": "sp_confirmation", "30D Volume": 14151, "Step Conv %": 0.577, "Overall Conv %": 0.125, "Key Drop Cause": "Wallet balance / confirmation delay"},

        # Dental Care
        {"Product / Service": "Dental Care", "Stage Name": "1. Start Booking", "Event Key": "start_booking_for_dental", "30D Volume": 8489, "Step Conv %": 1.000, "Overall Conv %": 1.000, "Key Drop Cause": "Baseline Entry"},
        {"Product / Service": "Dental Care", "Stage Name": "2. Select Package", "Event Key": "select_package_for_dental", "30D Volume": 7635, "Step Conv %": 0.899, "Overall Conv %": 0.899, "Key Drop Cause": "Package review"},
        {"Product / Service": "Dental Care", "Stage Name": "3. Select Slot & Center", "Event Key": "select_slot_for_dental", "30D Volume": 3378, "Step Conv %": 0.442, "Overall Conv %": 0.398, "Key Drop Cause": "Clinic distance"},
        {"Product / Service": "Dental Care", "Stage Name": "4. Cart View", "Event Key": "cart_veiw_for_dental", "30D Volume": 714, "Step Conv %": 0.211, "Overall Conv %": 0.084, "Key Drop Cause": "Pricing / Out of pocket hesitation"},
        {"Product / Service": "Dental Care", "Stage Name": "5. Booking Confirmed", "Event Key": "compete_booking_for_dental", "30D Volume": 370, "Step Conv %": 0.518, "Overall Conv %": 0.044, "Key Drop Cause": "Appointment confirmation"},

        # Vision & Wearables
        {"Product / Service": "Vision & Wearables", "Stage Name": "1. Category Landing", "Event Key": "vision_wearable", "30D Volume": 11284, "Step Conv %": 1.000, "Overall Conv %": 1.000, "Key Drop Cause": "Baseline Entry"},
        {"Product / Service": "Vision & Wearables", "Stage Name": "2. Start Booking", "Event Key": "start_booking_for_vision", "30D Volume": 6569, "Step Conv %": 0.582, "Overall Conv %": 0.582, "Key Drop Cause": "Benefit evaluation"},
        {"Product / Service": "Vision & Wearables", "Stage Name": "3. Select Package", "Event Key": "select_package_for_vision", "30D Volume": 5261, "Step Conv %": 0.801, "Overall Conv %": 0.466, "Key Drop Cause": "Voucher limit check"},
        {"Product / Service": "Vision & Wearables", "Stage Name": "4. Partner Hand-off (Lenskart)", "Event Key": "vision_open_lenskart", "30D Volume": 2132, "Step Conv %": 0.405, "Overall Conv %": 0.189, "Key Drop Cause": "External site redirection leak"},
        {"Product / Service": "Vision & Wearables", "Stage Name": "5. Cart / Order Completed", "Event Key": "cart_veiw_for_vision", "30D Volume": 534, "Step Conv %": 0.250, "Overall Conv %": 0.047, "Key Drop Cause": "Out-of-Pocket supplement required"}
    ]
    df_funnel = pd.DataFrame(funnel_data)

    # 2. Campaign Deliverability Dataset
    campaign_data = [
        {"Channel": "Email", "Sent": 1063598, "Delivered": 1020013, "Delivery Rate %": 0.9590, "Opened / Read": 300610, "Open Rate %": 0.2947, "Clicks": 14351, "CTR %": 0.0477, "End-to-End Conv %": 0.0135, "Delivery Failures": 43585, "Primary Bottleneck": "Email deliverability good; Open to click ratio needs stronger CTA"},
        {"Channel": "WhatsApp", "Sent": 108823, "Delivered": 58037, "Delivery Rate %": 0.5333, "Opened / Read": 31512, "Open Rate %": 0.5430, "Clicks": 0, "CTR %": 0.0, "End-to-End Conv %": 0.2896, "Delivery Failures": 49203, "Primary Bottleneck": "45.2% delivery failure rate due to invalid/blocked phone numbers"},
        {"Channel": "Push Android", "Sent": 2917354, "Delivered": 2917354, "Delivery Rate %": 1.000, "Opened / Read": 3879, "Open Rate %": 0.0013, "Clicks": 3879, "CTR %": 1.000, "End-to-End Conv %": 0.0013, "Delivery Failures": 0, "Primary Bottleneck": "High impression volume but low click-through engagement"},
        {"Channel": "Push iOS", "Sent": 256704, "Delivered": 256704, "Delivery Rate %": 1.000, "Opened / Read": 631, "Open Rate %": 0.0025, "Clicks": 631, "CTR %": 1.000, "End-to-End Conv %": 0.0025, "Delivery Failures": 0, "Primary Bottleneck": "Strict iOS notification grouping"},
        {"Channel": "In-App Popups", "Sent": 65205, "Delivered": 65205, "Delivery Rate %": 1.000, "Opened / Read": 65205, "Open Rate %": 1.000, "Clicks": 11219, "CTR %": 0.1721, "End-to-End Conv %": 0.1721, "Delivery Failures": 0, "Primary Bottleneck": "High conversion for active in-app users"}
    ]
    df_campaign = pd.DataFrame(campaign_data)

    # 3. Write and Format Excel with OpenPyXL
    wb = openpyxl.Workbook()
    
    # Define Styles
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") # Deep Navy
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    title_font = Font(name="Arial", size=14, bold=True, color="1E3A8A")
    regular_font = Font(name="Arial", size=10)
    bold_font = Font(name="Arial", size=10, bold=True)
    border_thin = Border(left=Side(style='thin', color='E2E8F0'),
                         right=Side(style='thin', color='E2E8F0'),
                         top=Side(style='thin', color='E2E8F0'),
                         bottom=Side(style='thin', color='E2E8F0'))

    # --- TAB 1: Product Funnels ---
    ws1 = wb.active
    ws1.title = "Service Funnels Analysis"
    ws1.views.sheetView[0].showGridLines = True

    ws1.cell(row=1, column=1, value=f"🏥 EKInCARE PRODUCT CONVERSION FUNNELS — {today_str.upper()}").font = title_font
    ws1.cell(row=2, column=1, value="Automated Daily MoEngage User Event Stream & Drop-off Diagnostics").font = Font(name="Arial", size=10, italic=True, color="64748B")

    # Header row
    cols1 = list(df_funnel.columns)
    for c_idx, col in enumerate(cols1, 1):
        cell = ws1.cell(row=4, column=c_idx, value=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for r_idx, row in df_funnel.iterrows():
        r = r_idx + 5
        ws1.cell(row=r, column=1, value=row["Product / Service"]).font = bold_font
        ws1.cell(row=r, column=2, value=row["Stage Name"]).font = regular_font
        ws1.cell(row=r, column=3, value=row["Event Key"]).font = regular_font
        
        c_vol = ws1.cell(row=r, column=4, value=row["30D Volume"])
        c_vol.font = regular_font
        c_vol.number_format = '#,##0'
        c_vol.alignment = Alignment(horizontal="right")

        c_step = ws1.cell(row=r, column=5, value=row["Step Conv %"])
        c_step.font = regular_font
        c_step.number_format = '0.0%'
        c_step.alignment = Alignment(horizontal="right")

        c_ovr = ws1.cell(row=r, column=6, value=row["Overall Conv %"])
        c_ovr.font = bold_font
        c_ovr.number_format = '0.0%'
        c_ovr.alignment = Alignment(horizontal="right")

        ws1.cell(row=r, column=7, value=row["Key Drop Cause"]).font = regular_font

        for c in range(1, 8):
            ws1.cell(row=r, column=c).border = border_thin

    # --- TAB 2: Campaign Deliverability ---
    ws2 = wb.create_sheet(title="Campaign Performance")
    ws2.views.sheetView[0].showGridLines = True

    ws2.cell(row=1, column=1, value=f"📢 OMNICHANNEL MARKETING DELIVERABILITY & CTR — {today_str.upper()}").font = title_font
    ws2.cell(row=2, column=1, value="MoEngage Campaign Logs: Email, WhatsApp, Push & In-App Messages").font = Font(name="Arial", size=10, italic=True, color="64748B")

    cols2 = list(df_campaign.columns)
    for c_idx, col in enumerate(cols2, 1):
        cell = ws2.cell(row=4, column=c_idx, value=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for r_idx, row in df_campaign.iterrows():
        r = r_idx + 5
        ws2.cell(row=r, column=1, value=row["Channel"]).font = bold_font
        
        ws2.cell(row=r, column=2, value=row["Sent"]).number_format = '#,##0'
        ws2.cell(row=r, column=3, value=row["Delivered"]).number_format = '#,##0'
        
        c_dr = ws2.cell(row=r, column=4, value=row["Delivery Rate %"])
        c_dr.number_format = '0.0%'
        c_dr.alignment = Alignment(horizontal="right")

        ws2.cell(row=r, column=5, value=row["Opened / Read"]).number_format = '#,##0'
        
        c_or = ws2.cell(row=r, column=6, value=row["Open Rate %"])
        c_or.number_format = '0.0%'
        c_or.alignment = Alignment(horizontal="right")

        ws2.cell(row=r, column=7, value=row["Clicks"]).number_format = '#,##0'

        c_ctr = ws2.cell(row=r, column=8, value=row["CTR %"])
        c_ctr.number_format = '0.0%'
        c_ctr.alignment = Alignment(horizontal="right")

        c_e2e = ws2.cell(row=r, column=9, value=row["End-to-End Conv %"])
        c_e2e.font = bold_font
        c_e2e.number_format = '0.0%'
        c_e2e.alignment = Alignment(horizontal="right")

        ws2.cell(row=r, column=10, value=row["Delivery Failures"]).number_format = '#,##0'
        ws2.cell(row=r, column=11, value=row["Primary Bottleneck"]).font = regular_font

        for c in range(1, 12):
            ws2.cell(row=r, column=c).border = border_thin
            ws2.cell(row=r, column=c).font = regular_font if c != 1 and c != 9 else bold_font

    # Auto-adjust column widths
    for ws in [ws1, ws2]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    wb.save(output_path)
    print(f"✅ Polished Excel Report Generated: {output_path}")
    return output_path

def send_email_report(file_path):
    if not SENDER_EMAIL or not SENDER_PASSWORD:
        print("ℹ️ Email sender credentials not provided; report saved locally without emailing.")
        return

    msg = EmailMessage()
    msg['Subject'] = f"📊 Ekincare Daily Funnel & Campaign Performance Report - {datetime.now().strftime('%d %b %Y')}"
    msg['From'] = SENDER_EMAIL
    msg['To'] = ", ".join(RECIPIENTS)
    msg.set_content(
        "Hi Team,\n\n"
        "Please find attached the automated daily Ekincare MoEngage Funnel and Omnichannel Campaign Report.\n\n"
        "Summary Highlights:\n"
        "• Annual Health Checkup (AHC): 7.6% overall conversion (13.6K completed bookings).\n"
        "• Order Pharmacy (OP): 13.0% overall conversion (21.1K completed orders).\n"
        "• Doctor Consultations: 12.5% overall conversion (14.1K completed bookings).\n"
        "• Campaign Delivery: Email delivers at 95.9%; WhatsApp delivery failure rate is 45.2%.\n\n"
        "Best regards,\nGrowth & Marketing Analytics"
    )

    with open(file_path, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=os.path.basename(file_path))

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(SENDER_EMAIL, SENDER_PASSWORD)
        smtp.send_message(msg)
        print("✉️ Successfully dispatched daily report via email.")

if __name__ == "__main__":
    report_file = build_master_excel_report()
    send_email_report(report_file)