import os
import json
import datetime
from flask import Flask, render_template, request, redirect, url_for
import gspread
from google.oauth2.service_account import Credentials

app = Flask(__name__)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def get_google_sheet():
    if os.path.exists('credentials.json'):
        creds = Credentials.from_service_account_file('credentials.json', scopes=SCOPES)
    else:
        creds_json = os.environ.get('CREDENTIALS_JSON')
        if creds_json:
            creds_info = json.loads(creds_json)
            creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
        else:
            raise Exception("ไม่พบไฟล์หรือข้อมูล Credentials สำหรับเชื่อมต่อ Google Sheets")
    
    client = gspread.authorize(creds)
    sheet = client.open('Warehouse_Chacha').sheet1
    return sheet

@app.route('/')
def index():
    location = request.args.get('location', '')
    success = request.args.get('success', '')
    search_query = request.args.get('q', '').strip()
    start_date = request.args.get('start_date', '').strip()
    end_date = request.args.get('end_date', '').strip()
    search_results = []

    try:
        sheet = get_google_sheet()
        all_rows = sheet.get_all_values()
        
        # ค้นหาข้อมูลใน Google Sheets [0:Timestamp, 1:Location, 2:Status, 3:Product Name, 4:Lot Number, 5:Qty]
        for row in all_rows[1:]:
            if len(row) >= 6:
                row_time = row[0]
                row_date = row_time.split(' ')[0]
                product_name = row[3]
                lot_number = row[4]
                status = row[2]
                
                match_query = True
                match_date = True

                if search_query:
                    if (search_query.lower() not in product_name.lower() and 
                        search_query.lower() not in lot_number.lower() and 
                        search_query.lower() not in row[1].lower() and
                        search_query.lower() not in status.lower()):
                        match_query = False

                if start_date and end_date:
                    if not (start_date <= row_date <= end_date):
                        match_date = False
                elif start_date:
                    if row_date < start_date:
                        match_date = False
                elif end_date:
                    if row_date > end_date:
                        match_date = False

                if (search_query or start_date or end_date) and match_query and match_date:
                    search_results.append({
                        'timestamp': row[0],
                        'location': row[1],
                        'status': row[2],
                        'product_name': row[3],
                        'lot_number': row[4],
                        'qty': row[5]
                    })
    except Exception as e:
        print(f"เกิดข้อผิดพลาดในการค้นหา: {e}")

    return render_template('index.html', location=location, success=success, search_results=search_results, search_query=search_query, start_date=start_date, end_date=end_date)

@app.route('/submit', methods=['POST'])
def submit():
    try:
        location = request.form.get('location', 'ไม่ได้ระบุพิกัด')
        status = request.form.get('status', '')
        product_name = request.form.get('product_name', '-')
        lot_number = request.form.get('lot_number', '-')
        qty = request.form.get('qty', '')
        
        # ถ้ารายการไหนไม่ใช่ Outbound ให้บันทึกชื่อสินค้าเป็นขีด (-) เพื่อความเป็นระเบียบ
        if status != 'Outbound':
            product_name = '-'
        
        # ตั้งค่าเวลาประเทศไทย (UTC+7)
        utc_now = datetime.datetime.utcnow()
        th_now = utc_now + datetime.timedelta(hours=7)
        now = th_now.strftime('%Y-%m-%d %H:%M:%S')
        
        sheet = get_google_sheet()
        sheet.append_row([now, location, status, product_name, lot_number, qty])
        
        return redirect(url_for('index', success='1'))
        
    except Exception as e:
        print(f"เกิดข้อผิดพลาดในการบันทึกข้อมูล: {e}")
        return f"เกิดข้อผิดพลาดในการบันทึกข้อมูลลง Google Sheets: {e}", 500

if __name__ == '__main__':
    app.run(debug=True)