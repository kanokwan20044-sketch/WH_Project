import os
import json
import datetime
from flask import Flask, render_template, request, redirect, url_for, jsonify
import gspread
from google.oauth2.service_account import Credentials

app = Flask(__name__)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def get_google_sheet(sheet_name='DATA_WH'):
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
    sheet = client.open('Warehouse_Chacha').worksheet(sheet_name)
    return sheet

@app.route('/')
def index():
    success = request.args.get('success', '')
    search_query = request.args.get('q', '').strip()
    start_date = request.args.get('start_date', '').strip()
    end_date = request.args.get('end_date', '').strip()
    search_results = []

    try:
        sheet = get_google_sheet('DATA_WH')
        all_rows = sheet.get_all_values()
        
        for row in all_rows[1:]:
            if len(row) >= 8:
                row_time = row[0]
                row_location = row[1]
                row_status = row[2]
                item_code = row[3]
                product_name = row[4]
                lot_serial = row[5]
                qty = row[6]
                unit = row[7]
                
                row_date = row_time.split(' ')[0] if ' ' in row_time else ''
                
                match_query = True
                match_date = True

                if search_query:
                    if (search_query.lower() not in item_code.lower() and 
                        search_query.lower() not in product_name.lower() and 
                        search_query.lower() not in lot_serial.lower() and 
                        search_query.lower() not in row_location.lower() and
                        search_query.lower() not in row_status.lower()):
                        match_query = False

                if start_date and end_date and row_date:
                    if not (start_date <= row_date <= end_date):
                        match_date = False

                if (search_query or start_date or end_date) and match_query and match_date:
                    search_results.append({
                        'timestamp': row_time,
                        'location': row_location,
                        'status': row_status,
                        'item_code': item_code,
                        'product_name': product_name,
                        'lot_serial': lot_serial,
                        'quantity': qty,
                        'unit': unit
                    })
    except Exception as e:
        print(f"เกิดข้อผิดพลาดในการค้นหา: {e}")

    return render_template('index.html', success=success, search_results=search_results, search_query=search_query, start_date=start_date, end_date=end_date)

# API ค้นหาข้อมูลสินค้าและ Location สำหรับ Outbound
@app.route('/search_outbound')
def search_outbound():
    code = request.args.get('code', '').strip()
    if not code:
        return jsonify({'success': False})
    
    try:
        # 1. ดึงข้อมูลจาก Master_code
        master_sheet = get_google_sheet('Master_code')
        master_rows = master_sheet.get_all_values()
        prod_name = ""
        unit = ""
        lot_serial = ""
        
        for row in master_rows[1:]:
            if len(row) >= 3 and row[0].strip().lower() == code.lower():
                prod_name = row[1]
                unit = row[2]
                lot_serial = row[3] if len(row) > 3 else ''
                break

        # 2. ค้นหา Location จาก DATA_WH ที่มีการเก็บสินค้านี้อยู่ (เช็คประวัติ Inbound ล่าสุด)
        data_sheet = get_google_sheet('DATA_WH')
        data_rows = data_sheet.get_all_values()
        found_location = "ไม่พบพิกัด (ยังไม่มีประวัติ Inbound)"
        
        for row in reversed(data_rows[1:]):
            if len(row) >= 4 and row[3].strip().lower() == code.lower():
                if row[2].strip() == 'Inbound':
                    found_location = row[1]
                    break

        return jsonify({
            'success': True,
            'product_name': prod_name if prod_name else 'ไม่พบรหัสสินค้าใน Master_code',
            'unit': unit if unit else 'Piece',
            'lot_serial': lot_serial,
            'location': found_location
        })

    except Exception as e:
        print(f"Error search_outbound: {e}")
        
    return jsonify({'success': False})

@app.route('/submit', methods=['POST'])
def submit():
    try:
        location = request.form.get('location', 'ไม่ได้ระบุพิกัด')
        status = request.form.get('status', 'Outbound')
        item_code = request.form.get('item_code', '')
        product_name = request.form.get('product_name', '')
        lot_serial = request.form.get('lot_serial', '')
        quantity = request.form.get('quantity', '1')
        unit = request.form.get('unit', '')
        
        utc_now = datetime.datetime.utcnow()
        th_now = utc_now + datetime.timedelta(hours=7)
        now = th_now.strftime('%Y-%m-%d %H:%M:%S')
        
        sheet = get_google_sheet('DATA_WH')
        sheet.append_row([now, location, status, item_code, product_name, lot_serial, quantity, unit])
        
        return redirect(url_for('index', success='1'))
        
    except Exception as e:
        print(f"เกิดข้อผิดพลาดในการบันทึกข้อมูล: {e}")
        return f"เกิดข้อผิดพลาดในการบันทึกข้อมูลลง Google Sheets: {e}", 500

if __name__ == '__main__':
    app.run(debug=True)