import os
import json
import datetime
from flask import Flask, render_template, request, redirect, url_for
import gspread
from google.oauth2.service_account import Credentials

app = Flask(__name__)

# กำหนดสิทธิ์การเข้าถึง Google Sheets และ Drive
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
    search_results = []

    if search_query:
        try:
            sheet = get_google_sheet()
            all_rows = sheet.get_all_values()
            for row in all_rows[1:]:
                if len(row) >= 4 and search_query.lower() in row[3].lower():
                    search_results.append({
                        'timestamp': row[0],
                        'location': row[1],
                        'status': row[2],
                        'product_name': row[3],
                        'qty': row[4],
                        'company': row[5] if len(row) > 5 else '',
                        'job_order': row[6] if len(row) > 6 else ''
                    })
        except Exception as e:
            print(f"เกิดข้อผิดพลาดในการค้นหา: {e}")

    return render_template('index.html', location=location, success=success, search_results=search_results, search_query=search_query)

@app.route('/submit', methods=['POST'])
def submit():
    try:
        location = request.form.get('location', 'ไม่ได้ระบุพิกัด')
        status = request.form.get('status', '')
        product_name = request.form.get('product_name', '')
        qty = request.form.get('qty', '')
        company = request.form.get('company', '')
        job_order = request.form.get('job_order', '')
        
        # บังคับดึงเวลาปัจจุบันบวกเพิ่ม 7 ชั่วโมงจาก UTC ตรงๆ เพื่อให้เป็นเวลาไทยแน่นอน
        utc_now = datetime.datetime.utcnow()
        th_now = utc_now + datetime.timedelta(hours=7)
        now = th_now.strftime('%Y-%m-%d %H:%M:%S')
        
        sheet = get_google_sheet()
        sheet.append_row([now, location, status, product_name, qty, company, job_order])
        
        return redirect(url_for('index', success='1'))
        
    except Exception as e:
        print(f"เกิดข้อผิดพลาดในการบันทึกข้อมูล: {e}")
        return f"เกิดข้อผิดพลาดในการบันทึกข้อมูลลง Google Sheets: {e}", 500

if __name__ == '__main__':
    app.run(debug=True)
    