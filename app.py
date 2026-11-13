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
    """ฟังก์ชันเชื่อมต่อ Google Sheets (ปลอดภัย ไม่ฮาร์ดโค้ดคีย์ลับในโค้ด)"""
    if os.path.exists('credentials.json'):
        creds = Credentials.from_service_account_file('credentials.json', scopes=SCOPES)
    else:
        # รองรับกรณีดึงจาก Environment Variables บน Render
        creds_json = os.environ.get('CREDENTIALS_JSON')
        if creds_json:
            creds_info = json.loads(creds_json)
            creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
        else:
            raise Exception("ไม่พบไฟล์หรือข้อมูล Credentials สำหรับเชื่อมต่อ Google Sheets")
    
    client = gspread.authorize(creds)
    # เปิดไฟล์ Google Sheets ชื่อ Warehouse_Chacha
    sheet = client.open('Warehouse_Chacha').sheet1
    return sheet

@app.route('/')
def index():
    # รับค่าพิกัดล็อกสินค้าจาก QR Code และสถานะความสำเร็จ
    location = request.args.get('location', '')
    success = request.args.get('success', '')
    return render_template('index.html', location=location, success=success)

@app.route('/submit', methods=['POST'])
def submit():
    try:
        # ดึงข้อมูลจากฟอร์ม WMS
        location = request.form.get('location', 'ไม่ได้ระบุพิกัด')
        status = request.form.get('status', '')
        product_name = request.form.get('product_name', '')
        qty = request.form.get('qty', '')
        company = request.form.get('company', '')
        job_order = request.form.get('job_order', '')
        
        # วันที่และเวลาปัจจุบัน
        now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # บันทึกข้อมูลลง Google Sheets (Warehouse_Chacha)
        sheet = get_google_sheet()
        sheet.append_row([now, location, status, product_name, qty, company, job_order])
        
        # บันทึกสำเร็จ กลับมาหน้าแรกพร้อมแสดงแถบแจ้งเตือน
        return redirect(url_for('index', success='1'))
        
    except Exception as e:
        print(f"เกิดข้อผิดพลาดในการบันทึกข้อมูล: {e}")
        return f"เกิดข้อผิดพลาดในการบันทึกข้อมูลลง Google Sheets: {e}", 500

if __name__ == '__main__':
    app.run(debug=True)
    