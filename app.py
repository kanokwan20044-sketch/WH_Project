import os
import json
import datetime
from flask import Flask, render_template, request, redirect, url_for, jsonify
import gspread
from google.oauth2.service_account import Credentials

app = Flask(__name__)

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]


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
    return client.open('Warehouse_Chacha').worksheet(sheet_name)


def parse_qty(val):
    try:
        return float(str(val).replace(',', '').strip())
    except Exception:
        return 0.0


def format_qty(num):
    return int(num) if float(num).is_integer() else round(float(num), 2)


def clean_txt(val):
    return ' '.join(str(val or '').replace('\u00a0', ' ').split()).strip()


@app.route('/')
def index():
    location = request.args.get('location', '').strip()
    success = request.args.get('success', '').strip()
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
                status = row[2]
                item_code = row[3]
                product_name = row[4]
                lot_serial = row[5]
                qty = row[6]
                unit = row[7]

                row_date = row_time.split(' ')[0] if ' ' in row_time else ''

                match_query = True
                match_date = True

                if search_query:
                    q_lower = search_query.lower()
                    if (
                        q_lower not in item_code.lower()
                        and q_lower not in product_name.lower()
                        and q_lower not in lot_serial.lower()
                        and q_lower not in row_location.lower()
                        and q_lower not in status.lower()
                    ):
                        match_query = False

                if start_date and end_date and row_date:
                    if not (start_date <= row_date <= end_date):
                        match_date = False
                elif start_date and row_date:
                    if row_date < start_date:
                        match_date = False
                elif end_date and row_date:
                    if row_date > end_date:
                        match_date = False

                if (search_query or start_date or end_date) and match_query and match_date:
                    search_results.append({
                        'timestamp': row_time,
                        'location': row_location,
                        'status': status,
                        'item_code': item_code,
                        'product_name': product_name,
                        'lot_serial': lot_serial,
                        'quantity': qty,
                        'unit': unit
                    })
    except Exception as e:
        print(f"เกิดข้อผิดพลาดในการค้นหา: {e}")

    return render_template(
        'index.html',
        location=location,
        success=success,
        search_results=search_results,
        search_query=search_query,
        start_date=start_date,
        end_date=end_date
    )


@app.route('/check_location_stock')
def check_location_stock():
    loc = request.args.get('location', '').strip()
    if not loc:
        return jsonify({'success': False, 'items': []})

    try:
        sheet = get_google_sheet('DATA_WH')
        all_rows = sheet.get_all_values()

        stock_map = {}
        for row in all_rows[1:]:
            if len(row) >= 8:
                row_loc = clean_txt(row[1])
                if row_loc.lower() != loc.lower():
                    continue

                status = clean_txt(row[2]).lower()
                item_code = clean_txt(row[3])
                product_name = clean_txt(row[4])
                lot_serial = clean_txt(row[5])
                qty = parse_qty(row[6])
                unit = clean_txt(row[7])

                key = (item_code.lower(), product_name.lower(), lot_serial.lower(), unit.lower())
                if key not in stock_map:
                    stock_map[key] = {
                        'item_code': item_code,
                        'product_name': product_name,
                        'lot_serial': lot_serial,
                        'unit': unit,
                        'balance': 0.0
                    }

                if status == 'inbound':
                    stock_map[key]['balance'] += qty
                elif status == 'outbound':
                    stock_map[key]['balance'] -= qty

        active_items = []
        for item in stock_map.values():
            if item['balance'] > 0:
                active_items.append({
                    'item_code': item['item_code'],
                    'product_name': item['product_name'],
                    'lot_serial': item['lot_serial'],
                    'quantity': format_qty(item['balance']),
                    'unit': item['unit']
                })

        return jsonify({
            'success': True,
            'location': loc,
            'has_stock': len(active_items) > 0,
            'items': active_items
        })
    except Exception as e:
        print(f"Error check_location_stock: {e}")
        return jsonify({'success': False, 'items': []})


@app.route('/get_product_info')
def get_product_info():
    raw_code = request.args.get('code', '').strip()
    code = raw_code.lower()
    if not code:
        return jsonify({'success': False, 'message': 'ไม่ได้ระบุรหัสสินค้า'})

    try:
        master_sheet = get_google_sheet('Master_code')
        rows = master_sheet.get_all_values()
        for row in rows[1:]:
            if not row or len(row) < 3:
                continue
            row_code = clean_txt(row[0]).lower()
            if row_code == code:
                fg_name = clean_txt(row[1])
                unit = clean_txt(row[2])
                kg_per_unit = clean_txt(row[3]) if len(row) > 3 and clean_txt(row[3]) else '1'
                return jsonify({
                    'success': True,
                    'item_code': clean_txt(row[0]),
                    'product_name': fg_name,
                    'description': fg_name,
                    'unit': unit,
                    'quantity': kg_per_unit
                })
    except Exception as e:
        print(f"Error fetching Master_code from Google Sheets: {e}")
        return jsonify({'success': False, 'message': f'เกิดข้อผิดพลาดในการดึงข้อมูลจาก Google Sheets: {e}'})

    return jsonify({'success': False, 'message': f'ไม่พบรหัส {raw_code} ในชีต Master_code'})


@app.route('/search_outbound')
def search_outbound():
    code = request.args.get('code', '').strip()
    if not code:
        return jsonify({'success': False})

    try:
        master_sheet = get_google_sheet('Master_code')
        master_rows = master_sheet.get_all_values()
        master_name = ""
        master_unit = ""

        for row in master_rows[1:]:
            if len(row) >= 3 and clean_txt(row[0]).lower() == code.lower():
                master_name = clean_txt(row[1])
                master_unit = clean_txt(row[2])
                break

        data_sheet = get_google_sheet('DATA_WH')
        data_rows = data_sheet.get_all_values()

        stock_by_loc = {}
        ordered_keys = []

        for row in data_rows[1:]:
            if len(row) >= 8 and clean_txt(row[3]).lower() == code.lower():
                row_loc = clean_txt(row[1]) or 'ไม่ระบุพิกัด'
                status = clean_txt(row[2]).lower()
                prod_name = clean_txt(row[4]) or master_name or 'ไม่ระบุชื่อสินค้า'
                lot_serial = clean_txt(row[5]) or '-'
                qty = parse_qty(row[6])
                unit = clean_txt(row[7]) or master_unit or 'Piece'

                key = (lot_serial.lower(), row_loc.lower())
                if key not in stock_by_loc:
                    stock_by_loc[key] = {
                        'product_name': prod_name,
                        'lot_serial': lot_serial,
                        'location': row_loc,
                        'unit': unit,
                        'balance': 0.0
                    }
                    ordered_keys.append(key)

                if status == 'inbound':
                    stock_by_loc[key]['balance'] += qty
                elif status == 'outbound':
                    stock_by_loc[key]['balance'] -= qty

        options = []
        for key in ordered_keys:
            info = stock_by_loc[key]
            if info['balance'] > 0:
                options.append({
                    'product_name': master_name or info['product_name'],
                    'lot_serial': info['lot_serial'],
                    'location': info['location'],
                    'balance': format_qty(info['balance']),
                    'unit': master_unit or info['unit']
                })

        if not options and ordered_keys:
            for key in reversed(ordered_keys):
                info = stock_by_loc[key]
                options.append({
                    'product_name': master_name or info['product_name'],
                    'lot_serial': info['lot_serial'],
                    'location': info['location'],
                    'balance': 0,
                    'unit': master_unit or info['unit']
                })

        return jsonify({
            'success': True,
            'item_code': code,
            'product_name': master_name if master_name else (options[0]['product_name'] if options else 'ไม่พบรหัสสินค้าใน Master_code'),
            'unit': master_unit if master_unit else (options[0]['unit'] if options else 'Piece'),
            'options': options
        })
    except Exception as e:
        print(f"Error search_outbound: {e}")

    return jsonify({'success': False})


@app.route('/submit', methods=['POST'])
def submit():
    try:
        location = clean_txt(request.form.get('location', '')) or 'ไม่ได้ระบุพิกัด'
        status = clean_txt(request.form.get('status', ''))
        item_code = clean_txt(request.form.get('item_code', ''))
        product_name = clean_txt(request.form.get('product_name', ''))
        lot_serial = clean_txt(request.form.get('lot_serial', ''))
        quantity = clean_txt(request.form.get('quantity', '1')) or '1'
        unit = clean_txt(request.form.get('unit', ''))

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