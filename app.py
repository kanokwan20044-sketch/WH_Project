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
                row_loc = row[1].strip()
                if row_loc.lower() != loc.lower():
                    continue

                status = row[2].strip().lower()
                item_code = row[3].strip()
                product_name = row[4].strip()
                lot_serial = row[5].strip()
                qty = parse_qty(row[6])
                unit = row[7].strip()

                key = (item_code, product_name, lot_serial, unit)
                if key not in stock_map:
                    stock_map[key] = 0.0

                if status == 'inbound':
                    stock_map[key] += qty
                elif status == 'outbound':
                    stock_map[key] -= qty

        active_items = []
        for (item_code, product_name, lot_serial, unit), balance in stock_map.items():
            if balance > 0:
                active_items.append({
                    'item_code': item_code,
                    'product_name': product_name,
                    'lot_serial': lot_serial,
                    'quantity': format_qty(balance),
                    'unit': unit
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
    code = request.args.get('code', '').strip()
    if not code:
        return jsonify({'success': False})

    try:
        master_sheet = get_google_sheet('Master_code')
        rows = master_sheet.get_all_values()
        for row in rows[1:]:
            if len(row) >= 3 and row[0].strip().lower() == code.lower():
                return jsonify({
                    'success': True,
                    'product_name': row[1],
                    'unit': row[2],
                    'lot_serial': row[3] if len(row) > 3 else '',
                    'pallet_no': row[4] if len(row) > 4 else '',
                    'description': row[5] if len(row) > 5 else row[1]
                })
    except Exception as e:
        print(f"Error fetching Master_code: {e}")

    return jsonify({'success': False})


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
            if len(row) >= 3 and row[0].strip().lower() == code.lower():
                master_name = row[1].strip()
                master_unit = row[2].strip()
                break

        data_sheet = get_google_sheet('DATA_WH')
        data_rows = data_sheet.get_all_values()

        lot_stock = {}
        inbound_records = []

        for row in data_rows[1:]:
            if len(row) >= 8 and row[3].strip().lower() == code.lower():
                row_loc = row[1].strip()
                status = row[2].strip().lower()
                prod_name = row[4].strip() or master_name or 'ไม่ระบุชื่อสินค้า'
                lot_serial = row[5].strip()
                qty = parse_qty(row[6])
                unit = row[7].strip() or master_unit or 'Piece'

                key = (lot_serial, row_loc, prod_name, unit)
                if key not in lot_stock:
                    lot_stock[key] = 0.0

                if status == 'inbound':
                    lot_stock[key] += qty
                    inbound_records.append({
                        'product_name': prod_name,
                        'lot_serial': lot_serial,
                        'location': row_loc if row_loc else 'ไม่ระบุพิกัด',
                        'unit': unit
                    })
                elif status == 'outbound':
                    lot_stock[key] -= qty

        options = []
        for (lot_serial, row_loc, prod_name, unit), balance in lot_stock.items():
            if balance > 0:
                options.append({
                    'product_name': prod_name,
                    'lot_serial': lot_serial if lot_serial else '-',
                    'location': row_loc if row_loc else 'ไม่ระบุพิกัด',
                    'balance': format_qty(balance),
                    'unit': unit
                })

        if not options and inbound_records:
            seen = set()
            for rec in reversed(inbound_records):
                uniq = (rec['lot_serial'], rec['location'], rec['product_name'], rec['unit'])
                if uniq not in seen:
                    seen.add(uniq)
                    options.append({
                        'product_name': rec['product_name'],
                        'lot_serial': rec['lot_serial'] if rec['lot_serial'] else '-',
                        'location': rec['location'],
                        'balance': 0,
                        'unit': rec['unit']
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
        location = request.form.get('location', '').strip() or 'ไม่ได้ระบุพิกัด'
        status = request.form.get('status', '').strip()
        item_code = request.form.get('item_code', '').strip()
        product_name = request.form.get('product_name', '').strip()
        lot_serial = request.form.get('lot_serial', '').strip()
        quantity = request.form.get('quantity', '1').strip()
        unit = request.form.get('unit', '').strip()

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