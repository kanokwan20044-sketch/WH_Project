import os
import json
import datetime
import tempfile
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl
import qrcode
from PIL import Image, ImageDraw, ImageFont, ImageTk
import gspread
from google.oauth2.service_account import Credentials


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def parse_qty(val):
    try:
        return float(str(val).replace(",", "").strip())
    except Exception:
        return 0.0


def fmt_qty(num):
    return int(num) if float(num).is_integer() else round(float(num), 2)


class ChachaWMSDesktopApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ระบบ WMS & สร้างป้าย QR Code - บริษัท ชาช่าไทย จำกัด เชียงราย")
        self.root.geometry("980x680")
        self.root.configure(bg="#f7fafc")

        self.excel_path = "RUN_Barcode_Template.xlsx"
        self.excel_sheet_name = "Data"
        self.master_data = {}
        self.outbound_options = []
        self.label_pil_image = None
        self.label_tk_image = None

        self.build_ui()
        self.load_excel_data()

    def get_worksheet(self, sheet_name="DATA_WH"):
        if os.path.exists("credentials.json"):
            creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
        else:
            creds_json = os.environ.get("CREDENTIALS_JSON")
            if not creds_json:
                raise Exception("ไม่พบไฟล์ credentials.json ในโฟลเดอร์โปรเจกต์")
            creds = Credentials.from_service_account_info(json.loads(creds_json), scopes=SCOPES)

        client = gspread.authorize(creds)
        return client.open("Warehouse_Chacha").worksheet(sheet_name)

    def build_ui(self):
        header = tk.Frame(self.root, bg="#1a365d", pady=10, padx=16)
        header.pack(fill="x")
        tk.Label(
            header,
            text="ระบบ WMS คลังสินค้า & พิมพ์ป้าย QR Code ติดสินค้า",
            font=("Tahoma", 14, "bold"),
            bg="#1a365d",
            fg="white",
        ).pack(side="left")
        tk.Label(
            header,
            text="บริษัท ชาช่าไทย จำกัด เชียงราย",
            font=("Tahoma", 10),
            bg="#1a365d",
            fg="#bee3f8",
        ).pack(side="right")

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=12, pady=10)

        self.tab_inbound = tk.Frame(self.notebook, bg="#f7fafc")
        self.tab_outbound = tk.Frame(self.notebook, bg="#f7fafc")
        self.tab_cycle = tk.Frame(self.notebook, bg="#f7fafc")
        self.tab_qr = tk.Frame(self.notebook, bg="#f7fafc")

        self.notebook.add(self.tab_inbound, text="  1. นำเข้า (Inbound)  ")
        self.notebook.add(self.tab_outbound, text="  2. เบิกออก (Outbound)  ")
        self.notebook.add(self.tab_cycle, text="  3. ตรวจสอบ (Cycle Count)  ")
        self.notebook.add(self.tab_qr, text="  4. สร้าง & ปริ้น QR Code  ")

        self.build_inbound_tab()
        self.build_outbound_tab()
        self.build_cycle_tab()
        self.build_qr_tab()

    # ================= 1. TAB INBOUND =================
    def build_inbound_tab(self):
        card = tk.Frame(self.tab_inbound, bg="white", bd=1, relief="solid", padx=20, pady=16)
        card.pack(fill="both", expand=True, padx=10, pady=10)

        tk.Label(card, text="นำเข้าจัดเก็บสินค้าเข้าล็อก (Inbound)", font=("Tahoma", 13, "bold"), bg="white", fg="#22543d").pack(anchor="w", pady=(0, 10))

        loc_row = tk.Frame(card, bg="white")
        loc_row.pack(fill="x", pady=4)
        tk.Label(loc_row, text="ยิง/กรอก Location ล็อก:", font=("Tahoma", 10, "bold"), bg="white", width=22, anchor="w").pack(side="left")
        self.in_loc = tk.Entry(loc_row, font=("Tahoma", 11), width=28)
        self.in_loc.pack(side="left", padx=6)
        self.in_loc.bind("<Return>", lambda e: self.check_inbound_location())
        tk.Button(
            loc_row, text="ตรวจสอบล็อก", font=("Tahoma", 10, "bold"), bg="#3182ce", fg="white",
            command=self.check_inbound_location
        ).pack(side="left", padx=4)
        tk.Button(
            loc_row, text="ล้างล็อก", font=("Tahoma", 9), bg="#edf2f7", fg="#2d3748",
            command=self.reset_inbound_lock
        ).pack(side="left", padx=4)

        self.in_alert_lbl = tk.Label(
            card, text="", font=("Tahoma", 10, "bold"), bg="#fff5f5", fg="#c53030",
            wraplength=860, justify="left", pady=8, padx=10
        )
        self.in_alert_lbl.pack(fill="x", pady=(6, 10))
        self.in_alert_lbl.pack_forget()

        self.in_fields_frame = tk.Frame(card, bg="white")
        self.in_fields_frame.pack(fill="x")

        self.in_vars = {
            "item_code": tk.StringVar(),
            "product_name": tk.StringVar(),
            "lot_serial": tk.StringVar(),
            "quantity": tk.StringVar(value="1"),
            "unit": tk.StringVar(),
        }

        row_defs = [
            ("ยิง/กรอก Item Code:", "item_code", True),
            ("ชื่อสินค้า (Product Name):", "product_name", False),
            ("ล็อตสินค้า (Lot Serial):", "lot_serial", False),
            ("จำนวน (Quantity):", "quantity", False),
            ("หน่วย (Unit):", "unit", False),
        ]

        self.in_entries = {}
        for idx, (lbl_txt, key, is_code) in enumerate(row_defs):
            r = tk.Frame(self.in_fields_frame, bg="white")
            r.pack(fill="x", pady=5)
            tk.Label(r, text=lbl_txt, font=("Tahoma", 10, "bold"), bg="white", width=22, anchor="w").pack(side="left")
            ent = tk.Entry(r, textvariable=self.in_vars[key], font=("Tahoma", 10), width=36)
            ent.pack(side="left", padx=6)
            self.in_entries[key] = ent
            if is_code:
                ent.bind("<Return>", lambda e: self.autofill_inbound_product())
                tk.Button(
                    r, text="ดึงข้อมูลสินค้า", font=("Tahoma", 9, "bold"), bg="#d69e2e", fg="white",
                    command=self.autofill_inbound_product
                ).pack(side="left", padx=4)

        self.btn_in_submit = tk.Button(
            card, text="บันทึกข้อมูล Inbound ลง Google Sheets", font=("Tahoma", 11, "bold"),
            bg="#2f855a", fg="white", command=self.submit_inbound, height=2
        )
        self.btn_in_submit.pack(fill="x", pady=(16, 0))

    def set_inbound_locked(self, locked, message=""):
        state = "disabled" if locked else "normal"
        for ent in self.in_entries.values():
            ent.config(state=state)
        self.btn_in_submit.config(state=state)
        if locked:
            self.in_alert_lbl.config(text=message)
            self.in_alert_lbl.pack(fill="x", pady=(6, 10))
        else:
            self.in_alert_lbl.pack_forget()

    def reset_inbound_lock(self):
        self.in_loc.delete(0, tk.END)
        for k, v in self.in_vars.items():
            v.set("1" if k == "quantity" else "")
        self.set_inbound_locked(False)
        self.in_loc.focus_set()

    def check_inbound_location(self):
        loc = self.in_loc.get().strip()
        if not loc:
            messagebox.showwarning("แจ้งเตือน", "กรุณายิงหรือกรอก Location ล็อกก่อนค่ะ")
            return

        try:
            ws = self.get_worksheet("DATA_WH")
            rows = ws.get_all_values()
            stock_map = {}
            for r in rows[1:]:
                if len(r) >= 8 and r[1].strip().lower() == loc.lower():
                    st = r[2].strip().lower()
                    code = r[3].strip()
                    name = r[4].strip()
                    lot = r[5].strip()
                    qty = parse_qty(r[6])
                    unit = r[7].strip()
                    key = (code, name, lot, unit)
                    stock_map[key] = stock_map.get(key, 0.0) + (qty if st == "inbound" else -qty if st == "outbound" else 0.0)

            active = [(c, n, l, b, u) for (c, n, l, u), b in stock_map.items() if b > 0]
            if active:
                lines = [f"🚫 ล็อก '{loc}' มีสินค้าอยู่แล้ว ไม่สามารถกรอกหรือบันทึกซ้ำได้:"]
                for c, n, l, b, u in active:
                    lines.append(f"  • รหัส: {c} | ชื่อ: {n} | Lot: {l or '-'} | คงเหลือ: {fmt_qty(b)} {u}")
                self.set_inbound_locked(True, "\n".join(lines))
            else:
                self.set_inbound_locked(False)
                messagebox.showinfo("ล็อกว่างพร้อมใช้งาน", f"ล็อก '{loc}' ไม่มีสินค้าคงค้าง สามารถนำเข้าสินค้าได้ค่ะ")
                self.in_entries["item_code"].focus_set()
        except Exception as e:
            messagebox.showerror("เชื่อมต่อ Google Sheets ไม่สำเร็จ", str(e))

    def autofill_inbound_product(self):
        raw = self.in_vars["item_code"].get().strip()
        if not raw:
            return
        if "#" in raw:
            parts = raw.split("#")
            self.in_vars["item_code"].set(parts[0].strip())
            if len(parts) >= 2:
                self.in_vars["product_name"].set(parts[1].strip())
            if len(parts) >= 3 and parts[2].strip():
                self.in_vars["quantity"].set(parts[2].strip())
            if len(parts) >= 4:
                self.in_vars["unit"].set(parts[3].strip())
            if len(parts) >= 5:
                self.in_vars["lot_serial"].set(parts[4].strip())
            return

        rec = self.master_data.get(raw.lower())
        if rec:
            self.in_vars["product_name"].set(rec["description"])
            self.in_vars["unit"].set(rec["unit"])
            if rec["lot_serial"]:
                self.in_vars["lot_serial"].set(rec["lot_serial"])
            return

        try:
            ws = self.get_worksheet("Master_code")
            for r in ws.get_all_values()[1:]:
                if len(r) >= 3 and r[0].strip().lower() == raw.lower():
                    self.in_vars["product_name"].set(r[1].strip())
                    self.in_vars["unit"].set(r[2].strip())
                    if len(r) > 3 and r[3].strip():
                        self.in_vars["lot_serial"].set(r[3].strip())
                    break
        except Exception:
            pass

    def submit_inbound(self):
        loc = self.in_loc.get().strip()
        code = self.in_vars["item_code"].get().strip()
        name = self.in_vars["product_name"].get().strip()
        lot = self.in_vars["lot_serial"].get().strip()
        qty = self.in_vars["quantity"].get().strip() or "1"
        unit = self.in_vars["unit"].get().strip()

        if not loc or not code or not name:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอก Location, Item Code และชื่อสินค้าให้ครบค่ะ")
            return

        try:
            ws = self.get_worksheet("DATA_WH")
            now_str = (datetime.datetime.utcnow() + datetime.timedelta(hours=7)).strftime("%Y-%m-%d %H:%M:%S")
            ws.append_row([now_str, loc, "Inbound", code, name, lot, qty, unit])
            messagebox.showinfo("บันทึกสำเร็จ", f"บันทึก Inbound สินค้า {code} เข้าล็อก {loc} เรียบร้อยแล้วค่ะ")
            self.reset_inbound_lock()
        except Exception as e:
            messagebox.showerror("บันทึกไม่สำเร็จ", str(e))

    # ================= 2. TAB OUTBOUND =================
    def build_outbound_tab(self):
        card = tk.Frame(self.tab_outbound, bg="white", bd=1, relief="solid", padx=20, pady=16)
        card.pack(fill="both", expand=True, padx=10, pady=10)

        tk.Label(card, text="เบิกจ่ายสินค้าออกตาม Lot และ Location (Outbound)", font=("Tahoma", 13, "bold"), bg="white", fg="#742a2a").pack(anchor="w", pady=(0, 10))

        search_row = tk.Frame(card, bg="white")
        search_row.pack(fill="x", pady=4)
        tk.Label(search_row, text="ยิง/พิมพ์ Item Code:", font=("Tahoma", 10, "bold"), bg="white", width=20, anchor="w").pack(side="left")
        self.out_code_entry = tk.Entry(search_row, font=("Tahoma", 11), width=28)
        self.out_code_entry.pack(side="left", padx=6)
        self.out_code_entry.bind("<Return>", lambda e: self.search_outbound_lots())
        tk.Button(
            search_row, text="ค้นหา Lot", font=("Tahoma", 10, "bold"), bg="#d69e2e", fg="white",
            command=self.search_outbound_lots
        ).pack(side="left", padx=4)

        tk.Label(card, text="เลือกรายการ Lot / Location ที่มีอยู่ในระบบ:", font=("Tahoma", 10, "bold"), bg="white", fg="#2b6cb0").pack(anchor="w", pady=(10, 4))
        self.out_lot_listbox = tk.Listbox(card, font=("Tahoma", 10), height=5, exportselection=False)
        self.out_lot_listbox.pack(fill="x", pady=(0, 10))
        self.out_lot_listbox.bind("<<ListboxSelect>>", self.on_outbound_lot_selected)

        self.out_summary_lbl = tk.Label(
            card, text="ข้อมูลสรุปที่เลือก: ยังไม่ได้เลือกรายการ", font=("Tahoma", 10, "bold"),
            bg="#f0fff4", fg="#22543d", justify="left", anchor="w", padx=10, pady=8
        )
        self.out_summary_lbl.pack(fill="x", pady=(0, 10))

        self.out_vars = {
            "location": tk.StringVar(),
            "product_name": tk.StringVar(),
            "lot_serial": tk.StringVar(),
            "quantity": tk.StringVar(value="1"),
            "unit": tk.StringVar(),
        }

        form_grid = tk.Frame(card, bg="white")
        form_grid.pack(fill="x")
        for lbl_txt, key in [
            ("พิกัด Location:", "location"),
            ("ชื่อสินค้า:", "product_name"),
            ("Lot Serial:", "lot_serial"),
            ("จำนวนที่เบิก:", "quantity"),
            ("หน่วย:", "unit"),
        ]:
            r = tk.Frame(form_grid, bg="white")
            r.pack(fill="x", pady=3)
            tk.Label(r, text=lbl_txt, font=("Tahoma", 10, "bold"), bg="white", width=20, anchor="w").pack(side="left")
            tk.Entry(r, textvariable=self.out_vars[key], font=("Tahoma", 10), width=36).pack(side="left", padx=6)

        tk.Button(
            card, text="บันทึกการเบิกออก (Outbound) ลง Google Sheets", font=("Tahoma", 11, "bold"),
            bg="#2f855a", fg="white", command=self.submit_outbound, height=2
        ).pack(fill="x", pady=(14, 0))

    def search_outbound_lots(self):
        raw = self.out_code_entry.get().strip()
        if not raw:
            return
        code = raw.split("#")[0].strip()
        self.out_code_entry.delete(0, tk.END)
        self.out_code_entry.insert(0, code)

        self.out_lot_listbox.delete(0, tk.END)
        self.outbound_options.clear()

        try:
            ws = self.get_worksheet("DATA_WH")
            rows = ws.get_all_values()
            stock_map = {}
            for r in rows[1:]:
                if len(r) >= 8 and r[3].strip().lower() == code.lower():
                    loc = r[1].strip() or "ไม่ระบุพิกัด"
                    st = r[2].strip().lower()
                    name = r[4].strip()
                    lot = r[5].strip()
                    qty = parse_qty(r[6])
                    unit = r[7].strip() or "Piece"
                    key = (lot, loc, name, unit)
                    stock_map[key] = stock_map.get(key, 0.0) + (qty if st == "inbound" else -qty if st == "outbound" else 0.0)

            for (lot, loc, name, unit), bal in stock_map.items():
                if bal > 0:
                    opt = {"lot": lot, "location": loc, "product_name": name, "balance": fmt_qty(bal), "unit": unit}
                    self.outbound_options.append(opt)
                    self.out_lot_listbox.insert(
                        tk.END, f"Lot: {lot or '-'}  |  Location: {loc}  |  คงเหลือ: {opt['balance']} {unit}  |  {name}"
                    )

            if not self.outbound_options:
                messagebox.showwarning("ไม่พบยอดคงเหลือ", f"ไม่พบยอดคงเหลือของรหัส {code} ในระบบค่ะ")
            else:
                self.out_lot_listbox.selection_set(0)
                self.on_outbound_lot_selected(None)
        except Exception as e:
            messagebox.showerror("ค้นหาไม่สำเร็จ", str(e))

    def on_outbound_lot_selected(self, event):
        sel = self.out_lot_listbox.curselection()
        if not sel:
            return
        opt = self.outbound_options[sel[0]]
        code = self.out_code_entry.get().strip()
        self.out_vars["location"].set(opt["location"])
        self.out_vars["product_name"].set(opt["product_name"])
        self.out_vars["lot_serial"].set(opt["lot"])
        self.out_vars["unit"].set(opt["unit"])
        self.out_summary_lbl.config(
            text=(
                f"ข้อมูลสรุปที่เลือก -> รหัส: {code} | ชื่อ: {opt['product_name']} | "
                f"Lot: {opt['lot'] or '-'} | Location: {opt['location']} | คงเหลือ: {opt['balance']} {opt['unit']}"
            )
        )

    def submit_outbound(self):
        code = self.out_code_entry.get().strip()
        loc = self.out_vars["location"].get().strip()
        name = self.out_vars["product_name"].get().strip()
        lot = self.out_vars["lot_serial"].get().strip()
        qty = self.out_vars["quantity"].get().strip() or "1"
        unit = self.out_vars["unit"].get().strip()

        if not code or not loc or not name:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณาค้นหาและเลือกรายการ Lot/Location ก่อนบันทึกค่ะ")
            return

        try:
            ws = self.get_worksheet("DATA_WH")
            now_str = (datetime.datetime.utcnow() + datetime.timedelta(hours=7)).strftime("%Y-%m-%d %H:%M:%S")
            ws.append_row([now_str, loc, "Outbound", code, name, lot, qty, unit])
            messagebox.showinfo("บันทึกสำเร็จ", f"บันทึกเบิกออกสินค้า {code} (Lot: {lot or '-'}) เรียบร้อยแล้วค่ะ")
            self.search_outbound_lots()
        except Exception as e:
            messagebox.showerror("บันทึกไม่สำเร็จ", str(e))

    # ================= 3. TAB CYCLE COUNT =================
    def build_cycle_tab(self):
        top = tk.Frame(self.tab_cycle, bg="white", bd=1, relief="solid", padx=16, pady=12)
        top.pack(fill="x", padx=10, pady=(10, 6))

        tk.Label(top, text="ค้นหาประวัติและผลตรวจนับ (Cycle Count)", font=("Tahoma", 12, "bold"), bg="white", fg="#2b6cb0").pack(anchor="w", pady=(0, 8))

        f1 = tk.Frame(top, bg="white")
        f1.pack(fill="x", pady=3)
        tk.Label(f1, text="คำค้นหา:", font=("Tahoma", 10, "bold"), bg="white", width=14, anchor="w").pack(side="left")
        self.cc_query = tk.Entry(f1, font=("Tahoma", 10), width=36)
        self.cc_query.pack(side="left", padx=6)

        f2 = tk.Frame(top, bg="white")
        f2.pack(fill="x", pady=4)
        tk.Label(f2, text="ตั้งแต่วันที่:", font=("Tahoma", 10, "bold"), bg="white", width=14, anchor="w").pack(side="left")
        self.cc_start = tk.Entry(f2, font=("Tahoma", 10), width=16)
        self.cc_start.insert(0, "YYYY-MM-DD")
        self.cc_start.pack(side="left", padx=(6, 14))

        tk.Label(f2, text="ถึงวันที่:", font=("Tahoma", 10, "bold"), bg="white").pack(side="left")
        self.cc_end = tk.Entry(f2, font=("Tahoma", 10), width=16)
        self.cc_end.insert(0, "YYYY-MM-DD")
        self.cc_end.pack(side="left", padx=6)

        tk.Button(
            f2, text="ค้นหาประวัติ", font=("Tahoma", 10, "bold"), bg="#3182ce", fg="white",
            command=self.run_cycle_search
        ).pack(side="left", padx=10)

        cols = ("timestamp", "status", "location", "item_code", "product_name", "lot_serial", "quantity", "unit")
        headers = ("วัน-เวลา", "สถานะ", "Location", "รหัสสินค้า", "ชื่อสินค้า", "Lot Serial", "จำนวน", "หน่วย")

        tree_frame = tk.Frame(self.tab_cycle, bg="white")
        tree_frame.pack(fill="both", expand=True, padx=10, pady=4)

        self.cc_tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=11)
        for col, h in zip(cols, headers):
            self.cc_tree.heading(col, text=h)
            self.cc_tree.column(col, width=110, anchor="center" if col in ("status", "quantity", "unit") else "w")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.cc_tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.cc_tree.xview)
        self.cc_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.cc_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self.cc_summary_lbl = tk.Label(
            self.tab_cycle,
            text="ยอดสรุป -> Inbound รวม: 0  |  Outbound รวม: 0  |  ยอดคงเหลือสุทธิ: 0",
            font=("Tahoma", 10, "bold"),
            bg="#ebf8ff",
            fg="#2b6cb0",
            pady=10,
            padx=14,
            anchor="w",
        )
        self.cc_summary_lbl.pack(fill="x", padx=10, pady=(4, 10))

    def run_cycle_search(self):
        q = self.cc_query.get().strip().lower()
        s_date = self.cc_start.get().strip()
        e_date = self.cc_end.get().strip()
        if s_date == "YYYY-MM-DD":
            s_date = ""
        if e_date == "YYYY-MM-DD":
            e_date = ""

        for row_id in self.cc_tree.get_children():
            self.cc_tree.delete(row_id)

        total_in = 0.0
        total_out = 0.0

        try:
            ws = self.get_worksheet("DATA_WH")
            for r in ws.get_all_values()[1:]:
                if len(r) < 8:
                    continue
                ts, loc, st, code, name, lot, qty_str, unit = r[:8]
                r_date = ts.split(" ")[0] if " " in ts else ""

                if q and not any(q in field.lower() for field in (ts, loc, st, code, name, lot)):
                    continue
                if s_date and r_date and r_date < s_date:
                    continue
                if e_date and r_date and r_date > e_date:
                    continue

                self.cc_tree.insert("", tk.END, values=(ts, st, loc, code, name, lot, qty_str, unit))
                q_num = parse_qty(qty_str)
                if st.strip().lower() == "inbound":
                    total_in += q_num
                elif st.strip().lower() == "outbound":
                    total_out += q_num

            net = total_in - total_out
            self.cc_summary_lbl.config(
                text=f"ยอดสรุป -> Inbound รวม: +{fmt_qty(total_in)}  |  Outbound รวม: -{fmt_qty(total_out)}  |  ยอดคงเหลือสุทธิ: {fmt_qty(net)}"
            )
        except Exception as e:
            messagebox.showerror("ค้นหาไม่สำเร็จ", str(e))

    # ================= 4. TAB QR CODE & PRINT LABEL =================
    def build_qr_tab(self):
        top = tk.Frame(self.tab_qr, bg="white", bd=1, relief="solid", padx=16, pady=10)
        top.pack(fill="x", padx=10, pady=(10, 6))

        self.qr_file_lbl = tk.Label(top, text=f"ไฟล์ข้อมูล: {self.excel_path} ({self.excel_sheet_name})", font=("Tahoma", 9), bg="white", fg="#718096")
        self.qr_file_lbl.pack(side="left")
        tk.Button(top, text="เลือกไฟล์ Excel", font=("Tahoma", 9), bg="#edf2f7", command=self.choose_excel_file).pack(side="right")

        body = tk.Frame(self.tab_qr, bg="white", bd=1, relief="solid", padx=16, pady=12)
        body.pack(fill="both", expand=True, padx=10, pady=4)

        left = tk.Frame(body, bg="white")
        left.pack(side="left", fill="both", expand=True)

        code_row = tk.Frame(left, bg="white")
        code_row.pack(fill="x", pady=4)
        tk.Label(code_row, text="ยิง/พิมพ์ Item Code:", font=("Tahoma", 10, "bold"), bg="white", width=18, anchor="w").pack(side="left")
        self.qr_code_entry = tk.Entry(code_row, font=("Tahoma", 11), width=24)
        self.qr_code_entry.pack(side="left", padx=4)
        self.qr_code_entry.bind("<Return>", lambda e: self.lookup_qr_item())
        tk.Button(
            code_row, text="ดึงข้อมูล", font=("Tahoma", 9, "bold"), bg="#3182ce", fg="white",
            command=self.lookup_qr_item
        ).pack(side="left", padx=4)

        self.qr_vars = {
            "description": tk.StringVar(),
            "lot_serial": tk.StringVar(),
            "pallet_no": tk.StringVar(),
            "quantity": tk.StringVar(value="1"),
            "unit": tk.StringVar(),
            "reference": tk.StringVar(),
        }

        for lbl_txt, key in [
            ("Description:", "description"),
            ("Lot Serial:", "lot_serial"),
            ("Pallet No.:", "pallet_no"),
            ("Quantity:", "quantity"),
            ("Unit:", "unit"),
            ("Reference:", "reference"),
        ]:
            r = tk.Frame(left, bg="white")
            r.pack(fill="x", pady=4)
            tk.Label(r, text=lbl_txt, font=("Tahoma", 10, "bold"), bg="white", width=18, anchor="w").pack(side="left")
            ent = tk.Entry(r, textvariable=self.qr_vars[key], font=("Tahoma", 10), width=32)
            ent.pack(side="left", padx=4)
            ent.bind("<KeyRelease>", lambda e: self.refresh_label_preview())

        right = tk.Frame(body, bg="white", padx=10)
        right.pack(side="right", fill="y")
        tk.Label(right, text="ตัวอย่างป้ายสำหรับปริ้นแปะสินค้า", font=("Tahoma", 10, "bold"), bg="white").pack(pady=(0, 6))
        self.label_preview_lbl = tk.Label(right, bg="#edf2f7", width=38, height=20, relief="groove")
        self.label_preview_lbl.pack()

        btns = tk.Frame(self.tab_qr, bg="#f7fafc")
        btns.pack(fill="x", padx=10, pady=(4, 10))
        tk.Button(
            btns, text="อัปเดต & บันทึกรูปป้าย QR", font=("Tahoma", 10, "bold"),
            bg="#2f855a", fg="white", command=self.save_label_image, height=2
        ).pack(side="left", expand=True, fill="x", padx=(0, 6))
        tk.Button(
            btns, text="สั่งปริ้นป้าย QR Code", font=("Tahoma", 10, "bold"),
            bg="#2b6cb0", fg="white", command=self.print_label_image, height=2
        ).pack(side="left", expand=True, fill="x", padx=(6, 0))

    def choose_excel_file(self):
        p = filedialog.askopenfilename(title="เลือกไฟล์ RUN_Barcode_Template.xlsx", filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if p:
            self.excel_path = p
            self.load_excel_data()

    def load_excel_data(self):
        if not os.path.exists(self.excel_path):
            return
        try:
            wb = openpyxl.load_workbook(self.excel_path, data_only=True)
            ws = wb[self.excel_sheet_name] if self.excel_sheet_name in wb.sheetnames else wb.active
            self.excel_sheet_name = ws.title
            self.master_data.clear()
            for r in ws.iter_rows(min_row=2, values_only=True):
                if not r or len(r) < 2 or not r[1]:
                    continue
                code = str(r[1]).strip()
                self.master_data[code.lower()] = {
                    "item_code": code,
                    "quantity": "" if r[2] is None else str(r[2]).strip(),
                    "unit": "" if len(r) < 4 or r[3] is None else str(r[3]).strip(),
                    "lot_serial": "" if len(r) < 5 or r[4] is None else str(r[4]).strip(),
                    "pallet_no": "" if len(r) < 6 or r[5] is None else str(r[5]).strip(),
                    "description": "" if len(r) < 7 or r[6] is None else str(r[6]).strip(),
                }
            self.qr_file_lbl.config(text=f"ไฟล์ข้อมูล: {os.path.basename(self.excel_path)} ({self.excel_sheet_name}) — {len(self.master_data)} รายการ")
        except Exception as e:
            messagebox.showerror("อ่าน Excel ไม่สำเร็จ", str(e))

    def lookup_qr_item(self):
        code = self.qr_code_entry.get().strip()
        if not code:
            return
        rec = self.master_data.get(code.lower())
        if not rec:
            messagebox.showwarning("ไม่พบรหัส", f"ไม่พบ Item Code: {code} ใน {self.excel_sheet_name}")
            return
        self.qr_vars["description"].set(rec["description"])
        self.qr_vars["lot_serial"].set(rec["lot_serial"])
        self.qr_vars["pallet_no"].set(rec["pallet_no"])
        self.qr_vars["quantity"].set(rec["quantity"] or "1")
        self.qr_vars["unit"].set(rec["unit"])
        self.refresh_label_preview()

    def build_label_pil(self):
        code = self.qr_code_entry.get().strip()
        desc = self.qr_vars["description"].get().strip()
        lot = self.qr_vars["lot_serial"].get().strip()
        pallet = self.qr_vars["pallet_no"].get().strip()
        qty = self.qr_vars["quantity"].get().strip()
        unit = self.qr_vars["unit"].get().strip()

        ref_text = f"{code}#{desc}#{qty}#{unit}#{lot}#{pallet}" if code else ""
        self.qr_vars["reference"].set(ref_text)
        if not ref_text:
            return None

        qr = qrcode.QRCode(box_size=5, border=2)
        qr.add_data(ref_text)
        qr.make(fit=True)
        qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGB").resize((190, 190))

        label_w, label_h = 560, 270
        label = Image.new("RGB", (label_w, label_h), "white")
        draw = ImageDraw.Draw(label)
        draw.rectangle([4, 4, label_w - 4, label_h - 4], outline="black", width=3)

        label.paste(qr_img, (18, 22))
        font = ImageFont.load_default()

        lines = [
            f"Item Code : {code}",
            f"Desc      : {desc[:38]}",
            f"Lot Serial: {lot or '-'}",
            f"Pallet No.: {pallet or '-'}",
            f"Qty / Unit: {qty} {unit}".strip(),
            f"Ref       : {ref_text[:40]}",
        ]
        y_pos = 26
        for line in lines:
            draw.text((224, y_pos), line, fill="black", font=font)
            y_pos += 34

        return label

    def refresh_label_preview(self):
        img = self.build_label_pil()
        if not img:
            return
        self.label_pil_image = img
        preview = img.resize((340, 165))
        self.label_tk_image = ImageTk.PhotoImage(preview)
        self.label_preview_lbl.config(image=self.label_tk_image, width=340, height=165)

    def save_label_image(self):
        img = self.build_label_pil()
        if not img:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณายิงหรือกรอก Item Code ก่อนค่ะ")
            return
        self.refresh_label_preview()
        out_dir = "qr_labels"
        os.makedirs(out_dir, exist_ok=True)
        code = self.qr_code_entry.get().strip()
        lot = self.qr_vars["lot_serial"].get().strip() or "NOLOT"
        path = os.path.join(out_dir, f"Label_{code}_{lot}.png")
        img.save(path)
        messagebox.showinfo("บันทึกสำเร็จ", f"บันทึกป้าย QR Code พร้อมข้อมูลติดสินค้าแล้วที่:\n{path}")

    def print_label_image(self):
        img = self.build_label_pil()
        if not img:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณายิงหรือกรอก Item Code ก่อนปริ้นค่ะ")
            return
        self.refresh_label_preview()
        tmp_file = os.path.join(tempfile.gettempdir(), "chacha_qr_label_print.png")
        img.save(tmp_file)
        try:
            os.startfile(tmp_file, "print")
        except Exception:
            subprocess.Popen(["cmd", "/c", "start", "", tmp_file], shell=True)


if __name__ == "__main__":
    root = tk.Tk()
    app = ChachaWMSDesktopApp(root)
    root.mainloop()