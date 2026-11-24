import os
import tkinter as tk
from tkinter import messagebox, filedialog
import openpyxl
import qrcode
from PIL import ImageTk


class BarcodeQRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("โปรแกรมสแกนบาร์โค้ดและสร้าง QR Code - ชาช่าไทย")
        self.root.geometry("620x540")
        self.root.configure(bg="#f7fafc")

        self.excel_path = "RUN_Barcode_Template.xlsx"
        self.sheet_name = "Data"
        self.master_data = {}
        self.qr_pil_image = None
        self.qr_tk_image = None

        self.build_ui()
        self.load_excel_data()

    def build_ui(self):
        title_lbl = tk.Label(
            self.root,
            text="ระบบยิงบาร์โค้ดและสร้าง QR Code สินค้า",
            font=("Tahoma", 15, "bold"),
            bg="#f7fafc",
            fg="#1a202c",
        )
        title_lbl.pack(pady=(14, 2))

        self.file_lbl = tk.Label(
            self.root,
            text=f"ฐานข้อมูล: {self.excel_path} ({self.sheet_name})",
            font=("Tahoma", 9),
            bg="#f7fafc",
            fg="#718096",
        )
        self.file_lbl.pack(pady=(0, 8))

        top_frame = tk.Frame(self.root, bg="#f7fafc")
        top_frame.pack(fill="x", padx=20)

        tk.Label(top_frame, text="ยิงหรือพิมพ์ Item code:", font=("Tahoma", 10, "bold"), bg="#f7fafc").grid(
            row=0, column=0, sticky="w", pady=4
        )
        self.entry_code = tk.Entry(top_frame, font=("Tahoma", 11), width=26)
        self.entry_code.grid(row=0, column=1, padx=6, pady=4, sticky="w")
        self.entry_code.bind("<Return>", lambda event: self.lookup_item())
        self.entry_code.focus_set()

        btn_search = tk.Button(
            top_frame,
            text="ค้นหา",
            font=("Tahoma", 10, "bold"),
            bg="#3182ce",
            fg="white",
            command=self.lookup_item,
            width=10,
        )
        btn_search.grid(row=0, column=2, padx=4, pady=4)

        btn_reload = tk.Button(
            top_frame,
            text="เลือกไฟล์ Excel",
            font=("Tahoma", 9),
            bg="#edf2f7",
            fg="#2d3748",
            command=self.choose_excel_file,
        )
        btn_reload.grid(row=0, column=3, padx=4, pady=4)

        body_frame = tk.Frame(self.root, bg="#ffffff", bd=1, relief="solid", padx=16, pady=12)
        body_frame.pack(fill="both", expand=True, padx=20, pady=10)

        info_frame = tk.Frame(body_frame, bg="#ffffff")
        info_frame.pack(side="left", fill="both", expand=True)

        self.vars = {
            "description": tk.StringVar(),
            "lot_serial": tk.StringVar(),
            "pallet_no": tk.StringVar(),
            "quantity": tk.StringVar(value="1"),
            "unit": tk.StringVar(),
            "reference": tk.StringVar(),
        }

        fields = [
            ("Description:", "description", True),
            ("Lot Serial:", "lot_serial", False),
            ("Pallet No.:", "pallet_no", False),
            ("Quantity:", "quantity", False),
            ("Unit:", "unit", False),
            ("Reference:", "reference", True),
        ]

        for idx, (label_text, key, readonly) in enumerate(fields):
            tk.Label(info_frame, text=label_text, font=("Tahoma", 10, "bold"), bg="#ffffff").grid(
                row=idx, column=0, sticky="w", pady=5
            )
            state = "readonly" if readonly else "normal"
            ent = tk.Entry(info_frame, textvariable=self.vars[key], font=("Tahoma", 10), width=26, state=state)
            ent.grid(row=idx, column=1, padx=8, pady=5, sticky="w")
            if not readonly:
                ent.bind("<KeyRelease>", lambda e: self.update_reference())

        qr_frame = tk.Frame(body_frame, bg="#ffffff", padx=10)
        qr_frame.pack(side="right", fill="y")

        tk.Label(qr_frame, text="ตัวอย่าง QR Code", font=("Tahoma", 10, "bold"), bg="#ffffff").pack(pady=(0, 6))
        self.qr_label = tk.Label(qr_frame, bg="#edf2f7", width=22, height=11, relief="groove")
        self.qr_label.pack()

        btn_frame = tk.Frame(self.root, bg="#f7fafc")
        btn_frame.pack(fill="x", padx=20, pady=(0, 14))

        btn_gen = tk.Button(
            btn_frame,
            text="สร้าง & บันทึกรูป QR Code",
            font=("Tahoma", 10, "bold"),
            bg="#2f855a",
            fg="white",
            command=self.generate_and_save_qr,
            height=2,
        )
        btn_gen.pack(side="left", expand=True, fill="x", padx=(0, 6))

        btn_clear = tk.Button(
            btn_frame,
            text="ล้างข้อมูล (ยิงตัวถัดไป)",
            font=("Tahoma", 10, "bold"),
            bg="#d69e2e",
            fg="white",
            command=self.clear_form,
            height=2,
        )
        btn_clear.pack(side="left", expand=True, fill="x", padx=(6, 0))

    def choose_excel_file(self):
        path = filedialog.askopenfilename(
            title="เลือกไฟล์ RUN_Barcode_Template.xlsx",
            filetypes=[("Excel files", "*.xlsx *.xlsm")],
        )
        if path:
            self.excel_path = path
            self.load_excel_data()

    def load_excel_data(self):
        if not os.path.exists(self.excel_path):
            messagebox.showwarning("ไม่พบไฟล์", f"ไม่พบไฟล์ {self.excel_path} ในโฟลเดอร์นี้")
            return

        try:
            wb = openpyxl.load_workbook(self.excel_path, data_only=True)
            ws = wb[self.sheet_name] if self.sheet_name in wb.sheetnames else wb.active
            self.sheet_name = ws.title

            self.master_data.clear()
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or len(row) < 2 or not row[1]:
                    continue
                code = str(row[1]).strip()
                self.master_data[code.lower()] = {
                    "item_code": code,
                    "quantity": "" if row[2] is None else str(row[2]).strip(),
                    "unit": "" if len(row) < 4 or row[3] is None else str(row[3]).strip(),
                    "lot_serial": "" if len(row) < 5 or row[4] is None else str(row[4]).strip(),
                    "pallet_no": "" if len(row) < 6 or row[5] is None else str(row[5]).strip(),
                    "description": "" if len(row) < 7 or row[6] is None else str(row[6]).strip(),
                }

            self.file_lbl.config(
                text=f"ฐานข้อมูล: {os.path.basename(self.excel_path)} ({self.sheet_name}) — โหลดแล้ว {len(self.master_data)} รายการ"
            )
        except Exception as e:
            messagebox.showerror("อ่านไฟล์ไม่สำเร็จ", f"เกิดข้อผิดพลาดในการอ่าน Excel:\n{e}")

    def lookup_item(self):
        code = self.entry_code.get().strip()
        if not code:
            return

        rec = self.master_data.get(code.lower())
        if not rec:
            messagebox.showwarning("ไม่พบรหัสสินค้า", f"ไม่พบ Item code: {code} ในชีต {self.sheet_name}")
            return

        self.vars["description"].set(rec["description"])
        self.vars["lot_serial"].set(rec["lot_serial"])
        self.vars["pallet_no"].set(rec["pallet_no"])
        self.vars["quantity"].set(rec["quantity"] or "1")
        self.vars["unit"].set(rec["unit"])
        self.update_reference()
        self.render_qr_preview()

    def update_reference(self):
        code = self.entry_code.get().strip()
        desc = self.vars["description"].get().strip()
        qty = self.vars["quantity"].get().strip()
        unit = self.vars["unit"].get().strip()
        lot = self.vars["lot_serial"].get().strip()
        pallet = self.vars["pallet_no"].get().strip()

        if not code:
            self.vars["reference"].set("")
            return

        ref_text = f"{code}#{desc}#{qty}#{unit}#{lot}#{pallet}"
        self.vars["reference"].set(ref_text)

    def render_qr_preview(self):
        ref_text = self.vars["reference"].get().strip()
        if not ref_text:
            return

        qr = qrcode.QRCode(box_size=5, border=2)
        qr.add_data(ref_text)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
        img = img.resize((160, 160))

        self.qr_pil_image = img
        self.qr_tk_image = ImageTk.PhotoImage(img)
        self.qr_label.config(image=self.qr_tk_image, width=160, height=160)

    def generate_and_save_qr(self):
        self.update_reference()
        ref_text = self.vars["reference"].get().strip()
        code = self.entry_code.get().strip()
        if not ref_text or not code:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณายิงหรือกรอก Item code ก่อนสร้าง QR Code")
            return

        self.render_qr_preview()
        out_dir = "qr_output"
        os.makedirs(out_dir, exist_ok=True)
        lot_part = self.vars["lot_serial"].get().strip() or "NOLOT"
        filename = f"{code}_{lot_part}.png"
        filepath = os.path.join(out_dir, filename)
        self.qr_pil_image.save(filepath)
        messagebox.showinfo("บันทึกสำเร็จ", f"สร้างและบันทึก QR Code เรียบร้อยแล้ว:\n{filepath}")

    def clear_form(self):
        self.entry_code.delete(0, tk.END)
        for key, var in self.vars.items():
            var.set("1" if key == "quantity" else "")
        self.qr_label.config(image="", width=22, height=11)
        self.entry_code.focus_set()


if __name__ == "__main__":
    root = tk.Tk()
    app = BarcodeQRApp(root)
    root.mainloop()