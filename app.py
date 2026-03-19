import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- CẤU HÌNH GIAO DIỆN ---
BG       = "#F1F5F9"
BG_CARD  = "#FFFFFF"
BG_HDR   = "#1B2A4A"  
BLUE_EXCEL = "0070C0" # Màu xanh dương của header Excel
FT_TITLE = ("Segoe UI", 16, "bold")
FT_LBL   = ("Segoe UI", 10, "bold")
FT_IN    = ("Segoe UI", 10)
FT_BTN   = ("Segoe UI", 11, "bold")

class DataExportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Hệ Thống Xuất Dữ Liệu")
        self.root.geometry("500x620")
        self.root.configure(bg=BG)
        self.file_path = ""

        # ─── HEADER: XUẤT DỮ LIỆU ───
        header = tk.Frame(self.root, bg=BG_HDR, height=70)
        header.pack(fill="x")
        tk.Label(header, text="XUẤT DỮ LIỆU", font=FT_TITLE, fg="white", bg=BG_HDR).place(relx=0.5, rely=0.5, anchor="center")

        card = tk.Frame(self.root, bg=BG_CARD, bd=0, highlightthickness=1, highlightbackground="#CBD5E1")
        card.pack(padx=30, pady=20, fill="both", expand=True)

        # 1. Chọn mẫu báo cáo
        tk.Label(card, text="Chọn mẫu báo cáo:", font=FT_LBL, bg=BG_CARD).pack(anchor="w", padx=25, pady=(15, 5))
        self.cbo_mau = ttk.Combobox(card, values=[f"Mau {i:02d}" for i in range(1, 80)] + ["Mau 79A"], font=FT_IN, state="readonly")
        self.cbo_mau.pack(fill="x", padx=25, pady=5)
        self.cbo_mau.set("Mau 01")

        # 2. MÃ ĐƠN VỊ (ĐÃ THAY 84003 THÀNH CHỮ MACSKCB)
        tk.Label(card, text="Mã đơn vị:", font=FT_LBL, bg=BG_CARD).pack(anchor="w", padx=25, pady=(10, 5))
        self.ent_ma_dv = tk.Entry(card, font=FT_IN, bd=1, relief="solid")
        self.ent_ma_dv.pack(fill="x", padx=25, pady=5, ipady=5)
        self.ent_ma_dv.insert(0, "MACSKCB") # <-- ĐÃ SỬA THEO YÊU CẦU CỦA BẠN

        # 3. MÃ CƠ SỞ KHÁM CHỮA BỆNH
        tk.Label(card, text="Mã cơ sở khám chữa bệnh:", font=FT_LBL, bg=BG_CARD).pack(anchor="w", padx=25, pady=(10, 5))
        self.ent_ma_cs = tk.Entry(card, font=FT_IN, bd=1, relief="solid")
        self.ent_ma_cs.pack(fill="x", padx=25, pady=5, ipady=5)
        self.ent_ma_cs.insert(0, "84003")

        # 4. Chọn file dữ liệu đầu vào
        tk.Label(card, text="File dữ liệu đầu vào (.xlsx):", font=FT_LBL, bg=BG_CARD).pack(anchor="w", padx=25, pady=(10, 5))
        file_frame = tk.Frame(card, bg=BG_CARD)
        file_frame.pack(fill="x", padx=25)
        self.lbl_file = tk.Label(file_frame, text="Chưa chọn file...", font=FT_IN, fg="#4A6FA5", bg="#F8FAFC", anchor="w", bd=1, relief="solid")
        self.lbl_file.pack(side="left", fill="x", expand=True, ipady=4)
        tk.Button(file_frame, text="Chọn File", bg="#2563EB", fg="white", font=("Segoe UI", 9, "bold"), command=self.browse_file, borderwidth=0, cursor="hand2").pack(side="right", padx=(5, 0))

        # 5. Nút bấm kết xuất
        tk.Button(card, text="🚀 XUẤT DỮ LIỆU EXCEL", bg="#16A34A", fg="white", font=FT_BTN, command=self.process_data, borderwidth=0, cursor="hand2", pady=10).pack(fill="x", padx=25, pady=25)

    def browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.file_path = path
            self.lbl_file.config(text=os.path.basename(path))

    def process_data(self):
        ma_dv = self.ent_ma_dv.get().strip() # Giá trị MACSKCB
        ma_cs = self.ent_ma_cs.get().strip() # Giá trị MA_CSKCB
        tu_ngay_val = "20260101"
        mau_sel = self.cbo_mau.get()

        if not self.file_path or not ma_dv or not ma_cs:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập đủ các mã và chọn file!")
            return

        try:
            df = pd.read_excel(self.file_path)

            # Lọc các cột cần thiết theo mẫu Ảnh 2
            cols_to_keep = ['STT', 'MA_KHOA', 'TEN_KHOA', 'BAN_KHAM', 'GIUONG_PD', 'GIUONG_TK', 'GIUONG_HSTC', 'GIUONG_HSCC']
            df_final = df[[c for c in cols_to_keep if c in df.columns]].copy()

            # Thêm các cột định dạng yêu cầu
            df_final['TU_NGAY'] = tu_ngay_val
            df_final['DEN_NGAY'] = df['DEN_NGAY'] if 'DEN_NGAY' in df.columns else ""
            df_final['MA_CSKCB'] = ma_cs
            df_final['MACSKCB'] = ma_dv

            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=f"Xuat_{mau_sel}_{ma_cs}.xlsx"
            )
            
            if save_path:
                sheet_name = mau_sel.replace(" ", "_").upper()
                df_final.to_excel(save_path, index=False, sheet_name=sheet_name)
                self._format_excel(save_path)
                messagebox.showinfo("Thành công", f"Đã xuất dữ liệu thành công!")

        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi xử lý: {e}")

    def _format_excel(self, path):
        wb = load_workbook(path)
        ws = wb.active
        header_fill = PatternFill("solid", fgColor=BLUE_EXCEL)
        header_font = Font(name="Calibri", bold=True, color="FFFFFF")
        border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
        
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = border
                cell.font = Font(name="Calibri")

        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except: pass
            ws.column_dimensions[column].width = max_length + 2

        wb.save(path)

if __name__ == "__main__":
    root = tk.Tk()
    app = DataExportApp(root)
    root.mainloop()