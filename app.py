import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- CẤU HÌNH MÀU SẮC ---
BG       = "#F1F5F9"
BG_CARD  = "#FFFFFF"
BG_INPUT = "#FFFFFF"
BG_HDR   = "#1B2A4A"
BLUE     = "#2563EB"
BLUE_HV  = "#1D4ED8"
TXT_H    = "#1B2A4A"
TXT_B    = "#1E3A5F"
TXT_M    = "#4A6FA5"
BORDER   = "#CBD5E1"
FT_TITLE = ("Segoe UI", 16, "bold")
FT_LBL   = ("Segoe UI", 10, "bold")
FT_IN    = ("Segoe UI", 10)
FT_BTN   = ("Segoe UI", 11, "bold")


class DataExportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Hệ Thống Kết Xuất Dữ Liệu")
        self.root.geometry("500x550")
        self.root.configure(bg=BG)
        self.file_path = ""

        # Header
        header = tk.Frame(self.root, bg=BG_HDR, height=70)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="✨ KẾT XUẤT DỮ LIỆU", font=FT_TITLE, fg="white", bg=BG_HDR).place(relx=0.5, rely=0.5, anchor="center")

        # Container (Card)
        card = tk.Frame(self.root, bg=BG_CARD, bd=0, highlightthickness=1, highlightbackground=BORDER)
        card.pack(padx=30, pady=30, fill="both", expand=True)

        # 1. Chọn mẫu
        tk.Label(card, text="Chọn mẫu báo cáo:", font=FT_LBL, fg=TXT_B, bg=BG_CARD).pack(anchor="w", padx=25, pady=(25, 5))
        self.cbo_mau = ttk.Combobox(card, values=self.mau_list(), font=FT_IN, state="readonly")
        self.cbo_mau.pack(fill="x", padx=25, pady=5)
        self.cbo_mau.set("Mau 01")

        # 2. Nhập mã đơn vị
        tk.Label(card, text="Mã đơn vị (Text):", font=FT_LBL, fg=TXT_B, bg=BG_CARD).pack(anchor="w", padx=25, pady=(15, 5))
        self.ent_ma = tk.Entry(card, font=FT_IN, bd=1, relief="solid", highlightthickness=0)
        self.ent_ma.pack(fill="x", padx=25, pady=5, ipady=5)
        self.ent_ma.insert(0, "84003")

        # 3. Chọn file dữ liệu
        tk.Label(card, text="File dữ liệu đầu vào (.xlsx):", font=FT_LBL, fg=TXT_B, bg=BG_CARD).pack(anchor="w", padx=25, pady=(15, 5))
        file_frame = tk.Frame(card, bg=BG_CARD)
        file_frame.pack(fill="x", padx=25)
        self.lbl_file = tk.Label(file_frame, text="Chưa chọn file...", font=FT_IN, fg=TXT_M, bg="#F8FAFC", anchor="w", bd=1, relief="solid")
        self.lbl_file.pack(side="left", fill="x", expand=True, ipady=4)
        btn_browse = tk.Button(file_frame, text="Chọn File", bg=BLUE, fg="white", font=("Segoe UI", 9, "bold"),
                               command=self.browse_file, borderwidth=0, cursor="hand2")
        btn_browse.pack(side="right", padx=(5, 0))

        # 4. Nút kết xuất
        self.btn_export = tk.Button(card, text="🚀 XUẤT DỮ LIỆU EXCEL", bg="#16A34A", fg="white", font=FT_BTN,
                                    command=self.process_data, borderwidth=0, cursor="hand2", pady=10)
        self.btn_export.pack(fill="x", padx=25, pady=35)

    def mau_list(self):
        lst = [f"Mau {i:02d}" for i in range(1, 80)]
        lst.append("Mau 79A")
        return lst

    def browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.file_path = path
            filename = os.path.basename(path)
            self.lbl_file.config(text=filename if len(filename) < 30 else filename[:27] + "...")

    def process_data(self):
        ma_dv = self.ent_ma.get().strip()
        mau_sel = self.cbo_mau.get()

        if not self.file_path:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn file dữ liệu!")
            return
        if not ma_dv:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập mã đơn vị!")
            return

        try:
            df = pd.read_excel(self.file_path)
            df.insert(0, "MA_DON_VI", ma_dv)
            df.insert(1, "MAU_BAO_CAO", mau_sel)

            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=f"Ket_Xuat_{mau_sel}_{ma_dv}.xlsx"
            )
            if save_path:
                df.to_excel(save_path, index=False)
                self._format_excel(save_path)
                messagebox.showinfo("Thành công", f"Đã xuất dữ liệu thành công ra file:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xử lý dữ liệu: {e}")

    def _format_excel(self, path):
        wb = load_workbook(path)
        ws = wb.active

        # Màu header
        hdr_fill = PatternFill("solid", fgColor="1B2A4A")
        hdr_font = Font(name="Segoe UI", bold=True, color="FFFFFF", size=11)

        # Màu xen kẽ cho data rows
        row_fill_even = PatternFill("solid", fgColor="EFF6FF")
        row_fill_odd  = PatternFill("solid", fgColor="FFFFFF")
        data_font     = Font(name="Segoe UI", size=10)

        # Border mỏng
        thin = Side(style="thin", color="CBD5E1")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        # Format header (row 1)
        for cell in ws[1]:
            cell.font      = hdr_font
            cell.fill      = hdr_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border    = border

        # Format data rows
        for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            fill = row_fill_even if row_idx % 2 == 0 else row_fill_odd
            for cell in row:
                cell.font      = data_font
                cell.fill      = fill
                cell.alignment = Alignment(vertical="center")
                cell.border    = border

        # Tự động điều chỉnh độ rộng cột
        for col_idx, col in enumerate(ws.columns, start=1):
            max_len = 0
            for cell in col:
                try:
                    val_len = len(str(cell.value)) if cell.value is not None else 0
                    max_len = max(max_len, val_len)
                except Exception:
                    pass
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 40)

        # Freeze header row
        ws.freeze_panes = "A2"

        # Chiều cao header
        ws.row_dimensions[1].height = 30

        wb.save(path)


if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style()
    style.theme_use('clam')
    style.configure("TCombobox", fieldbackground="white", background=BORDER)
    app = DataExportApp(root)
    root.mainloop()
