import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ====================== CẤU HÌNH MẪU ======================
MAU_MASTER_CONFIG = {
    "Mau 01": {
        "theme": "#2563EB", "sheet": "MAU_01",
        "out_cols": ["STT", "MA_KHOA", "TEN_KHOA", "BAN_KHAM", "GIUONG_PD", "GIUONG_TK", "GIUONG_HSTC", "GIUONG_HSCC", "TU_NGAY", "DEN_NGAY", "MA_CSKCB"],
        "mapping": {}, "fixed_val": {"TU_NGAY": "20260101"},
        "input_fields": [
            {"label": "Mã cơ sở (MA_CSKCB):", "key": "MA_CSKCB", "default": "84003"}
        ]
    },
    "Mau 02": {
        "theme": "#0EA5E9", "sheet": "MAU_02",
        "out_cols": ["STT", "MA_KHOA", "TEN_KHOA", "HO_TEN", "GIOI_TINH", "SO_DINH_DANH", "CHUCDANH_NN", "VI_TRI", "MACCHN", "NGAYCAP_CCHN", "NOICAP_CCHN", "PHAMVI_CM", "PHAMVI_CMBS", "DVKT_KHAC", "VB_PHANCONG", "THOIGIAN_DK", "THOIGIAN_NGAY", "THOIGIAN_TUAN", "CSKCB_KHAC", "CSKCB_CGKT", "QD_CGKT", "TU_NGAY", "DEN_NGAY", "MA_CSKCB"],
        "mapping": {"SO_CCCD": "SO_DINH_DANH"},
        "input_fields": [
            {"label": "Mã cơ sở (MA_CSKCB):", "key": "MA_CSKCB", "default": "84003"}
        ]
    },
    "Mau 03": {
        "theme": "#10B981", "sheet": "MAU_03",
        "out_cols": [
            "STT", "MA_THUOC", "TEN_HOAT_CHAT", "TEN_THUOC", "DON_VI_TINH", "HAM_LUONG", "DUONG_DUNG",
            "DANG_BAO_CHE", "SO_DANG_KY", "SO_LUONG", "DON_GIA", "DON_GIA_BH", "QUY_CACH",
            "NHA_SX", "NUOC_SX", "NHA_THAU", "TT_THAU", "TU_NGAY", "DEN_NGAY", "MA_CSKCB",
            "LOAI_THUOC", "LOAI_THAU", "HT_THAU", "MA_CSKCB_THUOC"
        ],
        "input_fields": [
            {"label": "Mã cơ sở (MA_CSKCB):", "key": "MA_CSKCB", "default": "84003"},
            {"label": "Mã cơ sở thuốc (MA_CSKCB_THUOC):", "key": "MA_CSKCB_THUOC", "default": "84003"}
        ]
    },
    "Mau 04": {
        "theme": "#F59E0B", "sheet": "MAU_04",
        "out_cols": [
            "STT", "MA_VAT_TU", "NHOM_VAT_TU", "TEN_VAT_TU", "MA_HIEU", "SO_LUU_HANH", "TINHNANG_KT",
            "QUY_CACH", "HANG_SX", "NUOC_SX", "DON_VI_TINH", "DON_GIA", "DON_GIA_BH", "TYLE_TT_BH",
            "SO_LUONG", "DINH_MUC", "NHA_THAU", "TT_THAU", "TU_NGAY_HD", "DEN_NGAY_HD", "MA_CSKCB",
            "LOAI_THAU", "HT_THAU", "MA_CSKCB_TBYT", "TU_NGAY", "DEN_NGAY"
        ],
        "input_fields": [
            {"label": "Mã cơ sở (MA_CSKCB):", "key": "MA_CSKCB", "default": "84003"},
            {"label": "Mã cơ sở TBYT (MA_CSKCB_TBYT):", "key": "MA_CSKCB_TBYT", "default": "84003"}
        ]
    },
    "Mau 05": {
        "theme": "#8B5CF6", "sheet": "MAU_05",
        "out_cols": ["STT", "MA_DICH_VU", "TEN_DICH_VU", "TEN_DVKT_GIA", "DON_GIA", "QUY_TRINH", "SO_LUONG_CGKT",
                     "CSKCB_CGKT", "CSKCB_CLS", "QD_DVKT", "QD_PD_GIA", "GHI_CHU", "TU_NGAY", "DEN_NGAY", "MA_CSKCB", "GIA_THANH_TOAN"],
        "input_fields": [
            {"label": "Mã cơ sở (MA_CSKCB):", "key": "MA_CSKCB", "default": "84003"}
        ]
    },
    "Mau 06": {
        "theme": "#475569", "sheet": "MAU_06",
        "out_cols": ["STT", "TEN_TB", "KY_HIEU", "CONGTY_SX", "NUOC_SX", "NAM_SX", "NAM_SD", "MA_MAY",
                     "SO_LUU_HANH", "HD_TU", "HD_DEN", "TU_NGAY", "DEN_NGAY", "MA_CSKCB"],
        "input_fields": [
            {"label": "Mã cơ sở (MA_CSKCB):", "key": "MA_CSKCB", "default": "84003"}
        ]
    }
}

class DataExportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Hệ Thống Xuất Dữ Liệu BHYT")
        self.root.geometry("740x820")
        self.root.configure(bg="#F0F9FF")
        self.file_path = ""
        self.current_entries = {}

        header = tk.Frame(root, bg="#0EA5E9", height=130)
        header.pack(fill="x")
        tk.Label(header, text="🌟 XUẤT DỮ LIỆU BHYT", 
                 font=("Segoe UI", 28, "bold"), 
                 fg="white", bg="#0EA5E9").place(relx=0.5, rely=0.5, anchor="center")

        main = tk.Frame(root, bg="white", padx=50, pady=40, relief="flat")
        main.pack(fill="both", expand=True, padx=35, pady=25)

        tk.Label(main, text="1. Chọn mẫu báo cáo", font=("Segoe UI", 12, "bold"), bg="white", fg="#1E40AF").pack(anchor="w")
        self.cbo_mau = ttk.Combobox(main, values=list(MAU_MASTER_CONFIG.keys()), 
                                    font=("Segoe UI", 11), state="readonly", height=15)
        self.cbo_mau.pack(fill="x", pady=(8, 25))
        self.cbo_mau.set("Mau 01")
        self.cbo_mau.bind("<<ComboboxSelected>>", self.update_form)

        self.fields_frame = tk.Frame(main, bg="white")
        self.fields_frame.pack(fill="x", pady=10)

        tk.Label(main, text="3. Chọn file dữ liệu đầu vào", font=("Segoe UI", 12, "bold"), bg="white", fg="#1E40AF").pack(anchor="w", pady=(30, 8))
        file_frame = tk.Frame(main, bg="white")
        file_frame.pack(fill="x")
        self.lbl_file = tk.Label(file_frame, text="Chưa chọn file...", bg="#E0F2FE", relief="solid", bd=1, 
                                 anchor="w", font=("Segoe UI", 10), height=2, padx=15)
        self.lbl_file.pack(side="left", fill="x", expand=True)
        tk.Button(file_frame, text="📁 Chọn File", bg="#0EA5E9", fg="white", font=("Segoe UI", 10, "bold"), 
                  width=15, command=self.browse_file).pack(side="right", padx=8)

        self.btn_export = tk.Button(main, text="🚀 BẮT ĐẦU XUẤT EXCEL", 
                                    font=("Segoe UI", 14, "bold"), 
                                    bg="#10B981", fg="white", height=2, 
                                    activebackground="#059669", command=self.process_data)
        self.btn_export.pack(fill="x", pady=(50, 10))

        self.update_form()

    def update_form(self, event=None):
        for widget in self.fields_frame.winfo_children():
            widget.destroy()
        self.current_entries.clear()

        mau = self.cbo_mau.get()
        cfg = MAU_MASTER_CONFIG[mau]

        for field in cfg.get("input_fields", []):
            tk.Label(self.fields_frame, text=field["label"], 
                     font=("Segoe UI", 10, "bold"), bg="white", fg="#1E3A8A").pack(anchor="w", pady=(12, 4))
            entry = tk.Entry(self.fields_frame, font=("Segoe UI", 11), relief="solid", bd=1)
            entry.pack(fill="x", ipady=9, pady=3)
            entry.insert(0, field["default"])
            self.current_entries[field["key"]] = entry

        self.btn_export.config(bg=cfg["theme"])

    def browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.file_path = path
            self.lbl_file.config(text=os.path.basename(path), fg="#0369A1")

    def process_data(self):
        if not self.file_path:
            messagebox.showwarning("Thiếu dữ liệu", "Vui lòng chọn file dữ liệu đầu vào!")
            return

        mau_sel = self.cbo_mau.get()
        cfg = MAU_MASTER_CONFIG[mau_sel]

        ma_cs = self.current_entries.get("MA_CSKCB", tk.StringVar(value="")).get().strip() or "84003"

        try:
            df_in = pd.read_excel(self.file_path, dtype=str)
            df_out = pd.DataFrame(columns=cfg["out_cols"])

            for col_out in cfg["out_cols"]:
                col_in = next((k for k, v in cfg.get("mapping", {}).items() if v == col_out), col_out)
                df_out[col_out] = df_in.get(col_in, "")

            if "MA_CSKCB" in df_out.columns:
                df_out["MA_CSKCB"] = ma_cs

            if "MA_CSKCB_THUOC" in df_out.columns and "MA_CSKCB_THUOC" in self.current_entries:
                df_out["MA_CSKCB_THUOC"] = self.current_entries["MA_CSKCB_THUOC"].get().strip() or ma_cs

            if "MA_CSKCB_TBYT" in df_out.columns and "MA_CSKCB_TBYT" in self.current_entries:
                df_out["MA_CSKCB_TBYT"] = self.current_entries["MA_CSKCB_TBYT"].get().strip() or ma_cs

            if "fixed_val" in cfg:
                for col, val in cfg["fixed_val"].items():
                    if col in df_out.columns:
                        df_out[col] = val

            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=f"Ket_Xuat_{mau_sel.replace(' ', '_')}_{ma_cs}.xlsx"
            )

            if save_path:
                df_out.to_excel(save_path, index=False, sheet_name=cfg["sheet"])
                self._format_excel(save_path, cfg["theme"].replace("#", ""))
                messagebox.showinfo("✅ Thành công", 
                    f"Đã xuất dữ liệu mẫu {mau_sel} thành công!\n\n"
                    f"Tên file: {os.path.basename(save_path)}")
        except Exception as e:
            messagebox.showerror("❌ Lỗi", f"Không thể xử lý:\n{str(e)}")

    def _format_excel(self, path, theme_color):
        wb = load_workbook(path)
        ws = wb.active
        h_fill = PatternFill("solid", fgColor=theme_color)
        h_font = Font(name="Calibri", bold=True, color="FFFFFF")
        border = Border(left=Side(style="thin"), right=Side(style="thin"), 
                       top=Side(style="thin"), bottom=Side(style="thin"))

        for cell in ws[1]:
            cell.fill = h_fill
            cell.font = h_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = border

        for col in ws.columns:
            max_len = max((len(str(cell.value or "")) for cell in col), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 50)

        wb.save(path)


if __name__ == "__main__":
    root = tk.Tk()
    app = DataExportApp(root)
    root.mainloop()