# -*- coding: utf-8 -*-
"""
VBTM Pro 2.2.4 - Ứng Dụng Desktop Chuẩn Hóa Văn Bản & Soát Lỗi Chính Tả AI Offline
Hỗ trợ xuất trực tiếp tệp Microsoft Word (.doc) chuẩn Nghị định 30/2020/NĐ-CP & CAND
"""

import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

SPELL_RULES = [
    (r"\bbổ xung\b", "bổ sung"),
    (r"\bsử lý\b", "xử lý"),
    (r"\bxử dụng\b", "sử dụng"),
    (r"\bchính sác\b", "chính xác"),
    (r"\bkiểm xoát\b", "kiểm soát"),
    (r"\bchiển khai\b", "triển khai"),
    (r"\bchủ chương\b", "chủ trương"),
    (r"\bkế hoạc\b", "kế hoạch"),
    (r"\bngiêm túc\b", "nghiêm túc"),
    (r"\bhổ trợ\b", "hỗ trợ"),
    (r"\btiêu chuẫn\b", "tiêu chuẩn"),
    (r"\bkhẩn chương\b", "khẩn trương"),
    (r"\bsát nhập\b", "sáp nhập"),
]


class VBTMProApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("VBTM Pro 2.2.4 - Chuẩn Hóa Văn Bản Nghị Định 30 & Chính Tả AI Offline")
        self.root.geometry("920x660")
        self.root.configure(bg="#0f172a")
        self._build_ui()

    def _build_ui(self):
        header = tk.Frame(self.root, bg="#1e293b", pady=12, padx=16)
        header.pack(fill="x")
        tk.Label(
            header,
            text="VBTM Pro 2.2.4 · Chuẩn Hóa Thể Thức Văn Bản Nghị Định 30, CAND & Soát Lỗi Chính Tả AI",
            font=("Segoe UI", 13, "bold"),
            fg="#38bdf8",
            bg="#1e293b",
        ).pack(side="left")

        toolbar = tk.Frame(self.root, bg="#0f172a", pady=10, padx=16)
        toolbar.pack(fill="x")

        tk.Button(
            toolbar,
            text="1. Chuẩn Hóa Văn Bản & Sửa Chính Tả AI",
            command=self.standardize_and_spellcheck,
            bg="#0284c7",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            padx=12,
            pady=6,
            relief="flat",
        ).pack(side="left", padx=4)

        tk.Button(
            toolbar,
            text="2. Xuất File Word (.doc Chuẩn A4 NĐ30)",
            command=self.export_word,
            bg="#059669",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            padx=12,
            pady=6,
            relief="flat",
        ).pack(side="left", padx=4)

        meta_frame = tk.Frame(self.root, bg="#0f172a", padx=16, pady=4)
        meta_frame.pack(fill="x")

        tk.Label(meta_frame, text="Cơ quan ban hành:", fg="#cbd5e1", bg="#0f172a").grid(row=0, column=0, sticky="w")
        self.org_var = tk.StringVar(value="PHÒNG KỸ THUẬT HÌNH SỰ")
        tk.Entry(meta_frame, textvariable=self.org_var, width=32, font=("Segoe UI", 10)).grid(row=0, column=1, padx=8, pady=4)

        tk.Label(meta_frame, text="Số/Ký hiệu:", fg="#cbd5e1", bg="#0f172a").grid(row=0, column=2, sticky="w")
        self.num_var = tk.StringVar(value="86/BC-PC09")
        tk.Entry(meta_frame, textvariable=self.num_var, width=20, font=("Segoe UI", 10)).grid(row=0, column=3, padx=8, pady=4)

        tk.Label(meta_frame, text="Tiêu đề / Loại VB:", fg="#cbd5e1", bg="#0f172a").grid(row=1, column=0, sticky="w")
        self.title_var = tk.StringVar(value="BÁO CÁO")
        tk.Entry(meta_frame, textvariable=self.title_var, width=32, font=("Segoe UI", 10)).grid(row=1, column=1, padx=8, pady=4)

        tk.Label(meta_frame, text="Địa danh, ngày tháng:", fg="#cbd5e1", bg="#0f172a").grid(row=1, column=2, sticky="w")
        self.date_var = tk.StringVar(value="Đà Nẵng, ngày 11 tháng 11 năm 2026")
        tk.Entry(meta_frame, textvariable=self.date_var, width=32, font=("Segoe UI", 10)).grid(row=1, column=3, padx=8, pady=4)

        editor_frame = tk.Frame(self.root, bg="#0f172a", padx=16, pady=8)
        editor_frame.pack(fill="both", expand=True)

        self.text_area = tk.Text(
            editor_frame,
            font=("Times New Roman", 13),
            wrap="word",
            bg="#ffffff",
            fg="#000000",
            padx=18,
            pady=14,
        )
        self.text_area.pack(fill="both", expand=True)
        self.text_area.insert(
            "1.0",
            "Thực hiện chủ chương cải cách hành chính và ứng dụng công nghệ thông tin năm 2026, "
            "đơn vị đã khẩn chương chiển khai kế hoạc bổ xung phần mềm chuẩn hóa văn bản theo "
            "Nghị định số 30/2020/NĐ-CP và quy chuẩn CAND , đảm bảo kiểm xoát chính sác thể thức văn bản",
        )

    def standardize_and_spellcheck(self):
        raw = self.text_area.get("1.0", "end-1c")
        fixed_count = 0
        for pattern, replacement in SPELL_RULES:
            raw, n = re.subn(pattern, replacement, raw, flags=re.IGNORECASE)
            fixed_count += n

        raw = re.sub(r"[ \t]+", " ", raw)
        raw = re.sub(r"\s+([,.;:!?])", r"\1", raw)
        raw = re.sub(r"([,.;:!?])([^\s\d/])", r"\1 \2", raw)
        raw = raw.strip()
        if raw and not raw.endswith("./."):
            raw = raw.rstrip(".") + "./."

        self.text_area.delete("1.0", "end")
        self.text_area.insert("1.0", raw)
        messagebox.showinfo(
            "VBTM Pro 2.2.4",
            f"Đã chuẩn hóa thể thức văn bản & tự động sửa {fixed_count} lỗi chính tả tiếng Việt!",
        )

    def export_word(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".doc",
            filetypes=[("Microsoft Word Document", "*.doc")],
            initialfile="VanBan_ChuanHoa_ND30.doc",
        )
        if not path:
            return
        body_paragraphs = "".join(
            f"<p style='text-indent:1.27cm;text-align:justify;margin:0 0 6pt 0;'>{line.strip()}</p>"
            for line in self.text_area.get("1.0", "end-1c").split("\n")
            if line.strip()
        )
        html_doc = f"""<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns="http://www.w3.org/TR/REC-html40">
<head><meta charset="utf-8">
<style>
@page Section1 {{ size:21.0cm 29.7cm; margin:2.0cm 1.5cm 2.0cm 3.0cm; }}
div.Section1 {{ page:Section1; font-family:"Times New Roman",serif; font-size:14pt; }}
table {{ width:100%; border-collapse:collapse; margin-bottom:14pt; }}
td {{ text-align:center; vertical-align:top; font-family:"Times New Roman",serif; }}
</style></head>
<body><div class="Section1">
<table><tr>
<td style="width:45%;"><b>{self.org_var.get()}</b><br>Số: {self.num_var.get()}</td>
<td style="width:55%;"><b>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</b><br><b>Độc lập - Tự do - Hạnh phúc</b><br><i>{self.date_var.get()}</i></td>
</tr></table>
<p style="text-align:center;font-weight:bold;font-size:14pt;">{self.title_var.get()}</p>
{body_paragraphs}
</div></body></html>"""
        with open(path, "w", encoding="utf-8-sig") as f:
            f.write(html_doc)
        messagebox.showinfo("VBTM Pro 2.2.4", f"Đã xuất tệp Word chuẩn Nghị định 30 tại:\n{path}")


if __name__ == "__main__":
    root = tk.Tk()
    app = VBTMProApp(root)
    root.mainloop()
