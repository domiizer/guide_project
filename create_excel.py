import tkinter as tk
from tkinter import filedialog, messagebox
import xlsxwriter
import os
import platform
import subprocess

class XLSXFileCreatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("สร้างไฟล์ Excel (เลือกสร้างได้ 1–3 ไฟล์)")

        self.file_paths = [tk.StringVar() for _ in range(3)]
        
        for i in range(3):
            frame = tk.Frame(root)
            frame.pack(pady=5)

            label = tk.Label(frame, text=f"ไฟล์ที่ {i+1}:")
            label.pack(side=tk.LEFT)

            entry = tk.Entry(frame, textvariable=self.file_paths[i], width=40)
            entry.pack(side=tk.LEFT, padx=5)

            button = tk.Button(frame, text="เลือกที่บันทึก", command=lambda idx=i: self.select_save_path(idx))
            button.pack(side=tk.LEFT)

        self.create_button = tk.Button(root, text="สร้างไฟล์", command=self.create_files)
        self.create_button.pack(pady=20)

    def select_save_path(self, idx):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title=f"เลือกที่จัดเก็บไฟล์ที่ {idx+1}"
        )
        if file_path:
            self.file_paths[idx].set(file_path)

    def create_files(self):
        created = 0
        for i, path_var in enumerate(self.file_paths):
            file_path = path_var.get()
            if file_path:
                self.write_sample_excel(file_path, i+1)
                self.open_file(file_path)
                created += 1

        if created == 0:
            messagebox.showwarning("ยังไม่ได้เลือกไฟล์", "กรุณาเลือกอย่างน้อย 1 ไฟล์เพื่อสร้าง")
        else:
            messagebox.showinfo("สำเร็จ", f"สร้างและเปิดไฟล์สำเร็จ {created} ไฟล์!")

    def write_sample_excel(self, file_path, file_index):
        workbook = xlsxwriter.Workbook(file_path)
        worksheet = workbook.add_worksheet("ข้อมูล")

        worksheet.write("A1", "ไฟล์ที่")
        worksheet.write("B1", "รายการ")
        worksheet.write("C1", "จำนวน")

        for i in range(1, 6):
            worksheet.write(i, 0, f"ไฟล์ {file_index}")
            worksheet.write(i, 1, f"รายการ {i}")
            worksheet.write(i, 2, i * 10)

        workbook.close()

    def open_file(self, file_path):
        try:
            if platform.system() == "Windows":
                os.startfile(file_path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.call(["open", file_path])
            else:  # Linux, etc.
                subprocess.call(["xdg-open", file_path])
        except Exception as e:
            messagebox.showerror("เปิดไฟล์ไม่สำเร็จ", f"ไม่สามารถเปิดไฟล์ได้:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = XLSXFileCreatorApp(root)
    root.mainloop()
