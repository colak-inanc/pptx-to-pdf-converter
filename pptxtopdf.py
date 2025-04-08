import tkinter as tk
from tkinter import filedialog, messagebox
import os
import shutil
from reportlab.pdfgen import canvas
from PIL import Image
import comtypes.client

class PPTXtoPDFConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Toplu PPT/PPTX → PDF Dönüştürücü")
        self.root.geometry("500x300")
        self.root.configure(bg="#f2f2f2")
        self.temp_folder = "_converted_images_"

        self.main_frame = tk.Frame(root, bg="#f2f2f2", padx=20, pady=20)
        self.main_frame.pack(expand=True, fill='both')

        self.title_label = tk.Label(
            self.main_frame,
            text="🎯 Toplu PPT/PPTX → PDF Dönüştürücü",
            font=("Segoe UI", 16, "bold"),
            bg="#f2f2f2",
            fg="#333"
        )
        self.title_label.pack(pady=10)

        self.select_button = tk.Button(
            self.main_frame,
            text="📂 PPT/PPTX Dosyalarını Seç",
            command=self.select_files,
            bg="#4CAF50",
            fg="white",
            font=("Segoe UI", 10),
            width=35
        )
        self.select_button.pack(pady=5)

        self.convert_button = tk.Button(
            self.main_frame,
            text="💾 PDF'lere Dönüştür",
            command=self.convert_all,
            bg="#2196F3",
            fg="white",
            font=("Segoe UI", 10),
            width=35,
            state='disabled'
        )
        self.convert_button.pack(pady=5)

        self.status_label = tk.Label(
            self.main_frame,
            text="",
            font=("Segoe UI", 9),
            bg="#f2f2f2",
            fg="#666"
        )
        self.status_label.pack(pady=10)

        self.selected_files = []

    def select_files(self):
        file_paths = filedialog.askopenfilenames(
            title="PPT/PPTX Dosyaları Seç",
            filetypes=[
                ("PowerPoint Dosyaları", "*.ppt *.pptx"),
                ("Tüm Desteklenen", "*.pptx *.ppt")
            ]
        )
        if file_paths:
            self.selected_files = file_paths
            self.convert_button.config(state='normal')
            self.status_label.config(text=f"{len(file_paths)} dosya seçildi.")

    def convert_all(self):
        if not self.selected_files:
            messagebox.showerror("Hata", "Lütfen önce dosyaları seçin!")
            return

        try:
            powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
            powerpoint.Visible = 1
            powerpoint.WindowState = 2

            for file_path in self.selected_files:
                self.convert_single(file_path, powerpoint)

            powerpoint.Quit()
            self.status_label.config(text="✅ Tüm dönüşümler tamamlandı!")
            messagebox.showinfo("Başarılı", "Tüm sunumlar başarıyla PDF'e dönüştürüldü.")

        except Exception as e:
            self.status_label.config(text="❌ Toplu dönüşüm hatası!")
            messagebox.showerror("Hata", f"Hata oluştu:\n{str(e)}")

    def convert_single(self, ppt_path, powerpoint):
        try:
            if os.path.exists(self.temp_folder):
                shutil.rmtree(self.temp_folder)
            os.makedirs(self.temp_folder)

            abs_path = os.path.abspath(ppt_path)
            presentation = powerpoint.Presentations.Open(abs_path, WithWindow=False)

            save_path = os.path.abspath(self.temp_folder).replace("/", "\\")
            presentation.SaveAs(save_path, 18)  # 18 = Export as PNG
            presentation.Close()

            pdf_path = os.path.splitext(abs_path)[0] + ".pdf"
            c = canvas.Canvas(pdf_path)

            for img_file in sorted(os.listdir(self.temp_folder)):
                img_path = os.path.join(self.temp_folder, img_file)
                with Image.open(img_path) as img:
                    c.setPageSize(img.size)
                    c.drawInlineImage(img_path, 0, 0)
                    c.showPage()

            c.save()
            shutil.rmtree(self.temp_folder)
            print(f"[✓] {os.path.basename(ppt_path)} → PDF tamam.")

        except Exception as e:
            print(f"[X] Hata: {ppt_path} | {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PPTXtoPDFConverter(root)
    root.mainloop()
