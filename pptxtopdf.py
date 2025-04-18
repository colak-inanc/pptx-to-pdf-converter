# Değişmeyen importlar
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar, Style
import os
import threading
import comtypes.client
import winsound

class PPTXtoPDFConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("📄 PPT/PPTX → PDF Dönüştürücü")
        self.root.geometry("720x600")
        self.root.configure(bg="#e8edf1")

        self.selected_files = []
        self.output_folder = ""
        self.use_original_folder = tk.BooleanVar(value=True)

        self.build_ui()

    def build_ui(self):
        # Ana kart görünümü
        self.card = tk.Frame(self.root, bg="white", padx=30, pady=30, bd=1, relief="solid")
        self.card.place(relx=0.5, rely=0.5, anchor="center")

        tk.Label(
            self.card,
            text="🧠 PPT/PPTX → PDF Dönüştürücü",
            font=("Segoe UI", 20, "bold"),
            bg="white",
            fg="#333"
        ).pack(pady=(10, 20))

        tk.Button(
            self.card, text="📂 PPT/PPTX Dosyalarını Seç", command=self.select_files,
            bg="#4CAF50", fg="white", font=("Segoe UI", 10, "bold"), width=40, height=2
        ).pack(pady=5)

        tk.Button(
            self.card, text="📁 Hedef Klasör Seç", command=self.select_output_folder,
            bg="#FF9800", fg="white", font=("Segoe UI", 10, "bold"), width=40, height=2
        ).pack(pady=5)

        tk.Checkbutton(
            self.card,
            text="📌 PDF'yi dosyanın bulunduğu klasöre kaydet",
            variable=self.use_original_folder,
            bg="white",
            font=("Segoe UI", 9),
            onvalue=True,
            offvalue=False
        ).pack(pady=(5, 10))

        self.file_frame = tk.Frame(self.card, bg="white")
        self.file_frame.pack(pady=5)

        tk.Button(
            self.card, text="♻️ Seçimi Temizle", command=self.clear_selection,
            bg="#9E9E9E", fg="white", font=("Segoe UI", 10), width=40, height=2
        ).pack(pady=5)

        self.convert_button = tk.Button(
            self.card, text="💾 PDF'lere Dönüştür", command=self.start_conversion,
            bg="#2196F3", fg="white", font=("Segoe UI", 10, "bold"), width=40, height=2,
            state='disabled'
        )
        self.convert_button.pack(pady=5)

        style = Style()
        style.theme_use('default')
        style.configure(
            "modern.Horizontal.TProgressbar",
            troughcolor="#ddd",
            background="#00C853",
            thickness=20,
            bordercolor="white"
        )

        self.progress = Progressbar(
            self.card, orient="horizontal", length=500,
            mode='determinate', style="modern.Horizontal.TProgressbar"
        )
        self.progress.pack(pady=10)

        self.status_label = tk.Label(self.card, text="", bg="white", font=("Segoe UI", 10), fg="#666")
        self.status_label.pack(pady=5)

    def update_file_list_labels(self):
        for widget in self.file_frame.winfo_children():
            widget.destroy()

        for f in self.selected_files:
            row = tk.Frame(self.file_frame, bg="white")
            row.pack(fill="x", pady=2)

            tk.Label(
                row,
                text="📄 " + os.path.basename(f),
                font=("Segoe UI", 9),
                bg="#f5f5f5",
                fg="black",
                anchor="w",
                relief="groove",
                padx=8,
                pady=4,
                width=45
            ).pack(side="left", padx=2)

            tk.Button(
                row,
                text="❌",
                command=lambda path=f: self.remove_file(path),
                font=("Segoe UI", 10, "bold"),
                bg="#FF5252",
                fg="white",
                relief="flat",
                width=3
            ).pack(side="right", padx=4)

    def select_files(self):
        paths = filedialog.askopenfilenames(
            title="PPT/PPTX Dosyaları Seç",
            filetypes=[("PowerPoint Dosyaları", "*.ppt *.pptx")]
        )
        if paths:
            self.selected_files = list(paths)
            self.update_file_list_labels()
            self.update_convert_state()

    def remove_file(self, path):
        if path in self.selected_files:
            self.selected_files.remove(path)
            self.update_file_list_labels()
            self.update_convert_state()

    def select_output_folder(self):
        folder = filedialog.askdirectory(title="PDF'ler için Klasör Seç")
        if folder:
            self.output_folder = folder
            self.use_original_folder.set(False)
            self.update_convert_state()

    def clear_selection(self):
        self.selected_files.clear()
        for widget in self.file_frame.winfo_children():
            widget.destroy()
        self.output_folder = ""
        self.convert_button.config(state='disabled')
        self.progress["value"] = 0
        self.status_label.config(text="Seçim temizlendi.")
        self.use_original_folder.set(True)

    def update_convert_state(self):
        if self.selected_files and (self.use_original_folder.get() or self.output_folder):
            self.convert_button.config(state='normal')
            self.status_label.config(text=f"{len(self.selected_files)} dosya seçildi.")
        else:
            self.convert_button.config(state='disabled')

    def start_conversion(self):
        threading.Thread(target=self.convert_all, daemon=True).start()

    def convert_all(self):
        try:
            self.progress["value"] = 0
            total = len(self.selected_files)
            powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
            powerpoint.Visible = 1
            powerpoint.WindowState = 2

            for i, file in enumerate(self.selected_files):
                self.convert_single(file, powerpoint)
                self.progress["value"] = ((i + 1) / total) * 100
                self.status_label.config(text=f"{i+1}/{total} dönüştürüldü.")
                self.root.update_idletasks()

            powerpoint.Quit()
            winsound.Beep(700, 300)
            self.status_label.config(text="✅ Tüm dosyalar başarıyla dönüştürüldü.")
            messagebox.showinfo("✔️ Tamamlandı", "Tüm sunumlar başarıyla PDF'e dönüştürüldü.")

        except Exception as e:
            with open("hata_log.txt", "a", encoding="utf-8") as log:
                log.write(str(e) + "\n")
            self.status_label.config(text="❌ Hata oluştu.")
            messagebox.showerror("Hata", f"Hata:\n{str(e)}")

    def convert_single(self, ppt_path, powerpoint):
        try:
            pres = powerpoint.Presentations.Open(os.path.abspath(ppt_path), WithWindow=False)
            pdf_filename = os.path.splitext(os.path.basename(ppt_path))[0] + ".pdf"

            if self.use_original_folder.get():
                output_path = os.path.join(os.path.dirname(ppt_path), pdf_filename)
            else:
                output_path = os.path.join(self.output_folder, pdf_filename)

            pres.SaveAs(os.path.normpath(output_path), 32)
            pres.Close()

        except Exception as e:
            with open("hata_log.txt", "a", encoding="utf-8") as log:
                log.write(f"{ppt_path} | {str(e)}\n")

if __name__ == "__main__":
    root = tk.Tk()
    app = PPTXtoPDFConverter(root)
    root.mainloop()
