import os
import pandas as pd
import time
import tkinter as tk
from tkinter import filedialog, ttk
from tkinter import messagebox

class FolderMakerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Folder Maker")
        self.root.geometry("600x500")
        
        # Variabel untuk menyimpan data
        self.excel_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.selected_sheet = tk.StringVar()
        self.selected_column = tk.StringVar()
        self.header_row = tk.StringVar(value="0")
        
        self.create_widgets()
    
    def create_widgets(self):
        # Frame untuk input file Excel
        excel_frame = ttk.LabelFrame(self.root, text="Langkah 1: Pilih File Excel", padding=10)
        excel_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Entry(excel_frame, textvariable=self.excel_path, width=50).pack(side="left", padx=5)
        ttk.Button(excel_frame, text="Pilih File", command=self.browse_excel).pack(side="left")
        ttk.Button(excel_frame, text="Baca File Excel", command=self.read_excel).pack(side="left", padx=5)
        
        # Frame untuk konfigurasi Excel
        config_frame = ttk.LabelFrame(self.root, text="Langkah 2: Pengaturan Excel", padding=10)
        config_frame.pack(fill="x", padx=10, pady=5)
        
        # Menambahkan label penjelasan
        ttk.Label(config_frame, 
                 text="Baris Judul: Pilih baris yang berisi judul kolom (0 = baris pertama)",
                 wraplength=550).grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        
        ttk.Label(config_frame, text="Pilih Baris Judul:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.header_combo = ttk.Combobox(config_frame, textvariable=self.header_row, 
                                        values=["0 (Baris Pertama)", "1 (Baris Kedua)", 
                                               "2 (Baris Ketiga)", "3 (Baris Keempat)"], 
                                        width=20)
        self.header_combo.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.header_combo.set("0 (Baris Pertama)")
        
        ttk.Label(config_frame, text="Pilih Sheet Excel:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.sheet_combo = ttk.Combobox(config_frame, textvariable=self.selected_sheet, width=20)
        self.sheet_combo.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        
        ttk.Label(config_frame, text="Pilih Kolom Data:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.column_combo = ttk.Combobox(config_frame, textvariable=self.selected_column, width=20)
        self.column_combo.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        
        # Frame untuk output folder
        output_frame = ttk.LabelFrame(self.root, text="Langkah 3: Pilih Lokasi Folder Output", padding=10)
        output_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(output_frame, 
                 text="Pilih lokasi dimana folder-folder baru akan dibuat:",
                 wraplength=550).pack(anchor="w", padx=5, pady=5)
        
        ttk.Entry(output_frame, textvariable=self.output_path, width=50).pack(side="left", padx=5)
        ttk.Button(output_frame, text="Pilih Lokasi", command=self.browse_output).pack(side="left")
        
        # Tombol proses
        process_frame = ttk.LabelFrame(self.root, text="Langkah 4: Proses Pembuatan Folder", padding=10)
        process_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(process_frame, 
                  text="Buat Folder", 
                  command=self.create_folders,
                  style='Accent.TButton').pack(pady=10)

        # Membuat style untuk tombol accent
        style = ttk.Style()
        style.configure('Accent.TButton', font=('Helvetica', 10, 'bold'))
    
    def browse_excel(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.excel_path.set(filename)
            self.update_excel_info()
    
    def browse_output(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_path.set(folder)
    
    def update_excel_info(self):
        try:
            excel = pd.ExcelFile(self.excel_path.get())
            self.sheet_combo['values'] = excel.sheet_names
            
            # Mengambil nomor header dari string combo (mengambil angka saja)
            header_text = self.header_row.get()
            header_value = int(header_text.split()[0])  # Mengambil angka di awal string
            
            # Baca sheet pertama untuk mendapatkan kolom
            df = pd.read_excel(self.excel_path.get(), 
                             sheet_name=excel.sheet_names[0],
                             header=header_value)
            self.column_combo['values'] = df.columns.tolist()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error membaca file Excel: {str(e)}")
    
    def read_excel(self):
        try:
            if not self.excel_path.get():
                messagebox.showwarning("Peringatan", "Silakan pilih file Excel terlebih dahulu!")
                return
            
            # Mengambil nomor header dari string combo (mengambil angka saja)
            header_text = self.header_row.get()
            header_value = int(header_text.split()[0])  # Mengambil angka di awal string
                
            excel = pd.ExcelFile(self.excel_path.get())
            self.sheet_combo['values'] = excel.sheet_names
            
            # Baca sheet pertama untuk mendapatkan kolom
            df = pd.read_excel(self.excel_path.get(), 
                             sheet_name=excel.sheet_names[0],
                             header=header_value)
            self.column_combo['values'] = df.columns.tolist()
            
            messagebox.showinfo("Sukses", 
                              "File Excel berhasil dibaca!\n\n" +
                              "Silakan pilih Sheet dan Kolom yang berisi data untuk nama folder.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membaca file Excel: {str(e)}")
    
    def create_folders(self):
        try:
            # Validasi input
            if not all([self.excel_path.get(), self.output_path.get(), 
                       self.selected_sheet.get(), self.selected_column.get()]):
                messagebox.showwarning("Peringatan", "Mohon lengkapi semua field!")
                return
            
            # Mengambil nomor header dari string combo
            header_text = self.header_row.get()
            header_value = int(header_text.split()[0])
            
            # Baca file Excel
            df = pd.read_excel(self.excel_path.get(),
                             sheet_name=self.selected_sheet.get(),
                             header=header_value)
            
            # Ambil nilai dari kolom yang dipilih dan pertahankan urutan
            folder_names = df[self.selected_column.get()].dropna().tolist()
            
            # Buat folder dengan delay bertahap untuk mempertahankan urutan
            for idx, num in enumerate(folder_names):
                num = str(int(num))  # Konversi ke integer untuk menghilangkan desimal
                folder_path = os.path.join(self.output_path.get(), num)
                try:
                    os.makedirs(folder_path, exist_ok=True)
                    # Gunakan waktu saat ini + indeks untuk memastikan urutan yang benar
                    created_time = time.time() + idx
                    os.utime(folder_path, (created_time, created_time))
                    print(f"Folder berhasil dibuat: {folder_path}")
                except Exception as e:
                    print(f"Gagal membuat folder {num}: {e}")
            
            messagebox.showinfo("Sukses", "Proses pembuatan folder selesai!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Terjadi kesalahan: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = FolderMakerApp(root)
    root.mainloop()