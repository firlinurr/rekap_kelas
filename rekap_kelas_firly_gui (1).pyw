import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd


class RekapMahasiswaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Aplikasi Rekap Nilai Mahasiswa")
        self.root.geometry("900x700")
        self.root.resizable(False, False)

        # Frame Input
        frame_input = ttk.LabelFrame(root, text="Input Data Mahasiswa", padding=(10, 10))
        frame_input.pack(pady=10, padx=10, fill="x")

        ttk.Label(frame_input, text="NIM:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.entry_nim = ttk.Entry(frame_input, width=20)
        self.entry_nim.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame_input, text="Nama Mahasiswa:").grid(row=0, column=2, sticky="w", padx=5, pady=5)
        self.entry_nama = ttk.Entry(frame_input, width=30)
        self.entry_nama.grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(frame_input, text="Mata Kuliah:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.entry_mk = ttk.Entry(frame_input, width=25)
        self.entry_mk.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(frame_input, text="Semester:").grid(row=1, column=2, sticky="w", padx=5, pady=5)
        self.entry_semester = ttk.Entry(frame_input, width=10)
        self.entry_semester.grid(row=1, column=3, padx=5, pady=5)

        ttk.Label(frame_input, text="Nilai:").grid(row=1, column=4, sticky="w", padx=5, pady=5)
        self.entry_nilai = ttk.Entry(frame_input, width=10)
        self.entry_nilai.grid(row=1, column=5, padx=5, pady=5)

        self.btn_tambah = ttk.Button(frame_input, text="Tambah Data", command=self.tambah_data)
        self.btn_tambah.grid(row=1, column=6, padx=10, pady=5)

        # Frame Tabel
        frame_tabel = ttk.LabelFrame(root, text="Data Mahasiswa", padding=(10, 10))
        frame_tabel.pack(pady=10, padx=10, fill="both", expand=True)

        # Tabel Data
        self.tree = ttk.Treeview(
            frame_tabel,
            columns=("NIM", "Nama", "Mata Kuliah", "Semester", "Nilai"),
            show="headings",
            height=15
        )
        self.tree.pack(side="left", fill="both", expand=True)

        # Scrollbar
        scrollbar = ttk.Scrollbar(frame_tabel, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscroll=scrollbar.set)

        self.tree.heading("NIM", text="NIM", anchor="center")
        self.tree.heading("Nama", text="Nama Mahasiswa", anchor="center")
        self.tree.heading("Mata Kuliah", text="Mata Kuliah", anchor="center")
        self.tree.heading("Semester", text="Semester", anchor="center")
        self.tree.heading("Nilai", text="Nilai", anchor="center")

        self.tree.column("NIM", width=100, anchor="center")
        self.tree.column("Nama", width=200, anchor="w")
        self.tree.column("Mata Kuliah", width=150, anchor="w")
        self.tree.column("Semester", width=80, anchor="center")
        self.tree.column("Nilai", width=80, anchor="center")

        # Frame Tombol Aksi
        frame_tombol = ttk.Frame(root)
        frame_tombol.pack(pady=10)

        self.btn_edit = ttk.Button(frame_tombol, text="Edit Data", command=self.edit_data, width=20)
        self.btn_edit.pack(side="left", padx=10)

        self.btn_hapus = ttk.Button(frame_tombol, text="Hapus Data", command=self.hapus_data, width=20)
        self.btn_hapus.pack(side="left", padx=10)

        self.btn_export = ttk.Button(frame_tombol, text="Export ke Excel", command=self.export_excel, width=20)
        self.btn_export.pack(side="left", padx=10)

        self.btn_import = ttk.Button(frame_tombol, text="Import dari CSV", command=self.import_csv, width=20)
        self.btn_import.pack(side="left", padx=10)

        self.btn_simpan_rekap = ttk.Button(frame_tombol, text="Simpan Rekap", command=self.simpan_rekap, width=20)
        self.btn_simpan_rekap.pack(side="left", padx=10)

        self.btn_statistik = ttk.Button(frame_tombol, text="Tampilkan Statistik", command=self.tampilkan_statistik, width=20)
        self.btn_statistik.pack(side="left", padx=10)

        # Statistik
        frame_statistik = ttk.LabelFrame(root, text="Statistik", padding=(10, 10))
        frame_statistik.pack(pady=10, padx=10, fill="x")

        self.label_total = ttk.Label(frame_statistik, text="Total Mahasiswa: 0")
        self.label_total.grid(row=0, column=0, sticky="w", padx=10, pady=5)

        self.label_rata_rata = ttk.Label(frame_statistik, text="Rata-Rata Nilai: 0")
        self.label_rata_rata.grid(row=0, column=1, sticky="w", padx=10, pady=5)

    def tambah_data(self):
        nim = self.entry_nim.get().strip()
        nama = self.entry_nama.get().strip()
        mk = self.entry_mk.get().strip()
        semester = self.entry_semester.get().strip()
        nilai = self.entry_nilai.get().strip()

        if not nim or not nama or not mk or not semester or not nilai:
            messagebox.showwarning("Peringatan", "Semua kolom harus diisi!")
            return

        if not nim.isdigit():
            messagebox.showerror("Error", "NIM harus berupa angka!")
            return

        try:
            nilai = float(nilai)
        except ValueError:
            messagebox.showerror("Error", "Nilai harus berupa angka!")
            return

        if not (0 <= nilai <= 100):
            messagebox.showerror("Error", "Nilai harus antara 0 dan 100!")
            return

        self.tree.insert("", "end", values=(nim, nama, mk, semester, nilai))
        self.reset_input()
        self.update_stats()

    def reset_input(self):
        self.entry_nim.delete(0, tk.END)
        self.entry_nama.delete(0, tk.END)
        self.entry_mk.delete(0, tk.END)
        self.entry_semester.delete(0, tk.END)
        self.entry_nilai.delete(0, tk.END)

    def simpan_rekap(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return

        try:
            # Mengambil data dari Treeview
            data = []
            for item in self.tree.get_children():
                row = self.tree.item(item)["values"]
                data.append(row)

            # Membuat DataFrame dari data
            df = pd.DataFrame(data, columns=["NIM", "Nama", "Mata Kuliah", "Semester", "Nilai"])

            # Menyimpan DataFrame ke file Excel
            df.to_excel(file_path, index=False)

            messagebox.showinfo("Sukses", "Rekap berhasil disimpan ke Excel!")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menyimpan rekap: {e}")

    def export_excel(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return

        try:
            # Mengambil data dari Treeview
            data = []
            for item in self.tree.get_children():
                row = self.tree.item(item)["values"]
                data.append(row)

            # Membuat DataFrame dari data
            df = pd.DataFrame(data, columns=["NIM", "Nama", "Mata Kuliah", "Semester", "Nilai"])

            # Menyimpan DataFrame ke file Excel
            df.to_excel(file_path, index=False)

            messagebox.showinfo("Sukses", "Data berhasil diexport ke Excel!")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menyimpan data: {e}")

    def update_stats(self):
        total_mahasiswa = len(self.tree.get_children())
        total_nilai = 0

        for item in self.tree.get_children():
            total_nilai += float(self.tree.item(item)["values"][4])

        if total_mahasiswa > 0:
            rata_rata = total_nilai / total_mahasiswa
        else:
            rata_rata = 0

        self.label_total.config(text=f"Total Mahasiswa: {total_mahasiswa}")
        self.label_rata_rata.config(text=f"Rata-Rata Nilai: {rata_rata:.2f}")

    def edit_data(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Peringatan", "Silakan pilih data yang akan diedit!")
            return

        item_values = self.tree.item(selected_item)["values"]

        self.entry_nim.delete(0, tk.END)
        self.entry_nim.insert(0, item_values[0])
        self.entry_nama.delete(0, tk.END)
        self.entry_nama.insert(0, item_values[1])
        self.entry_mk.delete(0, tk.END)
        self.entry_mk.insert(0, item_values[2])
        self.entry_semester.delete(0, tk.END)
        self.entry_semester.insert(0, item_values[3])
        self.entry_nilai.delete(0, tk.END)
        self.entry_nilai.insert(0, item_values[4])

        self.tree.delete(selected_item)
        self.update_stats()

    def hapus_data(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Peringatan", "Silakan pilih data yang akan dihapus!")
            return
        self.tree.delete(selected_item)
        self.update_stats()

    def import_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if not file_path:
            return

        try:
            data = pd.read_csv(file_path)
            for index, row in data.iterrows():
                self.tree.insert("", "end", values=(row["NIM"], row["Nama"], row["Mata Kuliah"], row["Semester"], row["Nilai"]))
            self.update_stats()
            messagebox.showinfo("Sukses", "Data berhasil diimpor dari CSV!")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal mengimpor data: {e}")

    def tampilkan_statistik(self):
        total_mahasiswa = len(self.tree.get_children())
        total_nilai = 0

        for item in self.tree.get_children():
            total_nilai += float(self.tree.item(item)["values"][4])

        if total_mahasiswa > 0:
            rata_rata = total_nilai / total_mahasiswa
        else:
            rata_rata = 0

        messagebox.showinfo("Statistik", f"Total Mahasiswa: {total_mahasiswa}\nRata-Rata Nilai: {rata_rata:.2f}")


if __name__ == "__main__":
    root = tk.Tk()
    app = RekapMahasiswaApp(root)
    root.mainloop()
