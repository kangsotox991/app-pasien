#!/usr/bin/env python3
"""
Excel Logbook Editor
Desktop application untuk membaca dan menyimpan file Excel.
Fitur: buka file, import data pasien, rename sheet otomatis sesuai tanggal,
       pengisian otomatis no CM, tanggal, dan alat medis EKG.
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import re
from collections import OrderedDict
from datetime import datetime
import io

try:
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    from openpyxl.drawing.image import Image as XlImage
    from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, OneCellAnchor, AnchorMarker
except ImportError:
    import subprocess, sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    from openpyxl.drawing.image import Image as XlImage
    from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, OneCellAnchor, AnchorMarker

BULAN_INDO = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
    5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
    9: "September", 10: "Oktober", 11: "November", 12: "Desember"
}


class ExcelEditorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Logbook Editor")
        self.root.geometry("1200x700")
        self.root.minsize(900, 500)

        self.workbook = None
        self.filepath = None
        self.original_filepath = None
        self.current_sheet_name = None
        self.modified = False
        self.patient_data = None

        self._build_ui()

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        # Menu bar
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Buka File Excel...", command=self.open_file, accelerator="Ctrl+O")
        file_menu.add_command(label="Simpan", command=self.save_file, accelerator="Ctrl+S")
        file_menu.add_command(label="Simpan Sebagai...", command=self.save_as, accelerator="Ctrl+Shift+S")
        file_menu.add_separator()
        file_menu.add_command(label="Keluar", command=self.on_close)
        menubar.add_cascade(label="File", menu=file_menu)

        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="Load Data Pasien...", command=self.load_data_pasien)
        tools_menu.add_command(label="Proses Data & Rename Sheets", command=self.rename_sheets_from_data)
        menubar.add_cascade(label="Tools", menu=tools_menu)

        self.root.config(menu=menubar)
        self.root.bind("<Control-o>", lambda e: self.open_file())
        self.root.bind("<Control-s>", lambda e: self.save_file())
        self.root.bind("<Control-S>", lambda e: self.save_as())
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # Toolbar
        toolbar = ttk.Frame(self.root)
        toolbar.pack(fill=tk.X, padx=8, pady=(8, 0))

        ttk.Button(toolbar, text="Buka File", command=self.open_file).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Simpan", command=self.save_file).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Simpan Sebagai", command=self.save_as).pack(side=tk.LEFT, padx=2)
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=8)
        ttk.Button(toolbar, text="Load Data Pasien", command=self.load_data_pasien).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Proses & Rename", command=self.rename_sheets_from_data).pack(side=tk.LEFT, padx=2)

        self.file_label = ttk.Label(toolbar, text="Belum ada file dibuka", foreground="gray")
        self.file_label.pack(side=tk.RIGHT, padx=8)

        self.status_label = ttk.Label(toolbar, text="", foreground="green")
        self.status_label.pack(side=tk.RIGHT, padx=8)

        # Sheet tabs
        self.tab_frame = ttk.Frame(self.root)
        self.tab_frame.pack(fill=tk.X, padx=8, pady=(8, 0))

        self.tab_canvas = tk.Canvas(self.tab_frame, height=32, highlightthickness=0)
        self.tab_scrollbar = ttk.Scrollbar(self.tab_frame, orient=tk.HORIZONTAL, command=self.tab_canvas.xview)
        self.tab_inner = ttk.Frame(self.tab_canvas)

        self.tab_canvas.configure(xscrollcommand=self.tab_scrollbar.set)
        self.tab_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.tab_canvas.pack(side=tk.TOP, fill=tk.X)
        self.tab_canvas.create_window((0, 0), window=self.tab_inner, anchor="nw")
        self.tab_inner.bind("<Configure>", lambda e: self.tab_canvas.configure(scrollregion=self.tab_canvas.bbox("all")))

        # Table area
        table_frame = ttk.Frame(self.root)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        self.tree = ttk.Treeview(table_frame, show="headings", selectmode="browse")

        vsb = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        # Status bar
        self.statusbar = ttk.Label(self.root, text="Siap", relief=tk.SUNKEN, anchor=tk.W)
        self.statusbar.pack(fill=tk.X, side=tk.BOTTOM)

        # Welcome screen
        self._show_welcome()

    def _show_welcome(self):
        for widget in self.tab_inner.winfo_children():
            widget.destroy()
        self.tree["columns"] = ()
        self.tree.delete(*self.tree.get_children())

    # ------------------------------------------------------------------ FILE OPS
    def open_file(self):
        path = filedialog.askopenfilename(
            title="Pilih File Excel",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            self.workbook = load_workbook(path)
            self.filepath = path
            self.original_filepath = path
            self.modified = False
            self.patient_data = None
            self.file_label.config(text=path.split("/")[-1].split("\\")[-1], foreground="black")
            self._render_tabs()
            self._switch_sheet(self.workbook.sheetnames[0])
            self.statusbar.config(text=f"File dibuka: {path}")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membuka file:\n{e}")

    def save_file(self):
        if not self.workbook or not self.filepath:
            self.save_as()
            return
        # Never overwrite original file — auto-generate new filename
        if self.filepath == self.original_filepath:
            base, ext = os.path.splitext(self.original_filepath)
            new_path = f"{base}_edited{ext}"
            counter = 1
            while os.path.exists(new_path):
                new_path = f"{base}_edited_{counter}{ext}"
                counter += 1
            self.filepath = new_path
        try:
            self.workbook.save(self.filepath)
            self.modified = False
            self.file_label.config(text=self.filepath.split("/")[-1].split("\\")[-1], foreground="black")
            self.status_label.config(text="Tersimpan", foreground="green")
            self.statusbar.config(text=f"File disimpan: {self.filepath} (file asli tidak diubah)")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menyimpan:\n{e}")

    def save_as(self):
        if not self.workbook:
            messagebox.showwarning("Peringatan", "Belum ada file yang dibuka")
            return
        path = filedialog.asksaveasfilename(
            title="Simpan Sebagai",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            self.workbook.save(path)
            self.filepath = path
            self.modified = False
            self.file_label.config(text=path.split("/")[-1].split("\\")[-1], foreground="black")
            self.status_label.config(text="Tersimpan", foreground="green")
            self.statusbar.config(text=f"File disimpan: {path}")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menyimpan:\n{e}")

    def on_close(self):
        if self.modified:
            ans = messagebox.askyesnocancel("Simpan?", "Ada perubahan yang belum disimpan. Simpan dulu?")
            if ans is None:
                return
            if ans:
                self.save_file()
        self.root.destroy()

    # ------------------------------------------------------------------ SHEET TABS
    def _render_tabs(self):
        for w in self.tab_inner.winfo_children():
            w.destroy()
        if not self.workbook:
            return
        for name in self.workbook.sheetnames:
            btn = tk.Button(
                self.tab_inner, text=name, padx=12, pady=4,
                relief=tk.RAISED if name != self.current_sheet_name else tk.SUNKEN,
                bg="#2563eb" if name == self.current_sheet_name else "#f0f0f0",
                fg="white" if name == self.current_sheet_name else "black",
                font=("Segoe UI", 9, "bold" if name == self.current_sheet_name else "normal"),
                command=lambda n=name: self._switch_sheet(n)
            )
            btn.pack(side=tk.LEFT, padx=1, pady=2)

    def _switch_sheet(self, name):
        self.current_sheet_name = name
        self._render_tabs()
        self._render_table()

    # ------------------------------------------------------------------ TABLE
    def _render_table(self):
        self.tree.delete(*self.tree.get_children())

        if not self.workbook or not self.current_sheet_name:
            return

        ws = self.workbook[self.current_sheet_name]
        max_col = ws.max_column or 1
        max_row = ws.max_row or 1

        cols = ["#"] + [get_column_letter(c) for c in range(1, max_col + 1)]
        self.tree["columns"] = cols
        for c in cols:
            w = 50 if c == "#" else 140
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, minwidth=40, stretch=True)

        for row_idx in range(1, max_row + 1):
            values = [str(row_idx)]
            for col_idx in range(1, max_col + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                val = cell.value
                if val is None:
                    val = ""
                values.append(str(val))
            self.tree.insert("", tk.END, values=values, tags=(f"row_{row_idx}",))

        self.statusbar.config(text=f"Sheet: {self.current_sheet_name} | {max_row} baris x {max_col} kolom")

    # ------------------------------------------------------------------ LOAD DATA PASIEN
    def load_data_pasien(self):
        """Load patient data from .txt file or paste — only loads into memory."""
        win = tk.Toplevel(self.root)
        win.title("Load Data Pasien")
        win.geometry("750x650")
        win.transient(self.root)
        win.grab_set()

        ttk.Label(win, text="Load Data Pasien", font=("Segoe UI", 14, "bold")).pack(pady=(12, 4))
        ttk.Label(win, text=(
            "Paste data pasien di bawah, atau load dari file .txt.\n"
            "Format per baris: DD/MM/YYYY Nama No. Reg XXXXXXX dx: diagnosa\n\n"
            "Data akan disimpan di memori. Untuk proses dan rename sheet,\n"
            "gunakan tombol 'Proses & Rename' di toolbar atau menu Tools."
        ), justify=tk.LEFT, foreground="gray").pack(padx=16, anchor="w")

        # Load from file button
        load_frame = ttk.Frame(win)
        load_frame.pack(fill=tk.X, padx=16, pady=(4, 0))

        def load_from_txt():
            path = filedialog.askopenfilename(
                title="Pilih File Data Pasien",
                filetypes=[("Text files", "*.txt"), ("CSV files", "*.csv"), ("All files", "*.*")],
                parent=win
            )
            if not path:
                return
            try:
                with open(path, "r", encoding="utf-8") as f:
                    content = f.read()
            except UnicodeDecodeError:
                with open(path, "r", encoding="latin-1") as f:
                    content = f.read()
            text.delete("1.0", tk.END)
            text.insert("1.0", content)
            result_label.config(
                text=f"File dimuat: {path.split('/')[-1].split(chr(92))[-1]}",
                foreground="green"
            )

        ttk.Button(load_frame, text="Load dari File .txt", command=load_from_txt).pack(side=tk.LEFT)
        ttk.Label(load_frame, text="  atau paste langsung di bawah", foreground="gray").pack(side=tk.LEFT)

        text_frame = ttk.Frame(win)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=16, pady=8)

        text = tk.Text(text_frame, font=("Consolas", 10), wrap=tk.NONE)
        text_vsb = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text.yview)
        text_hsb = ttk.Scrollbar(text_frame, orient=tk.HORIZONTAL, command=text.xview)
        text.configure(yscrollcommand=text_vsb.set, xscrollcommand=text_hsb.set)

        text.grid(row=0, column=0, sticky="nsew")
        text_vsb.grid(row=0, column=1, sticky="ns")
        text_hsb.grid(row=1, column=0, sticky="ew")
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)

        # Pre-fill if data already loaded
        if self.patient_data:
            text.insert("1.0", self.patient_data)

        btn_frame = ttk.Frame(win)
        btn_frame.pack(fill=tk.X, padx=16, pady=(0, 12))

        result_label = ttk.Label(btn_frame, text="", foreground="blue")
        result_label.pack(side=tk.LEFT)

        def preview():
            raw = text.get("1.0", tk.END).strip()
            if not raw:
                result_label.config(text="Tidak ada data", foreground="red")
                return
            groups = self._parse_patient_data(raw)
            if not groups:
                result_label.config(text="Format data tidak dikenali", foreground="red")
                return
            dates = list(groups.keys())
            total = sum(len(v) for v in groups.values())
            sheet_info = f" Sheet tersedia: {len(self.workbook.sheetnames)}" if self.workbook else ""
            result_label.config(
                text=f"Ditemukan {len(dates)} tanggal, {total} pasien.{sheet_info}",
                foreground="blue"
            )

        def do_load():
            raw = text.get("1.0", tk.END).strip()
            if not raw:
                messagebox.showwarning("Peringatan", "Tidak ada data", parent=win)
                return
            groups = self._parse_patient_data(raw)
            if not groups:
                messagebox.showerror("Error", "Format data tidak dikenali", parent=win)
                return

            self.patient_data = raw
            dates = list(groups.keys())
            total = sum(len(v) for v in groups.values())

            self.statusbar.config(text=f"Data pasien dimuat: {len(dates)} tanggal, {total} pasien")
            self.status_label.config(text="Data dimuat", foreground="blue")

            messagebox.showinfo(
                "Sukses",
                f"Data pasien berhasil dimuat!\n"
                f"{len(dates)} tanggal, {total} pasien.\n\n"
                f"Untuk proses dan rename sheet, klik 'Proses & Rename' di toolbar.",
                parent=win
            )
            win.destroy()

        ttk.Button(btn_frame, text="Preview", command=preview).pack(side=tk.RIGHT, padx=4)
        ttk.Button(btn_frame, text="Simpan Data", command=do_load).pack(side=tk.RIGHT, padx=4)
        ttk.Button(btn_frame, text="Batal", command=win.destroy).pack(side=tk.RIGHT, padx=4)

    # ------------------------------------------------------------------ PROSES DATA & RENAME SHEETS
    def rename_sheets_from_data(self):
        """Copy template sheet per tanggal, isi data pasien, dan rename."""
        if not self.workbook:
            messagebox.showwarning("Peringatan", "Buka file Excel terlebih dahulu")
            return
        if not self.patient_data:
            messagebox.showwarning(
                "Peringatan",
                "Belum ada data pasien yang dimuat.\n"
                "Klik 'Load Data Pasien' terlebih dahulu."
            )
            return

        groups = self._parse_patient_data(self.patient_data)
        if not groups:
            messagebox.showerror("Error", "Format data pasien tidak dikenali")
            return

        dates = list(groups.keys())
        total = sum(len(v) for v in groups.values())

        # Use first sheet as template
        template_ws = self.workbook.worksheets[0]
        template_name = template_ws.title

        # Confirm before processing
        msg = (
            f"Data: {len(dates)} tanggal, {total} pasien\n"
            f"Template sheet: '{template_name}'\n\n"
            f"Akan dibuat {len(dates)} sheet baru dari template,\n"
            f"diisi data pasien otomatis, dan di-rename sesuai tanggal.\n\n"
            f"Lanjutkan?"
        )
        if not messagebox.askyesno("Konfirmasi Proses", msg):
            return

        # Create copies of template for each date
        for i, date_key in enumerate(dates):
            day, month, year = date_key
            patients = groups[date_key]

            # Copy template sheet
            new_ws = self.workbook.copy_worksheet(template_ws)
            self._copy_images(template_ws, new_ws)

            # Rename sheet
            sheet_name = f"{day} {BULAN_INDO[month]}"
            sheet_name = re.sub(r'[\\/*?\[\]:]', '', sheet_name)[:31]

            # Handle duplicate names
            existing = [ws.title for ws in self.workbook.worksheets]
            if sheet_name in existing:
                sheet_name = f"{day} {BULAN_INDO[month]} {year}"
                sheet_name = re.sub(r'[\\/*?\[\]:]', '', sheet_name)[:31]
            new_ws.title = sheet_name

            # Fill patient data into the sheet
            self._fill_sheet_data(new_ws, date_key, patients)

        # Remove original template sheets (the ones that were there before)
        original_sheets = list(self.workbook.worksheets[:len(self.workbook.sheetnames) - len(dates)])
        for ws in original_sheets:
            # Only remove if it's one of the original template sheets
            if ws.title in [s.title for s in self.workbook.worksheets[:len(self.workbook.sheetnames) - len(dates)]]:
                self.workbook.remove(ws)

        self.modified = True
        self.status_label.config(text="Belum disimpan", foreground="orange")
        self.current_sheet_name = self.workbook.sheetnames[0]
        self._render_tabs()
        self._render_table()

        messagebox.showinfo(
            "Sukses",
            f"Berhasil memproses {len(dates)} sheet!\n"
            f"Data pasien telah diisi otomatis (no CM, tanggal, alat medis)."
        )

    def _fill_sheet_data(self, ws, date_key, patients):
        """Fill patient data into a sheet.

        1. Tanggal di kolom B sebelah Briefing (row 16)
        2. No. Reg di belakang teks 'no CM'
        3. No. Reg pasien EKG terakhir di belakang teks alat medis
        """
        day, month, year = date_key

        # Build date object for the date cell
        date_obj = datetime(year, month, day)

        # Extract No. Reg from each patient line
        reg_numbers = []
        last_ekg_reg = None
        for patient_line in patients:
            reg_match = re.search(r'No\.\s*Reg\s+(\d+)', patient_line)
            if reg_match:
                reg_num = reg_match.group(1)
                reg_numbers.append(reg_num)
                # Check if this patient has EKG in diagnosis
                if 'ekg' in patient_line.lower():
                    last_ekg_reg = reg_num

        reg_list_str = ", ".join(reg_numbers)
        max_col = ws.max_column or 1
        max_row = min(ws.max_row or 1, 40)

        # Scan all cells to find and fill data
        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                val = ws.cell(row=row, column=col).value
                if val is None or not isinstance(val, str):
                    continue

                # 1. Fill "no CM" cells with No. Reg list
                if val.rstrip().endswith("no CM"):
                    ws.cell(row=row, column=col).value = (
                        f"{val} {reg_list_str}."
                    )

                # 2. Fill alat medis (syrenge) cell with last EKG patient's No. Reg
                if "agar siap pakai dengan" in val.lower():
                    if last_ekg_reg:
                        ws.cell(row=row, column=col).value = (
                            f"{val.rstrip()} {last_ekg_reg}."
                        )

        # 3. Fill tanggal di kolom sebelah Briefing
        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                val = ws.cell(row=row, column=col).value
                if val is None:
                    continue
                if isinstance(val, str) and val.strip().lower() == "briefing":
                    # Tanggal is in the column before Briefing (col - 1)
                    if col > 1:
                        ws.cell(row=row, column=col - 1).value = date_obj

    def _parse_patient_data(self, raw_text):
        """Parse patient data text and group by date.
        Returns OrderedDict: {(day, month, year): [lines...]}
        """
        groups = OrderedDict()
        pattern = re.compile(
            r'(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})\s+(.+)'
        )

        for line in raw_text.strip().splitlines():
            line = line.strip()
            if not line:
                continue
            m = pattern.match(line)
            if m:
                day = int(m.group(1))
                month = int(m.group(2))
                year = int(m.group(3))
                rest = m.group(4).strip()
                key = (day, month, year)
                if key not in groups:
                    groups[key] = []
                groups[key].append(rest)

        return groups

    def _copy_images(self, source_ws, target_ws):
        """Copy all images from source worksheet to target worksheet."""
        for img in source_ws._images:
            try:
                img.ref.seek(0)
                img_data = img.ref.read()
                img.ref.seek(0)
            except Exception:
                try:
                    img_data = img._data()
                except Exception:
                    continue

            new_img = XlImage(io.BytesIO(img_data))
            new_img.width = img.width
            new_img.height = img.height

            anchor = img.anchor
            if isinstance(anchor, TwoCellAnchor):
                new_anchor = TwoCellAnchor()
                new_anchor._from = AnchorMarker(
                    col=anchor._from.col,
                    colOff=anchor._from.colOff,
                    row=anchor._from.row,
                    rowOff=anchor._from.rowOff
                )
                new_anchor.to = AnchorMarker(
                    col=anchor.to.col,
                    colOff=anchor.to.colOff,
                    row=anchor.to.row,
                    rowOff=anchor.to.rowOff
                )
                new_img.anchor = new_anchor
            elif isinstance(anchor, OneCellAnchor):
                new_anchor = OneCellAnchor()
                new_anchor._from = AnchorMarker(
                    col=anchor._from.col,
                    colOff=anchor._from.colOff,
                    row=anchor._from.row,
                    rowOff=anchor._from.rowOff
                )
                new_img.anchor = new_anchor

            target_ws.add_image(new_img)


def main():
    root = tk.Tk()

    # Style
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Treeview", rowheight=28, font=("Segoe UI", 10))
    style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))

    app = ExcelEditorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
