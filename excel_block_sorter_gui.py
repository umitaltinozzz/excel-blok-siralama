from __future__ import annotations

from pathlib import Path
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from excel_block_sorter import (
    DEFAULT_OUTPUT_DIR,
    SortConfig,
    columns_to_indices,
    process_files,
    result_line,
)


def default_desktop_output_dir() -> Path:
    desktop = Path.home() / "Desktop"
    if desktop.exists():
        return desktop / "Excel_Sirali_Ciktilar"
    return Path.cwd() / DEFAULT_OUTPUT_DIR


class ExcelBlockSorterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel Blok Siralama")
        self.geometry("900x620")
        self.minsize(820, 560)

        self.selected_paths: list[Path] = []
        self.log_queue: queue.Queue[str] = queue.Queue()
        self.worker: threading.Thread | None = None

        self.sort_column_var = tk.StringVar(value="J")
        self.first_row_var = tk.StringVar(value="auto")
        self.sheet_var = tk.StringVar(value="")
        self.output_var = tk.StringVar(value=str(default_desktop_output_dir()))
        self.direction_var = tk.StringVar(value="desc")
        self.suffix_var = tk.StringVar(value="_sirali")
        self.output_name_var = tk.StringVar(value="")

        self._build_ui()
        self.after(100, self._poll_log_queue)

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        header = ttk.Frame(self, padding=(18, 16, 18, 8))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        title = ttk.Label(header, text="Excel Blok Siralama", font=("Segoe UI", 16, "bold"))
        title.grid(row=0, column=0, sticky="w")
        subtitle = ttk.Label(
            header,
            text="Toplam satirlarini sabit tutar, alt detay satirlarini secilen kolona gore siralar.",
        )
        subtitle.grid(row=1, column=0, sticky="w", pady=(4, 0))

        body = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        body.grid(row=1, column=0, sticky="nsew", padx=18, pady=8)

        left = ttk.Frame(body, padding=12)
        right = ttk.Frame(body, padding=12)
        body.add(left, weight=2)
        body.add(right, weight=3)

        self._build_file_panel(left)
        self._build_settings_panel(right)

        footer = ttk.Frame(self, padding=(18, 8, 18, 16))
        footer.grid(row=2, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(footer, mode="indeterminate")
        self.progress.grid(row=0, column=0, sticky="ew", padx=(0, 12))

        self.run_button = ttk.Button(footer, text="Calistir", command=self.run_sorting)
        self.run_button.grid(row=0, column=1, sticky="e")

    def _build_file_panel(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        label = ttk.Label(parent, text="Excel secimi", font=("Segoe UI", 11, "bold"))
        label.grid(row=0, column=0, sticky="w")

        list_frame = ttk.Frame(parent)
        list_frame.grid(row=1, column=0, sticky="nsew", pady=(8, 8))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.path_list = tk.Listbox(list_frame, height=12, activestyle="none")
        self.path_list.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.path_list.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.path_list.configure(yscrollcommand=scrollbar.set)

        buttons = ttk.Frame(parent)
        buttons.grid(row=2, column=0, sticky="ew")
        buttons.columnconfigure((0, 1, 2), weight=1)

        ttk.Button(buttons, text="Dosya ekle", command=self.add_files).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(buttons, text="Klasor ekle", command=self.add_folder).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(buttons, text="Temizle", command=self.clear_paths).grid(row=0, column=2, sticky="ew", padx=(6, 0))

        output_frame = ttk.LabelFrame(parent, text="Cikti klasoru", padding=10)
        output_frame.grid(row=3, column=0, sticky="ew", pady=(14, 0))
        output_frame.columnconfigure(0, weight=1)
        ttk.Entry(output_frame, textvariable=self.output_var).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(output_frame, text="Masaustu", command=self.use_desktop_output).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(output_frame, text="Sec", command=self.choose_output).grid(row=0, column=2)

    def _build_settings_panel(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)

        settings = ttk.LabelFrame(parent, text="Ayarlar", padding=12)
        settings.grid(row=0, column=0, sticky="ew")
        settings.columnconfigure(1, weight=1)
        settings.columnconfigure(3, weight=1)

        ttk.Label(settings, text="Siralama kolonu").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
        ttk.Entry(settings, textvariable=self.sort_column_var, width=14).grid(row=0, column=1, sticky="w", pady=5)

        ttk.Label(settings, text="Sayfa adi").grid(row=0, column=2, sticky="w", padx=(18, 8), pady=5)
        ttk.Entry(settings, textvariable=self.sheet_var).grid(row=0, column=3, sticky="ew", pady=5)

        ttk.Label(settings, text="Yon").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=5)
        direction = ttk.Frame(settings)
        direction.grid(row=1, column=1, columnspan=3, sticky="w", pady=5)
        ttk.Radiobutton(direction, text="Buyukten kucuge", variable=self.direction_var, value="desc").grid(row=0, column=0)
        ttk.Radiobutton(direction, text="Kucukten buyuge", variable=self.direction_var, value="asc").grid(row=0, column=1, padx=(18, 0))

        ttk.Label(settings, text="Yeni dosya adi").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=5)
        ttk.Entry(settings, textvariable=self.output_name_var).grid(row=2, column=1, columnspan=3, sticky="ew", pady=5)

        ttk.Label(settings, text="Coklu dosya eki").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=5)
        ttk.Entry(settings, textvariable=self.suffix_var, width=14).grid(row=3, column=1, sticky="w", pady=5)

        help_box = ttk.LabelFrame(parent, text="Calisma mantigi", padding=12)
        help_box.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        help_text = (
            "1. Basliklari atlayip ilk gercek toplam satirini kendisi bulur.\n"
            "2. Siralama kolonu tek veya coklu yazilabilir: J ya da J,K.\n"
            "3. Toplam bloklarini toplam degerine gore siralar; Genel Toplam altta kalir.\n"
            "4. Her toplam blogunun alt detay satirlari da secilen kolonlara gore siralanir.\n"
            "5. Dosyadaki son tablo kolonunu otomatik bulur ve satiri komple tasir."
        )
        ttk.Label(help_box, text=help_text, justify="left").grid(row=0, column=0, sticky="w")

        log_frame = ttk.LabelFrame(parent, text="Islem gunlugu", padding=8)
        log_frame.grid(row=2, column=0, sticky="nsew", pady=(12, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_text = tk.Text(log_frame, height=10, wrap="word", state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        log_scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        log_scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=log_scrollbar.set)

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Excel dosyasi sec",
            filetypes=[("Excel dosyalari", "*.xlsx"), ("Tum dosyalar", "*.*")],
        )
        self._add_paths(Path(path) for path in paths)

    def add_folder(self):
        folder = filedialog.askdirectory(title="Excel klasoru sec")
        if folder:
            self._add_paths([Path(folder)])

    def clear_paths(self):
        self.selected_paths.clear()
        self.path_list.delete(0, tk.END)

    def choose_output(self):
        folder = filedialog.askdirectory(title="Cikti klasoru sec")
        if folder:
            self.output_var.set(folder)

    def use_desktop_output(self):
        self.output_var.set(str(default_desktop_output_dir()))

    def _add_paths(self, paths):
        existing = {path.resolve() for path in self.selected_paths}
        for path in paths:
            resolved = path.resolve()
            if resolved in existing:
                continue
            self.selected_paths.append(path)
            existing.add(resolved)
            self.path_list.insert(tk.END, str(path))

    def validate_config(self) -> tuple[SortConfig, Path, str, str | None] | None:
        if not self.selected_paths:
            messagebox.showerror("Eksik secim", "En az bir Excel dosyasi veya klasoru secin.")
            return None

        try:
            sort_columns = columns_to_indices(self.sort_column_var.get())
        except ValueError as exc:
            messagebox.showerror("Hatali kolon", str(exc))
            return None

        output_name = self.output_name_var.get().strip() or None
        if output_name and len(self.selected_paths) != 1:
            messagebox.showerror("Hatali dosya adi", "Yeni dosya adi sadece tek Excel dosyasi secildiginde kullanilir.")
            return None
        if output_name and self.selected_paths[0].is_dir():
            messagebox.showerror("Hatali dosya adi", "Klasor seciminde yeni dosya adi yerine coklu dosya eki kullanilir.")
            return None

        output_dir = Path(self.output_var.get().strip() or DEFAULT_OUTPUT_DIR)
        suffix = self.suffix_var.get().strip() or "_sirali"
        config = SortConfig(
            sort_column=sort_columns[0],
            first_row=None,
            last_column=None,
            descending=self.direction_var.get() == "desc",
            sheet_name=self.sheet_var.get().strip() or None,
            sort_columns=sort_columns,
        )
        return config, output_dir, suffix, output_name

    def run_sorting(self):
        validated = self.validate_config()
        if validated is None:
            return

        config, output_dir, suffix, output_name = validated
        self.run_button.configure(state="disabled")
        self.progress.start(12)
        self._clear_log()
        self._log("Islem basladi.")

        self.worker = threading.Thread(
            target=self._run_sorting_worker,
            args=(list(self.selected_paths), output_dir, config, suffix, output_name),
            daemon=True,
        )
        self.worker.start()

    def _run_sorting_worker(self, paths, output_dir, config, suffix, output_name):
        try:
            total_results = 0
            skipped_results = 0
            for path in paths:
                self.log_queue.put(f"Taraniyor: {path}")
                selected_output_name = output_name if len(paths) == 1 and path.is_file() else None
                results = process_files(
                    path,
                    output_dir,
                    config,
                    suffix=suffix,
                    output_name=selected_output_name,
                )
                if not results:
                    self.log_queue.put(f"SKIP {path}: islenecek .xlsx dosyasi bulunamadi")
                    skipped_results += 1
                    continue
                for result in results:
                    total_results += 1
                    if result.skipped:
                        skipped_results += 1
                    self.log_queue.put(result_line(result, dry_run=False))
            self.log_queue.put(f"DONE|{total_results}|{skipped_results}")
        except Exception as exc:
            self.log_queue.put(f"ERROR|{exc}")

    def _poll_log_queue(self):
        while True:
            try:
                message = self.log_queue.get_nowait()
            except queue.Empty:
                break

            if message.startswith("DONE|"):
                _, total, skipped = message.split("|", 2)
                self._finish_run(int(total), int(skipped))
            elif message.startswith("ERROR|"):
                self._finish_run(0, 1, message.split("|", 1)[1])
            else:
                self._log(message)

        self.after(100, self._poll_log_queue)

    def _finish_run(self, total, skipped, error=None):
        self.progress.stop()
        self.run_button.configure(state="normal")
        if error:
            self._log(f"HATA: {error}")
            messagebox.showerror("Islem hatasi", error)
            return

        self._log(f"Islem bitti. Toplam: {total}, sorunlu/atlanmis: {skipped}.")
        if skipped:
            messagebox.showwarning("Islem tamamlandi", f"Islem bitti. {skipped} dosya/blok atlandi.")
        else:
            messagebox.showinfo("Islem tamamlandi", "Tum dosyalar basariyla islendi.")

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state="disabled")

    def _log(self, message):
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")


def main():
    app = ExcelBlockSorterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
