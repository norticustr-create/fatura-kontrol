"""
Supplier Invoice Auditor - Windows GUI Application.

A lightweight, modern, native Windows graphical user interface for
Turkish e-invoice (UBL 2.1 XML) auditing and Excel reporting.
"""

import ctypes
import os
from pathlib import Path
import subprocess
import sys
import threading
from typing import Optional, List, Dict, Any

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from batch import collect_files, process_batch, batch_summary
from excel_export import export_to_excel


def set_dpi_awareness():
    """Enables Per-Monitor DPI awareness on Windows for crisp font rendering."""
    if sys.platform == "win32":
        try:
            # PROCESS_PER_MONITOR_DPI_AWARE (1)
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass


class InvoiceAuditorGUI:
    """Main graphical interface for Supplier Invoice Auditor."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Tedarikçi Fatura Denetçisi (UBL 2.1 e-Fatura)")
        self.root.geometry("700x580")
        self.root.minsize(640, 520)

        # State variables
        self.source_var = tk.StringVar()
        self.output_var = tk.StringVar(value="invoice_audit_report.xlsx")
        self.status_var = tk.StringVar(value="Denetim için dosya, klasör veya ZIP arşivi seçin.")
        self.is_processing = False
        self.last_saved_path: Optional[str] = None

        # Summary count variables
        self.total_count_var = tk.StringVar(value="0")
        self.pass_count_var = tk.StringVar(value="0")
        self.review_count_var = tk.StringVar(value="0")
        self.error_count_var = tk.StringVar(value="0")

        self._configure_styles()
        self._build_ui()

    def _configure_styles(self):
        """Sets up ttk styles and fonts."""
        style = ttk.Style()
        available_themes = style.theme_names()
        if "vista" in available_themes:
            style.theme_use("vista")
        elif "clam" in available_themes:
            style.theme_use("clam")

        # Configure custom button styles
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), padding=6)
        style.configure("Action.TButton", font=("Segoe UI", 9), padding=4)
        style.configure("Header.TLabel", font=("Segoe UI", 13, "bold"), foreground="#0F172A")
        style.configure("SubHeader.TLabel", font=("Segoe UI", 9), foreground="#475569")
        style.configure("Section.TLabel", font=("Segoe UI", 9, "bold"), foreground="#1E293B")
        style.configure("Status.TLabel", font=("Segoe UI", 9), foreground="#334155")

    def _build_ui(self):
        """Constructs all layout frames and widgets."""
        main_container = ttk.Frame(self.root, padding=16)
        main_container.pack(fill=tk.BOTH, expand=True)

        # 1. Header Frame
        header_frame = ttk.Frame(main_container)
        header_frame.pack(fill=tk.X, pady=(0, 12))

        title_lbl = ttk.Label(
            header_frame,
            text="e-Fatura Toplu Denetim & Excel Raporlama",
            style="Header.TLabel",
        )
        title_lbl.pack(anchor=tk.W)

        subtitle_lbl = ttk.Label(
            header_frame,
            text="UBL 2.1 XML faturalarını doğrulayın, KDV ve ödenecek tutar matematiksel denetimlerini yapın.",
            style="SubHeader.TLabel",
        )
        subtitle_lbl.pack(anchor=tk.W, pady=(2, 0))

        sep1 = ttk.Separator(main_container, orient=tk.HORIZONTAL)
        sep1.pack(fill=tk.X, pady=(0, 12))

        # 2. Source Selection Frame
        src_label = ttk.Label(main_container, text="1. Girdi Kaynağı (XML / ZIP / Klasör):", style="Section.TLabel")
        src_label.pack(anchor=tk.W, pady=(0, 4))

        src_entry_frame = ttk.Frame(main_container)
        src_entry_frame.pack(fill=tk.X, pady=(0, 4))

        self.src_entry = ttk.Entry(src_entry_frame, textvariable=self.source_var, font=("Segoe UI", 9))
        self.src_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

        src_btn_frame = ttk.Frame(main_container)
        src_btn_frame.pack(fill=tk.X, pady=(0, 10))

        self.btn_select_file = ttk.Button(
            src_btn_frame,
            text="📄 Dosya Seç (XML / ZIP)",
            style="Action.TButton",
            command=self._choose_source_file,
        )
        self.btn_select_file.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_select_folder = ttk.Button(
            src_btn_frame,
            text="📁 Klasör Seç",
            style="Action.TButton",
            command=self._choose_source_folder,
        )
        self.btn_select_folder.pack(side=tk.LEFT)

        # 3. Output Selection Frame
        out_label = ttk.Label(main_container, text="2. Çıktı Raporu (Excel Dosyası):", style="Section.TLabel")
        out_label.pack(anchor=tk.W, pady=(0, 4))

        out_entry_frame = ttk.Frame(main_container)
        out_entry_frame.pack(fill=tk.X, pady=(0, 14))

        self.out_entry = ttk.Entry(out_entry_frame, textvariable=self.output_var, font=("Segoe UI", 9))
        self.out_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

        self.btn_select_output = ttk.Button(
            out_entry_frame,
            text="Değiştir...",
            style="Action.TButton",
            command=self._choose_output_file,
        )
        self.btn_select_output.pack(side=tk.LEFT)

        # 4. Action & Progress Frame
        self.btn_start = ttk.Button(
            main_container,
            text="🚀 Faturaları Denetle ve Raporla",
            style="Primary.TButton",
            command=self.start_processing,
        )
        self.btn_start.pack(fill=tk.X, pady=(0, 10))

        self.progress_bar = ttk.Progressbar(main_container, orient=tk.HORIZONTAL, mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=(0, 4))

        self.status_label = ttk.Label(
            main_container,
            textvariable=self.status_var,
            style="Status.TLabel",
        )
        self.status_label.pack(anchor=tk.W, pady=(0, 12))

        # 5. Summary Dashboard Cards Frame
        dash_label = ttk.Label(main_container, text="Denetim Özeti:", style="Section.TLabel")
        dash_label.pack(anchor=tk.W, pady=(0, 4))

        cards_frame = ttk.Frame(main_container)
        cards_frame.pack(fill=tk.X, pady=(0, 14))
        for col in range(4):
            cards_frame.columnconfigure(col, weight=1)

        # Total Card
        self._build_card(cards_frame, 0, "TOPLAM", self.total_count_var, "#1E293B", "#F8FAFC")
        # Pass Card
        self._build_card(cards_frame, 1, "BAŞARILI (PASS)", self.pass_count_var, "#15803D", "#F0FDF4")
        # Review Card
        self._build_card(cards_frame, 2, "İNCELEME (REVIEW)", self.review_count_var, "#B45309", "#FFFBEB")
        # Error Card
        self._build_card(cards_frame, 3, "HATALI (ERROR)", self.error_count_var, "#B91C1C", "#FEF2F2")

        # 6. Post-Process Action Buttons Frame
        post_btn_frame = ttk.Frame(main_container)
        post_btn_frame.pack(fill=tk.X, pady=(4, 0))

        self.btn_open_excel = ttk.Button(
            post_btn_frame,
            text="📊 Excel Raporunu Aç",
            style="Action.TButton",
            state=tk.DISABLED,
            command=self.open_excel_report,
        )
        self.btn_open_excel.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_open_folder = ttk.Button(
            post_btn_frame,
            text="📁 Klasörde Göster",
            style="Action.TButton",
            state=tk.DISABLED,
            command=self.open_output_folder,
        )
        self.btn_open_folder.pack(side=tk.LEFT)

    def _build_card(
        self,
        parent: ttk.Frame,
        col: int,
        title: str,
        var: tk.StringVar,
        accent_color: str,
        bg_color: str,
    ):
        """Constructs an individual colored metric card."""
        card = tk.Frame(parent, bg=bg_color, highlightbackground="#CBD5E1", highlightthickness=1, padx=10, pady=8)
        card.grid(row=0, column=col, padx=4, sticky="nsew")

        lbl_title = tk.Label(
            card,
            text=title,
            font=("Segoe UI", 8, "bold"),
            fg=accent_color,
            bg=bg_color,
        )
        lbl_title.pack(anchor=tk.CENTER)

        lbl_num = tk.Label(
            card,
            textvariable=var,
            font=("Segoe UI", 16, "bold"),
            fg=accent_color,
            bg=bg_color,
        )
        lbl_num.pack(anchor=tk.CENTER, pady=(2, 0))

    def _choose_source_file(self):
        """Prompts user to select an XML or ZIP file."""
        filetypes = [
            ("e-Fatura Dosyaları (XML, ZIP)", "*.xml;*.zip"),
            ("UBL 2.1 XML Faturaları", "*.xml"),
            ("ZIP Arşivleri", "*.zip"),
            ("Tüm Dosyalar", "*.*"),
        ]
        chosen = filedialog.askopenfilename(
            title="e-Fatura XML veya ZIP Dosyası Seçin",
            filetypes=filetypes,
        )
        if chosen:
            self.source_var.set(chosen)
            self._suggest_output_path(chosen)

    def _choose_source_folder(self):
        """Prompts user to select a folder of XML files."""
        chosen = filedialog.askdirectory(title="XML Faturaları İçeren Klasörü Seçin")
        if chosen:
            self.source_var.set(chosen)
            self._suggest_output_path(chosen)

    def _choose_output_file(self):
        """Prompts user to choose or change the output Excel report path."""
        filetypes = [("Excel Çalışma Kitabı", "*.xlsx"), ("Tüm Dosyalar", "*.*")]
        current_out = self.output_var.get().strip() or "invoice_audit_report.xlsx"
        chosen = filedialog.asksaveasfilename(
            title="Excel Raporu Kayıt Yeri Seçin",
            initialfile=Path(current_out).name,
            initialdir=str(Path(current_out).parent) if Path(current_out).parent.exists() else None,
            defaultextension=".xlsx",
            filetypes=filetypes,
        )
        if chosen:
            if not chosen.lower().endswith(".xlsx"):
                chosen += ".xlsx"
            self.output_var.set(chosen)

    def _suggest_output_path(self, source_path_str: str):
        """Intelligently updates the default output file to match the source folder."""
        try:
            src = Path(source_path_str)
            if src.suffix.lower() in [".xml", ".zip"] or src.is_file():
                parent_dir = src.parent
                default_name = f"{src.stem}_denetim_raporu.xlsx"
            else:
                parent_dir = src
                default_name = "fatura_denetim_raporu.xlsx"
            suggested = parent_dir / default_name
            self.output_var.set(str(suggested))
        except Exception:
            pass

    def start_processing(self):
        """Validates inputs and triggers background audit processing."""
        if self.is_processing:
            return

        source = self.source_var.get().strip()
        if not source:
            messagebox.showwarning(
                "Kaynak Eksik",
                "Lütfen denetlenecek bir XML dosyası, ZIP arşivi veya klasör seçin.",
            )
            return

        src_path = Path(source)
        if not src_path.exists():
            messagebox.showerror(
                "Kaynak Bulunamadı",
                f"Belirtilen girdi kaynağı mevcut değil:\n\n{source}",
            )
            return

        output_path = self.output_var.get().strip()
        if not output_path:
            output_path = "invoice_audit_report.xlsx"
            self.output_var.set(output_path)

        if not output_path.lower().endswith(".xlsx"):
            output_path += ".xlsx"
            self.output_var.set(output_path)

        # Reset counts and state
        self.is_processing = True
        self.btn_start.config(state=tk.DISABLED)
        self.btn_open_excel.config(state=tk.DISABLED)
        self.btn_open_folder.config(state=tk.DISABLED)
        self.total_count_var.set("0")
        self.pass_count_var.set("0")
        self.review_count_var.set("0")
        self.error_count_var.set("0")
        self.progress_bar["value"] = 0
        self.status_var.set("Faturalar taranıyor...")

        # Run audit in background thread
        thread = threading.Thread(
            target=self._process_worker,
            args=(source, output_path),
            daemon=True,
        )
        thread.start()

    def _process_worker(self, source: str, output_path: str):
        """Worker thread executing collection, parsing, auditing, and export."""
        try:
            # 1. Collect files
            files = collect_files(source)
            if not files:
                self.root.after(
                    0,
                    self._on_error,
                    "Fatura Bulunamadı",
                    "Seçilen kaynakta geçerli bir e-fatura XML dosyası bulunamadı.",
                )
                return

            total_files = len(files)
            self.root.after(0, self._on_collection_done, total_files)

            # 2. Process batch with callback
            def progress_callback(cur_idx: int, tot: int, res: Dict[str, Any]):
                self.root.after(0, self._on_progress_update, cur_idx, tot, res)

            results = process_batch(files, progress_callback=progress_callback)
            summary = batch_summary(results)

            # 3. Export to Excel
            saved_file = export_to_excel(results, output_path)

            # 4. Notify completion
            self.root.after(0, self._on_complete, saved_file, summary)

        except PermissionError:
            self.root.after(
                0,
                self._on_error,
                "Dosya Erişim Hatası",
                f"Excel raporu kaydedilemedi!\n\nDosya başka bir programda (örneğin Excel) açık olabilir. "
                f"Lütfen dosyayı kapatıp tekrar deneyin:\n\n{output_path}",
            )
        except Exception as e:
            self.root.after(
                0,
                self._on_error,
                "İşlem Hatası",
                f"Denetim sırasında beklenmeyen bir hata oluştu:\n\n{type(e).__name__}: {str(e)}",
            )

    def _on_collection_done(self, total_files: int):
        """Updates progress bar max value after files are discovered."""
        self.progress_bar["maximum"] = total_files
        self.progress_bar["value"] = 0
        self.total_count_var.set(str(total_files))
        self.status_var.set(f"{total_files} adet fatura bulundu. Denetim başlıyor...")

    def _on_progress_update(self, cur_idx: int, total_files: int, result: Dict[str, Any]):
        """Updates progress bar and live counters as each invoice finishes."""
        self.progress_bar["value"] = cur_idx
        fname = result.get("original_name") or result.get("file_name") or "Fatura"
        status = result.get("status", "")
        self.status_var.set(f"İşleniyor ({cur_idx}/{total_files}): {fname} [{status}]")

    def _on_complete(self, saved_file: Any, summary: Dict[str, int]):
        """Handles successful processing finish."""
        self.is_processing = False
        self.last_saved_path = str(saved_file)
        self.btn_start.config(state=tk.NORMAL)
        self.btn_open_excel.config(state=tk.NORMAL)
        self.btn_open_folder.config(state=tk.NORMAL)

        self.total_count_var.set(str(summary["total"]))
        self.pass_count_var.set(str(summary["pass"]))
        self.review_count_var.set(str(summary["review"]))
        self.error_count_var.set(str(summary["error"]))

        status_text = (
            f"Tamamlandı! Toplam: {summary['total']} | "
            f"Başarılı: {summary['pass']} | İnceleme: {summary['review']} | Hatalı: {summary['error']}"
        )
        self.status_var.set(status_text)

        messagebox.showinfo(
            "Denetim Tamamlandı",
            f"Fatura denetimi başarıyla tamamlandı.\n\n"
            f"Toplam Fatura : {summary['total']}\n"
            f"Başarılı (PASS) : {summary['pass']}\n"
            f"İnceleme (REVIEW) : {summary['review']}\n"
            f"Hatalı (ERROR) : {summary['error']}\n\n"
            f"Rapor Kaydedildi:\n{saved_file}",
        )

    def _on_error(self, title: str, message: str):
        """Handles processing errors cleanly on the UI thread."""
        self.is_processing = False
        self.btn_start.config(state=tk.NORMAL)
        self.status_var.set(f"Hata: {title}")
        messagebox.showerror(title, message)

    def open_excel_report(self):
        """Opens the generated Excel file using default system spreadsheet viewer."""
        if not self.last_saved_path or not Path(self.last_saved_path).exists():
            messagebox.showwarning("Rapor Yok", "Açılacak bir Excel raporu bulunamadı.")
            return

        try:
            if sys.platform == "win32":
                os.startfile(self.last_saved_path)
            elif sys.platform == "darwin":
                subprocess.run(["open", self.last_saved_path])
            else:
                subprocess.run(["xdg-open", self.last_saved_path])
        except Exception as e:
            messagebox.showerror("Açma Hatası", f"Excel raporu açılamadı:\n{e}")

    def open_output_folder(self):
        """Reveals the generated Excel file in Windows Explorer."""
        if not self.last_saved_path or not Path(self.last_saved_path).exists():
            messagebox.showwarning("Klasör Yok", "Rapor klasörü bulunamadı.")
            return

        try:
            folder = Path(self.last_saved_path).resolve().parent
            if sys.platform == "win32":
                subprocess.run(["explorer.exe", f"/select,{str(Path(self.last_saved_path).resolve())}"])
            elif sys.platform == "darwin":
                subprocess.run(["open", "-R", self.last_saved_path])
            else:
                subprocess.run(["xdg-open", str(folder)])
        except Exception as e:
            messagebox.showerror("Klasör Hatası", f"Klasör açılamadı:\n{e}")


def launch_gui():
    """Application entry point for GUI."""
    set_dpi_awareness()
    root = tk.Tk()
    app = InvoiceAuditorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    launch_gui()
