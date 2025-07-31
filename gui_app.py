import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from downloader import PDFDownloader
from converter import PDFConverter
import webbrowser

class NSSDecisionDownloader(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("NSS Decision Downloader")
        self.geometry("650x700")

        # Variables to store file paths
        self.input_file_path = tk.StringVar()
        self.destination_dir = tk.StringVar()

        # File type (xlsx or xml) detected automatically from the input file.
        # Start with no default value until a file is selected.
        self.file_type = tk.StringVar(value="")
        self.file_type_label_var = tk.StringVar(value="")
        # Update file type when the input path changes (e.g., user edits the
        # entry manually)
        self.input_file_path.trace_add(
            "write", lambda *_: self.detect_file_type())

        # Excel filtering option enabled by default
        self.use_excel_filter = tk.BooleanVar()
        self.use_excel_filter.set(True)

        # Filtering options

        # Decision type filtering options removed from the GUI.
        # (The variables remain defined but will not be used.)
        self.use_decision_type_filter = tk.BooleanVar()
        self.decision_type = tk.StringVar(value="Meritorní")

        # Flags to track if the processes should be running or stopped
        self.downloading = False
        self.converting = False

        # Progress bar
        self.progress_bar = ttk.Progressbar(self, orient="horizontal", mode="determinate", length=500)
        self.progress_bar.pack(pady=10)

        # Label with link to the author's website
        credit_label = tk.Label(self, text="Made by Oldřich Tristan Florian", fg="blue", cursor="hand2")
        credit_label.pack(pady=10)
        credit_label.bind("<Button-1>", lambda e: self.open_link("https://otflorian.com"))

        # UI Elements for selecting files and folders
        tk.Label(self, text="Select Input File (XLSX open data or XML export)").pack(pady=5)
        tk.Entry(self, textvariable=self.input_file_path, width=60).pack(pady=5)
        tk.Button(self, text="Browse", command=self.browse_input_file).pack(pady=5)

        # Display the detected file type
        self.detected_type_label = tk.Label(
            self, textvariable=self.file_type_label_var)
        self.detected_type_label.pack(pady=5)

        # Notice label for Excel filtering status
        self.filter_notice_var = tk.StringVar(value="")
        self.filter_notice_label = tk.Label(
            self, textvariable=self.filter_notice_var)
        self.filter_notice_label.pack(pady=2)

        tk.Label(self, text="Select Destination Folder").pack(pady=5)
        tk.Entry(self, textvariable=self.destination_dir, width=60).pack(pady=5)
        tk.Button(self, text="Browse", command=self.browse_destination).pack(pady=5)

        # Buttons for starting/stopping download and conversion
        button_frame = tk.Frame(self)
        button_frame.pack(pady=20)
        self.download_button = tk.Button(button_frame, text="Start Download", command=self.toggle_download)
        self.download_button.pack(side=tk.LEFT, padx=10)
        self.convert_button = tk.Button(button_frame, text="Start Conversion", command=self.toggle_conversion)
        self.convert_button.pack(side=tk.LEFT, padx=10)

        # Frame to hold the text box and scrollbar for progress output
        text_frame = tk.Frame(self)
        text_frame.pack(pady=10, fill=tk.BOTH, expand=True)
        self.progress_text = tk.Text(text_frame, height=20, width=70, wrap=tk.WORD)
        self.progress_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = tk.Scrollbar(text_frame, command=self.progress_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.progress_text['yscrollcommand'] = scrollbar.set

        # Summary storage
        self.summary = {"downloaded": 0, "skipped": 0, "replaced": 0, "failed": 0}

        # Initialize the UI state
        self.update_file_type()

    def open_link(self, url):
        """Open the given URL in the default web browser"""
        webbrowser.open_new(url)

    def browse_input_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Supported files", "*.xlsx *.xml"), ("All files", "*.*")])
        if file_path:
            self.input_file_path.set(file_path)
            # detect_file_type will be triggered by trace on input_file_path
            self.detect_file_type()

    def browse_destination(self):
        directory = filedialog.askdirectory()
        self.destination_dir.set(directory)

    def update_file_type(self):
        """Adjust UI elements based on selected file type."""
        if self.file_type.get() == "xml":
            self.use_excel_filter.set(False)
            self.file_type_label_var.set("File Type: Search Results Export")
            self.filter_notice_var.set("")
        elif self.file_type.get() == "xlsx":
            self.use_excel_filter.set(True)
            self.file_type_label_var.set("File Type: Open Data")
            self.filter_notice_var.set("Excel filtering is applied")
        else:
            self.use_excel_filter.set(False)
            self.file_type_label_var.set("")
            self.filter_notice_var.set("")

    def detect_file_type(self, *_):
        """Detect file type from the input file extension."""
        path = self.input_file_path.get().lower()
        if path.endswith(".xml"):
            self.file_type.set("xml")
        elif path.endswith(".xlsx"):
            self.file_type.set("xlsx")
        else:
            self.file_type.set("")
        self.update_file_type()

    def toggle_download(self):
        if self.downloading:
            self.downloading = False
            self.download_button.config(text="Start Download")
            self.progress_text.insert(tk.END, "\nDownload process has been stopped.\n")
            self.show_summary()
        else:
            self.downloading = True
            self.download_button.config(text="Stop Download")
            self.start_download_thread()

    def toggle_conversion(self):
        if self.converting:
            self.converting = False
            self.convert_button.config(text="Start Conversion")
            self.progress_text.insert(tk.END, "\nConversion process has been stopped.\n")
            self.show_summary()
        else:
            self.converting = True
            self.convert_button.config(text="Stop Conversion")
            self.start_conversion_thread()

    def start_download_thread(self):
        thread = threading.Thread(target=self.start_download)
        thread.start()

    def start_conversion_thread(self):
        thread = threading.Thread(target=self.start_conversion)
        thread.start()

    def update_download_progress(self, current, total, file_name, status):
        if not self.downloading:
            return
        self.progress_bar["value"] = (current / total) * 100
        self.progress_text.insert(tk.END, f"{file_name}: {status.capitalize()}\n")
        self.progress_text.yview(tk.END)
        self.update_idletasks()
        self.summary[status] += 1

    def update_conversion_progress(self, current, total, file_name, status):
        if not self.converting:
            return
        self.progress_bar["value"] = (current / total) * 100
        self.progress_text.insert(tk.END, f"{file_name}: {status.capitalize()}\n")
        self.progress_text.yview(tk.END)
        self.update_idletasks()
        self.summary[status] += 1

    def start_download(self):
        if not self.input_file_path.get() or not self.destination_dir.get():
            messagebox.showwarning("Input Error", "Please select both an input file and a destination folder.")
            return

        self.progress_text.delete(1.0, tk.END)
        self.progress_text.insert(tk.END, "Starting download...\n")
        self.progress_bar["value"] = 0

        # Decision Type filtering has been removed; always pass None.
        decision_type = None
        use_excel_filter = self.use_excel_filter.get() if self.file_type.get() == "xlsx" else False

        self.summary = {"downloaded": 0, "skipped": 0, "replaced": 0, "failed": 0}

        downloader = PDFDownloader(
            self.input_file_path.get(),
            self.destination_dir.get(),
            self.update_download_progress,
            decision_type=decision_type,
            use_excel_filter=use_excel_filter,
            stop_callback=lambda: not self.downloading,
            file_type=self.file_type.get()
        )
        summary = downloader.download_pdfs()

        if not self.downloading:
            return

        self.progress_text.insert(tk.END, "\nDownload process has finished.\n")
        self.progress_bar["value"] = 100
        self.show_summary()
        self.downloading = False
        self.download_button.config(text="Start Download")

    def start_conversion(self):
        if not self.destination_dir.get():
            messagebox.showwarning("Input Error", "Please select a destination folder.")
            return

        pdf_dir = os.path.join(self.destination_dir.get(), "pdf")
        txt_dir = os.path.join(self.destination_dir.get(), "txt")

        self.progress_text.delete(1.0, tk.END)
        self.progress_text.insert(tk.END, "Starting conversion...\n")
        self.progress_bar["value"] = 0

        self.summary = {"converted": 0, "skipped": 0, "failed": 0}

        converter = PDFConverter(
            pdf_dir,
            txt_dir,
            self.update_conversion_progress,
            stop_callback=lambda: not self.converting
        )
        summary = converter.convert_pdfs()

        if not self.converting:
            return

        self.progress_text.insert(tk.END, "\nConversion process has finished.\n")
        self.progress_bar["value"] = 100
        self.show_summary()
        self.converting = False
        self.convert_button.config(text="Start Conversion")

    def show_summary(self):
        self.progress_text.insert(tk.END, f"\nSummary:\n")
        self.progress_text.yview(tk.END)
        if self.converting:
            self.progress_text.insert(tk.END, f"Converted: {self.summary.get('converted', 0)}\n")
            self.progress_text.insert(tk.END, f"Skipped: {self.summary.get('skipped', 0)}\n")
            self.progress_text.insert(tk.END, f"Failed: {self.summary.get('failed', 0)}\n")
        else:
            self.progress_text.insert(tk.END, f"Downloaded: {self.summary.get('downloaded', 0)}\n")
            self.progress_text.insert(tk.END, f"Skipped: {self.summary.get('skipped', 0)}\n")
            self.progress_text.insert(tk.END, f"Replaced: {self.summary.get('replaced', 0)}\n")
            self.progress_text.insert(tk.END, f"Failed: {self.summary.get('failed', 0)}\n")
        self.progress_text.yview(tk.END)

if __name__ == "__main__":
    app = NSSDecisionDownloader()
    app.mainloop()
