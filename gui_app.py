import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkcalendar import DateEntry
from downloader import PDFDownloader
from converter import PDFConverter
import webbrowser

class NSSDecisionDownloader(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("NSS Decision Downloader")
        self.geometry("650x700")

        # Variables to store file paths
        self.xlsx_file_path = tk.StringVar()
        self.destination_dir = tk.StringVar()

        # Excel filtering option enabled by default
        self.use_excel_filter = tk.BooleanVar()
        self.use_excel_filter.set(True)

        # Filtering options
        self.use_date_filter = tk.BooleanVar()
        self.start_date = "2003-01-01"  # Default start date
        self.end_date = tk.StringVar()

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
        tk.Label(self, text="Select Excel File (Otevřená data k soudní činnosti)").pack(pady=5)
        tk.Entry(self, textvariable=self.xlsx_file_path, width=60).pack(pady=5)
        tk.Button(self, text="Browse", command=self.browse_xlsx).pack(pady=5)

        tk.Label(self, text="Select Destination Folder").pack(pady=5)
        tk.Entry(self, textvariable=self.destination_dir, width=60).pack(pady=5)
        tk.Button(self, text="Browse", command=self.browse_destination).pack(pady=5)

        # Date Filter with checkbox and date fields on the same line
        date_filter_frame = tk.Frame(self)
        date_filter_frame.pack(pady=5, fill=tk.X)
        date_filter_inner_frame = tk.Frame(date_filter_frame)
        date_filter_inner_frame.pack(anchor=tk.CENTER, expand=True)
        tk.Checkbutton(date_filter_inner_frame, variable=self.use_date_filter, command=self.toggle_date_filter).pack(side=tk.LEFT, padx=5)
        tk.Label(date_filter_inner_frame, text="Start Date:").pack(side=tk.LEFT)
        self.start_date_entry = DateEntry(date_filter_inner_frame, date_pattern='yyyy-mm-dd',
                                          year=2003, month=1, day=1)
        self.start_date_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(date_filter_inner_frame, text="End Date:").pack(side=tk.LEFT)
        self.end_date_entry = DateEntry(date_filter_inner_frame, textvariable=self.end_date, date_pattern='yyyy-mm-dd')
        self.end_date_entry.pack(side=tk.LEFT, padx=5)

        # Checkbox for using Excel filtering (enabled by default)
        tk.Checkbutton(self, text="Use filtering from Excel", variable=self.use_excel_filter).pack(pady=5)

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

        # Initialize the filter state
        self.toggle_date_filter()

    def open_link(self, url):
        """Open the given URL in the default web browser"""
        webbrowser.open_new(url)

    def browse_xlsx(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        self.xlsx_file_path.set(file_path)

    def browse_destination(self):
        directory = filedialog.askdirectory()
        self.destination_dir.set(directory)

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
        if not self.xlsx_file_path.get() or not self.destination_dir.get():
            messagebox.showwarning("Input Error", "Please select both an Excel file and a destination folder.")
            return

        self.progress_text.delete(1.0, tk.END)
        self.progress_text.insert(tk.END, "Starting download...\n")
        self.progress_bar["value"] = 0

        # Gather filtering options
        start_date = self.start_date_entry.get_date() if self.use_date_filter.get() else None
        end_date = self.end_date_entry.get_date() if self.use_date_filter.get() else None
        # Decision Type filtering has been removed; always pass None.
        decision_type = None
        use_excel_filter = self.use_excel_filter.get()

        self.summary = {"downloaded": 0, "skipped": 0, "replaced": 0, "failed": 0}

        downloader = PDFDownloader(
            self.xlsx_file_path.get(),
            self.destination_dir.get(),
            self.update_download_progress,
            start_date=start_date,
            end_date=end_date,
            decision_type=decision_type,
            use_excel_filter=use_excel_filter,
            stop_callback=lambda: not self.downloading
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

    def toggle_date_filter(self):
        state = tk.NORMAL if self.use_date_filter.get() else tk.DISABLED
        self.start_date_entry.config(state=state)
        self.end_date_entry.config(state=state)

if __name__ == "__main__":
    app = NSSDecisionDownloader()
    app.mainloop()
