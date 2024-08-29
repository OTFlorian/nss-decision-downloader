import fitz  # PyMuPDF
import os

class PDFConverter:
    def __init__(self, pdf_dir, txt_dir, progress_callback=None, stop_callback=None):
        self.pdf_dir = pdf_dir
        self.txt_dir = txt_dir
        self.progress_callback = progress_callback  # Callback function for progress updates
        self.stop_callback = stop_callback  # Callback to check if the conversion should stop

        # Create the txt directory if it doesn't exist
        if not os.path.exists(self.txt_dir):
            os.makedirs(self.txt_dir)

        # Summary counts
        self.converted = 0
        self.skipped = 0
        self.failed = 0

    def convert_pdfs(self):
        files = [f for f in os.listdir(self.pdf_dir) if f.endswith(".pdf")]
        total_files = len(files)

        for index, pdf_file in enumerate(files):
            # Check if the stop signal is set
            if self.stop_callback and self.stop_callback():
                break

            status = self._convert_pdf_to_text(os.path.join(self.pdf_dir, pdf_file))
            # Call the progress callback after each conversion
            if self.progress_callback:
                self.progress_callback(index + 1, total_files, pdf_file, status)

        return {
            "converted": self.converted,
            "skipped": self.skipped,
            "failed": self.failed
        }

    def _convert_pdf_to_text(self, pdf_file_path):
        base_name = os.path.splitext(os.path.basename(pdf_file_path))[0]
        txt_file_path = os.path.join(self.txt_dir, f"{base_name}.txt")

        try:
            # Check if the file is empty
            if os.path.getsize(pdf_file_path) == 0:
                self.skipped += 1
                return "skipped"

            # Open the PDF file
            with fitz.open(pdf_file_path) as doc:
                text = ""
                for page in doc:
                    text += page.get_text()

            # Save the text to the txt file
            with open(txt_file_path, 'w', encoding='utf-8') as f:
                f.write(text)

            self.converted += 1
            return "converted"
        except Exception as e:
            print(f"Failed to convert {pdf_file_path}: {e}")
            self.failed += 1
            return "failed"

if __name__ == "__main__":
    # Example usage in a console environment
    converter = PDFConverter("pdf", "txt")
    summary = converter.convert_pdfs()

    print("\nSummary of the conversion process:")
    print(f"Converted files: {summary['converted']}")
    print(f"Skipped files: {summary['skipped']}")
    print(f"Failed conversions: {summary['failed']}")
