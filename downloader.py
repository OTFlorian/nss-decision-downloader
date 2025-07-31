import pandas as pd
import requests
import re
import os
import openpyxl  # Import openpyxl to handle filtered rows in Excel

class PDFDownloader:
    def __init__(self, file_path, destination_dir, progress_callback=None,
                 decision_type=None, use_excel_filter=False, stop_callback=None,
                 file_type="xlsx"):
        self.file_path = file_path
        self.pdf_dir = os.path.join(destination_dir, "pdf")
        self.new_downloads = []
        self.skipped_downloads = []
        self.replaced_downloads = []
        self.failed_downloads = []
        self.progress_callback = progress_callback  # Callback function for progress updates
        self.stop_callback = stop_callback  # Callback to check if the download should stop
        self.decision_type = decision_type
        self.use_excel_filter = use_excel_filter  # Flag to determine if Excel filters should be respected
        self.file_type = file_type  # 'xlsx' for open data, 'xml' for export from search engine

        # Create the pdf directory if it doesn't exist
        if not os.path.exists(self.pdf_dir):
            os.makedirs(self.pdf_dir)

    def download_pdfs(self):
        df = self._load_dataframe()

        # Filter the DataFrame by "Typ rozhodnutí" if provided
        if self.decision_type:
            df = df[df['Typ rozhodnutí'] == self.decision_type]

        # Extract the filtered URLs from the specific column
        urls = df['Odkaz ECLI'].dropna().tolist()

        # Download each PDF
        for index, url in enumerate(urls):
            # Check if the stop signal is set
            if self.stop_callback and self.stop_callback():
                break

            file_name, status = self._download_pdf(url)
            # Call the progress callback after each download
            if self.progress_callback:
                self.progress_callback(index + 1, len(urls), file_name, status)

        return {
            "new_downloads": self.new_downloads,
            "skipped_downloads": self.skipped_downloads,
            "replaced_downloads": self.replaced_downloads,
            "failed_downloads": self.failed_downloads
        }

    def _get_visible_rows(self, df, sheet):
        """Filter the DataFrame to only include visible rows (non-hidden)"""
        visible_rows = []
        for row_idx, row in enumerate(sheet.iter_rows(min_row=2), start=2):  # Assuming the first row is headers
            # Check if the row is hidden
            if not sheet.row_dimensions[row_idx].hidden:
                visible_rows.append([cell.value for cell in row])

        # Reconstruct DataFrame with visible rows only
        return pd.DataFrame(visible_rows, columns=[cell.value for cell in sheet[1]])

    def _load_dataframe(self):
        """Load the input file and return a unified DataFrame with an 'Odkaz ECLI' column."""
        if self.file_type == "xml":
            df = pd.read_xml(self.file_path, parser="etree")
            # Normalise column names to match the XLSX structure
            if 'Datum' in df.columns:
                df.rename(columns={'Datum': 'Datum rozhodnutí'}, inplace=True)
            if 'Druh dokumentu' in df.columns:
                df.rename(columns={'Druh dokumentu': 'Typ rozhodnutí'}, inplace=True)
            if 'ECLI' in df.columns:
                df['Odkaz ECLI'] = "http://vyhledavac.nssoud.cz/GetContent?ECLI=" + df['ECLI']
        else:
            if self.use_excel_filter:
                wb = openpyxl.load_workbook(self.file_path, data_only=True)
                sheet = wb.active
                df = pd.DataFrame(sheet.values)
                df = self._get_visible_rows(df, sheet)
            else:
                df = pd.read_excel(self.file_path, sheet_name='List1')
        return df

    def _download_pdf(self, url):
        try:
            # Extract the relevant part of the URL using regex
            match = re.search(r'ECLI:CZ:NSS:(\d{4}):(.*)', url)
            if match:
                # Extracted string after "ECLI:CZ:NSS:XXXX:"
                file_identifier = match.group(2)
                # Replace all dots with spaces
                file_name = file_identifier.replace('.', ' ') + ".pdf"

                # Define the full path for the PDF file in the pdf directory
                pdf_file_path = os.path.join(self.pdf_dir, file_name)

                # Check if the file already exists
                if os.path.exists(pdf_file_path):
                    # Check if the file is empty (0 kB)
                    if os.path.getsize(pdf_file_path) == 0:
                        self.replaced_downloads.append(file_name)
                        return file_name, "replaced"
                    else:
                        self.skipped_downloads.append(file_name)
                        return file_name, "skipped"

                # Download the content
                response = requests.get(url)
                if response.status_code == 200:
                    # Save the file, replacing any existing file with the same name
                    with open(pdf_file_path, 'wb') as f:
                        f.write(response.content)
                    if file_name not in self.replaced_downloads:
                        self.new_downloads.append(file_name)
                    return file_name, "downloaded"
                else:
                    self.failed_downloads.append(file_name)
                    return file_name, "failed"
            else:
                self.failed_downloads.append(url)
                return url, "failed"
        except Exception as e:
            self.failed_downloads.append(url)
            return url, "failed"

if __name__ == "__main__":
    # Example usage in a console environment
    downloader = PDFDownloader(
        "otevrena_data_NSS-2024-08-15-small.xlsx", ".",
        decision_type='Meritorní',
        use_excel_filter=True,  # Set to True to test the filtering functionality
        file_type="xlsx"
    )
    summary = downloader.download_pdfs()

    print("\nSummary of the download process:")
    print(f"Newly downloaded files: {len(summary['new_downloads'])}")
    print(f"Skipped existing files: {len(summary['skipped_downloads'])}")
    print(f"Replaced empty files: {len(summary['replaced_downloads'])}")
    print(f"Failed downloads: {len(summary['failed_downloads'])}")
