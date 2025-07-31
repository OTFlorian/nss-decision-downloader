# Czech Supreme Administrative Court (NSS) Decision Downloader

This application is designed to download and convert Czech Supreme Administrative Court decisions. The tool provides a GUI that respects Excel-applied filters, downloads the decisions, and converts the downloaded PDFs into plain text files.

## Features

- **Download Supreme Administrative Court Decisions**: Automatically download decisions based on visible (non-hidden) Excel rows.
- **Excel Filtering Support**: Automatically respects filters or visible rows directly applied in the Excel sheet when using the open data XLSX file.
- **XML Export Support**: You can also load an XML export from the official search engine and construct download URLs from ECLI values.
- **Convert PDFs to Text**: Convert the downloaded PDF files to TXT format for easier analysis or processing.
- **User-Friendly GUI**: Simple graphical interface with options to start/stop the download and conversion processes.
- **Progress Tracking**: Real-time progress updates and the ability to stop processes at any time.

## Requirements

- **Python** >= 3.10
- **Required Python packages**: see the `requirements.txt` file.

## Installation

1. Clone this repository:

   ```bash
   git clone https://github.com/OTFlorian/nss-decision-downloader.git
   cd nss-decision-downloader
   ```

2. Install the required dependencies using the provided requirements file:

   ```bash
   pip install -r requirements.txt
   ```

3. Run the application:

   ```bash
   python gui_app.py
   ```

## Usage

1. **Select the Input File**: Click the "Browse" button to select either the open data XLSX file or an XML export containing the court decisions' metadata.
    - The XLSX file can be obtained at the [Supreme Administrative Court website](https://www.nssoud.cz/informace-pro-verejnost/otevrena-data-k-soudni-cinnosti).
    - XML exports can be downloaded directly from the [search engine](https://vyhledavac.nssoud.cz/). However, **you need to enable ECLI** in the search results.
2. **Select the Destination Folder**: Choose where to save the downloaded PDFs and the converted TXT files.
3. **Filter Options**:
    - Filtering from Excel is always applied when using the open data XLSX file. A label below the file type confirms this behavior.
4. **Start Download**: Click "Start Download" to begin downloading the filtered decisions. You can stop the process at any time.
5. **Start Conversion**: After downloading, click "Start Conversion" to convert the PDFs to TXT format.

## Files

- `downloader.py`: Contains the logic for downloading the PDF files based on filters and Excel visibility settings. It supports both the open data XLSX file and XML exports from the search engine.
- `converter.py`: Handles the conversion of downloaded PDFs to TXT format.
- `gui_app.py`: The GUI application that integrates downloading and conversion.
- `requirements.txt`: Lists all the Python packages required to run the application.

## Example Citations

Here are some examples of how the software can be cited.

- **ČSN ISO 690** (Czech):

  FLORIAN, Tristan. *NSS Decision Downloader* [software]. 2025 [cit. 2025-03-30]. Dostupné z: https://github.com/OTFlorian/nss-decision-downloader.

- **Chicago**:

  Florian, Tristan. *NSS Decision Downloader*. 2025. GitHub. https://github.com/OTFlorian/nss-decision-downloader

- **APA**:
  
  Florian, T. (2025). *NSS Decision Downloader* [Computer software]. GitHub. https://github.com/OTFlorian/nss-decision-downloader

## Credits

Created by [Oldřich Tristan Florian](https://otflorian.com), who publishes under the name *Tristan Florian*.

## License

This project is licensed under a custom license. Please see the [LICENSE](LICENSE) file for details.

## Purpose

The official search engine of the Czech Supreme Administrative Court (NSS) does not support bulk downloading of decisions, which poses a limitation for research and data analysis. This application addresses that issue by working with the _Otevřená data k soudní činnosti_ dataset or with XML exports from the search engine. When using the open data file, you can apply filters directly in Excel and the tool automatically downloads all visible decisions. The downloaded PDFs can be converted into plain text files, enabling content-based quantitative analysis. Future plans include experimenting with machine learning models trained on NSS case law and associated metadata.
