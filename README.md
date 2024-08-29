# Czech Supreme Administrative Court (NSS) Decision Downloader

This application is designed to download and convert Czech Supreme Administrative Court decisions. The tool provides a GUI that allows users to filter the decisions by date, type, or Excel-applied filters, download them, and then convert the downloaded PDFs into plain text files.

## Features

- **Download Supreme Administrative Court Decisions**: Automatically download decisions based on a specified date range, decision type, or visible (non-hidden) Excel rows.
- **Excel Filtering Support**: Option to respect filters or visible rows directly applied in the Excel sheet.
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

1. **Select the Excel File**: Click the "Browse" button to select the Excel file containing the court decisions' metadata.
    - The Excel file can be obtained at the [Supreme Administrative Court website](https://www.nssoud.cz/informace-pro-verejnost/otevrena-data-k-soudni-cinnosti).
2. **Select the Destination Folder**: Choose where to save the downloaded PDFs and the converted TXT files.
3. **Filter Options**:
    - You can filter decisions by date by checking the respective checkbox and setting the desired values.
    - Enable the "Use filtering from Excel" checkbox to only process visible (non-hidden) rows in the Excel sheet.
4. **Start Download**: Click "Start Download" to begin downloading the filtered decisions. You can stop the process at any time.
5. **Start Conversion**: After downloading, click "Start Conversion" to convert the PDFs to TXT format.

## Files

- `downloader.py`: Contains the logic for downloading the PDF files based on filters and Excel visibility settings.
- `converter.py`: Handles the conversion of downloaded PDFs to TXT format.
- `gui_app.py`: The GUI application that integrates downloading and conversion.
- `requirements.txt`: Lists all the Python packages required to run the application.

## Example Citations

If you use this project, give appropriate credit by mentioning the author's publishing name, *Tristan Florian*, and the tool (see the [LICENSE](LICENSE) file).

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
