# MellowUPIreportsFormater

MellowUPIreportsFormater is a Python tool designed to simplify the formatting of UPI merchant transaction reports. It streamlines the process of validating and converting these reports into PDF format. The tool focuses on organizing data, ensuring accuracy, and facilitating efficient documentation printing.

## Features

- Formats UPI merchant transaction reports for validation.
- Converts formatted reports to PDF.
- Organizes data to enhance clarity and accuracy.
- Streamlines documentation printing process.
- Supports various data manipulation and formatting tasks.
- Allows users to choose either single or multiple CSV or ZIP files for processing.

## Compatibility

Current Compatibility: PhonePe UPI Transaction Reports Only.

MellowUPIreportsFormater is currently optimized for processing PhonePe UPI reports. If you're interested in using the tool for other UPI platforms, please let us know!


## Dependencies

Before using MellowUPIreportsFormater, ensure you have the following dependencies installed:

- `tkinter`: Python's standard GUI library.
- `openpyxl`: A library for reading and writing Excel files.
- `os`: Provides functions for interacting with the operating system.
- `datetime`: Offers classes for working with dates and times.
- `zipfile`: Allows manipulation of ZIP archives.
- `win32com.client`: Enables interaction with COM objects (for Windows users).
- `win32print`: Facilitates interaction with printers on Windows systems.

You can install these dependencies using the following command:

```bash
pip install openpyxl tkinter win32com pywin32
```

## Usage

### Using Python

1. Install the required dependencies using the provided command.

2. Run the main script by executing:

   ```bash
   python main.py
   ```

3. Follow the on-screen instructions to select either a single or multiple CSV or ZIP files containing UPI merchant transaction reports, to validate and format them, and generate PDFs.

### Using Executable (exe)

1. Download the latest release from the [Releases](https://github.com/yourusername/MellowUPIreportsFormater/releases) section.

2. Double-click the `MellowUPIreportsFormater.exe` file.

3. Follow the on-screen instructions to select either a single or multiple CSV or ZIP files containing UPI merchant transaction reports, to validate and format them, and generate PDFs.

## Contributing

Contributions are welcome! If you find any issues or want to enhance the tool for other UPI platforms, feel free to open an issue or submit a pull request.

## License

This project is licensed under the [MIT License](LICENSE).

## How to Contribute

1. Fork the repository.

2. Create a new branch for your feature or bug fix.

3. Make your changes and test thoroughly.

4. Submit a pull request with a clear description of your changes.

---
