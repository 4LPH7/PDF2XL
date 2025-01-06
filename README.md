---

# PDF Table Extractor to Excel

This Python script extracts table data from a PDF file, cleans it, and writes the cleaned data into an Excel file with formatting applied. It uses the `pdfplumber` library to extract tables and the `openpyxl` library to create and format the Excel file.

---

## Features

- **Extract Tables from PDF**: Extracts table data from a PDF file using `pdfplumber`.
- **Clean Data**: Removes empty or `None` values and filters out rows with no meaningful data.
- **Write to Excel**: Writes the cleaned data to an Excel file with formatting:
  - Bold and centered headers with a yellow background.
  - Left-aligned data cells.
  - Borders around all cells.
  - Automatic column width adjustment.
- **Error Handling**: Checks for missing input files and handles cases where no valid table data is found.

---

## Requirements

- Python 3.x
- Libraries: `pdfplumber`, `openpyxl`, `pandas` (optional, not used in the script but imported)

---

## Installation

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/yourusername/pdf-table-extractor.git
   cd pdf-table-extractor
   ```

2. **Install Dependencies**:
   Install the required Python libraries using pip:
   ```bash
   pip install pdfplumber openpyxl pandas
   ```

---

## Usage

1. **Prepare the Input PDF**:
   - Place the PDF file (`test.pdf`) in the project directory, or update the `pdf_path` variable in the script to point to your PDF file.

2. **Run the Script**:
   Execute the script:
   ```bash
   python script_name.py
   ```
   Replace `script_name.py` with the name of your script file.

3. **Check the Output**:
   - The script will generate an Excel file (`output.xlsx`) in the same directory.
   - Open the file to view the extracted and formatted table data.

---

## Customization

- **Input PDF Path**:
  Update the `pdf_path` variable in the `main()` function to point to your desired PDF file.

- **Output Excel Path**:
  Update the `output_path` variable in the `main()` function to specify where the Excel file should be saved.

- **Formatting**:
  Modify the formatting in the `write_cleaned_data_to_excel` function to suit your needs. For example:
  - Change the header font, alignment, or background color.
  - Adjust the border style or cell alignment.



## Error Handling

- If the input PDF file is not found, the script will print an error message and exit.
- If no valid table data is found in the PDF, the script will notify the user and exit.

---

## Dependencies

- **pdfplumber**: For extracting tables from PDFs.
- **openpyxl**: For creating and formatting Excel files.
- **pandas**: Although imported, it is not used in this script. You can remove it if not needed elsewhere.

---

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

## Contributing

Contributions are welcome! If you find any issues or have suggestions for improvements, please open an issue or submit a pull request.

---

## Author

- GitHub: [@4LPH7](https://github.com/4LPH7)

Feel free to contribute or suggest improvements!

---
### Show your support

Give a ⭐ if you like this website!

<a href="https://buymeacoffee.com/arulartadg" target="_blank"><img src="https://cdn.buymeacoffee.com/buttons/v2/default-violet.png" alt="Buy Me A Coffee" height= "60px" width= "217px" ></a>


