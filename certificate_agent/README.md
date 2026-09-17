# Certificate Generation Agent

An automated tool to generate PDF certificates from an Excel data source and a PDF template.

## Purpose
This agent takes an Excel file containing student data and a PDF certificate template containing placeholders `{{NAME}}` and `{{REG_NO}}`. It automatically generates individual PDF certificates for each student while preserving the exact layout, dimensions, and background artwork of the original template.

## Folder Structure
- `main.py`: The main script to generate certificates.
- `requirements.txt`: Python dependencies.
- `README.md`: This documentation.

## Required Input
Place the following files in the directory where you run the script:
1. **Excel File** (.xlsx or .xls) containing the student data.
2. **PDF Template** (.pdf) containing `{{NAME}}` and `{{REG_NO}}`.

The script uses automatic file detection and will scan the current directory for valid inputs.

### Supported Excel Columns
The agent looks for columns indicating Registration Number and Name. It supports variations (case-insensitive, spaces ignored):
- **Registration Number**: `Reg No`, `RegNo`, `REG NO`, `Registration No`, `Registration Number`
- **Name**: `Name`, `Student Name`, `StudentName`

## Installation
Ensure you have Python installed. Install dependencies using:
```bash
pip install -r requirements.txt
```

## How to Run
Execute the main script from your project directory (the folder containing the Excel and PDF files):
```bash
python certificate_agent/main.py
```

## Output Structure
A new folder named `certificates/` will be created in your current working directory.
- Individual PDFs: Generated certificates for each valid student.
- `generation_report.csv`: A report of all processed records, their success status, and corresponding filenames.

## Features
- **Automatic File Detection:** No need to hard-code input filenames.
- **Filename Rules:** Certificates are saved as `<RegNo>_<StudentName>.pdf`. Invalid Windows characters are automatically sanitized.
- **Placeholder Handling:** Accurately finds and redacts placeholders (`{{NAME}}`, `{{REG_NO}}`) directly from the PDF layers and replaces them with student data.
- **Font Sizing:** Dynamically reads the original template's font size and shrinks text automatically for unusually long names to prevent overlap or clipping.
- **Rerun Behavior:** Running the script multiple times is safe. The `certificates/` folder is cleared of old outputs before generation begins.
- **Validation & Error Handling:** Ensures PDFs are valid, correct page count/dimensions are maintained, and inserted text is readable in the output. Displays descriptive errors if inputs are invalid.
