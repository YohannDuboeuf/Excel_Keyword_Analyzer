# 🔍 Excel Keyword Analyzer

Excel Keyword Analyzer is a Python tool to **scan Excel files (`.xls`, `.xlsm`) for a specific keyword** in either **VBA macros** or **cell formulas**.

---

## 🚀 Features

- 🔎 Search for a **custom keyword** in:
  - VBA macros embedded in Excel files.
  - Excel formulas (e.g., `=SUM(...)`, `=IF(...)`).
- 📂 Supports `.xls` and `.xlsm` files.
- 🔁 Automatically converts `.xls` files to `.xlsx` using LibreOffice.
- ⚙️ Uses **parallel processing** (multi-threading) for performance.
- 💻 Clean **command-line interface** with support for `--keyword` or `-k`.

---

## 📦 Installation

### 1. Requirements

- Python 3.7+
- [LibreOffice](https://www.libreoffice.org/download/) (used for `.xls` to `.xlsx` conversion)
- Pip (Python package installer)

### 2. Clone the repository

```bash
git clone https://github.com/your-username/excel-keyword-analyzer.git
cd excel-keyword-analyzer
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

## 📁 Project Structure

```
excel-keyword-analyzer/
│
├── assets/                      # Working directory for input/output files
│   ├── excel/                   # Input: Excel files to scan (.xls, .xlsm)
│   ├── macro/                   # Temp folder: extracted macro files (.txt)
│   └── macro_trouves/           # Output: matched Excel/macro files
│
├── main.py                      # Main script: keyword analysis logic
├── requirements.txt             # Dependencies for the project
├── README.md                    # Documentation (you’re reading it)
└── LICENSE   
```

## Usage

### Run the script

To run the script, navigate to the project directory and execute the following command:

```bash
python excel_keyword_analyzer.py --keyword "your_keyword"
```

Replace `"your_keyword"` with the keyword you want to search for.