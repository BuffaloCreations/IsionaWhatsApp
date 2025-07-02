# Isiona WhatsApp Bot

This project contains a WhatsApp automation script for sending messages to dental clinics.

## Setup

### 1. Create Virtual Environment
```bash
python3 -m venv venv
```

### 2. Activate Virtual Environment
```bash
# On macOS/Linux:
source venv/bin/activate

# On Windows:
venv\Scripts\activate
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

## Usage

1. Make sure your virtual environment is activated
2. Update the Excel file path in `Main.py` if needed
3. Run the script:
```bash
python Main.py
```

## Deactivate Virtual Environment
When you're done working, deactivate the virtual environment:
```bash
deactivate
```

## Dependencies
- pandas: For Excel file processing
- pywhatkit: For WhatsApp automation
- openpyxl: For Excel file reading 