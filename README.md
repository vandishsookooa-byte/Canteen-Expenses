# Canteen Expenses Dashboard

A professional Flask-based dashboard for managing and analysing canteen expenditure across multiple nationalities.

## Setup

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Place your `Canteen.xlsx` file in the project root **or** set the environment variable:
```bash
# Windows
set EXCEL_PATH=C:\Users\transport\Desktop\Canteen Expenses\Canteen.xlsx

# Linux / macOS
export EXCEL_PATH=/path/to/Canteen.xlsx
```

3. Run the app:
```bash
python app.py
```

4. Open in browser: `http://localhost:5000`  
   From other machines on the same network: `http://<your-ip>:5000`

## Excel File Structure

| Sheet | Columns |
|-------|---------|
| Bangladeshi | PERIOD, ITEMS, QTY, UNIT, UNIT PRICE, TOTAL |
| Malagasy | PERIOD, ITEMS, QTY, UNIT, UNIT PRICE, TOTAL |
| Indian | PERIOD, ITEMS, QTY, UNIT, UNIT PRICE, TOTAL |
| Srilankan | PERIOD, ITEMS, QTY, UNIT, UNIT PRICE, TOTAL |
| EMPLOYEES | PERIOD, BANGLADESHI, INDIAN, MALAGASY, SRILANKAN |
