# Academic Product Margin Intelligence Dashboard

A Streamlit web application built from the supplied margin workbook. It is designed to run locally in Visual Studio Code and provides an executive dashboard, product/grade analytics, operating-cost views, item-level drilldowns, and scenario controls.

## 1. Open in Visual Studio Code
Open the `margin_dashboard_app` folder in VS Code.

## 2. Create a virtual environment
### Windows
```bash
python -m venv .venv
.venv\Scripts\activate
```

### macOS / Linux
```bash
python3 -m venv .venv
source .venv/bin/activate
```

## 3. Install dependencies
```bash
pip install -r requirements.txt
```

## 4. Start the web application
```bash
streamlit run app.py
```

Your browser should open automatically. If not, use the local URL shown in the VS Code terminal (normally `http://localhost:8501`).

## What the application includes
- Executive KPI cards: landing value, product cost, gross margin, operating cost, net margin, below-target items
- Academic Year / Product Type / Grade / Subject filters
- Product profitability and target-health charts
- Grade-wise margin analysis and treemap
- Operating-cost composition and scenario inputs
- Searchable item-level profitability table
- CSV export of the filtered analysis
- Scenario target margin and operating-cost controls
- Upload option to analyze a replacement workbook with the same structure

## Data logic note
The source workbook contains formula cells whose cached results are not populated when read outside Excel, and the item-level Academic Year formula appears shifted. The app therefore recomputes the margin model from underlying raw fields and derives Academic Year from `Edition` / `Product Name`.

## Bundled source workbook
The supplied Excel file is included under `data/margin_model.xlsx`, so the application runs immediately after dependencies are installed.
