# Merchandise Planner Streamlit app

## Run locally

From this folder, install the dependencies and start the app:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run streamlit_app.py
```

The app can use the default planner data or accept an optional fixture CSV and SKU master CSV. It displays the generated reports, a schematic store floor plan, and fixture elevation pages; downloads include the Excel workbook, planogram CSV, and fixture elevation PDF.

## Deploy with Streamlit Community Cloud

1. Push this folder to a GitHub repository. Include `streamlit_app.py`, `merch_planner.py`, `planogram_visual.py`, `fixtures_input.csv`, `assumptions.json`, and `requirements.txt`. Do not include `.venv`.
2. In [Streamlit Community Cloud](https://share.streamlit.io/), create an app from that repository.
3. Set the app's main file path to `streamlit_app.py` and deploy.
