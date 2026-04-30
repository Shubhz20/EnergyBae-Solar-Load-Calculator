# Energybae - Solar Load Calculator

Automate the manual MSEDCL bill -> Excel template workflow for the Energybae
sales team. Upload a bill, Gemini extracts the fields, the user verifies, and
the Solar Load template is filled without touching any formulas.

## What it does

1. **Upload** an MSEDCL electricity bill (PDF / PNG / JPG).
2. **Extract** the key fields with Gemini 1.5 Flash:
   `consumer_name`, `consumer_number`, `billing_unit`, `connected_load_kw`,
   `tariff_category`, `avg_monthly_consumption`.
3. **Verify** the extracted values in editable inputs (human-in-the-loop).
4. **Fill** the Solar Load Excel Template via `openpyxl`, writing only the
   designated input cells. Every existing formula, named range, and style is
   preserved.
5. **Download** the filled `.xlsx`.

## Project structure

```
EnergyBae/
├── app.py                              # Streamlit UI
├── requirements.txt
├── .env.example
├── .streamlit/
│   └── secrets.toml.example
├── README.md
├── SUBMISSION_NOTE.md
├── templates/
│   └── solar_load_template.xlsx        # bundled Solar Load template
├── samples/
│   └── SolarLoad_SampleOutput.xlsx     # demo of a filled output
├── scripts/
│   └── build_template.py               # regenerate the template
├── tests/
│   └── test_excel_handler.py
└── utils/
    ├── __init__.py
    ├── extractor.py                    # Gemini 1.5 Flash bill parser
    └── excel_handler.py                # openpyxl writer (formula-safe)
```

## About the Excel template

The repo ships a working **`templates/solar_load_template.xlsx`** so the app
produces real output out of the box. It has the same input cells the parser
expects (`C5–C10`) plus seven downstream formulas: annual consumption, system
size (kWp), generation, savings, system cost, payback period, and 25-year
ROI. The app uses this bundled file automatically.

When the official Energybae template is available, drop it into
`templates/solar_load_template.xlsx` (overwriting the bundled one) and update
`DEFAULT_CELL_MAP` in `utils/excel_handler.py` to match the real input cells.
Sales staff can also override per-session by uploading a different `.xlsx`
in the sidebar.

To regenerate the bundled template:
```bash
python scripts/build_template.py
```

## Setup

```bash
python -m venv .venv
source .venv/bin/activate           # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

Get a Gemini API key at <https://aistudio.google.com/app/apikey>.

### Local dev

```bash
cp .env.example .env                # paste GOOGLE_API_KEY
streamlit run app.py
```

…or use Streamlit's secrets file:

```bash
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# edit and add GOOGLE_API_KEY
streamlit run app.py
```

### Deployment (Streamlit Community Cloud)

The end-users (sales team) **never see or enter the API key**. The admin
configures it once:

1. Push the repo to GitHub.
2. <https://share.streamlit.io> -> **New app** -> point at this repo.
3. **Settings -> Secrets** -> paste:
   ```toml
   GOOGLE_API_KEY = "AIza...your_key..."
   ```
4. Save. The app reads the key from `st.secrets` automatically; the sidebar
   shows "Gemini API key loaded" and no input field.

Same pattern works on Render, Railway, Fly, or any host that supports env
vars - just set `GOOGLE_API_KEY` in the host's environment.

### Models

The brief specified Gemini 1.5 Flash. Google has rotated model aliases over
time, so `utils/extractor.py` walks a candidate list
(`gemini-1.5-flash-latest` -> `gemini-flash-latest` -> `gemini-1.5-flash` ->
2.0/2.5 flash) and uses the first one that resolves for the active key. If
all named candidates fail, it queries `list_models()` and picks any flash
model that supports `generateContent`.

## Adapting to the real template

The default field-to-cell mapping lives in `utils/excel_handler.py`:

```python
DEFAULT_CELL_MAP = {
    "consumer_name":           "C5",
    "consumer_number":         "C6",
    "billing_unit":            "C7",
    "tariff_category":         "C8",
    "connected_load_kw":       "C9",
    "avg_monthly_consumption": "C10",
}
```

Open the Solar Load Template once, note which cells the team actually types
into, and update this dict. Nothing else needs to change. Cells that contain
formulas (`=...`) are skipped automatically as a second safety net.

## Design choices (interview defense)

1. **Formula integrity** - `openpyxl` is used over Pandas specifically to
   preserve the proprietary ROI / system-sizing formulas. Only input cells are
   written; formulas are detected and skipped.
2. **Human-in-the-loop** - 100% accuracy in financial billing matters, so
   every extracted field is shown in an editable input before export.
3. **Regional context** - The Gemini prompt is engineered for MSEDCL formats:
   it handles HP -> kW conversion (`1 HP = 0.7457 kW`) and computes the
   6-month consumption average when the history table is present.
4. **Deployment readiness** - Streamlit is cross-platform and deploys to the
   cloud (Streamlit Community Cloud, Render, Fly) in minutes; the team can
   use it on a phone while on-site with a customer.
5. **Efficiency** - Reduces 15-30 minutes of manual entry to under 30
   seconds, removing a real bottleneck in the sales cycle.

## What I'd improve next

- Auto-detect template cell positions by named ranges or label-matching
  instead of a hard-coded map.
- Add support for batch uploads (folder of bills -> folder of filled sheets).
- Cache Gemini calls and add a JSON-schema validator for stricter outputs.
