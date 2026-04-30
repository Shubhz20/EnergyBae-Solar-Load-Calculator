# Submission Note - AI Intern Task

I built a Streamlit web app that takes an MSEDCL bill (PDF or image), uses
**Gemini 1.5 Flash** to extract `consumer_name`, `consumer_number`,
`billing_unit`, `connected_load_kw`, `tariff_category`, and
`avg_monthly_consumption` as strict JSON, shows the values in editable text
boxes for the sales team to verify, and then writes them into the Solar Load
Excel Template using **openpyxl** so all of the existing ROI and
system-sizing formulas stay intact.

The stack is intentionally minimal - Streamlit for the UI, `openpyxl` for
formula-safe Excel writes (Pandas would have flattened formulas, which is why
I avoided it), and `google-generativeai` for the bill parsing. The Gemini
prompt is engineered for MSEDCL formats: it converts HP to kW where needed
and computes the 6-month consumption average from the history table when the
bill provides it.

What I would improve next: auto-detect the template's input cells via named
ranges or label-matching instead of a hard-coded `field -> cell` map, add a
batch mode for processing a folder of bills, and cache Gemini responses so
repeat uploads of the same bill skip the API call.
