# Engineering Data Tools — Streamlit v9/v5

## Contents
- Tab 1: APC1 → OIT, based on uploaded `1114 APC1_to_OIT_v9.py`
- Tab 2: QA PSV Compare, based on uploaded `1114 QApsv_compare_list_v5.py`

The original mapping, header aliases, PSV parsing, comparison normalization, report sheets and formatting are retained. The desktop file picker and auto-open behavior are replaced with browser uploads and download buttons.

## Local run (Anaconda)
```bash
conda create -n engineering_tools python=3.11 -y
conda activate engineering_tools
pip install -r requirements.txt
streamlit run app.py
```

## VS Code
Open this folder, select the `engineering_tools` interpreter, open Terminal, and run the commands above.

## Streamlit Community Cloud
1. Create a GitHub repository.
2. Upload all project files to the repository root (including `modules/` and `.streamlit/`).
3. Visit https://share.streamlit.io/ and sign in with GitHub.
4. Create app; choose repository and branch; set Main file path to `app.py`.
5. Deploy and share the generated URL.

## Important
- This adaptation uses the new v9 APC1 and v5 QA scripts supplied for this task.
- Test with non-confidential sample workbooks before production use.
- A public Streamlit app is accessible to anyone who has its URL. Do not upload company-confidential PDK/APC1/OIT/webDM data unless approved by company IT/InfoSec.
