# Server Hardware Validation Report Generator

Local Streamlit app that reads a validated Excel workbook, accepts a chassis image or a PDF of group visuals, and generates a PowerPoint plus an Outlook-ready email.

## Run locally
1. Install Python 3.10+.
2. In this folder, run: `python -m pip install -r requirements.txt`
3. Run: `streamlit run app.py`
4. Open the local URL shown by Streamlit.

## Input assumptions
- First worksheet is the table of contents.
- Remaining worksheets are named `Group 1`, `Group 2`, etc.
- Group sheets use columns: Component, Part Number, Description, Quantity, Supported (Yes/No/Refer comment/N/A), Comment.
- For a PDF, page order is matched to group order. A single image is reused for every group.

## Notes
- This is a first working baseline. Review generated output before sending.
- Component descriptions, statuses and comments are read from the workbook; the app does not independently validate hardware.
