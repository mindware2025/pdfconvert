# Amal PDF to Excel

This functionality is now integrated into the main project Streamlit app as the Operations tool `Comm Generator`.

It processes:

- One `SOB` PDF
- One IBM `PO` / `Commercial Invoice` PDF
- One Excel workbook output with:
- `comm-inv`
- `pack_list`

## Run

Start the main project app from the project root:

```powershell
pip install -r requirements.txt
streamlit run app.py
```

Then open:

- `Operations`
- `Comm Generator`

## Notes

- `amal/app.py` was removed because the feature now runs from the root app.
- The processing logic remains in the `amal` module files such as `processor.py`, `ibm_parser.py`, `sob_parser.py`, and `workbook_builder.py`.
