# Amal PDF to Excel

This functionality is integrated into the main project Streamlit app as the Operations tool `Commercial Invoice & Packing List`.

It supports:

- One or more `SOB` PDF files paired with IBM `PO` / `Commercial Invoice` PDFs
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
- `Commercial Invoice & Packing List`

## Notes

- The standalone `amal/app.py` entrypoint is no longer needed and can be removed.
- The processing logic lives in `processor.py`, `ibm_parser.py`, `sob_parser.py`, `pdf_utils.py`, and `workbook_builder.py`.
