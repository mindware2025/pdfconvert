from PyPDF2 import PdfReader


def extract_text_from_pdf(uploaded_file) -> str:
    uploaded_file.seek(0)
    reader = PdfReader(uploaded_file)

    pages = []
    for page in reader.pages:
        pages.append(page.extract_text() or "")

    uploaded_file.seek(0)
    return "\n".join(pages).strip()
