import fitz  # PyMuPDF

def extract_ibm_terms_text(file_like) -> str:
    doc = fitz.open(stream=file_like.read(), filetype="pdf")
    found_terms = False
    ibm_terms_lines = []
    useful_resources_lines = []
    capture_useful = False
    useful_resources_captured = False  # Track if we've already captured useful resources

    for page in doc:
        lines = (page.get_text("text") or page.get_text()).splitlines()
        for line in lines:
            line = line.strip()
            if "IBM Terms and Conditions" in line:
                found_terms = True
                capture_useful = False
                continue  # skip the header itself
            if "Useful/Important web resources:" in line and not useful_resources_captured:
                capture_useful = True
                useful_resources_lines.append(line)
                useful_resources_captured = True  # Mark as captured to prevent duplicates
                continue
            if found_terms and line:
                # Stop at page numbers or footer
                if line.lower().startswith("page ") and line.count(" ") <= 3:
                    continue
                ibm_terms_lines.append(line)
            elif capture_useful and line:
                useful_resources_lines.append(line)

    # Reconstruct paragraphs as in your logic
    reconstructed_ibm_terms = []
    current_paragraph = []
    for line in ibm_terms_lines:
        if not line:
            continue
        if (line.startswith("IBM International") or 
            line.startswith("The quote or order") or
            line.startswith("Unless specifically") or
            line.startswith("The terms of the IBM") or
            line.startswith("If you have any trouble")):
            if current_paragraph:
                reconstructed_ibm_terms.append(" ".join(current_paragraph))
                current_paragraph = []
            current_paragraph = [line]
        else:
            if current_paragraph:
                current_paragraph.append(line)
            else:
                current_paragraph = [line]
    if current_paragraph:
        reconstructed_ibm_terms.append(" ".join(current_paragraph))

    all_content = reconstructed_ibm_terms
    if useful_resources_lines:
        all_content.append("")
        all_content.extend(useful_resources_lines)
    return "\n\n".join(all_content)
