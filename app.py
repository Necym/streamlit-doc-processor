import streamlit as st
import pandas as pd
from docx import Document
import re
from io import BytesIO

# ───────────────────────── Helpers: header detection ─────────────────────────

def _norm(s: str) -> str:
    # normalize type/header text (lowercase, letters only) to survive emojis/spacing
    return re.sub(r'[^a-z]+', '', (s or '').lower())

def _get_header_indices(header_cells):
    labels = [_norm(c.text) for c in header_cells]
    need = {'id': 'id', 'type': 'type', 'sourcetext': 'sourcetext', 'translation': 'translation'}
    idx = {}
    for key, token in need.items():
        for i, lab in enumerate(labels):
            if token in lab:
                idx[key] = i
                break
    return idx if len(idx) == 4 else None

# ───────────────────────── Excel parsing (unchanged) ─────────────────────────

def extract_prompt_answers_and_explanation(df_row):
    question_text = str(df_row['Question'])
    explanation = str(df_row.get('Explanation', '') or '')

    # Split prompt vs answers on first "A."
    parts = re.split(r'(?=\bA\.)', question_text, maxsplit=1)
    prompt = parts[0].strip()

    answers = []
    if len(parts) > 1:
        raw_answers = re.split(r'\n(?=[A-Z]\.)', parts[1].strip())
        for answer in raw_answers:
            clean = re.sub(r'^[A-Z]\.\s*', '', answer).strip()
            if clean:
                answers.append(clean)

    return prompt, answers, explanation

# ───────────────────────── Core processor (row fix only) ─────────────────────────

def scan_word_document_version_qp(word_file, excel_file, sheet_name, question_limit):
    # Read Excel sheet by name (user-provided)
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
    except ValueError as e:
        # typical "Worksheet named 'X' not found" -> surface as a friendly error
        raise RuntimeError(f"Excel sheet '{sheet_name}' was not found.") from e

    doc = Document(word_file)

    # Find the localization table (by headers) and get column indices
    table = None
    col = None
    for t in doc.tables:
        if not t.rows:
            continue
        idx = _get_header_indices(t.rows[0].cells)
        if idx:
            table = t
            col = idx
            break
    if table is None:
        raise RuntimeError("No table with headers [ID, Type, Source Text, Translation] found.")

    def get_vals(row_obj):
        cells = row_obj.cells
        return (
            cells[col['id']].text.strip(),
            cells[col['type']].text.strip(),
            cells[col['sourcetext']].text.strip(),
            cells[col['translation']].text.strip()
        )

    def put_translation(row_obj, text):
        row_obj.cells[col['translation']].text = text

    total_rows = len(table.rows)
    q_count = 0
    i = 1  # skip header row at 0

    while i < total_rows and q_count < question_limit and q_count < len(df):
        _, type_text, _, _ = get_vals(table.rows[i])
        if _norm(type_text) == 'questionprompt':  # anchor
            q_index = q_count + 1  # 1-based
            excel_prompt, excel_answers, excel_explanation = extract_prompt_answers_and_explanation(df.iloc[q_index - 1])

            # PROMPT → this row's Translation
            put_translation(table.rows[i], excel_prompt)

            # ANSWERS → take next up to 4 rows whose Type starts with "radio button" and ends with "normal state"
            answers_needed = min(4, len(excel_answers))
            answers_filled = 0
            j = 1
            while answers_filled < answers_needed and (i + j) < total_rows:
                _, t_next, _, _ = get_vals(table.rows[i + j])
                tnorm = _norm(t_next)
                if tnorm == 'questionprompt':
                    # next question started unexpectedly; stop this block
                    break
                # match "radio button ... normal state" without relying on numbers/punctuation
                if tnorm.startswith('radiobutton') and tnorm.endswith('normalstate'):
                    put_translation(table.rows[i + j], excel_answers[answers_filled])
                    answers_filled += 1
                j += 1

            # Skip any "copyright 1" rows that might follow
            k = i + j
            while k < total_rows:
                _, t_k, _, _ = get_vals(table.rows[k])
                tknorm = _norm(t_k)
                if tknorm == 'copyright':
                    k += 1
                    continue
                # EXPLANATION → first "rounded rectangular caption" after answers/copyrights
                if tknorm == 'roundedrectangularcaption':
                    put_translation(table.rows[k], excel_explanation)
                    k += 1
                    break
                # Stop if we hit another question
                if tknorm == 'questionprompt':
                    break
                k += 1

            q_count += 1
            # Continue scanning from where we left off
            i = max(k, i + j, i + 1)
        else:
            i += 1

    # Return updated document as BytesIO
    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out, f"Processed {q_count} question(s) with sheet '{sheet_name}'."

# ───────────────────────── Streamlit app ─────────────────────────

st.title("Document Processor")

version_choice = st.selectbox("Select Version", ["Version A", "Version B", "Version QP (by Type)"])

# Let the user type the Excel sheet/tab name (default to your example "1Q1")
sheet_name_input = st.text_input("Excel sheet/tab name", value="1Q1", help="Type the exact tab name in your Excel file.")

word_file = st.file_uploader("Upload Word Document", type=["docx"])
excel_file = st.file_uploader("Upload Excel Document", type=["xlsx"])

if word_file and excel_file:
    question_limit = st.number_input("How many questions would you like to process?", min_value=1, value=5)
    if st.button("Process"):
        # read files once
        word_bytes = word_file.read()
        excel_bytes = excel_file.read()
        try:
            if version_choice == "Version A":
                output_buffer, output_message = scan_word_document_version_a(BytesIO(word_bytes), BytesIO(excel_bytes), question_limit)
            elif version_choice == "Version B":
                output_buffer, output_message = scan_word_document_version_b(BytesIO(word_bytes), BytesIO(excel_bytes), question_limit)
            else:
                output_buffer, output_message = scan_word_document_version_qp(BytesIO(word_bytes), BytesIO(excel_bytes), sheet_name_input, question_limit)

            st.write(output_message)
            st.success("Processing complete. Download the updated Word document below.")
            st.download_button(
                label="Download updated document",
                data=output_buffer,
                file_name="updated_document.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except RuntimeError as e:
            st.error(str(e))
        except Exception as e:
            # surface unexpected errors for easier debugging
            st.exception(e)
