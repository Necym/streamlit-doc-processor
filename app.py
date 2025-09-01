import streamlit as st
import pandas as pd
from docx import Document
import re
from io import BytesIO

# ───────────────────────────── Globals ─────────────────────────────
# We'll set this from the UI before calling any processor.
SHEET_NAME_SELECTED = None

# ─────────────────────── Helpers: header detection ───────────────────────

def _norm(s: str) -> str:
    """
    Normalize strings from the Word table to make matching robust.
    - Lowercase
    - Remove non-letters (digits, punctuation, spaces, emojis)
    Examples:
      "Question Prompt" -> "questionprompt"
      "Radio Button 1 - Normal state" -> "radiobuttonnormalstate"
      "copyright 1" -> "copyright"
    """
    return re.sub(r'[^a-z]+', '', (s or '').lower())

def _get_header_indices(header_cells):
    """
    Find indices for the required columns by normalized label containment.
    Returns: dict with keys id/type/sourcetext/translation or None if not found.
    """
    labels = [_norm(c.text) for c in header_cells]
    need = {'id': 'id', 'type': 'type', 'sourcetext': 'sourcetext', 'translation': 'translation'}
    idx = {}
    for key, token in need.items():
        for i, lab in enumerate(labels):
            if token in lab:
                idx[key] = i
                break
    return idx if len(idx) == 4 else None

# ───────────────────── Excel parsing (unchanged behavior) ─────────────────────

def extract_prompt_answers_and_explanation(df_row):
    """
    Keep the original parsing behavior:
      - Prompt is everything before the first 'A.' in df['Question']
      - Answers are split on newlines with labels A./B./C./...
      - Explanation comes from df['Explanation']
    """
    question_text = str(df_row['Question'])
    explanation = str(df_row.get('Explanation', '') or '')

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

# ───────────────────── Core processor (Type-anchored, row-only fix) ─────────────────────

def scan_word_document_version_qp(word_file, excel_file, sheet_name, question_limit):
    """
    New robust implementation that:
      - Anchors on Type == 'Question Prompt' rows
      - Fills next four 'Radio Button ... Normal state' rows with answers
      - Skips 'copyright' rows
      - Fills the next 'Rounded Rectangular Caption' row with explanation
      - Maps questions sequentially to Excel rows (1->1, 2->2, ...)
      - Writes ONLY into the Translation column (same as original)
    """
    # Read Excel sheet by user-provided name
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
    except ValueError as e:
        # Typical message: "Worksheet named 'X' not found"
        raise RuntimeError(f"Excel sheet '{sheet_name}' was not found.") from e

    doc = Document(word_file)

    # Locate the localization table and get column indices
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
    i = 1  # skip header row (index 0)

    while i < total_rows and q_count < question_limit and q_count < len(df):
        _, type_text, _, _ = get_vals(table.rows[i])
        if _norm(type_text) == 'questionprompt':
            # Sequential mapping to Excel (1-based)
            q_index = q_count + 1
            excel_prompt, excel_answers, excel_explanation = extract_prompt_answers_and_explanation(df.iloc[q_index - 1])

            # PROMPT → Translation on the same row
            put_translation(table.rows[i], excel_prompt)

            # ANSWERS → fill next up to 4 rows where Type is "Radio Button ... Normal state"
            answers_needed = min(4, len(excel_answers))
            answers_filled = 0
            j = 1
            while answers_filled < answers_needed and (i + j) < total_rows:
                _, t_next, _, _ = get_vals(table.rows[i + j])
                tnorm = _norm(t_next)
                if tnorm == 'questionprompt':
                    # Next question encountered early; stop this block
                    break
                if tnorm.startswith('radiobutton') and tnorm.endswith('normalstate'):
                    put_translation(table.rows[i + j], excel_answers[answers_filled])
                    answers_filled += 1
                j += 1

            # EXPLANATION → first "Rounded Rectangular Caption" after answers; skip "copyright"
            k = i + j
            while k < total_rows:
                _, t_k, _, _ = get_vals(table.rows[k])
                tknorm = _norm(t_k)
                if tknorm == 'copyright':
                    k += 1
                    continue
                if tknorm == 'roundedrectangularcaption':
                    put_translation(table.rows[k], excel_explanation)
                    k += 1
                    break
                if tknorm == 'questionprompt':
                    # explanation missing; move on defensively
                    break
                k += 1

            q_count += 1
            # Advance to next scanning position (avoid double processing)
            i = max(k, i + j, i + 1)
        else:
            i += 1

    # Return updated document
    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out, f"Processed {q_count} question(s) with sheet '{sheet_name}'."

# ───────────────────── Back-compat wrappers (Option B) ─────────────────────

def scan_word_document_version_a(word_file, excel_file, question_limit):
    """
    Backward-compatible alias for old 'Version A'.
    Uses the sheet name the user typed in the UI (SHEET_NAME_SELECTED),
    falling back to 'Simulated' if not set.
    """
    sheet = SHEET_NAME_SELECTED or "Simulated"
    return scan_word_document_version_qp(word_file, excel_file, sheet, question_limit)

def scan_word_document_version_b(word_file, excel_file, question_limit):
    """
    Backward-compatible alias for old 'Version B'.
    Uses the sheet name the user typed in the UI (SHEET_NAME_SELECTED),
    falling back to 'Simulated' if not set.
    """
    sheet = SHEET_NAME_SELECTED or "Simulated"
    return scan_word_document_version_qp(word_file, excel_file, sheet, question_limit)

# ───────────────────────────── Streamlit UI ─────────────────────────────

st.title("Document Processor")

version_choice = st.selectbox("Select Version", ["Version A", "Version B", "Version QP (by Type)"])

# Let the user type the Excel sheet/tab name (default to your example "1Q1")
sheet_name_input = st.text_input(
    "Excel sheet/tab name",
    value="1Q1",
    help="Type the exact tab name in your Excel file (e.g., 1Q1)."
)

word_file = st.file_uploader("Upload Word Document", type=["docx"])
excel_file = st.file_uploader("Upload Excel Document", type=["xlsx"])
question_limit = st.number_input("How many questions would you like to process?", min_value=1, value=5)

if word_file and excel_file and st.button("Process"):
    # Read the files once
    word_bytes = word_file.read()
    excel_bytes = excel_file.read()

    # Store the selected sheet name for wrappers A/B
    global SHEET_NAME_SELECTED
    SHEET_NAME_SELECTED = sheet_name_input

    try:
        if version_choice == "Version A":
            output_buffer, output_message = scan_word_document_version_a(BytesIO(word_bytes), BytesIO(excel_bytes), question_limit)
        elif version_choice == "Version B":
            output_buffer, output_message = scan_word_document_version_b(BytesIO(word_bytes), BytesIO(excel_bytes), question_limit)
        else:
            # Direct call for the explicit QP option
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
        # Surface unexpected errors for easier debugging
        st.exception(e)
