import streamlit as st
import pandas as pd
from docx import Document
import re
from io import BytesIO

# ───────────────────────── Helpers: header detection / normalization ─────────────────────────

def _norm(s: str) -> str:
    """
    Normalize strings from the Word table to make matching robust.
    - Lowercase
    - Remove non-letters (digits, punctuation, spaces, emojis)
    Examples:
      "Question Prompt" -> "questionprompt"
      "Radio Button 1 - Normal state" -> "radiobuttonnormalstate"
      "copyright 1" -> "copyright"
      "Text Box" -> "textbox"
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

# ───────────────────────── Excel parsing (kept like your old code) ─────────────────────────

def extract_prompt_answers_and_explanation(df_row):
    """
    Original behavior:
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

def get_correct_feedback(df_row):
    """
    For Version B: prefer a dedicated 'Correct'/'Correct Feedback' column if present,
    otherwise fall back to 'Explanation'.
    """
    candidates = ['Correct', 'Correct Feedback', 'CorrectFeedback', 'Correct_Explanation', 'CorrectFeedbackText']
    for c in candidates:
        if c in df_row.index:
            val = str(df_row.get(c, '') or '')
            if val.strip():
                return val
    return str(df_row.get('Explanation', '') or '')

# ───────────────────────── Core processors (Type-anchored, row-only fixes) ─────────────────────────

def scan_word_document_version_a(word_file, excel_file, sheet_name, question_limit):
    """
    Version A: Fill Prompt, Answers, and EXPLANATION ('Rounded Rectangular Caption').
    """
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
    except ValueError as e:
        raise RuntimeError(f"Excel sheet '{sheet_name}' was not found.") from e

    doc = Document(word_file)

    # Locate localization table and columns
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
    i = 1  # skip header row

    while i < total_rows and q_count < question_limit and q_count < len(df):
        _, type_text, _, _ = get_vals(table.rows[i])
        if _norm(type_text) == 'questionprompt':
            q_index = q_count + 1
            excel_prompt, excel_answers, excel_explanation = extract_prompt_answers_and_explanation(df.iloc[q_index - 1])

            # PROMPT
            put_translation(table.rows[i], excel_prompt)

            # ANSWERS: next up to 4 Radio Button ... Normal state
            answers_needed = min(4, len(excel_answers))
            answers_filled = 0
            j = 1
            while answers_filled < answers_needed and (i + j) < total_rows:
                _, t_next, _, _ = get_vals(table.rows[i + j])
                tnorm = _norm(t_next)
                if tnorm == 'questionprompt':
                    break
                if tnorm.startswith('radiobutton') and tnorm.endswith('normalstate'):
                    put_translation(table.rows[i + j], excel_answers[answers_filled])
                    answers_filled += 1
                j += 1

            # EXPLANATION: first Rounded Rectangular Caption after answers; skip copyright
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
                    break
                k += 1

            q_count += 1
            i = max(k, i + j, i + 1)
        else:
            i += 1

    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out, f"Processed {q_count} question(s) [Version A] with sheet '{sheet_name}'."

def scan_word_document_version_b(word_file, excel_file, sheet_name, question_limit):
    """
    Version B: Fill Prompt, Answers, and CORRECT FEEDBACK (second Text Box after answers).
    """
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
    except ValueError as e:
        raise RuntimeError(f"Excel sheet '{sheet_name}' was not found.") from e

    doc = Document(word_file)

    # Locate localization table and columns
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
    i = 1  # skip header row

    while i < total_rows and q_count < question_limit and q_count < len(df):
        _, type_text, _, _ = get_vals(table.rows[i])
        if _norm(type_text) == 'questionprompt':
            q_index = q_count + 1
            excel_prompt, excel_answers, _excel_expl = extract_prompt_answers_and_explanation(df.iloc[q_index - 1])
            correct_feedback = get_correct_feedback(df.iloc[q_index - 1])

            # PROMPT
            put_translation(table.rows[i], excel_prompt)

            # ANSWERS: next up to 4 Radio Button ... Normal state
            answers_needed = min(4, len(excel_answers))
            answers_filled = 0
            j = 1
            while answers_filled < answers_needed and (i + j) < total_rows:
                _, t_next, _, _ = get_vals(table.rows[i + j])
                tnorm = _norm(t_next)
                if tnorm == 'questionprompt':
                    break
                if tnorm.startswith('radiobutton') and tnorm.endswith('normalstate'):
                    put_translation(table.rows[i + j], excel_answers[answers_filled])
                    answers_filled += 1
                j += 1

            # CORRECT FEEDBACK:
            # After answers, skip copyright rows, then find two consecutive Text Box rows:
            #   first is the "Correct!" label, second is the feedback we should overwrite.
            k = i + j
            label_found = False
            while k < total_rows:
                _, t_k, _, _ = get_vals(table.rows[k])
                tknorm = _norm(t_k)

                if tknorm == 'questionprompt':  # next question -> stop block
                    break
                if tknorm == 'copyright':
                    k += 1
                    continue
                if tknorm == 'textbox':
                    if not label_found:
                        label_found = True  # this is the "Correct!" label row
                    else:
                        # this is the second Text Box -> the Correct feedback row
                        put_translation(table.rows[k], correct_feedback)
                        k += 1
                        break
                k += 1

            q_count += 1
            i = max(k, i + j, i + 1)
        else:
            i += 1

    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out, f"Processed {q_count} question(s) [Version B] with sheet '{sheet_name}'."

# ───────────────────────── Streamlit UI (keeps Option A / Option B) ─────────────────────────

st.title("Document Processor")

version_choice = st.selectbox("Select Version", ["Version A", "Version B"])

# Let the user type the Excel sheet/tab name (e.g., 1Q1)
sheet_name_input = st.text_input(
    "Excel sheet/tab name",
    value="1Q1",
    help="Type the exact tab name in your Excel file (e.g., 1Q1)."
)

word_file = st.file_uploader("Upload Word Document", type=["docx"])
excel_file = st.file_uploader("Upload Excel Document", type=["xlsx"])
question_limit = st.number_input("How many questions would you like to process?", min_value=1, value=10)

if word_file and excel_file and st.button("Process"):
    word_bytes = word_file.read()
    excel_bytes = excel_file.read()
    try:
        if version_choice == "Version A":
            output_buffer, output_message = scan_word_document_version_a(
                BytesIO(word_bytes), BytesIO(excel_bytes), sheet_name_input, question_limit
            )
        else:  # Version B
            output_buffer, output_message = scan_word_document_version_b(
                BytesIO(word_bytes), BytesIO(excel_bytes), sheet_name_input, question_limit
            )

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
        st.exception(e)
