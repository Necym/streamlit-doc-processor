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

# ───────────────────────── Excel parsing (same behavior as your old code) ─────────────────────────

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

def get_feedback_value(df_row, feedback_column_name: str):
    """
    Returns the desired feedback text from the chosen Excel column.
    Defaults to 'Explanation' when the specified column is missing/empty.
    """
    col = (feedback_column_name or 'Explanation').strip()
    if col in df_row.index:
        val = str(df_row.get(col, '') or '')
        if val.strip():
            return val
    # Fallback
    return str(df_row.get('Explanation', '') or '')

# ───────────────────────── Version A (Type-anchored, fills Explanation) ─────────────────────────

def scan_word_document_version_a(word_file, excel_file, sheet_name, question_limit):
    """
    Version A: Fill Prompt, Answers, and EXPLANATION ('Rounded Rectangular Caption')
    using the fixed, Type-anchored sequential mapping.
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

# ───────────────────────── Version B (Type-anchored, fills Correct feedback) ─────────────────────────

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
            correct_feedback = get_feedback_value(df.iloc[q_index - 1], 'Correct')

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

                if tknorm == 'questionprompt':
                    break
                if tknorm == 'copyright':
                    k += 1
                    continue
                if tknorm == 'textbox':
                    if not label_found:
                        label_found = True  # the "Correct!" label row
                    else:
                        # feedback row
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

# ───────────────────────── Universal (config-driven: anchor + offsets) ─────────────────────────

def scan_word_document_universal(word_file,
                                 excel_file,
                                 sheet_name: str,
                                 anchor_type_value: str,
                                 anchor_keyword: str,
                                 prompt_offset: int,
                                 answer_offset: int,
                                 explanation_offset: int,
                                 question_limit: int,
                                 feedback_column_name: str = 'Explanation'):
    """
    Universal mapping:
      - Find anchor rows where Type matches 'anchor_type_value' (normalized exact match).
      - If 'anchor_keyword' is provided, ALSO require that Source Text contains it (case-insensitive).
      - PROMPT at:         row (anchor + prompt_offset)
      - ANSWERS start at:  row (anchor + answer_offset), then consecutive rows for each answer
      - FEEDBACK at:       row (anchor + explanation_offset), from Excel column 'feedback_column_name'
      - Questions map sequentially: the first anchor -> Excel row 1, second -> row 2, ...
      - Only writes to Translation column.
      - Safeguard: don't overwrite the next anchor (except when writing prompt to the anchor itself if prompt_offset == 0).
    """
    if prompt_offset < 0 or answer_offset < 0 or explanation_offset < 0:
        raise RuntimeError("Offsets must be non-negative integers (>= 0).")

    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
    except ValueError as e:
        raise RuntimeError(f"Excel sheet '{sheet_name}' was not found.") from e

    if 'Question' not in df.columns:
        raise RuntimeError("Excel sheet must contain a 'Question' column.")

    doc = Document(word_file)

    # Find localization table
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

    anchor_norm = _norm(anchor_type_value)
    total_rows = len(table.rows)
    q_count = 0
    i = 1  # skip header

    while i < total_rows and q_count < question_limit and q_count < len(df):
        _, type_text, source_text, _ = get_vals(table.rows[i])
        if _norm(type_text) == anchor_norm:
            # Optional keyword validation (Source Text contains keyword, case-insensitive)
            if anchor_keyword and anchor_keyword.strip():
                if anchor_keyword.lower() not in source_text.lower():
                    i += 1
                    continue

            # Map to Excel row (1-based)
            q_index = q_count + 1
            df_row = df.iloc[q_index - 1]
            excel_prompt, excel_answers, _excel_expl = extract_prompt_answers_and_explanation(df_row)
            feedback_text = get_feedback_value(df_row, feedback_column_name)

            last_written = i

            # PROMPT
            p_idx = i + prompt_offset
            if 0 <= p_idx < total_rows:
                _, t_p, _, _ = get_vals(table.rows[p_idx])
                if p_idx == i or _norm(t_p) != anchor_norm:
                    put_translation(table.rows[p_idx], excel_prompt)
                    last_written = max(last_written, p_idx)

            # ANSWERS
            for a_idx, ans in enumerate(excel_answers):
                tgt = i + answer_offset + a_idx
                if not (0 <= tgt < total_rows):
                    break
                _, t_tgt, _, _ = get_vals(table.rows[tgt])
                if tgt != i and _norm(t_tgt) == anchor_norm:
                    break
                put_translation(table.rows[tgt], ans)
                last_written = max(last_written, tgt)

            # FEEDBACK / EXPLANATION
            f_idx = i + explanation_offset
            if 0 <= f_idx < total_rows:
                _, t_f, _, _ = get_vals(table.rows[f_idx])
                if f_idx == i or _norm(t_f) != anchor_norm:
                    put_translation(table.rows[f_idx], feedback_text)
                    last_written = max(last_written, f_idx)

            q_count += 1
            i = max(last_written + 1, i + 1)
        else:
            i += 1

    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out, f"Processed {q_count} question(s) [Universal] with sheet '{sheet_name}'."

# ───────────────────────── Streamlit UI ─────────────────────────

st.title("Document Processor — Multi-Version")

version_choice = st.selectbox(
    "Choose processor",
    ["Version A", "Version B", "Universal"],
    help="Universal adds configurable anchor/offsets. A and B use fixed row mappings."
)

# Always: Excel sheet/tab
sheet_name_input = st.text_input(
    "Excel sheet/tab name",
    value="1Q1",
    help="Type the exact tab name in your Excel file (e.g., 1Q1)."
)

# Show Universal-specific config only when selected
if version_choice == "Universal":
    st.subheader("Universal Configuration")
    anchor_type_value = st.text_input(
        "Anchor Type value (exact text under the Type column that marks a question block)",
        value="Question Prompt",
        help='Example: "Question Prompt" or any Type label that reliably appears per question.'
    )

    anchor_keyword = st.text_input(
        "Optional keyword to validate the anchor in Source Text (case-insensitive)",
        value="",
        help="Leave blank to skip keyword validation."
    )

    prompt_offset = st.number_input(
        "Rows after ANCHOR where the PROMPT lives",
        min_value=0,
        value=0,
        help="0 means the prompt is on the anchor row; 2 means at i+2, etc."
    )

    answer_offset = st.number_input(
        "Rows after ANCHOR where the FIRST ANSWER lives",
        min_value=0,
        value=1,
        help="1 means first answer at i+1, then i+2, i+3..."
    )

    explanation_offset = st.number_input(
        "Rows after ANCHOR where the EXPLANATION / FEEDBACK lives",
        min_value=0,
        value=6,
        help="6 means explanation/feedback at i+6."
    )

    feedback_column_name = st.text_input(
        "Excel column to use for the final feedback row",
        value="Explanation",
        help="Use 'Explanation' (default) or another column name like 'Correct'."
    )
else:
    # Provide dummies so the process block can reference them safely
    anchor_type_value = ""
    anchor_keyword = ""
    prompt_offset = 0
    answer_offset = 0
    explanation_offset = 0
    feedback_column_name = "Explanation"

# Files and controls (common)
word_file = st.file_uploader("Upload Word Document (.docx)", type=["docx"])
excel_file = st.file_uploader("Upload Excel Document (.xlsx)", type=["xlsx"])
question_limit = st.number_input(
    "How many questions would you like to process?",
    min_value=1,
    value=10
)

if word_file and excel_file and st.button("Process"):
    word_bytes = word_file.read()
    excel_bytes = excel_file.read()
    try:
        if version_choice == "Version A":
            output_buffer, output_message = scan_word_document_version_a(
                BytesIO(word_bytes), BytesIO(excel_bytes), sheet_name_input, int(question_limit)
            )
        elif version_choice == "Version B":
            output_buffer, output_message = scan_word_document_version_b(
                BytesIO(word_bytes), BytesIO(excel_bytes), sheet_name_input, int(question_limit)
            )
        else:
            output_buffer, output_message = scan_word_document_universal(
                BytesIO(word_bytes),
                BytesIO(excel_bytes),
                sheet_name_input,
                anchor_type_value,
                anchor_keyword,
                int(prompt_offset),
                int(answer_offset),
                int(explanation_offset),
                int(question_limit),
                feedback_column_name.strip() or 'Explanation'
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
