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
      - Explanation comes from df['Explanation'] (or other column if chosen)
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

# ───────────────────────── Universal processor (anchor + configurable offsets) ─────────────────────────

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
    # Validate offsets (allow zero for prompt; answers/feedback may also be zero if desired)
    if prompt_offset < 0 or answer_offset < 0 or explanation_offset < 0:
        raise RuntimeError("Offsets must be non-negative integers (>= 0).")

    # Load Excel
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
    except ValueError as e:
        raise RuntimeError(f"Excel sheet '{sheet_name}' was not found.") from e

    if 'Question' not in df.columns:
        raise RuntimeError("Excel sheet must contain a 'Question' column.")

    # Load Word
    doc = Document(word_file)

    # Find the localization table & columns
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
    i = 1  # skip header row

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

            last_written = i  # track the farthest row we wrote to (to advance safely)

            # ── PROMPT ──
            p_idx = i + prompt_offset
            if 0 <= p_idx < total_rows:
                # Allow writing to anchor itself (prompt_offset == 0), but prevent writing onto the NEXT anchor
                _, t_p, _, _ = get_vals(table.rows[p_idx])
                if p_idx == i or _norm(t_p) != anchor_norm:
                    put_translation(table.rows[p_idx], excel_prompt)
                    last_written = max(last_written, p_idx)

            # ── ANSWERS ──
            for a_idx, ans in enumerate(excel_answers):
                tgt = i + answer_offset + a_idx
                if not (0 <= tgt < total_rows):
                    break
                _, t_tgt, _, _ = get_vals(table.rows[tgt])
                # Do not overwrite the next anchor
                if tgt != i and _norm(t_tgt) == anchor_norm:
                    break
                put_translation(table.rows[tgt], ans)
                last_written = max(last_written, tgt)

            # ── FEEDBACK / EXPLANATION ──
            f_idx = i + explanation_offset
            if 0 <= f_idx < total_rows:
                _, t_f, _, _ = get_vals(table.rows[f_idx])
                if f_idx == i or _norm(t_f) != anchor_norm:
                    put_translation(table.rows[f_idx], feedback_text)
                    last_written = max(last_written, f_idx)

            q_count += 1
            # Advance beyond what we wrote (at least move one row)
            i = max(last_written + 1, i + 1)
        else:
            i += 1

    # Return updated .docx
    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out, f"Processed {q_count} question(s) with sheet '{sheet_name}'."

# ───────────────────────── Streamlit UI (Universal) ─────────────────────────

st.title("Document Processor — Universal (Config-Driven)")

# Excel tab name
sheet_name_input = st.text_input(
    "Excel sheet/tab name",
    value="1Q1",
    help="Type the exact tab name in your Excel file (e.g., 1Q1)."
)

# Config: anchor finding
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

# Offsets (relative to the ANCHOR row)
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

# Files
word_file = st.file_uploader("Upload Word Document (.docx)", type=["docx"])
excel_file = st.file_uploader("Upload Excel Document (.xlsx)", type=["xlsx"])

# Controls
question_limit = st.number_input(
    "How many questions would you like to process?",
    min_value=1,
    value=10
)

if word_file and excel_file and st.button("Process"):
    word_bytes = word_file.read()
    excel_bytes = excel_file.read()
    try:
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
