import streamlit as st
import pandas as pd
from docx import Document
import re
import json
import hashlib
from io import BytesIO

# ───────────────────────── Helpers: header detection / normalization ─────────────────────────

def _norm(s: str) -> str:
    """Normalize table labels for robust matching (keep digits so RB1..RB4 stay distinct)."""
    return re.sub(r'[^a-z0-9]+', '', (s or '').lower())

def _get_header_indices(header_cells):
    """Return indices for id/type/sourcetext/translation or None."""
    labels = [_norm(c.text) for c in header_cells]
    need = {'id': 'id', 'type': 'type', 'sourcetext': 'sourcetext', 'translation': 'translation'}
    idx = {}
    for key, token in need.items():
        for i, lab in enumerate(labels):
            if token in lab:
                idx[key] = i
                break
    return idx if len(idx) == 4 else None

# ───────────────────────── Debug helpers ─────────────────────────

DEBUG_ENABLED = False
DEBUG_LOG = []
DEBUG_EVENT_LIMIT = 5000  # safety cap

def d(msg: str):
    if DEBUG_ENABLED and len(DEBUG_LOG) < DEBUG_EVENT_LIMIT:
        DEBUG_LOG.append(str(msg))

def _short(s: str, n: int = 120) -> str:
    s = s or ""
    return (s[:n] + "…") if len(s) > n else s

# ───────────────────────── Excel parsing (same as your original behavior) ─────────────────────────

def extract_prompt_answers_and_explanation(df_row):
    """Split df['Question'] at first 'A.'; explanation from df['Explanation']."""
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
    """Use chosen feedback column; fallback to 'Explanation'."""
    col = (feedback_column_name or 'Explanation').strip()
    if col in df_row.index:
        val = str(df_row.get(col, '') or '')
        if val.strip():
            return val
    return str(df_row.get('Explanation', '') or '')

# ───────────────────────── Version A (Type-anchored, fills Explanation) ─────────────────────────

def scan_word_document_version_a(word_file, excel_file, sheet_name, question_limit):
    d(f"[A] Loading Excel sheet: {sheet_name}")
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
    except ValueError as e:
        raise RuntimeError(f"Excel sheet '{sheet_name}' was not found.") from e

    doc = Document(word_file)

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

    d(f"[A] Header indices: {col}")

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
    d(f"[A] Total table rows: {total_rows}")
    q_count = 0
    i = 1  # skip header

    while i < total_rows and q_count < question_limit and q_count < len(df):
        _, type_text, _, _ = get_vals(table.rows[i])
        tnorm = _norm(type_text)
        if tnorm == 'questionprompt':
            d(f"[A] Anchor (Question Prompt) found at row {i}: type='{type_text}' norm='{tnorm}'")
            q_index = q_count + 1
            excel_prompt, excel_answers, excel_explanation = extract_prompt_answers_and_explanation(df.iloc[q_index - 1])
            d(f"[A] Q{q_index}: answers parsed={len(excel_answers)}")

            put_translation(table.rows[i], excel_prompt)
            d(f"[A]   Wrote PROMPT at row {i}")

            answers_needed = min(4, len(excel_answers))
            answers_filled = 0
            j = 1
            while answers_filled < answers_needed and (i + j) < total_rows:
                _, t_next, _, _ = get_vals(table.rows[i + j])
                tnorm_next = _norm(t_next)
                if tnorm_next == 'questionprompt':
                    d(f"[A]   Stop answers at row {i+j}: next Question Prompt")
                    break
                if tnorm_next.startswith('radiobutton') and tnorm_next.endswith('normalstate'):
                    put_translation(table.rows[i + j], excel_answers[answers_filled])
                    d(f"[A]   Wrote ANSWER {answers_filled+1} at row {i+j} (type='{t_next}', norm='{tnorm_next}')")
                    answers_filled += 1
                j += 1

            # Explanation after answers
            k = i + j
            while k < total_rows:
                _, t_k, _, _ = get_vals(table.rows[k])
                tknorm = _norm(t_k)
                if tknorm == 'copyright':
                    k += 1
                    continue
                if tknorm == 'roundedrectangularcaption':
                    put_translation(table.rows[k], excel_explanation)
                    d(f"[A]   Wrote EXPLANATION at row {k}")
                    k += 1
                    break
                if tknorm == 'questionprompt':
                    d(f"[A]   Stop explanation at row {k}: next Question Prompt")
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
    d(f"[B] Loading Excel sheet: {sheet_name}")
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
    except ValueError as e:
        raise RuntimeError(f"Excel sheet '{sheet_name}' was not found.") from e

    doc = Document(word_file)

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

    d(f"[B] Header indices: {col}")

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
    d(f"[B] Total table rows: {total_rows}")
    q_count = 0
    i = 1  # skip header

    while i < total_rows and q_count < question_limit and q_count < len(df):
        _, type_text, _, _ = get_vals(table.rows[i])
        tnorm = _norm(type_text)
        if tnorm == 'questionprompt':
            d(f"[B] Anchor (Question Prompt) found at row {i}: type='{type_text}' norm='{tnorm}'")
            q_index = q_count + 1
            excel_prompt, excel_answers, _excel_expl = extract_prompt_answers_and_explanation(df.iloc[q_index - 1])
            correct_feedback = get_feedback_value(df.iloc[q_index - 1], 'Correct')
            d(f"[B] Q{q_index}: answers parsed={len(excel_answers)}")

            put_translation(table.rows[i], excel_prompt)
            d(f"[B]   Wrote PROMPT at row {i}")

            answers_needed = min(4, len(excel_answers))
            answers_filled = 0
            j = 1
            while answers_filled < answers_needed and (i + j) < total_rows:
                _, t_next, _, _ = get_vals(table.rows[i + j])
                tnorm_next = _norm(t_next)
                if tnorm_next == 'questionprompt':
                    d(f"[B]   Stop answers at row {i+j}: next Question Prompt")
                    break
                if tnorm_next.startswith('radiobutton') and tnorm_next.endswith('normalstate'):
                    put_translation(table.rows[i + j], excel_answers[answers_filled])
                    d(f"[B]   Wrote ANSWER {answers_filled+1} at row {i+j} (type='{t_next}', norm='{tnorm_next}')")
                    answers_filled += 1
                j += 1

            # Correct feedback rows (Text Box x 2)
            k = i + j
            label_found = False
            while k < total_rows:
                _, t_k, _, _ = get_vals(table.rows[k])
                tknorm = _norm(t_k)
                if tknorm == 'questionprompt':
                    d(f"[B]   Stop feedback at row {k}: next Question Prompt")
                    break
                if tknorm == 'copyright':
                    k += 1
                    continue
                if tknorm == 'textbox':
                    if not label_found:
                        label_found = True
                        d(f"[B]   Found 'Correct!' label at row {k}")
                    else:
                        put_translation(table.rows[k], correct_feedback)
                        d(f"[B]   Wrote CORRECT FEEDBACK at row {k}")
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

# ───────────────────────── Universal (anchor + configurable offsets) ─────────────────────────

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
    d(f"[U] Loading Excel sheet: {sheet_name}")
    # Allow prompt_offset to be negative; keep others non-negative
    if answer_offset < 0 or explanation_offset < 0:
        raise RuntimeError("Answer and explanation offsets must be non-negative. Prompt offset may be negative.")

    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
    except ValueError as e:
        raise RuntimeError(f"Excel sheet '{sheet_name}' was not found.") from e

    if 'Question' not in df.columns:
        raise RuntimeError("Excel sheet must contain a 'Question' column.")

    doc = Document(word_file)

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

    d(f"[U] Header indices: {col}")

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
    d(f"[U] Anchor config: anchor_type_value='{anchor_type_value}', anchor_norm='{anchor_norm}', "
      f"keyword='{anchor_keyword}', offsets: P={prompt_offset}, A={answer_offset}, E={explanation_offset}")

    total_rows = len(table.rows)
    d(f"[U] Total table rows: {total_rows}")
    q_count = 0
    i = 1  # skip header

    def is_radio_normal(tnorm: str) -> bool:
        return tnorm.startswith('radiobutton') and tnorm.endswith('normalstate')

    def find_block_end(start_idx: int) -> int:
        """End is the next anchor (RB1) after start_idx, else end of table."""
        k = start_idx + 1
        while k < total_rows:
            _, t_k, _, _ = get_vals(table.rows[k])
            if _norm(t_k) == anchor_norm:  # RB1 of next question
                return k
            k += 1
        return total_rows

    while i < total_rows and q_count < question_limit and q_count < len(df):
        id_text, type_text, source_text, trans_text = get_vals(table.rows[i])
        tnorm = _norm(type_text)

        if q_count == 0 and i < 10:
            d(f"[U] row {i}: type='{_short(type_text)}' norm='{tnorm}', source='{_short(source_text)}'")

        if tnorm == anchor_norm:
            d(f"[U] Anchor FOUND at row {i}: type='{type_text}', norm='{tnorm}', source='{_short(source_text)}'")

            if anchor_keyword and anchor_keyword.strip():
                if anchor_keyword.lower() not in source_text.lower():
                    d(f"[U]   Skipped anchor at row {i}: keyword '{anchor_keyword}' not in source.")
                    i += 1
                    continue

            block_end = find_block_end(i)
            d(f"[U]   Block end at row {block_end}")

            q_index = q_count + 1
            df_row = df.iloc[q_index - 1]
            excel_prompt, excel_answers, _excel_expl = extract_prompt_answers_and_explanation(df_row)
            feedback_text = get_feedback_value(df_row, feedback_column_name)
            d(f"[U]   Q{q_index}: answers parsed={len(excel_answers)}")

            last_written = i

            # PROMPT
            p_idx = i + prompt_offset
            if 0 <= p_idx < total_rows and p_idx < block_end:
                put_translation(table.rows[p_idx], excel_prompt)
                d(f"[U]   Wrote PROMPT at row {p_idx} (offset {prompt_offset})")
                last_written = max(last_written, p_idx)
            else:
                d(f"[U]   PROMPT target out of bounds or past block end: row {p_idx}")

            # ANSWERS: collect only radio-button rows within the block starting at (i + answer_offset)
            answer_rows = []
            k = i + answer_offset
            if not (0 <= k < total_rows):
                d(f"[U]   Answer start out of bounds: row {k}")
            else:
                while k < block_end and len(answer_rows) < len(excel_answers):
                    _, t_k, _, _ = get_vals(table.rows[k])
                    tkn = _norm(t_k)
                    if is_radio_normal(tkn):
                        answer_rows.append(k)
                    elif len(answer_rows) > 0:
                        # Once the contiguous radio block ends, stop collecting
                        break
                    k += 1

            d(f"[U]   Answer rows detected: {answer_rows}")

            for a_idx, ans in enumerate(excel_answers):
                if a_idx >= len(answer_rows):
                    d(f"[U]   Not enough radio-button rows for ANSWER {a_idx+1}; stopping")
                    break
                tgt = answer_rows[a_idx]
                put_translation(table.rows[tgt], ans)
                d(f"[U]   Wrote ANSWER {a_idx+1} at row {tgt}")
                last_written = max(last_written, tgt)

            # FEEDBACK within block
            f_idx = i + explanation_offset
            if 0 <= f_idx < total_rows and f_idx < block_end:
                if f_idx == i and (answer_offset == 0):
                    d(f"[U]   WARNING: feedback row {f_idx} == anchor and answer_offset=0 → feedback may overwrite Answer 1")
                put_translation(table.rows[f_idx], feedback_text)
                d(f"[U]   Wrote FEEDBACK at row {f_idx} (offset {explanation_offset})")
                last_written = max(last_written, f_idx)
            else:
                d(f"[U]   FEEDBACK target out of bounds or past block end: row {f_idx}")

            q_count += 1
            i = max(last_written + 1, i + 1)
        else:
            i += 1

    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out, f"Processed {q_count} question(s) [Universal] with sheet '{sheet_name}'."

# ───────────────────────── Preset helpers (local-only JSON) ─────────────────────────

PRESET_SCHEMA_VERSION = 1

DEFAULTS = {
    "version_choice": "Version A",
    "sheet_name_input": "1Q1",
    "anchor_type_value": "Question Prompt",
    "anchor_keyword": "",
    "prompt_offset": 0,
    "answer_offset": 1,
    "explanation_offset": 6,
    "feedback_column_name": "Explanation",
    "preset_name": ""
}

def ensure_defaults():
    for k, v in DEFAULTS.items():
        st.session_state.setdefault(k, v)

def apply_preset_dict(preset: dict):
    """Apply preset into session_state (called BEFORE widgets render)."""
    st.session_state["version_choice"] = preset.get("mode", DEFAULTS["version_choice"])
    st.session_state["sheet_name_input"] = preset.get("sheet_name", DEFAULTS["sheet_name_input"])
    st.session_state["anchor_type_value"] = preset.get("anchor_type_value", DEFAULTS["anchor_type_value"])
    st.session_state["anchor_keyword"] = preset.get("anchor_keyword", DEFAULTS["anchor_keyword"])
    st.session_state["prompt_offset"] = int(preset.get("prompt_offset", DEFAULTS["prompt_offset"]) or 0)
    st.session_state["answer_offset"] = int(preset.get("answer_offset", DEFAULTS["answer_offset"]) or 1)
    st.session_state["explanation_offset"] = int(preset.get("explanation_offset", DEFAULTS["explanation_offset"]) or 6)
    st.session_state["feedback_column_name"] = preset.get("feedback_column_name", DEFAULTS["feedback_column_name"])
    st.session_state["preset_name"] = preset.get("name", DEFAULTS["preset_name"])

def current_config_as_preset():
    """Collect the current UI state into a preset dict."""
    return {
        "name": st.session_state.get("preset_name", "").strip() or "unnamed",
        "schema_version": PRESET_SCHEMA_VERSION,
        "mode": st.session_state.get("version_choice", DEFAULTS["version_choice"]),
        "sheet_name": st.session_state.get("sheet_name_input", DEFAULTS["sheet_name_input"]),
        "anchor_type_value": st.session_state.get("anchor_type_value", DEFAULTS["anchor_type_value"]),
        "anchor_keyword": st.session_state.get("anchor_keyword", DEFAULTS["anchor_keyword"]),
        "prompt_offset": int(st.session_state.get("prompt_offset", DEFAULTS["prompt_offset"]) or 0),
        "answer_offset": int(st.session_state.get("answer_offset", DEFAULTS['answer_offset']) or 1),
        "explanation_offset": int(st.session_state.get("explanation_offset", DEFAULTS["explanation_offset"]) or 6),
        "feedback_column_name": st.session_state.get("feedback_column_name", DEFAULTS["feedback_column_name"]),
    }

# ───────────────────────── Sidebar: Presets FIRST (fast + re-apply) ─────────────────────────

st.sidebar.header("Presets")

# Debug controls (no behavior change, only logging)
DEBUG_ENABLED = st.sidebar.checkbox("Enable debug logging", value=False)
DEBUG_EVENT_LIMIT = st.sidebar.number_input("Max debug lines", min_value=200, value=5000, step=100)

# One-time defaults
ensure_defaults()

# Load preset JSON (top of app, applies before main widgets)
uploaded = st.sidebar.file_uploader("Load preset JSON", type=["json"], key="preset_uploader_top")

# Clear the last token when no file is selected so remove → re-add will re-apply
if uploaded is None:
    st.session_state.pop("last_preset_token", None)

apply_now = st.sidebar.button("Apply/Refresh preset now")

if uploaded is not None:
    raw = uploaded.getvalue()
    token = hashlib.md5(raw).hexdigest()
    last = st.session_state.get("last_preset_token")

    # Apply if it's a new file OR the user explicitly requests a refresh
    if (token != last) or apply_now:
        try:
            data = json.loads(raw)
            apply_preset_dict(data)
            st.session_state["last_preset_token"] = token
            st.sidebar.success("Preset applied.")
        except Exception as e:
            st.sidebar.error(f"Failed to load preset: {e}")

# Download current preset
preset = current_config_as_preset()
st.sidebar.download_button(
    label="Download current preset JSON",
    data=json.dumps(preset, indent=2).encode("utf-8"),
    file_name=f"{preset['name'] or 'preset'}.json",
    mime="application/json"
)
st.sidebar.caption("Presets are local JSON files. Upload to apply instantly. Use 'Apply/Refresh' to re-apply the same file.")

# ───────────────────────── Main UI ─────────────────────────

st.title("Document Processor — Multi-Version (Debug Instrumented)")

# Version selector
version_choice = st.selectbox(
    "Choose processor",
    ["Version A", "Version B", "Universal"],
    key="version_choice",
    help="Universal adds configurable anchor/offsets. A and B use fixed row mappings."
)

# Sheet/tab
sheet_name_input = st.text_input(
    "Excel sheet/tab name",
    value=st.session_state.get("sheet_name_input", "1Q1"),
    key="sheet_name_input",
    help="Type the exact tab name in your Excel file (e.g., 1Q1)."
)

# Universal-only fields
if version_choice == "Universal":
    st.subheader("Universal Configuration")
    st.text_input(
        "Anchor Type value (exact text under the Type column that marks a question block)",
        value=st.session_state.get("anchor_type_value", "Question Prompt"),
        key="anchor_type_value"
    )
    st.text_input(
        "Optional keyword to validate the anchor in Source Text (case-insensitive)",
        value=st.session_state.get("anchor_keyword", ""),
        key="anchor_keyword"
    )
    st.number_input(
        "Rows after ANCHOR where the PROMPT lives (can be negative)",
        value=int(st.session_state.get("prompt_offset", 0) or 0),
        key="prompt_offset"
    )
    st.number_input(
        "Rows after ANCHOR where the FIRST ANSWER lives",
        min_value=0,
        value=int(st.session_state.get("answer_offset", 1) or 1),
        key="answer_offset"
    )
    st.number_input(
        "Rows after ANCHOR where the EXPLANATION / FEEDBACK lives",
        min_value=0,
        value=int(st.session_state.get("explanation_offset", 6) or 6),
        key="explanation_offset"
    )
    st.text_input(
        "Excel column to use for the final feedback row",
        value=st.session_state.get("feedback_column_name", "Explanation"),
        key="feedback_column_name"
    )

# Files & controls (common)
word_file = st.file_uploader("Upload Word Document (.docx)", type=["docx"])
excel_file = st.file_uploader("Upload Excel Document (.xlsx)", type=["xlsx"])
question_limit = st.number_input("How many questions would you like to process?", min_value=1, value=10)

# Process
if word_file and excel_file and st.button("Process"):
    # reset debug buffer
    if DEBUG_ENABLED:
        DEBUG_LOG.clear()
        d("=== DEBUG START ===")

    word_bytes = word_file.read()
    excel_bytes = excel_file.read()
    try:
        if version_choice == "Version A":
            output_buffer, output_message = scan_word_document_version_a(
                BytesIO(word_bytes), BytesIO(excel_bytes), st.session_state["sheet_name_input"], int(question_limit)
            )
        elif version_choice == "Version B":
            output_buffer, output_message = scan_word_document_version_b(
                BytesIO(word_bytes), BytesIO(excel_bytes), st.session_state["sheet_name_input"], int(question_limit)
            )
        else:
            output_buffer, output_message = scan_word_document_universal(
                BytesIO(word_bytes),
                BytesIO(excel_bytes),
                st.session_state["sheet_name_input"],
                st.session_state.get("anchor_type_value", "Question Prompt"),
                st.session_state.get("anchor_keyword", ""),
                int(st.session_state.get("prompt_offset", 0) or 0),
                int(st.session_state.get("answer_offset", 1) or 1),
                int(st.session_state.get("explanation_offset", 6) or 6),
                int(question_limit),
                st.session_state.get("feedback_column_name", "Explanation")
            )

        st.write(output_message)
        st.success("Processing complete. Download the updated Word document below.")
        st.download_button(
            label="Download updated document",
            data=output_buffer,
            file_name="updated_document.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

        if DEBUG_ENABLED:
            with st.expander("Debug log"):
                st.code("\n".join(DEBUG_LOG) if DEBUG_LOG else "(no debug entries)")
    except RuntimeError as e:
        st.error(str(e))
        if DEBUG_ENABLED:
            with st.expander("Debug log"):
                st.code("\n".join(DEBUG_LOG) if DEBUG_LOG else "(no debug entries)")
    except Exception as e:
        st.exception(e)
        if DEBUG_ENABLED:
            with st.expander("Debug log"):
                st.code("\n".join(DEBUG_LOG) if DEBUG_LOG else "(no debug entries)")
