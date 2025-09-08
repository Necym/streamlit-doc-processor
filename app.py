import streamlit as st
import pandas as pd
from docx import Document
import re
import json
from io import BytesIO

# ───────────────────────── Helpers: header detection / normalization ─────────────────────────

def _norm(s: str) -> str:
    """Normalize table labels for robust matching."""
    return re.sub(r'[^a-z]+', '', (s or '').lower())

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

# ───────────────────────── Excel parsing (same behavior as your old code) ─────────────────────────

def extract_prompt_answers_and_explanation(df_row):
    """Old behavior: split df['Question'] at first 'A.'; explanation from df['Explanation']."""
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
    i = 1  # skip header

    while i < total_rows and q_count < question_limit and q_count < len(df):
        _, type_text, _, _ = get_vals(table.rows[i])
        if _norm(type_text) == 'questionprompt':
            q_index = q_count + 1
            excel_prompt, excel_answers, excel_explanation = extract_prompt_answers_and_explanation(df.iloc[q_index - 1])

            # Prompt
            put_translation(table.rows[i], excel_prompt)

            # Answers (next up to 4 "Radio Button ... Normal state")
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

            # Explanation (first "Rounded Rectangular Caption" after answers; skip copyright)
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
    i = 1  # skip header

    while i < total_rows and q_count < question_limit and q_count < len(df):
        _, type_text, _, _ = get_vals(table.rows[i])
        if _norm(type_text) == 'questionprompt':
            q_index = q_count + 1
            excel_prompt, excel_answers, _excel_expl = extract_prompt_answers_and_explanation(df.iloc[q_index - 1])
            correct_feedback = get_feedback_value(df.iloc[q_index - 1], 'Correct')

            # Prompt
            put_translation(table.rows[i], excel_prompt)

            # Answers
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

            # Correct feedback: skip copyright; then two Text Box rows ("Correct!" label, then feedback)
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
                        label_found = True
                    else:
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
    if prompt_offset < 0 or answer_offset < 0 or explanation_offset < 0:
        raise RuntimeError("Offsets must be non-negative integers (>= 0).")

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
            # Optional keyword check
            if anchor_keyword and anchor_keyword.strip():
                if anchor_keyword.lower() not in source_text.lower():
                    i += 1
                    continue

            q_index = q_count + 1
            df_row = df.iloc[q_index - 1]
            excel_prompt, excel_answers, _excel_expl = extract_prompt_answers_and_explanation(df_row)
            feedback_text = get_feedback_value(df_row, feedback_column_name)

            last_written = i

            # Prompt at anchor + offset
            p_idx = i + prompt_offset
            if 0 <= p_idx < total_rows:
                _, t_p, _, _ = get_vals(table.rows[p_idx])
                if p_idx == i or _norm(t_p) != anchor_norm:
                    put_translation(table.rows[p_idx], excel_prompt)
                    last_written = max(last_written, p_idx)

            # Answers start at anchor + answer_offset
            for a_idx, ans in enumerate(excel_answers):
                tgt = i + answer_offset + a_idx
                if not (0 <= tgt < total_rows):
                    break
                _, t_tgt, _, _ = get_vals(table.rows[tgt])
                if tgt != i and _norm(t_tgt) == anchor_norm:
                    break
                put_translation(table.rows[tgt], ans)
                last_written = max(last_written, tgt)

            # Feedback/explanation at anchor + explanation_offset
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

# ───────────────────────── Preset helpers ─────────────────────────

PRESET_SCHEMA_VERSION = 1

def current_config_as_preset():
    """Collect the current UI state into a preset dict."""
    return {
        "name": st.session_state.get("preset_name", "").strip() or "unnamed",
        "schema_version": PRESET_SCHEMA_VERSION,
        "mode": st.session_state.get("version_choice", "Version A"),
        "sheet_name": st.session_state.get("sheet_name_input", "1Q1"),
        # Universal fields (kept even for A/B; harmless if unused)
        "anchor_type_value": st.session_state.get("anchor_type_value", "Question Prompt"),
        "anchor_keyword": st.session_state.get("anchor_keyword", ""),
        "prompt_offset": int(st.session_state.get("prompt_offset", 0) or 0),
        "answer_offset": int(st.session_state.get("answer_offset", 1) or 1),
        "explanation_offset": int(st.session_state.get("explanation_offset", 6) or 6),
        "feedback_column_name": st.session_state.get("feedback_column_name", "Explanation"),
    }

def apply_preset_to_state(preset: dict):
    """Apply a preset into st.session_state and rerun."""
    if not isinstance(preset, dict):
        st.error("Invalid preset format.")
        return
    # Minimal validation
    mode = preset.get("mode", "Version A")
    st.session_state["version_choice"] = mode
    st.session_state["sheet_name_input"] = preset.get("sheet_name", "1Q1")

    st.session_state["anchor_type_value"] = preset.get("anchor_type_value", "Question Prompt")
    st.session_state["anchor_keyword"] = preset.get("anchor_keyword", "")
    st.session_state["prompt_offset"] = int(preset.get("prompt_offset", 0) or 0)
    st.session_state["answer_offset"] = int(preset.get("answer_offset", 1) or 1)
    st.session_state["explanation_offset"] = int(preset.get("explanation_offset", 6) or 6)
    st.session_state["feedback_column_name"] = preset.get("feedback_column_name", "Explanation")

    # Name (optional, for UI)
    st.session_state["preset_name"] = preset.get("name", "")

    st.rerun()

# Initialize session containers
if "saved_presets" not in st.session_state:
    st.session_state["saved_presets"] = {}  # name -> dict
if "preset_name" not in st.session_state:
    st.session_state["preset_name"] = ""

# ───────────────────────── UI ─────────────────────────

st.title("Document Processor — Multi-Version with Presets")

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
        "Rows after ANCHOR where the PROMPT lives",
        min_value=0,
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

# Presets block (always visible)
with st.expander("Presets (Save / Load)"):
    st.caption("💡 Presets saved to *session* appear in the dropdown below. For long-term use, download JSON and re-upload later.")
    # Save current as preset
    st.text_input("Preset name", key="preset_name", placeholder="e.g., Storyline-Q1-Universal")
    cols = st.columns(2)
    with cols[0]:
        if st.button("Save to session presets"):
            preset = current_config_as_preset()
            st.session_state["saved_presets"][preset["name"]] = preset
            st.success(f"Saved preset '{preset['name']}' to session.")
    with cols[1]:
        # Download current config as JSON
        preset = current_config_as_preset()
        json_bytes = json.dumps(preset, indent=2).encode("utf-8")
        st.download_button(
            label="Download current preset JSON",
            data=json_bytes,
            file_name=f"{preset['name'] or 'preset'}.json",
            mime="application/json"
        )

    # Apply a session preset
    if st.session_state["saved_presets"]:
        names = sorted(st.session_state["saved_presets"].keys())
        sel = st.selectbox("Apply a session preset", names, key="apply_preset_select")
        if st.button("Apply selected preset"):
            apply_preset_to_state(st.session_state["saved_presets"][sel])

    # Load from uploaded JSON
    uploaded = st.file_uploader("Load preset JSON", type=["json"], key="preset_uploader")
    if uploaded:
        try:
            data = json.loads(uploaded.getvalue())
            st.success("Preset loaded. Applying …")
            apply_preset_to_state(data)
        except Exception as e:
            st.error(f"Failed to load preset: {e}")

# Files & controls (common)
word_file = st.file_uploader("Upload Word Document (.docx)", type=["docx"])
excel_file = st.file_uploader("Upload Excel Document (.xlsx)", type=["xlsx"])
question_limit = st.number_input("How many questions would you like to process?", min_value=1, value=10)

# Process
if word_file and excel_file and st.button("Process"):
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
    except RuntimeError as e:
        st.error(str(e))
    except Exception as e:
        st.exception(e)
