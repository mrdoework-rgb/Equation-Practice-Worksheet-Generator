"""GCSE Physics formula worksheet generator (Streamlit)."""
import io
import json
import math
import os
import random
import re
import tempfile
import urllib.error
import urllib.request
from fractions import Fraction
from pathlib import Path
from uuid import uuid4

import streamlit as st
import streamlit.components.v1 as components
from docx import Document
from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, Twips
from docx.text.paragraph import Paragraph
from physics_data import (EQUATIONS, UNIT_GROUPS, UNIT_INFO, fmt_num, is_exact_to_sig,
                          normalise_text, numbers_in_text, plain_symbol, safe_eval, units_in_text)

DEFAULT_TEMPLATE = Path(__file__).parent / "Formulae_Questions_worksheet_template.docx"
NBSP = " "
BLANK = "________"
QUESTION_PT = 8
MAX_ATTEMPTS = 3
FORBIDDEN_CONVERSION_UNITS = {"km/h", "h"}
SECTIONS = [
    dict(key="substitute", placeholder="substitute", heading="Substitute", n=8, kind="single", find="subject", conv=0, answer_line="inline"),
    dict(key="unit_conversions", placeholder="unit conversions", heading="Unit Conversions", n=6, kind="generated", answer_line="inline"),
    dict(key="substitute_unit_conversions", placeholder="substitute unit conversions", heading="Substitute (convert units)", n=5, kind="single", find="subject", conv=1, answer_line="inline"),
    dict(key="rearrange", placeholder="rearrange", heading="Rearrange", n=4, kind="single", find="other", conv=0, answer_line="inline"),
    dict(key="rearrange_unit_conversions", placeholder="rearrange unit conversions", heading="Rearrange (convert units)", n=5, kind="single", find="other", conv=1, answer_line="inline"),
    dict(key="word_questions", placeholder="word questions", heading="Word Questions", n=5, kind="single", find="any", conv=None, answer_line="one"),
    dict(key="multistep_questions", placeholder="multistep questions", heading="Multistep Questions", n=4, kind="multi", answer_line="two"),
]
SPEC = {s["key"]: s for s in SECTIONS}
FALLBACK_HOST = {"unit_conversions": "substitute", "multistep_questions": "word_questions"}


def build_prompt(eq, year_group="", ability="", notes="", previous_error=None):
    quantity_lines = []
    for key in eq.keys:
        others = eq.allowed_units(key)[1:]
        extra = f"; other allowed units: {', '.join(others)}" if others else "; no other units allowed"
        quantity_lines.append(f'- key "{key}" = {eq.qname[key]} ({plain_symbol(eq.symbol[key])}); SI unit: {eq.si_unit(key)}{extra}')
    section_lines = []
    for spec in SECTIONS:
        if spec["kind"] == "generated":
            continue
        free = math.ceil(spec["n"] / 2)
        if spec["key"] == "multistep_questions":
            rule = "Each needs 2 or more calculation steps; the last step must use the equation above."
            context = "Every question has a short real-world context."
        elif spec["key"] == "word_questions":
            rule = "One step with the equation above; information is presented in a short 2-3 sentence context."
            context = "Every question has the longer context."
        else:
            find = {"subject": f'"find" must be "{eq.subject}" (no rearranging)',
                    "other": f'"find" must NOT be "{eq.subject}" (rearrange the equation)'}[spec["find"]]
            conv = {0: "All values are in SI units.", 1: "Exactly ONE given value is in a non-SI unit."}[spec["conv"]]
            rule = f"{find}. {conv} Very brief, one-sentence questions."
            context = (f"First {free} questions have no scenario. Remaining {spec['n'] - free} have a very brief real-world scenario.")
        section_lines.append(f'- "{spec["key"]}": exactly {spec["n"]} questions. {rule} {context}')
    retry = f"\nYour previous answer was rejected: {previous_error}. Return the full corrected JSON.\n" if previous_error else ""
    return f"""You are an experienced GCSE Physics teacher making short calculation practice for lower-attaining and SEND students. Return one raw JSON object only.

EQUATION
Name: {eq.name}
In words: {eq.words}
In letters: {plain_symbol(eq.letters)}
Subject: {eq.subject}
Quantities (use exact keys):
{chr(10).join(quantity_lines)}

GROUP: {year_group or 'GCSE'}; ability: {ability or 'mixed'}.{(' Extra teacher instructions: ' + notes) if notes else ''}

SECTIONS AND RULES
{chr(10).join(section_lines)}

RULES
- Write question text exactly as students read it. Each value is a number, a space, and an allowed unit symbol. Do not spell out units.
- Do not ask students to convert km/h to m/s or hours (h) to seconds (s). Do not use km/h or h as given units.
- Do not tell the answer, working, or conversion in a question. Answers are in SI units.
- Use realistic positive values; answers should be exact to at most 3 significant figures.
- Use at most 5 different non-SI units across the worksheet. The app creates the separate conversion section.
- Questions in a section must differ.
- For single questions give "find" and "givens": one entry per other quantity, each {{"symbol": key, "value": number as written, "unit": unit symbol}}. Every value/unit in the text must match a given.
- For multistep questions give "text", "answer_quantity", "calculation_si" (one arithmetic expression in SI), "answer", and "working" (at least 2 short steps).
{retry}
Return ONLY JSON with exactly these keys: {', '.join(s['key'] for s in SECTIONS if s['kind'] != 'generated')}"""


def response_schema():
    given = {"type": "OBJECT", "properties": {"symbol": {"type": "STRING"}, "value": {"type": "NUMBER"}, "unit": {"type": "STRING"}}, "required": ["symbol", "value", "unit"]}
    single = {"type": "OBJECT", "properties": {"text": {"type": "STRING"}, "find": {"type": "STRING"}, "givens": {"type": "ARRAY", "items": given}}, "required": ["text", "find", "givens"]}
    multi = {"type": "OBJECT", "properties": {"text": {"type": "STRING"}, "answer_quantity": {"type": "STRING"}, "calculation_si": {"type": "STRING"}, "answer": {"type": "NUMBER"}, "working": {"type": "ARRAY", "items": {"type": "STRING"}}}, "required": ["text", "answer_quantity", "calculation_si", "answer", "working"]}
    props = {spec["key"]: {"type": "ARRAY", "items": multi if spec["kind"] == "multi" else single} for spec in SECTIONS if spec["kind"] != "generated"}
    return {"type": "OBJECT", "properties": props, "required": list(props)}


def get_gemini_api_key():
    key = os.getenv("GEMINI_API_KEY")
    if key:
        return key
    try:
        return st.secrets["GEMINI_API_KEY"]
    except Exception:
        return None


def clean_json_text(text):
    text = text.strip()
    if text.startswith("```"):
        lines = text.splitlines()
        text = "\n".join(lines[1:-1]) if len(lines) >= 2 else text.strip("`")
    return text.strip()


def call_gemini(prompt, api_key):
    model = os.getenv("GEMINI_MODEL") or "gemini-2.5-flash"
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    payload = {"contents": [{"role": "user", "parts": [{"text": prompt}]}], "generationConfig": {"temperature": 0.4, "responseMimeType": "application/json", "responseSchema": response_schema()}}
    request = urllib.request.Request(url, data=json.dumps(payload).encode(), method="POST", headers={"Content-Type": "application/json", "x-goog-api-key": api_key})
    try:
        with urllib.request.urlopen(request, timeout=300) as response:
            response_data = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        raise RuntimeError(f"Gemini API request failed ({exc.code}): {exc.read().decode('utf-8', errors='ignore')}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Unable to reach Gemini API: {exc}") from exc
    try:
        text = response_data["candidates"][0]["content"]["parts"][0]["text"]
        return json.loads(clean_json_text(text))
    except (KeyError, IndexError, TypeError) as exc:
        raise ValueError("Gemini response did not contain usable content.") from exc
    except json.JSONDecodeError as exc:
        raise ValueError(f"Gemini did not return valid JSON: {exc}") from exc


def _close(a, b, rel=1e-9):
    return abs(a - b) <= rel * max(1.0, abs(a), abs(b))


def _fmt_answer(value, unit):
    text = f"{fmt_num(value)} {unit}".strip()
    return text + (" (3 s.f.)" if not is_exact_to_sig(value) else "")


def process_single(eq, spec, index, raw):
    text = normalise_text(raw["text"]).strip()
    find = raw["find"]
    values_si, conversions = {}, []
    for given in raw["givens"]:
        key = given["symbol"]
        unit = normalise_text(given["unit"]).strip()
        value = float(given["value"])
        info = UNIT_INFO[unit]
        si_value = Fraction(str(value)) * info["factor"]
        values_si[key] = si_value
        if unit != info["si"]:
            conversions.append((value, unit, si_value, info["si"]))
    result = eq.evaluate(find, values_si)
    if not math.isfinite(result):
        raise ValueError("answer is not finite")
    steps = [f"{fmt_num(v, 6)} {u} = {fmt_num(float(s), 6)} {si}" for v, u, s, si in conversions]
    steps.append(f"{eq.rearranged(find)} = {eq.substituted(find, values_si)} = **{_fmt_answer(result, eq.si_unit(find))}**")
    return {"text": text, "answer": "; ".join(steps), "result": result, "units": units_in_text(text)}


def process_multi(eq, spec, index, raw):
    text = normalise_text(raw["text"]).strip()
    key = raw["answer_quantity"]
    calculation = raw["calculation_si"]
    working = [normalise_text(w).strip() for w in raw["working"] if str(w).strip()]
    result = safe_eval(calculation)
    if not math.isfinite(result) or not working:
        raise ValueError("invalid answer or working")
    return {"text": text, "answer": " → ".join(working) + f" → **{_fmt_answer(result, eq.si_unit(key))}**", "result": result, "units": units_in_text(text)}


def build_unit_conversion_items(eq, needed_units, n, rng):
    chosen = list(needed_units)
    pool = []
    for dim in [eq.dim[k] for k in eq.keys] + ["length", "time", "mass"]:
        pool += [u[0] for u in UNIT_GROUPS[dim]["others"] if u[0] not in FORBIDDEN_CONVERSION_UNITS]
    for unit in pool:
        if len(chosen) >= n:
            break
        if unit not in chosen:
            chosen.append(unit)
    items = []
    for unit in chosen:
        info = UNIT_INFO[unit]
        value = rng.choice(info["examples"])
        si_value = Fraction(str(value)) * info["factor"]
        items.append({"text": f"Convert {fmt_num(value, 6)} {unit} to {info['si']}.", "answer": f"{fmt_num(float(si_value), 6)} {info['si']}"})
    return items


def build_basics(eq):
    letters = "abcdefgh"
    def line(parts): return "   ".join(parts)
    def cap(s): return s[0].upper() + s[1:]
    q1 = line(f"{letters[i]}){NBSP}{eq.symbol[k]}{NBSP}{BLANK}" for i, k in enumerate(eq.keys))
    a1 = line(f"{letters[i]}){NBSP}{eq.symbol[k]}{NBSP}={NBSP}{eq.qname[k]}" for i, k in enumerate(eq.keys))
    q2 = line(f"{letters[i]}){NBSP}{cap(eq.qname[k]).replace(' ', NBSP)}{NBSP}{BLANK}" for i, k in enumerate(eq.keys))
    a2 = line(f"{letters[i]}){NBSP}{eq.si_unit(k)}" for i, k in enumerate(eq.keys))
    q3 = ("Write the equation using letters.   ", f"{BLANK}{BLANK}", eq.letters)
    q4 = line(f"{letters[i]}){NBSP}{eq.symbol[k]}{NBSP}={NBSP}{BLANK}" for i, k in enumerate(eq.keys))
    a4 = line(f"{letters[i]}){NBSP}{eq.rearranged(k).replace(' ', NBSP)}" for k in eq.keys)
    return [("What do these letters represent?   ", q1, a1), ("What are the units for:   ", q2, a2), q3, ("Rearrange the equation:   ", q4, a4)]


def build_model(eq, data, seed=None):
    if not isinstance(data, dict):
        raise ValueError("Gemini JSON must be an object.")
    sections, needed, warnings = {}, [], []
    for spec in SECTIONS:
        if spec["kind"] == "generated":
            continue
        raw_items = data.get(spec["key"])
        if not isinstance(raw_items, list):
            warnings.append(f'{spec["heading"]}: section data was missing or malformed; no questions included.')
            sections[spec["key"]] = []
            continue
        process = process_multi if spec["kind"] == "multi" else process_single
        items, seen_texts = [], set()
        for index, raw in enumerate(raw_items, start=1):
            try:
                item = process(eq, spec, index, raw)
                if not item["text"]:
                    raise ValueError("question text is empty")
                normalized_question = item["text"].lower()
                if normalized_question in seen_texts:
                    raise ValueError("duplicate question")
                seen_texts.add(normalized_question)
            except Exception as exc:
                # A malformed response should only discard this item, not the whole sheet.
                warnings.append(f'{spec["heading"]} question {index} skipped: {exc}')
                continue
            items.append(item)
            for _, unit in item["units"]:
                if unit != UNIT_INFO[unit]["si"] and unit not in needed:
                    needed.append(unit)
        if len(items) < spec["n"]:
            warnings.append(f'{spec["heading"]}: included {len(items)} of {spec["n"]} questions.')
        sections[spec["key"]] = items
    sections["unit_conversions"] = build_unit_conversion_items(eq, needed, SPEC["unit_conversions"]["n"], random.Random(seed if seed is not None else eq.name))
    return {"equation": eq, "sections": sections, "basics": build_basics(eq), "warnings": warnings}


def generate_content(eq, api_key, year_group="", ability="", notes="", on_response=None):
    error = None
    for attempt in range(1, MAX_ATTEMPTS + 1):
        data = call_gemini(build_prompt(eq, year_group, ability, notes, error), api_key)
        if on_response is not None:
            on_response(data, attempt)
        try:
            return build_model(eq, data), data
        except ValueError as exc:
            error = str(exc)
    raise ValueError(f"Gemini's questions failed validation {MAX_ATTEMPTS} times. Last problem: {error}")


def append_text(paragraph, raw_text, bold=None, size=None):
    text, last = str(raw_text or ""), 0
    def run(chunk, is_bold=False, sub=False):
        r = paragraph.add_run(chunk)
        if size: r.font.size = Pt(size)
        if bold or is_bold: r.font.bold = True
        if sub: r.font.subscript = True
    for match in re.finditer(r"(\*\*.+?\*\*|<sub>.+?</sub>)", text):
        if match.start() > last: run(text[last:match.start()])
        token = match.group(1)
        run(token[2:-2], is_bold=True) if token.startswith("**") else run(token[5:-6], sub=True)
        last = match.end()
    if last < len(text): run(text[last:])


def _all_paragraphs(doc):
    return [Paragraph(p, doc._body) for p in doc.element.body.iter(qn("w:p"))]


def _placeholder_name(paragraph):
    text = re.sub(r"\s+", " ", paragraph.text.replace(" ", " ")).strip()
    match = re.fullmatch(r"\[{1,2}\s*(.+?)\s*\]{1,2}", text)
    return match.group(1).lower() if match else None


def _new_paragraph(target, space_after=0.5, keep_next=False):
    p = target.insert_paragraph_before("")
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing = 0.9
    p.paragraph_format.keep_with_next = keep_next
    return p


def _numbered(target, label, text, size=QUESTION_PT):
    p = _new_paragraph(target)
    p.paragraph_format.left_indent = Cm(0.55)
    p.paragraph_format.first_line_indent = Cm(-0.55)
    p.paragraph_format.tab_stops.add_tab_stop(Twips(420), WD_TAB_ALIGNMENT.LEFT)
    append_text(p, f"{label}\t", size=size)
    append_text(p, text, size=size)
    return p


def _answer_line(target, label, width_twips):
    p = _new_paragraph(target)
    p.paragraph_format.left_indent = Cm(0.55)
    p.paragraph_format.tab_stops.add_tab_stop(Twips(max(800, width_twips - 500)), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.LINES)
    append_text(p, f"{label}\t", size=7.5)


def _render_section(target, spec, items, answers):
    heading = _new_paragraph(target, space_after=1, keep_next=True)
    append_text(heading, spec["heading"], bold=True, size=9)
    width = 9000
    node = target._p.getparent()
    while node is not None and node.tag != qn("w:tc"):
        node = node.getparent()
    if node is not None:
        tcw = node.find(qn("w:tcPr") + "/" + qn("w:tcW"))
        if tcw is not None and tcw.get(qn("w:w")):
            width = int(tcw.get(qn("w:w")))
    is_conversion = spec["key"] == "unit_conversions"
    for i, item in enumerate(items):
        label = f"{'abcdefgh'[i]}" if is_conversion and i < 8 else f"{i + 1}"
        label += ")"
        if answers:
            _numbered(target, label, item["answer"], size=7.5)
        elif spec["answer_line"] == "inline":
            _numbered(target, label, f"{item['text']}  {BLANK}")
        else:
            _numbered(target, label, item["text"])
            if spec["answer_line"] == "two": _answer_line(target, "Working:", width)
            _answer_line(target, "Answer:", width)
    _new_paragraph(target, space_after=0)


def _remove_page_breaks(paragraph):
    ppr = paragraph._p.pPr
    if ppr is not None:
        node = ppr.find(qn("w:pageBreakBefore"))
        if node is not None: ppr.remove(node)
    for node in list(paragraph._p.iter()):
        if node.tag == qn("w:br") and node.get(qn("w:type")) == "page":
            node.getparent().remove(node)
        elif node.tag == qn("w:lastRenderedPageBreak"):
            node.getparent().remove(node)


def _render_basics(target, eq, basics, answers):
    p = _new_paragraph(target, space_after=1)
    append_text(p, eq.words[0].upper() + eq.words[1:] + "   " + eq.letters, bold=True, size=10)
    for i, (lead, blanks, answer) in enumerate(basics):
        p = _numbered(target, f"{i + 1})", lead, size=7.5)
        append_text(p, answer if answers else blanks, size=7.5)


def generate_documents(model, template_path):
    eq = model["equation"]
    outputs, warnings = {}, []
    for mode in ("worksheet", "answers"):
        answers = mode == "answers"
        doc = Document(str(template_path))
        for section in doc.sections:
            section.page_width, section.page_height = Cm(21), Cm(29.7)
            section.top_margin = section.bottom_margin = Cm(0.65)
            section.left_margin = section.right_margin = Cm(0.8)
        targets, paragraphs = {}, _all_paragraphs(doc)
        for p in paragraphs:
            name = _placeholder_name(p)
            if name and name not in targets: targets[name] = p
            _remove_page_breaks(p)
        if "title" in targets:
            for r in list(targets["title"]._p.findall(qn("w:r"))): targets["title"]._p.remove(r)
            append_text(targets["title"], eq.title + (" – Answers" if answers else ""), bold=True, size=14)
        else:
            warnings.append("The template has no [[Title]] placeholder.")
        if "basics" in targets:
            _render_basics(targets["basics"], eq, model["basics"], answers)
            targets["basics"]._p.getparent().remove(targets["basics"]._p)
        else:
            warnings.append("The template has no [[Basics]] placeholder.")
        hosts = {}
        for spec in SECTIONS:
            name = spec["placeholder"]
            if name in targets:
                host = name
            else:
                fallback = SPEC[FALLBACK_HOST[spec["key"]]]["placeholder"] if spec["key"] in FALLBACK_HOST else None
                if spec["key"] == "unit_conversions" and fallback in targets:
                    nxt = targets[fallback]._p.getnext()
                    if nxt is not None and nxt.tag == qn("w:p") and not "".join(nxt.itertext()).strip():
                        targets[name] = Paragraph(nxt, doc._body)
                        host = name
                    else: host = fallback
                elif fallback in targets: host = fallback
                else:
                    warnings.append(f"The template has no [[{spec['heading']}]] placeholder; section omitted.")
                    continue
            hosts.setdefault(host, []).append(spec)
        for host, specs in hosts.items():
            for spec in sorted(specs, key=SECTIONS.index):
                _render_section(targets[host], spec, model["sections"][spec["key"]], answers)
            targets[host]._p.getparent().remove(targets[host]._p)
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        outputs[mode] = buf
    outputs["warnings"] = sorted(set(warnings))
    return outputs


def prepare_template_path(uploaded_template):
    if uploaded_template is None:
        if DEFAULT_TEMPLATE.exists(): return DEFAULT_TEMPLATE
        raise FileNotFoundError(f"The default template '{DEFAULT_TEMPLATE.name}' was not found next to the app.")
    temp_dir = Path(tempfile.gettempdir()) / "formula_worksheet_templates"
    temp_dir.mkdir(exist_ok=True)
    path = temp_dir / f"{uuid4().hex}.docx"
    path.write_bytes(uploaded_template.getvalue())
    return path


def main():
    st.set_page_config(page_title="Physics Formula Worksheet Generator", page_icon="⚡")
    st.title("Physics Formula Worksheet Generator")
    st.write("Choose a GCSE Physics equation. Gemini creates questions; the app calculates answers and creates a Word worksheet and answer sheet.")
    equation_name = st.selectbox("Equation", list(EQUATIONS))
    eq = EQUATIONS[equation_name]
    st.caption(f"{eq.words} | {plain_symbol(eq.letters)}")
    year_group = st.text_input("Year group (optional)", placeholder="e.g. Year 10")
    ability = st.text_input("Ability (optional)", placeholder="Lower-attaining / mixed / high ability")
    notes = st.text_area("Extra instructions (optional)", placeholder="e.g. Use sporting contexts. Keep numbers small.", height=80)
    st.markdown("Create a Google Gemini API key at [Google AI Studio](https://aistudio.google.com/).")
    api_key_input = st.text_input("Gemini API key", type="password")
    uploaded_template = st.file_uploader("Optional Word template (.docx)", type=["docx"])
    show_json = st.checkbox("Show Gemini JSON response (debug)", value=False)
    for key, default in (("fw_worksheet", None), ("fw_answers", None), ("fw_name", "worksheet"), ("fw_warnings", [])):
        st.session_state.setdefault(key, default)
    if st.button("Generate Worksheet", type="primary"):
        api_key = api_key_input.strip() or get_gemini_api_key()
        if not api_key:
            st.error("Enter a Gemini API key or set GEMINI_API_KEY.")
        else:
            json_output = st.container() if show_json else None
            def show_response(data, attempt):
                if json_output is not None:
                    json_output.markdown(f"Gemini JSON response (attempt {attempt})")
                    json_output.json(data)
            try:
                with st.spinner("Generating questions and checking answers..."):
                    model, raw = generate_content(eq, api_key, year_group, ability, notes, on_response=show_response)
                result = generate_documents(model, prepare_template_path(uploaded_template))
                st.session_state.update(fw_worksheet=result["worksheet"].getvalue(), fw_answers=result["answers"].getvalue(), fw_name=eq.name, fw_warnings=result["warnings"] + model.get("warnings", []))
                st.success("Worksheet generated successfully!")
            except Exception as exc:
                st.error(f"❌ An error occurred: {exc}")
    for warning in st.session_state["fw_warnings"]: st.warning(warning)
    mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    if st.session_state["fw_worksheet"] is not None:
        st.download_button("Download Worksheet", data=st.session_state["fw_worksheet"], file_name=f"{st.session_state['fw_name']}.docx", mime=mime)
    if st.session_state["fw_answers"] is not None:
        st.download_button("Download Answers", data=st.session_state["fw_answers"], file_name=f"{st.session_state['fw_name']} - Answers.docx", mime=mime)


if __name__ == "__main__":
    main()
