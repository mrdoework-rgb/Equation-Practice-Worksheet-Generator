"""GCSE Physics formula worksheet generator (Streamlit).

Flow
  1. Teacher picks an equation from the GCSE physics equation sheet.
  2. Gemini writes the numerical questions as JSON (text + the numbers and units used).
  3. Python validates the JSON, does ALL of the maths itself (so the answer key is
     never an LLM guess), and writes the unit-conversion section deterministically so it
     always covers every conversion used elsewhere on the sheet.
  4. The content is poured into the Word template, once as a worksheet and once as an
     answer key.
"""

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
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, Twips
from docx.text.paragraph import Paragraph
from PIL import Image, ImageDraw, ImageFont

from physics_data import (
    EQUATIONS, UNIT_GROUPS, UNIT_INFO, fmt_num, is_exact_to_sig, normalise_text,
    numbers_in_text, plain_symbol, safe_eval, units_in_text,
)

DEFAULT_TEMPLATE = Path(__file__).parent / "Formulae_Questions_worksheet_template.docx"
NBSP = " "
BLANK = "________"
QUESTION_PT = 10   # font size of questions; lower it if sections spill onto extra pages
MAX_ATTEMPTS = 3

# --------------------------------------------------------------------------- #
# What each section of the worksheet contains
# --------------------------------------------------------------------------- #
# kind:  "single" = one-equation question, answers computed by Python
#        "multi"  = multi-step question, answer computed from an arithmetic string
#        "generated" = written by Python, not by the LLM
# find:  "subject" (substitute only), "other" (needs rearranging), "any"
# conv:  how many givens must need a unit conversion (None = any number)
# placeholder: the [[...]] text in the template (case-insensitive)
SECTIONS = [
    dict(key="substitute", placeholder="substitute", heading="Substitute", n=8, kind="single", find="subject", conv=0, answer_line="inline"),
    dict(key="unit_conversions", placeholder="unit conversions", heading="Unit Conversions", n=6, kind="generated", answer_line="inline"),
    dict(key="substitute_unit_conversions", placeholder="substitute unit conversions", heading="Substitute (convert units)", n=5, kind="single", find="subject", conv=1, answer_line="inline"),
    dict(key="rearrange", placeholder="rearrange", heading="Rearrange", n=5, kind="single", find="other", conv=0, answer_line="inline"),
    dict(key="rearrange_unit_conversions", placeholder="rearrange unit conversions", heading="Rearrange (convert units)", n=5, kind="single", find="other", conv=1, answer_line="inline"),
    dict(key="word_questions", placeholder="word questions", heading="Word Questions", n=6, kind="single", find="any", conv=None, answer_line="one"),
    dict(key="multistep_questions", placeholder="multistep questions", heading="Multistep Questions", n=5, kind="multi", answer_line="two"),
]
SPEC = {s["key"]: s for s in SECTIONS}
# If the template has no placeholder for a section, put it with this one instead.
FALLBACK_HOST = {"unit_conversions": "substitute", "multistep_questions": "word_questions"}


# --------------------------------------------------------------------------- #
# Prompt and JSON schema
# --------------------------------------------------------------------------- #

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
            rule = ("Each needs 2 or more calculation steps, e.g. use another equation from the GCSE physics equation sheet to get a value, "
                    "take a mean of several values, or work out the number that gets substituted. The last step must use the equation above.")
            context = "Every question has a short real-world context."
        elif spec["key"] == "word_questions":
            rule = "One step with the equation above (rearranging and unit conversions both allowed), but the information is presented in a longer context of 2-3 sentences."
            context = "Every question has the longer context."
        else:
            find = {"subject": f'"find" must be "{eq.subject}" (the subject of the equation; no rearranging)',
                    "other": f'"find" must NOT be "{eq.subject}" (the equation must be rearranged)'}[spec["find"]]
            conv = {0: "All values are given in SI units.", 1: "Exactly ONE given value is in a non-SI unit (needs converting to SI first)."}[spec["conv"]]
            rule = f"{find}. {conv} Very brief questions, one sentence."
            context = (f"The first {free} questions are plain, with NO scenario, just the quantities and values. "
                       f"The remaining {spec['n'] - free} add a VERY brief real-world scenario of a few words (e.g. a tractor, a cyclist, a crane).")
        section_lines.append(f'- "{spec["key"]}": exactly {spec["n"]} questions. {rule} {context}')
    retry = f"\nYOUR PREVIOUS ANSWER WAS REJECTED: {previous_error}\nFix this problem and return the full corrected JSON.\n" if previous_error else ""
    return f"""You are an experienced GCSE Physics teacher writing calculation practice for lower-attaining and SEND students. Keep language very simple and questions very short. Return one raw JSON object only.

EQUATION
Name: {eq.name}
In words: {eq.words}
In letters: {plain_symbol(eq.letters)}
Subject of the equation: "{eq.subject}"
Quantities (use these exact keys):
{chr(10).join(quantity_lines)}

GROUP: {year_group or 'GCSE'}; ability: {ability or 'mixed'}.{(' Extra teacher instructions: ' + notes) if notes else ''}

SECTIONS (JSON keys) AND RULES
{chr(10).join(section_lines)}

RULES FOR EVERY QUESTION
- "text": the question exactly as students will read it. Write every value as a number, a space, then a unit symbol taken ONLY from the allowed units above (e.g. "20 cm", "5 kN", "36 km/h", "2 h"). Never spell units out in words, never use scientific notation or commas in numbers.
- Do not say which unit to answer in. Answers are always in SI units.
- Never mention the answer, the working, or the unit conversion in the question.
- Use realistic GCSE values and choose numbers so that the answer is positive and exact to at most 3 significant figures.
- Across the WHOLE worksheet use at most 5 different non-SI units in total (reuse the same ones). A separate unit-conversion section is created by the app, so do not write one.
- Questions within a section must all be different.
- For "single" questions ("substitute", "substitute_unit_conversions", "rearrange", "rearrange_unit_conversions", "word_questions") give: "find" (the key of the quantity asked for) and "givens": one entry for EVERY other quantity in the equation, each as {{"symbol": key, "value": number as written in the text, "unit": unit symbol as written in the text}}. Every number with a unit in the text must appear in "givens" and vice versa.
- For "multistep_questions" give: "text", "answer_quantity" (key of the quantity the final answer is, from the list above), "calculation_si" (ONE arithmetic expression that evaluates to the final answer in SI units, using only numbers, + - * / ** ( ) and sqrt(); convert every value to SI inside the expression), "answer" (the number it evaluates to) and "working" (2-5 short strings, one per step, written for students).
{retry}
Return ONLY the JSON object with exactly these keys: {', '.join(s['key'] for s in SECTIONS if s['kind'] != 'generated')}."""


def response_schema():
    given = {"type": "OBJECT", "properties": {"symbol": {"type": "STRING"}, "value": {"type": "NUMBER"}, "unit": {"type": "STRING"}},
             "required": ["symbol", "value", "unit"]}
    single = {"type": "OBJECT", "properties": {"text": {"type": "STRING"}, "find": {"type": "STRING"},
                                                "givens": {"type": "ARRAY", "items": given}},
              "required": ["text", "find", "givens"]}
    multi = {"type": "OBJECT", "properties": {"text": {"type": "STRING"}, "answer_quantity": {"type": "STRING"},
                                               "calculation_si": {"type": "STRING"}, "answer": {"type": "NUMBER"},
                                               "working": {"type": "ARRAY", "items": {"type": "STRING"}}},
             "required": ["text", "answer_quantity", "calculation_si", "answer", "working"]}
    props = {s["key"]: {"type": "ARRAY", "items": multi if s["kind"] == "multi" else single}
             for s in SECTIONS if s["kind"] != "generated"}
    return {"type": "OBJECT", "properties": props, "required": list(props)}


# --------------------------------------------------------------------------- #
# Gemini
# --------------------------------------------------------------------------- #

def get_gemini_api_key():
    api_key = os.getenv("GEMINI_API_KEY")
    if api_key:
        return api_key
    try:
        return st.secrets["GEMINI_API_KEY"]
    except Exception:
        return None


def clean_json_text(raw_text):
    text = raw_text.strip()
    if text.startswith("```"):
        lines = text.splitlines()
        text = "\n".join(lines[1:-1]) if len(lines) >= 2 else text.strip("`")
    return text.strip()


def call_gemini(prompt, api_key):
    model = os.getenv("GEMINI_MODEL") or "gemini-2.5-flash"
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    payload = {
        "contents": [{"role": "user", "parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.4, "responseMimeType": "application/json", "responseSchema": response_schema()},
    }
    request = urllib.request.Request(
        url, data=json.dumps(payload).encode("utf-8"), method="POST",
        headers={"Content-Type": "application/json", "x-goog-api-key": api_key},  # key in a header, not the URL
    )
    try:
        with urllib.request.urlopen(request, timeout=120) as response:
            response_data = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"Gemini API request failed ({exc.code}): {detail}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Unable to reach Gemini API: {exc}") from exc
    try:
        text = response_data["candidates"][0]["content"]["parts"][0]["text"]
    except (KeyError, IndexError, TypeError) as exc:
        raise ValueError("Gemini response did not contain a usable content block.") from exc
    try:
        return json.loads(clean_json_text(text))
    except json.JSONDecodeError as exc:
        raise ValueError(f"Gemini did not return valid JSON: {exc}") from exc


# --------------------------------------------------------------------------- #
# Validation and maths: LLM JSON  ->  worksheet model
# --------------------------------------------------------------------------- #

def _close(a, b, rel=1e-9):
    return abs(a - b) <= rel * max(1.0, abs(a), abs(b))


def _fmt_answer(value, unit):
    text = f"{fmt_num(value)} {unit}".strip()
    return text + (" (3 s.f.)" if not is_exact_to_sig(value) else "")


def process_single(eq, spec, index, raw):
    where = f'"{spec["key"]}" question {index}'
    if not isinstance(raw, dict) or not isinstance(raw.get("givens"), list) or not isinstance(raw.get("text"), str):
        raise ValueError(f"{where} is not in the required format.")
    text = normalise_text(raw["text"]).strip()
    find = raw.get("find")
    if find not in eq.keys:
        raise ValueError(f'{where}: "find" must be one of {eq.keys}.')
    if spec["find"] == "subject" and find != eq.subject:
        raise ValueError(f'{where}: "find" must be "{eq.subject}" (no rearranging in this section).')
    if spec["find"] == "other" and find == eq.subject:
        raise ValueError(f'{where}: "find" must not be "{eq.subject}" (this section needs rearranging).')

    expected = set(eq.keys) - {find}
    seen, values_si, conversions, given_pairs = set(), {}, [], []
    numbers = numbers_in_text(text)
    for given in raw["givens"]:
        try:
            key, unit, value = given["symbol"], normalise_text(given["unit"]).strip(), float(given["value"])
        except (KeyError, TypeError, ValueError) as exc:
            raise ValueError(f"{where}: a given value is malformed.") from exc
        if key not in expected or key in seen:
            raise ValueError(f"{where}: givens must contain each of {sorted(expected)} exactly once.")
        seen.add(key)
        if unit not in eq.allowed_units(key):
            raise ValueError(f'{where}: unit "{unit}" is not allowed for {eq.qname[key]}; use one of {eq.allowed_units(key)}.')
        if not any(_close(value, n) for n in numbers) or unit not in text:
            raise ValueError(f'{where}: the text does not contain "{fmt_num(value, 6)} {unit}" for {eq.qname[key]}.')
        info = UNIT_INFO[unit]
        si_value = Fraction(str(value)) * info["factor"]
        values_si[key] = si_value
        given_pairs.append((value, unit))
        if unit != info["si"]:
            conversions.append((value, unit, si_value, info["si"]))
    if seen != expected:
        raise ValueError(f"{where}: givens must contain each of {sorted(expected)} exactly once.")
    for number, unit in units_in_text(text):
        if not any(_close(number, v) and unit == u for v, u in given_pairs):
            raise ValueError(f'{where}: the text contains "{fmt_num(number, 6)} {unit}" which is not one of the givens.')
    if spec["conv"] is not None and len(conversions) != spec["conv"]:
        raise ValueError(f"{where}: needs exactly {spec['conv']} value(s) in a non-SI unit but has {len(conversions)}.")

    result = eq.evaluate(find, values_si)
    if not (math.isfinite(result) and 1e-6 <= result <= 1e9):
        raise ValueError(f"{where}: the answer ({result}) is not a sensible size.")
    si_unit = eq.si_unit(find)
    steps = [f"{fmt_num(v, 6)} {u} = {fmt_num(float(s), 6)} {si}" for v, u, s, si in conversions]
    steps.append(f"{eq.rearranged(find)} = {eq.substituted(find, values_si)} = **{_fmt_answer(result, si_unit)}**")
    return {"text": text, "answer": "; ".join(steps), "result": result, "units": units_in_text(text)}


def process_multi(eq, spec, index, raw):
    where = f'"{spec["key"]}" question {index}'
    try:
        text = normalise_text(raw["text"]).strip()
        key = raw["answer_quantity"]
        claimed = float(raw["answer"])
        working = [normalise_text(w).strip() for w in raw["working"] if str(w).strip()]
        calculation = raw["calculation_si"]
    except (KeyError, TypeError, ValueError) as exc:
        raise ValueError(f"{where} is not in the required format.") from exc
    if key not in eq.keys:
        raise ValueError(f'{where}: "answer_quantity" must be one of {eq.keys}.')
    if len(working) < 2:
        raise ValueError(f'{where}: "working" needs at least 2 steps.')
    result = safe_eval(calculation)
    if not _close(result, claimed, 0.01):
        raise ValueError(f'{where}: calculation_si evaluates to {fmt_num(result, 6)} but "answer" says {fmt_num(claimed, 6)}.')
    if not (1e-6 <= result <= 1e9):
        raise ValueError(f"{where}: the answer ({result}) is not a sensible size.")
    answer = " → ".join(working) + f" → **{_fmt_answer(result, eq.si_unit(key))}**"
    return {"text": text, "answer": answer, "result": result, "units": units_in_text(text)}


def build_unit_conversion_items(eq, needed_units, n, rng):
    """`n` conversions into SI that include every non-SI unit used elsewhere on the sheet."""
    chosen = list(needed_units)
    if len(chosen) > n:
        raise ValueError(f"The questions use {len(chosen)} different non-SI units ({', '.join(chosen)}); at most {n} are allowed. Reuse the same units.")
    pool = []
    for dim in [eq.dim[k] for k in eq.keys] + ["length", "time", "mass"]:
        pool += [u[0] for u in UNIT_GROUPS[dim]["others"]]
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
        items.append({"text": f"What is {fmt_num(value, 6)} {unit} in {info['si']}?",
                      "answer": f"{fmt_num(float(si_value), 6)} {info['si']}"})
    return items


def build_basics(eq):
    """The four fixed 'Basics' questions (questions text, answer text)."""
    letters = "abcdefgh"

    def line(parts):
        return "   ".join(parts)

    def cap(s):
        return s[0].upper() + s[1:]

    q1 = line(f"{letters[i]}){NBSP}{eq.symbol[k]}{NBSP}{BLANK}" for i, k in enumerate(eq.keys))
    a1 = line(f"{letters[i]}){NBSP}{eq.symbol[k]}{NBSP}={NBSP}{eq.qname[k]}" for i, k in enumerate(eq.keys))
    q2 = line(f"{letters[i]}){NBSP}{cap(eq.qname[k]).replace(' ', NBSP)}{NBSP}{BLANK}" for i, k in enumerate(eq.keys))
    a2 = line(f"{letters[i]}){NBSP}{eq.si_unit(k)}" for i, k in enumerate(eq.keys))
    if eq.triangle:
        q3 = ("Complete the formula triangle.   ", "", "")
    else:
        q3 = ("Write the equation using letters.   ", f"{BLANK}{BLANK}", eq.letters)
    q4 = line(f"{letters[i]}){NBSP}{eq.symbol[k]}{NBSP}={NBSP}{BLANK}" for i, k in enumerate(eq.keys))
    a4 = line(f"{letters[i]}){NBSP}{eq.rearranged(k).replace(' ', NBSP)}" for i, k in enumerate(eq.keys))
    return [
        ("What do these letters represent?   ", q1, a1),
        ("What are the units for:   ", q2, a2),
        q3,
        ("Rearrange the equation:   ", q4, a4),
    ]


def build_model(eq, data, seed=None):
    """Validate the LLM JSON and turn it into everything the renderer needs. Raises ValueError."""
    if not isinstance(data, dict):
        raise ValueError("Gemini JSON must be an object.")
    sections, needed = {}, []
    for spec in SECTIONS:
        if spec["kind"] == "generated":
            continue
        raw_items = data.get(spec["key"])
        if not isinstance(raw_items, list) or len(raw_items) != spec["n"]:
            raise ValueError(f'"{spec["key"]}" must be an array of exactly {spec["n"]} questions.')
        process = process_multi if spec["kind"] == "multi" else process_single
        items = [process(eq, spec, i + 1, raw) for i, raw in enumerate(raw_items)]
        if len({i["text"].lower() for i in items}) != len(items):
            raise ValueError(f'"{spec["key"]}" contains duplicate questions.')
        for item in items:
            for _, unit in item["units"]:
                if unit != UNIT_INFO[unit]["si"] and unit not in needed:
                    needed.append(unit)
        sections[spec["key"]] = items
    sections["unit_conversions"] = build_unit_conversion_items(
        eq, needed, SPEC["unit_conversions"]["n"], random.Random(seed if seed is not None else eq.name))
    return {"equation": eq, "sections": sections, "basics": build_basics(eq)}


def generate_content(eq, api_key, year_group="", ability="", notes=""):
    """Ask Gemini, validate, and retry with the error message if the JSON is unusable."""
    error, last_data = None, None
    for _ in range(MAX_ATTEMPTS):
        last_data = call_gemini(build_prompt(eq, year_group, ability, notes, error), api_key)
        try:
            return build_model(eq, last_data), last_data
        except ValueError as exc:
            error = str(exc)
    raise ValueError(f"Gemini's questions failed validation {MAX_ATTEMPTS} times. Last problem: {error}")


# --------------------------------------------------------------------------- #
# Formula triangle picture
# --------------------------------------------------------------------------- #

def _font(size):
    for path in ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "DejaVuSans.ttf", "arial.ttf"):
        try:
            return ImageFont.truetype(path, size)
        except OSError:
            continue
    return ImageFont.load_default(size=size)


def triangle_png(letters=None):
    """Blank formula triangle, or one with (top, bottom_left, bottom_right) filled in."""
    w, h, scale = 440, 380, 2
    img = Image.new("RGB", (w * scale, h * scale), "white")
    d = ImageDraw.Draw(img)
    s = scale
    apex, left, right, mid_y = (220 * s, 20 * s), (20 * s, 360 * s), (420 * s, 360 * s), 200 * s
    t = (mid_y - apex[1]) / (left[1] - apex[1])
    mid_left, mid_right = apex[0] - 200 * s * t, apex[0] + 200 * s * t
    width = 6 * s
    d.line([apex, left, right, apex], fill="black", width=width, joint="curve")
    d.line([(mid_left, mid_y), (mid_right, mid_y)], fill="black", width=width)
    d.line([(220 * s, mid_y), (220 * s, 360 * s)], fill="black", width=width)
    if letters:
        font = _font(95 * s)
        for text, xy in zip(letters, [(220 * s, 130 * s), (145 * s, 285 * s), (295 * s, 285 * s)]):
            d.text(xy, plain_symbol(text), font=font, fill="black", anchor="mm")
    out = io.BytesIO()
    img.save(out, format="PNG")
    out.seek(0)
    return out


# --------------------------------------------------------------------------- #
# Word output
# --------------------------------------------------------------------------- #

def append_text(paragraph, raw_text, bold=None, size=None):
    """Add text to a paragraph, understanding **bold** and <sub>subscript</sub>."""
    text = str(raw_text or "")
    last = 0

    def run(chunk, **fmt):
        r = paragraph.add_run(chunk)
        if size:
            r.font.size = Pt(size)
        if bold or fmt.get("bold"):
            r.font.bold = True
        if fmt.get("sub"):
            r.font.subscript = True
        return r

    for match in re.finditer(r"(\*\*.+?\*\*|<sub>.+?</sub>)", text):
        if match.start() > last:
            run(text[last:match.start()])
        token = match.group(1)
        if token.startswith("**"):
            run(token[2:-2], bold=True)
        else:
            run(token[5:-6], sub=True)
        last = match.end()
    if last < len(text):
        run(text[last:])


def _all_paragraphs(doc):
    """Every paragraph in the document, including those in table cells (each once, merged cells included)."""
    return [Paragraph(p, doc._body) for p in doc.element.body.iter(qn("w:p"))]


def _placeholder_name(paragraph):
    text = re.sub(r"\s+", " ", paragraph.text.replace(" ", " ")).strip()
    match = re.fullmatch(r"\[{1,2}\s*(.+?)\s*\]{1,2}", text)   # tolerate a mistyped single bracket
    return match.group(1).lower() if match else None


def _cell_width_twips(paragraph):
    node = paragraph._p.getparent()
    while node is not None and node.tag != qn("w:tc"):
        node = node.getparent()
    if node is not None:
        width = node.find(qn("w:tcPr") + "/" + qn("w:tcW"))
        if width is not None and width.get(qn("w:w")):
            return int(width.get(qn("w:w")))
    return 9000


def _new_paragraph(target, space_after=3, keep_next=False):
    p = target.insert_paragraph_before("")
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.keep_with_next = keep_next
    return p


def _clear_runs(paragraph):
    for r in paragraph._p.findall(qn("w:r")):
        paragraph._p.remove(r)


def _numbered(target, label, text, size=QUESTION_PT):
    p = _new_paragraph(target)
    p.paragraph_format.left_indent = Cm(0.8)
    p.paragraph_format.first_line_indent = Cm(-0.8)
    append_text(p, f"{label}\t", size=size)
    append_text(p, text, size=size)
    return p


def _answer_line(target, label, width_twips):
    p = _new_paragraph(target, space_after=3)
    p.paragraph_format.left_indent = Cm(0.8)
    p.paragraph_format.tab_stops.add_tab_stop(Twips(width_twips - 250), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.LINES)
    append_text(p, f"{label}\t", size=QUESTION_PT)


def _render_section(target, spec, items, answers):
    """Insert one section (heading + questions or answers) in front of `target`."""
    heading = _new_paragraph(target, space_after=4, keep_next=True)
    append_text(heading, spec["heading"], bold=True, size=12)
    width = _cell_width_twips(target)
    unit_conversions = spec["key"] == "unit_conversions"
    for i, item in enumerate(items):
        label = f"{'abcdefgh'[i]})" if unit_conversions else f"{i + 1})"
        if answers:
            _numbered(target, label, item["answer"])
            continue
        if spec["answer_line"] == "inline":
            _numbered(target, label, f"{item['text']}  {BLANK}")
        else:
            _numbered(target, label, item["text"])
            if spec["answer_line"] == "two":
                _answer_line(target, "Working:", width)
            _answer_line(target, "Answer:", width)
    spacer = _new_paragraph(target, space_after=2)
    return spacer


def _render_basics(target, eq, basics, answers):
    for line_text in (eq.words[0].upper() + eq.words[1:], eq.letters):
        p = _new_paragraph(target, space_after=4)
        append_text(p, line_text, bold=True, size=12)
    for i, (lead, blanks, answer) in enumerate(basics):
        p = _numbered(target, f"{i + 1})", lead)
        if i == 2 and eq.triangle:
            letters = [eq.symbol[k] for k in eq.triangle] if answers else None
            p.add_run().add_picture(triangle_png(letters), height=Cm(2.4))
            for run in p.runs[:-1]:   # lift the question text up beside the middle of the triangle
                position = OxmlElement("w:position")
                position.set(qn("w:val"), "30")
                run._r.get_or_add_rPr().append(position)
        else:
            append_text(p, answer if answers else blanks, size=11)


def generate_documents(model, template_path):
    """Return {'worksheet': BytesIO, 'answers': BytesIO, 'warnings': [...]}."""
    eq = model["equation"]
    outputs, warnings = {}, []
    for mode in ("worksheet", "answers"):
        answers = mode == "answers"
        doc = Document(str(template_path))
        targets = {}
        for p in _all_paragraphs(doc):
            name = _placeholder_name(p)
            if name and name not in targets:
                targets[name] = p

        # --- title ---
        title = eq.title + (" – Answers" if answers else "")
        if "title" in targets:
            _clear_runs(targets["title"])
            append_text(targets["title"], title, bold=True, size=20)
        else:
            warnings.append("The template has no [[Title]] placeholder.")

        # --- basics ---
        if "basics" in targets:
            _render_basics(targets["basics"], eq, model["basics"], answers)
            targets["basics"]._p.getparent().remove(targets["basics"]._p)
        else:
            warnings.append("The template has no [[Basics]] placeholder.")

        # --- question sections, grouped by the paragraph they will be inserted in front of ---
        hosts = {}  # host placeholder name -> [section specs]
        for spec in SECTIONS:
            name = spec["placeholder"]
            if name in targets:
                host = name
            else:
                fallback = SPEC[FALLBACK_HOST[spec["key"]]]["placeholder"] if spec["key"] in FALLBACK_HOST else None
                if spec["key"] == "unit_conversions" and fallback in targets:
                    nxt = targets[fallback]._p.getnext()
                    if nxt is not None and nxt.tag == qn("w:p") and not "".join(nxt.itertext()).strip():
                        targets[name] = Paragraph(nxt, doc._body)   # the empty paragraph under [[Substitute]]
                        host = name
                    else:
                        host = fallback
                elif fallback in targets:
                    host = fallback
                else:
                    warnings.append(f"The template has no [[{spec['heading']}]] placeholder, so that section was left out.")
                    continue
            hosts.setdefault(host, []).append(spec)
        for host, specs in hosts.items():
            for spec in sorted(specs, key=SECTIONS.index):
                _render_section(targets[host], spec, model["sections"][spec["key"]], answers)
            targets[host]._p.getparent().remove(targets[host]._p)

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        outputs[mode] = buffer
    outputs["warnings"] = sorted(set(warnings))
    return outputs


def prepare_template_path(uploaded_template):
    if uploaded_template is None:
        if DEFAULT_TEMPLATE.exists():
            return DEFAULT_TEMPLATE
        raise FileNotFoundError(f"The default template '{DEFAULT_TEMPLATE.name}' was not found next to the app.")
    temp_dir = Path(tempfile.gettempdir()) / "formula_worksheet_templates"
    temp_dir.mkdir(exist_ok=True)
    temp_path = temp_dir / f"{uuid4().hex}.docx"
    temp_path.write_bytes(uploaded_template.getvalue())
    return temp_path


# --------------------------------------------------------------------------- #
# Streamlit UI
# --------------------------------------------------------------------------- #

def render_section_label(text):
    st.markdown(f"<strong>{text}</strong>", unsafe_allow_html=True)


def main():
    st.set_page_config(page_title="Physics Formula Worksheet Generator", page_icon="⚡")
    st.title("Physics Formula Worksheet Generator")
    st.write(
        "Choose an equation from the GCSE physics equation sheet. Gemini writes the questions, the app checks and "
        "calculates every answer itself, then builds a Word worksheet and a matching answer sheet."
    )

    render_section_label("Which equation?")
    equation_name = st.selectbox("Equation", list(EQUATIONS), label_visibility="collapsed")
    eq = EQUATIONS[equation_name]
    st.caption(f"{eq.words}   |   {plain_symbol(eq.letters)}")

    render_section_label("Year group (optional)")
    year_group = st.text_input("Year group", placeholder="e.g. Year 10", label_visibility="collapsed")
    render_section_label("Ability of the group (optional)")
    ability = st.text_input("Ability", placeholder="e.g. Lower-attaining / mixed / high ability", label_visibility="collapsed")
    render_section_label("Anything else Gemini should know? (optional)")
    notes = st.text_area("Notes", placeholder="e.g. Use sporting contexts. Keep numbers small.", height=80, label_visibility="collapsed")

    st.markdown("Watch this video to learn how to get a free Gemini API key")
    video = "https://www.youtube.com/embed/Uyn-P2nRvDA"
    if hasattr(st, "iframe"):   # newer Streamlit; components.iframe is deprecated
        st.iframe(video, height=158, width=280)
    else:
        components.iframe(video, height=158, width=280, scrolling=False)
    st.markdown("Create your Google Gemini API key at [https://aistudio.google.com/](https://aistudio.google.com/).")
    api_key_input = st.text_input("Gemini API key", type="password", placeholder="Enter your Gemini API key",
                                  help="This value is used for the request when provided.")
    uploaded_template = st.file_uploader("Optional: upload your own Word template (.docx)", type=["docx"])
    show_json = st.checkbox("Show Gemini JSON response (debug)", value=False)

    for key, default in (("fw_worksheet", None), ("fw_answers", None), ("fw_name", "worksheet"), ("fw_warnings", [])):
        st.session_state.setdefault(key, default)

    with st.expander("Template guidance", expanded=False):
        st.write(
            "Placeholders the app fills: [[Title]], [[Basics]], [[Substitute]], [[Unit Conversions]], [[Substitute Unit Conversions]], "
            "[[Rearrange]], [[Rearrange Unit Conversions]], [[Word Questions]], [[Multistep Questions]]. "
            "If [[Unit Conversions]] is missing it goes in the empty paragraph under [[Substitute]] (or straight after it); "
            "if [[Multistep Questions]] is missing it goes under the word questions."
        )
        if DEFAULT_TEMPLATE.exists():
            st.download_button("Download starter template", data=DEFAULT_TEMPLATE.read_bytes(), file_name=DEFAULT_TEMPLATE.name,
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    if st.button("Generate Worksheet", type="primary"):
        api_key = api_key_input.strip() or get_gemini_api_key()
        if not api_key:
            st.error("Enter a Gemini API key (or set the GEMINI_API_KEY environment variable / Streamlit secret) first.")
        else:
            try:
                with st.spinner("Asking Gemini for questions and checking every answer..."):
                    model, raw_json = generate_content(eq, api_key, year_group, ability, notes)
                if show_json:
                    st.subheader("Gemini JSON response")
                    st.json(raw_json)
                result = generate_documents(model, prepare_template_path(uploaded_template))
                st.session_state.update(
                    fw_worksheet=result["worksheet"].getvalue(), fw_answers=result["answers"].getvalue(),
                    fw_name=eq.name, fw_warnings=result["warnings"],
                )
                st.success("Worksheet generated successfully!")
            except Exception as exc:
                st.error(f"❌ An error occurred: {exc}")

    for warning in st.session_state["fw_warnings"]:
        st.warning(warning)
    mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    if st.session_state["fw_worksheet"] is not None:
        st.download_button("📥 Download Worksheet", data=st.session_state["fw_worksheet"],
                           file_name=f"{st.session_state['fw_name']}.docx", mime=mime)
    if st.session_state["fw_answers"] is not None:
        st.download_button("📥 Download Answers", data=st.session_state["fw_answers"],
                           file_name=f"{st.session_state['fw_name']} - Answers.docx", mime=mime)
        st.caption("Single-step answers are calculated by the app. Multistep working is written by Gemini, so skim it before using.")


if __name__ == "__main__":
    main()
