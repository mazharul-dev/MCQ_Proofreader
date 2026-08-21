from __future__ import annotations

import base64
from dataclasses import dataclass, field
from html import escape
from io import BytesIO
import re
import unicodedata
from typing import Any

from docx import Document
from docx.table import Table
from lxml import etree


OPTION_LABELS = ["\u0995", "\u0996", "\u0997", "\u0998"]
ANSWER_INDEX = {
    "a": 0,
    "b": 1,
    "c": 2,
    "d": 3,
    "\u0995": 0,
    "\u0996": 1,
    "\u0997": 2,
    "\u0998": 3,
    "1": 0,
    "2": 1,
    "3": 2,
    "4": 3,
    "\u09e7": 0,
    "\u09e8": 1,
    "\u09e9": 2,
    "\u09ea": 3,
}
ANSWER_TOKEN_RE = re.compile(r"(?<![a-z])[abcd](?![a-z])|[\u0995\u0996\u0997\u0998\u09e7-\u09ea1-4]")
MATH_TOKEN_RE = re.compile(r"\s+|\d+(?:\.\d+)?|[A-Za-z]+|[+\-−=×÷*/<>≤≥±∞√(),.;:]|.")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
V_NS = "urn:schemas-microsoft-com:vml"
MML_NS = "http://www.w3.org/1998/Math/MathML"
NSMAP = {"w": W_NS, "m": M_NS, "a": A_NS, "r": R_NS, "v": V_NS}

SKIP_MATH_TAGS = {
    "accPr",
    "barPr",
    "boxPr",
    "ctrlPr",
    "dPr",
    "fPr",
    "groupChrPr",
    "jc",
    "lang",
    "limLoc",
    "naryPr",
    "rFonts",
    "rPr",
    "radPr",
    "sSubPr",
    "sSubSupPr",
    "sSupPr",
    "scr",
    "sty",
    "sz",
    "szCs",
}


@dataclass
class CellPart:
    type: str
    text: str = ""
    html: str = ""

    def to_dict(self) -> dict[str, str]:
        return {"type": self.type, "text": self.text, "html": self.html}


@dataclass
class CellContent:
    parts: list[CellPart] = field(default_factory=list)

    @property
    def text(self) -> str:
        return "".join(part.text for part in self.parts)

    @property
    def html(self) -> str:
        return "".join(part.html if part.type != "text" else escape(part.text) for part in self.parts)

    def to_parts(self) -> list[dict[str, str]]:
        return [part.to_dict() for part in self.parts]


@dataclass
class DuplicateTarget:
    serial: str
    uid: str

    def to_dict(self) -> dict[str, str]:
        return {"serial": self.serial, "uid": self.uid}


@dataclass
class MCQItem:
    uid: str
    serial: str
    category: str
    question: CellContent
    options: list[CellContent]
    explanation: CellContent
    answer_raw: str
    answer_index: int | None
    source_table: int
    warnings: list[str] = field(default_factory=list)
    duplicate_of: DuplicateTarget | None = None
    duplicate_targets: list[DuplicateTarget] = field(default_factory=list)

    @property
    def answer_label(self) -> str:
        if self.answer_index is None:
            return ""
        return OPTION_LABELS[self.answer_index]

    @property
    def answer_content(self) -> CellContent:
        if self.answer_index is None or self.answer_index >= len(self.options):
            return CellContent()
        return self.options[self.answer_index]

    @property
    def answer_text(self) -> str:
        return self.answer_content.text

    def to_dict(self) -> dict[str, Any]:
        return {
            "uid": self.uid,
            "serial": self.serial,
            "category": self.category,
            "question": self.question.text,
            "questionHtml": self.question.html,
            "questionParts": self.question.to_parts(),
            "options": [
                {
                    "key": key,
                    "label": label,
                    "text": content.text,
                    "html": content.html,
                    "parts": content.to_parts(),
                }
                for key, label, content in zip(["A", "B", "C", "D"], OPTION_LABELS, self.options)
            ],
            "explanation": self.explanation.text,
            "explanationHtml": self.explanation.html,
            "explanationParts": self.explanation.to_parts(),
            "answerRaw": self.answer_raw,
            "answerIndex": self.answer_index,
            "answerLabel": self.answer_label,
            "answerText": self.answer_text,
            "answerHtml": self.answer_content.html,
            "answerParts": self.answer_content.to_parts(),
            "duplicateOf": self.duplicate_of.to_dict() if self.duplicate_of else None,
            "duplicateTargets": [target.to_dict() for target in self.duplicate_targets],
            "sourceTable": self.source_table,
            "warnings": self.warnings,
        }


def parse_docx_bytes(data: bytes) -> dict[str, Any]:
    """Parse an uploaded DOCX from memory and return preview-ready MCQ data."""
    if not data:
        raise ValueError("The uploaded file is empty.")

    document = Document(BytesIO(data))
    questions: list[MCQItem] = []
    warnings: list[str] = []

    for table_index, table in enumerate(document.tables):
        if len(table.rows) < 8:
            warnings.append(
                f"Table {table_index + 1} skipped: expected at least 8 rows, found {len(table.rows)}."
            )
            continue

        item = _parse_question_table(table, table_index, document.part)
        if not item.question.text:
            warnings.append(f"Table {table_index + 1} skipped: question cell [1,0] is empty.")
            continue

        questions.append(item)

    duplicate_pairs = _mark_duplicates(questions)

    return {
        "total": len(questions),
        "duplicateCount": len(duplicate_pairs),
        "duplicatePairs": duplicate_pairs,
        "questions": [question.to_dict() for question in questions],
        "warnings": warnings,
    }


def _parse_question_table(table: Table, table_index: int, document_part: Any) -> MCQItem:
    serial = _cell_content(table, 0, 0, document_part).text
    category = _cell_content(table, 0, 1, document_part).text
    question = _cell_content(table, 1, 0, document_part)
    options = [_cell_content(table, row_index, 0, document_part) for row_index in range(2, 6)]
    explanation = _cell_content(table, 6, 0, document_part)
    answer_raw = _cell_content(table, 7, 0, document_part).text
    answer_index = _answer_index(answer_raw)
    warnings: list[str] = []

    if not serial:
        serial = str(table_index + 1)
        warnings.append("Serial cell [0,0] was empty; table order was used.")

    if answer_index is None:
        warnings.append(f"Answer cell [7,0] value '{answer_raw}' could not be mapped to A-D.")

    if any(not option.text for option in options):
        warnings.append("One or more option cells [2,0] to [5,0] are empty.")

    return MCQItem(
        uid=f"question-{table_index + 1}",
        serial=serial,
        category=category,
        question=question,
        options=options,
        explanation=explanation,
        answer_raw=answer_raw,
        answer_index=answer_index,
        source_table=table_index + 1,
        warnings=warnings,
    )


def _cell_content(table: Table, row_index: int, column_index: int, document_part: Any) -> CellContent:
    try:
        row = table.rows[row_index]
    except IndexError:
        return CellContent()

    if column_index >= len(row.cells):
        return CellContent()

    parts = _extract_cell_parts(row.cells[column_index], document_part)
    return CellContent(_clean_parts(parts))


def _extract_cell_parts(cell: Any, document_part: Any) -> list[CellPart]:
    parts: list[CellPart] = []
    has_content = False

    for child in cell._tc.iterchildren():
        if _local_name(child) != "p":
            continue

        paragraph_parts = _merge_adjacent_text(_extract_paragraph_parts(child, document_part))
        if not paragraph_parts:
            continue

        if has_content:
            parts.append(CellPart(type="text", text="\n"))
        parts.extend(paragraph_parts)
        has_content = True

    return parts


def _extract_paragraph_parts(paragraph: Any, document_part: Any) -> list[CellPart]:
    parts: list[CellPart] = []

    for child in paragraph.iterchildren():
        parts.extend(_extract_node_parts(child, document_part))

    return parts


def _extract_node_parts(element: Any, document_part: Any) -> list[CellPart]:
    local_name = _local_name(element)

    if local_name == "r":
        return _extract_run_parts(element, document_part)
    if local_name == "hyperlink":
        return _extract_children_parts(element, document_part)
    if local_name in {"oMath", "oMathPara"}:
        return [_math_part(element)]
    if local_name == "drawing":
        return _image_parts(element, document_part)
    if local_name == "pict":
        return _image_parts(element, document_part)

    return []


def _extract_run_parts(run: Any, document_part: Any) -> list[CellPart]:
    parts: list[CellPart] = []

    for child in run.iterchildren():
        local_name = _local_name(child)
        if local_name == "t":
            parts.append(CellPart(type="text", text=child.text or ""))
        elif local_name == "tab":
            parts.append(CellPart(type="text", text=" "))
        elif local_name in {"br", "cr"}:
            parts.append(CellPart(type="text", text="\n"))
        elif local_name in {"drawing", "pict"}:
            parts.extend(_image_parts(child, document_part))
        elif local_name in {"oMath", "oMathPara"}:
            parts.append(_math_part(child))
        elif list(child):
            parts.extend(_extract_children_parts(child, document_part))

    return parts


def _extract_children_parts(element: Any, document_part: Any) -> list[CellPart]:
    parts: list[CellPart] = []

    for child in element.iterchildren():
        parts.extend(_extract_node_parts(child, document_part))

    return parts


def _image_parts(element: Any, document_part: Any) -> list[CellPart]:
    parts: list[CellPart] = []
    relationship_ids: list[str] = []
    width_px, height_px = _image_size(element)

    for child in element.iter():
        local_name = _local_name(child)
        if local_name == "blip":
            relationship_id = child.get(f"{{{R_NS}}}embed") or child.get(f"{{{R_NS}}}link")
            if relationship_id:
                relationship_ids.append(relationship_id)
        elif local_name == "imagedata":
            relationship_id = child.get(f"{{{R_NS}}}id")
            if relationship_id:
                relationship_ids.append(relationship_id)

    for relationship_id in dict.fromkeys(relationship_ids):
        related_part = document_part.related_parts.get(relationship_id)
        blob = getattr(related_part, "blob", None)
        content_type = getattr(related_part, "content_type", "image/png")
        if not blob:
            continue

        encoded = base64.b64encode(blob).decode("ascii")
        style = ""
        if width_px and height_px:
            style = (
                f' style="width:{width_px:.1f}px;height:{height_px:.1f}px;'
                'max-width:100%;object-fit:contain;"'
            )
        html = (
            f'<img class="docx-image" src="data:{escape(content_type, quote=True)};base64,{encoded}" '
            f'alt="Embedded DOCX image" loading="lazy"{style} />'
        )
        parts.append(CellPart(type="image", text="", html=html))

    return parts


def _image_size(element: Any) -> tuple[float | None, float | None]:
    for child in element.iter():
        if _local_name(child) != "extent":
            continue

        try:
            width_emu = float(child.get("cx"))
            height_emu = float(child.get("cy"))
        except (TypeError, ValueError):
            continue

        if width_emu > 0 and height_emu > 0:
            return width_emu / 914400 * 96, height_emu / 914400 * 96

    return None, None


def _math_part(element: Any) -> CellPart:
    plain = _clean_plain_text(_omml_plain_text(element))
    mathml = _omml_mathml_document(element)
    return CellPart(type="math", text=plain, html=mathml)


def _omml_mathml_document(element: Any) -> str:
    body = _omml_children_mathml(element)
    if not body:
        body = _math_tokens(_omml_plain_text(element))

    return f'<math class="docx-math" xmlns="{MML_NS}"><mrow>{body}</mrow></math>'


def _omml_to_mathml(element: Any) -> str:
    local_name = _local_name(element)

    if local_name in SKIP_MATH_TAGS:
        return ""
    if local_name == "t":
        return _math_tokens(element.text or "")
    if local_name == "r":
        return _omml_children_mathml(element)
    if local_name == "sSup":
        return f"<msup>{_math_group(_child_by_local(element, 'e'))}{_math_group(_child_by_local(element, 'sup'))}</msup>"
    if local_name == "sSub":
        return f"<msub>{_math_group(_child_by_local(element, 'e'))}{_math_group(_child_by_local(element, 'sub'))}</msub>"
    if local_name == "sSubSup":
        return (
            f"<msubsup>{_math_group(_child_by_local(element, 'e'))}"
            f"{_math_group(_child_by_local(element, 'sub'))}{_math_group(_child_by_local(element, 'sup'))}</msubsup>"
        )
    if local_name == "f":
        return f"<mfrac>{_math_group(_child_by_local(element, 'num'))}{_math_group(_child_by_local(element, 'den'))}</mfrac>"
    if local_name == "rad":
        degree = _child_by_local(element, "deg")
        expression = _math_group(_child_by_local(element, "e"))
        if degree is not None and _omml_plain_text(degree).strip():
            return f"<mroot>{expression}{_math_group(degree)}</mroot>"
        return f"<msqrt>{expression}</msqrt>"
    if local_name == "d":
        begin = _delimiter_value(element, "begChr", "(")
        end = _delimiter_value(element, "endChr", ")")
        return f"<mrow>{_math_operator(begin)}{_math_group(_child_by_local(element, 'e'))}{_math_operator(end)}</mrow>"
    if local_name == "nary":
        operator = _delimiter_value(element, "chr", "\u2211")
        expression = _math_group(_child_by_local(element, "e"))
        sub = _child_by_local(element, "sub")
        sup = _child_by_local(element, "sup")
        if sub is not None and sup is not None:
            return f"<mrow><munderover>{_math_operator(operator)}{_math_group(sub)}{_math_group(sup)}</munderover>{expression}</mrow>"
        if sub is not None:
            return f"<mrow><munder>{_math_operator(operator)}{_math_group(sub)}</munder>{expression}</mrow>"
        if sup is not None:
            return f"<mrow><mover>{_math_operator(operator)}{_math_group(sup)}</mover>{expression}</mrow>"
        return f"<mrow>{_math_operator(operator)}{expression}</mrow>"
    if local_name == "bar":
        return f'<mover accent="true">{_math_group(_child_by_local(element, "e"))}<mo>¯</mo></mover>'
    if local_name == "limLow":
        return f"<munder>{_math_group(_child_by_local(element, 'e'))}{_math_group(_child_by_local(element, 'lim'))}</munder>"
    if local_name == "limUpp":
        return f"<mover>{_math_group(_child_by_local(element, 'e'))}{_math_group(_child_by_local(element, 'lim'))}</mover>"
    if local_name == "eqArr":
        rows = "".join(f"<mtr><mtd>{_math_group(child)}</mtd></mtr>" for child in _children_by_local(element, "e"))
        return f"<mtable>{rows}</mtable>"

    return _omml_children_mathml(element)


def _omml_children_mathml(element: Any) -> str:
    return "".join(_omml_to_mathml(child) for child in element.iterchildren())


def _math_group(element: Any | None) -> str:
    if element is None:
        return "<mrow></mrow>"
    return f"<mrow>{_omml_children_mathml(element)}</mrow>"


def _math_tokens(text: str) -> str:
    tokens: list[str] = []

    for token in MATH_TOKEN_RE.findall(text):
        if not token:
            continue
        if token.isspace():
            tokens.append('<mspace width="0.25em"></mspace>')
        elif re.fullmatch(r"\d+(?:\.\d+)?", token):
            tokens.append(f"<mn>{escape(token)}</mn>")
        elif re.fullmatch(r"[A-Za-z]+", token):
            tokens.append(f"<mi>{escape(token)}</mi>")
        elif re.fullmatch(r"[+\-−=×÷*/<>≤≥±∞√(),.;:]", token):
            tokens.append(_math_operator(token))
        else:
            tokens.append(f"<mtext>{escape(token)}</mtext>")

    return "".join(tokens)


def _math_operator(value: str) -> str:
    return f"<mo>{escape(value)}</mo>" if value else ""


def _omml_plain_text(element: Any) -> str:
    local_name = _local_name(element)

    if local_name in SKIP_MATH_TAGS:
        return ""
    if local_name == "t":
        return element.text or ""
    if local_name == "f":
        numerator = _omml_plain_text(_child_by_local(element, "num"))
        denominator = _omml_plain_text(_child_by_local(element, "den"))
        return f"({numerator})/({denominator})"
    if local_name == "sSup":
        base = _omml_plain_text(_child_by_local(element, "e"))
        superscript = _omml_plain_text(_child_by_local(element, "sup"))
        return f"{base}^{superscript}"
    if local_name == "sSub":
        base = _omml_plain_text(_child_by_local(element, "e"))
        subscript = _omml_plain_text(_child_by_local(element, "sub"))
        return f"{base}_{subscript}"
    if local_name == "sSubSup":
        base = _omml_plain_text(_child_by_local(element, "e"))
        subscript = _omml_plain_text(_child_by_local(element, "sub"))
        superscript = _omml_plain_text(_child_by_local(element, "sup"))
        return f"{base}_{subscript}^{superscript}"
    if local_name == "d":
        begin = _delimiter_value(element, "begChr", "(")
        end = _delimiter_value(element, "endChr", ")")
        return f"{begin}{_omml_plain_text(_child_by_local(element, 'e'))}{end}"
    if element is None:
        return ""

    return "".join(_omml_plain_text(child) for child in element.iterchildren())


def _delimiter_value(element: Any, local_name: str, default: str) -> str:
    for child in element.iter():
        if _local_name(child) == local_name:
            return child.get(f"{{{M_NS}}}val") or child.get("val") or default
    return default


def _child_by_local(element: Any | None, local_name: str) -> Any | None:
    if element is None:
        return None

    for child in element.iterchildren():
        if _local_name(child) == local_name:
            return child

    return None


def _children_by_local(element: Any, local_name: str) -> list[Any]:
    return [child for child in element.iterchildren() if _local_name(child) == local_name]


def _clean_parts(parts: list[CellPart]) -> list[CellPart]:
    cleaned = _merge_adjacent_text(parts)

    for part in cleaned:
        if part.type == "text":
            part.text = part.text.replace("\r\n", "\n").replace("\r", "\n").replace("\u00a0", " ")
            part.text = re.sub(r"[ \t]+", " ", part.text)
        elif part.type == "math":
            part.text = _clean_plain_text(part.text)

    if cleaned and cleaned[0].type == "text":
        cleaned[0].text = cleaned[0].text.lstrip()
    if cleaned and cleaned[-1].type == "text":
        cleaned[-1].text = cleaned[-1].text.rstrip()

    return [part for part in _merge_adjacent_text(cleaned) if part.type != "text" or part.text]


def _merge_adjacent_text(parts: list[CellPart]) -> list[CellPart]:
    merged: list[CellPart] = []

    for part in parts:
        if part.type == "text" and merged and merged[-1].type == "text":
            merged[-1].text += part.text
        else:
            merged.append(CellPart(type=part.type, text=part.text, html=part.html))

    return merged


def _clean_plain_text(text: str) -> str:
    text = (text or "").replace("\r\n", "\n").replace("\r", "\n").replace("\u00a0", " ")
    lines = [re.sub(r"[ \t]+", " ", line).strip() for line in text.split("\n")]

    cleaned_lines: list[str] = []
    for line in lines:
        if line:
            cleaned_lines.append(line)
        elif cleaned_lines and cleaned_lines[-1] != "":
            cleaned_lines.append("")

    while cleaned_lines and cleaned_lines[-1] == "":
        cleaned_lines.pop()

    return "\n".join(cleaned_lines).strip()


def _answer_index(raw_answer: str) -> int | None:
    answer = unicodedata.normalize("NFKC", raw_answer or "").strip().lower()
    answer = answer.strip("()[]{}.:।-–— ")

    if answer in ANSWER_INDEX:
        return ANSWER_INDEX[answer]

    match = ANSWER_TOKEN_RE.search(answer)
    if not match:
        return None

    return ANSWER_INDEX.get(match.group(0))


def _mark_duplicates(questions: list[MCQItem]) -> list[dict[str, Any]]:
    seen: dict[str, MCQItem] = {}
    duplicate_pairs: list[dict[str, Any]] = []

    for question in questions:
        normalized = _normalize_question(question.question.text)
        if not normalized:
            continue

        first = seen.get(normalized)
        if first is None:
            seen[normalized] = question
            continue

        question.duplicate_of = DuplicateTarget(serial=first.serial, uid=first.uid)
        question.duplicate_targets.append(DuplicateTarget(serial=first.serial, uid=first.uid))
        first.duplicate_targets.append(DuplicateTarget(serial=question.serial, uid=question.uid))
        duplicate_pairs.append(
            {
                "source": {"serial": first.serial, "uid": first.uid},
                "repeat": {"serial": question.serial, "uid": question.uid},
            }
        )

    return duplicate_pairs


def _normalize_question(text: str) -> str:
    normalized = unicodedata.normalize("NFKC", text or "").casefold()
    normalized = normalized.replace("’", "'").replace("‘", "'").replace("“", '"').replace("”", '"')
    return re.sub(r"[\W_]+", "", normalized, flags=re.UNICODE)


def _local_name(element: Any) -> str:
    if element is None:
        return ""
    return etree.QName(element).localname
