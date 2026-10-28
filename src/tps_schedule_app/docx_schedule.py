from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
import io
import re
import zipfile
from xml.etree import ElementTree as ET

from .models import PersonSchedule


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
NS = {"w": WORD_NS}
DAY_ORDINAL_RE = re.compile(r"(.*?\b\d+)(st|nd|rd|th)(\b.*)", re.IGNORECASE)


@dataclass(frozen=True)
class TemplateMetadata:
    title_prefix: str
    sample_name: str
    day_labels: list[str]
    filename_pattern: str


def load_template_metadata(template_path: str | Path) -> TemplateMetadata:
    template = Path(template_path)
    with zipfile.ZipFile(template) as archive:
        document_xml = archive.read("word/document.xml")

    title_text, day_labels = _extract_template_text(document_xml)
    match = re.match(r"^(.*?Schedule for )(.+)$", title_text)
    if match:
        title_prefix = match.group(1)
        sample_name = match.group(2).strip()
    else:
        title_prefix = "Schedule for "
        sample_name = title_text.strip()

    filename_pattern = _infer_filename_pattern(template.name, sample_name)
    return TemplateMetadata(
        title_prefix=title_prefix,
        sample_name=sample_name,
        day_labels=day_labels,
        filename_pattern=filename_pattern,
    )


def build_day_label_map(
    sheet_names: list[str],
    template_day_labels: list[str],
    overrides: dict[str, str] | None = None,
) -> dict[str, str]:
    mapping = {sheet_name: sheet_name.title() for sheet_name in sheet_names}
    for index, label in enumerate(template_day_labels[: len(sheet_names)]):
        mapping[sheet_names[index]] = label
    if overrides:
        mapping.update(overrides)
    return mapping


def write_schedule_docx(
    person: PersonSchedule,
    template_path: str | Path,
    output_path: str | Path,
    title_prefix: str,
    day_label_map: dict[str, str],
) -> None:
    template = Path(template_path)
    output = Path(output_path)
    with zipfile.ZipFile(template) as archive:
        original_document = archive.read("word/document.xml")
        rendered_document = _render_document_xml(
            template_xml=original_document,
            person=person,
            title_prefix=title_prefix,
            day_label_map=day_label_map,
        )

        with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as target:
            for info in archive.infolist():
                payload = rendered_document if info.filename == "word/document.xml" else archive.read(info.filename)
                target.writestr(info, payload)


def sanitize_filename(filename: str) -> str:
    invalid = '<>:"/\\|?*'
    cleaned = filename.translate(str.maketrans({character: "_" for character in invalid}))
    cleaned = cleaned.strip().rstrip(".")
    return cleaned or "schedule.docx"


def _extract_template_text(document_xml: bytes) -> tuple[str, list[str]]:
    root = ET.fromstring(document_xml)
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("The DOCX template is missing a document body.")

    paragraphs = []
    for paragraph in body.findall("w:p", NS):
        text = _paragraph_text(paragraph)
        if not text:
            continue
        is_list = paragraph.find("w:pPr/w:numPr", NS) is not None
        paragraphs.append((text, is_list))

    if not paragraphs:
        raise ValueError("The DOCX template does not contain any text paragraphs.")

    title_text = paragraphs[0][0]
    day_labels = [text for text, is_list in paragraphs[1:] if not is_list]
    return title_text, day_labels


def _paragraph_text(paragraph: ET.Element) -> str:
    return "".join(
        text_node.text or ""
        for text_node in paragraph.iter(f"{{{WORD_NS}}}t")
    ).strip()


def _infer_filename_pattern(template_filename: str, sample_name: str) -> str:
    if sample_name and template_filename.startswith(sample_name):
        return "{name}" + template_filename[len(sample_name) :]
    return "{name} Schedule.docx"


def _render_document_xml(
    template_xml: bytes,
    person: PersonSchedule,
    title_prefix: str,
    day_label_map: dict[str, str],
) -> bytes:
    _register_template_namespaces(template_xml)
    root = ET.fromstring(template_xml)
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("The DOCX template is missing a document body.")

    section_properties = None
    if len(body) and body[-1].tag == _w("sectPr"):
        section_properties = deepcopy(body[-1])

    body.clear()
    body.append(_build_title_paragraph(f"{title_prefix}{person.name}"))

    if person.days:
        for index, day in enumerate(person.days):
            if index > 0:
                body.append(_build_blank_paragraph())
            day_label = day_label_map.get(day.sheet_name, day.sheet_name)
            body.append(_build_day_heading_paragraph(day_label))
            if day.events:
                for event in day.events:
                    body.append(_build_event_paragraph(event.time_label, event.description))
            else:
                body.append(_build_event_paragraph("", "No scheduled items."))
    else:
        body.append(_build_blank_paragraph())
        body.append(_build_event_paragraph("", "No scheduled items."))

    if section_properties is not None:
        body.append(section_properties)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _register_template_namespaces(template_xml: bytes) -> None:
    for event, namespace in ET.iterparse(io.BytesIO(template_xml), events=("start-ns",)):
        prefix, uri = namespace
        ET.register_namespace(prefix or "", uri)


def _build_title_paragraph(title_text: str) -> ET.Element:
    paragraph = ET.Element(_w("p"))
    properties = ET.SubElement(paragraph, _w("pPr"))
    ET.SubElement(properties, _w("jc"), {_w("val"): "center"})
    paragraph_run_properties = ET.SubElement(properties, _w("rPr"))
    _apply_run_style(paragraph_run_properties, bold=True, size=28)

    run = ET.SubElement(paragraph, _w("r"))
    run_properties = ET.SubElement(run, _w("rPr"))
    _apply_run_style(run_properties, bold=True, size=28)
    text = ET.SubElement(run, _w("t"))
    text.text = title_text
    return paragraph


def _build_day_heading_paragraph(label: str) -> ET.Element:
    paragraph = ET.Element(_w("p"))
    properties = ET.SubElement(paragraph, _w("pPr"))
    paragraph_run_properties = ET.SubElement(properties, _w("rPr"))
    _apply_run_style(paragraph_run_properties, bold=True)

    match = DAY_ORDINAL_RE.fullmatch(label)
    if match is None:
        _append_text_run(paragraph, label, bold=True)
        return paragraph

    _append_text_run(paragraph, match.group(1), bold=True)
    _append_text_run(paragraph, match.group(2), bold=True, superscript=True)
    _append_text_run(paragraph, match.group(3), bold=True)
    return paragraph


def _build_event_paragraph(time_label: str, description: str) -> ET.Element:
    paragraph = ET.Element(_w("p"))
    properties = ET.SubElement(paragraph, _w("pPr"))
    ET.SubElement(properties, _w("pStyle"), {_w("val"): "ListParagraph"})
    numbering = ET.SubElement(properties, _w("numPr"))
    ET.SubElement(numbering, _w("ilvl"), {_w("val"): "0"})
    ET.SubElement(numbering, _w("numId"), {_w("val"): "1"})
    paragraph_run_properties = ET.SubElement(properties, _w("rPr"))
    _apply_run_style(paragraph_run_properties, bold=True)

    if time_label:
        _append_text_run(paragraph, time_label)
        _append_tab_run(paragraph)
        _append_tab_run(paragraph)
    _append_text_run(paragraph, description)
    return paragraph


def _build_blank_paragraph() -> ET.Element:
    paragraph = ET.Element(_w("p"))
    properties = ET.SubElement(paragraph, _w("pPr"))
    paragraph_run_properties = ET.SubElement(properties, _w("rPr"))
    _apply_run_style(paragraph_run_properties, bold=True)
    return paragraph


def _append_text_run(
    paragraph: ET.Element,
    text_value: str,
    *,
    bold: bool = False,
    size: int | None = None,
    superscript: bool = False,
) -> None:
    run = ET.SubElement(paragraph, _w("r"))
    run_properties = ET.SubElement(run, _w("rPr"))
    _apply_run_style(run_properties, bold=bold, size=size, superscript=superscript)
    text = ET.SubElement(run, _w("t"))
    if text_value.startswith(" ") or text_value.endswith(" "):
        text.set(XML_SPACE, "preserve")
    text.text = text_value


def _append_tab_run(paragraph: ET.Element) -> None:
    run = ET.SubElement(paragraph, _w("r"))
    run_properties = ET.SubElement(run, _w("rPr"))
    _apply_run_style(run_properties)
    ET.SubElement(run, _w("tab"))


def _apply_run_style(
    run_properties: ET.Element,
    *,
    bold: bool = False,
    size: int | None = None,
    superscript: bool = False,
) -> None:
    ET.SubElement(
        run_properties,
        _w("rFonts"),
        {
            _w("ascii"): "Times New Roman",
            _w("hAnsi"): "Times New Roman",
            _w("cs"): "Times New Roman",
        },
    )
    if bold:
        ET.SubElement(run_properties, _w("b"))
        ET.SubElement(run_properties, _w("bCs"))
    if size is not None:
        ET.SubElement(run_properties, _w("sz"), {_w("val"): str(size)})
        ET.SubElement(run_properties, _w("szCs"), {_w("val"): str(size)})
    if superscript:
        ET.SubElement(run_properties, _w("vertAlign"), {_w("val"): "superscript"})


def _w(tag_name: str) -> str:
    return f"{{{WORD_NS}}}{tag_name}"
