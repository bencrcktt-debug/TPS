from pathlib import Path
import shutil
import sys
import unittest
import uuid
import zipfile


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC = PROJECT_ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from tps_schedule_app.app import generate_documents
from tps_schedule_app.excel_schedule import load_workbook_schedule


class TestScheduleGenerator(unittest.TestCase):
    def test_load_workbook_schedule_extracts_people_and_ranges(self):
        tmp_path = _workspace_temp_dir()
        try:
            workbook_path = _create_test_workbook(tmp_path / "schedule.xlsx")

            workbook = load_workbook_schedule(workbook_path)

            self.assertEqual(workbook.sheet_names, ["WED", "THURS"])
            self.assertEqual(len(workbook.people), 1)
            self.assertEqual(workbook.people[0].name, "Alice Smith")
            self.assertEqual(
                [event.time_label for event in workbook.people[0].days[0].events],
                ["9:15 am", "9:45 \u2013 10:30 am"],
            )
            self.assertEqual(
                [event.description for event in workbook.people[0].days[1].events],
                ["Panel Discussion"],
            )
        finally:
            shutil.rmtree(tmp_path, ignore_errors=True)

    def test_generate_documents_writes_docx_per_person(self):
        tmp_path = _workspace_temp_dir()
        try:
            workbook_path = _create_test_workbook(tmp_path / "schedule.xlsx")
            template_path = _create_test_template(tmp_path / "Alice Smith TPS 2026.docx")
            output_dir = tmp_path / "output"

            created_files = generate_documents(workbook_path, template_path, output_dir)

            self.assertEqual(len(created_files), 1)
            output_file = created_files[0]
            self.assertEqual(output_file.name, "Alice Smith TPS 2026.docx")

            with zipfile.ZipFile(output_file) as archive:
                document_xml = archive.read("word/document.xml").decode("utf-8")

            self.assertIn("Texas Policy Summit 2026 - Schedule for Alice Smith", document_xml)
            self.assertIn("Wednesday, April 8", document_xml)
            self.assertIn("9:15 am", document_xml)
            self.assertIn("9:45 \u2013 10:30 am", document_xml)
            self.assertIn("Opening Session", document_xml)
            self.assertIn("Thursday, April 9", document_xml)
            self.assertIn("8:30 \u2013 9:15 am", document_xml)
            self.assertIn("Panel Discussion", document_xml)
        finally:
            shutil.rmtree(tmp_path, ignore_errors=True)


def _create_test_workbook(path: Path) -> Path:
    shared_strings = [
        "W 9:00 am",
        "W 9:15 am",
        "W 9:30 am",
        "W 9:45 am",
        "W 10:00 am",
        "W 10:15 am",
        "W 10:30 am",
        "TH 8:00 am",
        "TH 8:15 am",
        "TH 8:30 am",
        "TH 8:45 am",
        "TH 9:00 am",
        "TH 9:15 am",
        "Alice",
        "Smith",
        "Arrive by 9:15 am",
        "Opening Session",
        "Panel Discussion",
    ]

    workbook_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="WED" sheetId="1" r:id="rId1"/>
    <sheet name="THURS" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>
"""
    workbook_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>
"""
    shared_strings_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="18" uniqueCount="18">
  {items}
</sst>
""".format(
        items="".join(f"<si><t>{value}</t></si>" for value in shared_strings)
    )
    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="3" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0"/>
  </cellXfs>
</styleSheet>
"""
    sheet1_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:H9"/>
  <sheetData>
    <row r="1">
      <c r="B1" t="s"><v>0</v></c>
      <c r="C1" t="s"><v>1</v></c>
      <c r="D1" t="s"><v>2</v></c>
      <c r="E1" t="s"><v>3</v></c>
      <c r="F1" t="s"><v>4</v></c>
      <c r="G1" t="s"><v>5</v></c>
      <c r="H1" t="s"><v>6</v></c>
    </row>
    <row r="9">
      <c r="A9" t="s"><v>13</v></c>
      <c r="B9" t="s"><v>14</v></c>
      <c r="C9" s="2" t="s"><v>15</v></c>
      <c r="E9" s="3" t="s"><v>16</v></c>
      <c r="F9" s="4"/>
      <c r="G9" s="4"/>
    </row>
  </sheetData>
</worksheet>
"""
    sheet2_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:G9"/>
  <sheetData>
    <row r="1">
      <c r="B1" t="s"><v>7</v></c>
      <c r="C1" t="s"><v>8</v></c>
      <c r="D1" t="s"><v>9</v></c>
      <c r="E1" t="s"><v>10</v></c>
      <c r="F1" t="s"><v>11</v></c>
      <c r="G1" t="s"><v>12</v></c>
    </row>
    <row r="9">
      <c r="A9" t="s"><v>13</v></c>
      <c r="B9" t="s"><v>14</v></c>
      <c r="D9" s="3" t="s"><v>17</v></c>
      <c r="E9" s="4"/>
      <c r="F9" s="4"/>
    </row>
  </sheetData>
</worksheet>
"""
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
"""
    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"""

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr("xl/workbook.xml", workbook_xml)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        archive.writestr("xl/sharedStrings.xml", shared_strings_xml)
        archive.writestr("xl/styles.xml", styles_xml)
        archive.writestr("xl/worksheets/sheet1.xml", sheet1_xml)
        archive.writestr("xl/worksheets/sheet2.xml", sheet2_xml)

    return path


def _create_test_template(path: Path) -> Path:
    document_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Texas Policy Summit 2026 - Schedule for Alice Smith</w:t></w:r></w:p>
    <w:p><w:r><w:t>Wednesday, April 8th</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>9:15 am</w:t></w:r></w:p>
    <w:p><w:r><w:t>Thursday, April 9th</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>
  </w:body>
</w:document>
"""
    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph">
    <w:name w:val="List Paragraph"/>
  </w:style>
</w:styles>
"""
    numbering_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val=""/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>
"""
    doc_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>
"""
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>
"""
    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr("word/document.xml", document_xml)
        archive.writestr("word/styles.xml", styles_xml)
        archive.writestr("word/numbering.xml", numbering_xml)
        archive.writestr("word/_rels/document.xml.rels", doc_rels)

    return path


def _workspace_temp_dir() -> Path:
    root = PROJECT_ROOT / "test-output"
    root.mkdir(parents=True, exist_ok=True)
    temp_dir = root / uuid.uuid4().hex
    temp_dir.mkdir()
    return temp_dir


if __name__ == "__main__":
    unittest.main()
