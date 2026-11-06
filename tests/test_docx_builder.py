import io
import unittest
import xml.etree.ElementTree as ET
import zipfile

from docx import Document

from core.docx_builder import build_docx


class DocxBuilderTests(unittest.TestCase):
    def test_generated_docx_is_valid_and_keeps_heading_levels(self):
        report = """# Ⅰ. 기술 개요
### Ⅰ-1. 기술명 및 기술 정의
- 핵심 기술 설명

| 구분 | 내용 |
| --- | --- |
| 기술 | 테스트 |
"""

        docx_bytes = build_docx(report, patent_name="복수 특허 포트폴리오")

        document = Document(io.BytesIO(docx_bytes))
        paragraph_text = [paragraph.text for paragraph in document.paragraphs]
        self.assertTrue(any("Ⅰ. 기술 개요" in text for text in paragraph_text))
        self.assertTrue(any("Ⅰ-1. 기술명" in text for text in paragraph_text))

        with zipfile.ZipFile(io.BytesIO(docx_bytes)) as archive:
            for name in archive.namelist():
                if name.endswith(".xml"):
                    ET.fromstring(archive.read(name))


if __name__ == "__main__":
    unittest.main()
