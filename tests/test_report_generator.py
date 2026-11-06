import unittest
from types import SimpleNamespace
from unittest.mock import patch

from core.report_generator import format_patent_documents, generate_vf_report_streaming


class _FakeResponses:
    def __init__(self, events):
        self.events = events
        self.kwargs = None

    def create(self, **kwargs):
        self.kwargs = kwargs
        return iter(self.events)


class _FakeClient:
    def __init__(self, events):
        self.responses = _FakeResponses(events)


class ReportGeneratorTests(unittest.TestCase):
    def test_format_patent_documents_keeps_every_document(self):
        documents = [
            ("first.pdf", "첫 번째 기술 " * 1000),
            ("second.docx", "두 번째 기술 " * 1000),
        ]

        context = format_patent_documents(documents, max_chars=4000)

        self.assertIn("특허 명세서 1/2: first.pdf", context)
        self.assertIn("특허 명세서 2/2: second.docx", context)
        self.assertIn("첫 번째 기술", context)
        self.assertIn("두 번째 기술", context)
        self.assertLess(len(context), 5000)

    def test_responses_stream_yields_text_and_uses_requested_model(self):
        fake_client = _FakeClient(
            [
                SimpleNamespace(type="response.output_text.delta", delta="첫째"),
                SimpleNamespace(type="response.output_text.delta", delta=" 둘째"),
                SimpleNamespace(type="response.completed"),
            ]
        )

        with patch("core.report_generator.OpenAI", return_value=fake_client):
            result = "".join(
                generate_vf_report_streaming(
                    api_key="test-key",
                    patent_documents=[("a.pdf", "특허 본문")],
                    section_name="Ⅰ. 기술 개요",
                    prev_sections="",
                    model="gpt-test",
                )
            )

        self.assertEqual(result, "첫째 둘째")
        self.assertEqual(fake_client.responses.kwargs["model"], "gpt-test")
        self.assertFalse(fake_client.responses.kwargs["store"])
        self.assertTrue(fake_client.responses.kwargs["stream"])
        self.assertIn("a.pdf", fake_client.responses.kwargs["input"])
        self.assertIn('# Ⅰ. 기술 개요', fake_client.responses.kwargs["input"])

    def test_responses_stream_raises_on_error_event(self):
        fake_client = _FakeClient(
            [SimpleNamespace(type="error", message="요청 실패")]
        )

        with patch("core.report_generator.OpenAI", return_value=fake_client):
            with self.assertRaisesRegex(RuntimeError, "요청 실패"):
                list(
                    generate_vf_report_streaming(
                        api_key="test-key",
                        patent_documents=[("a.pdf", "특허 본문")],
                        section_name="Ⅰ. 기술 개요",
                        prev_sections="",
                    )
                )

    def test_responses_stream_raises_on_incomplete_event(self):
        incomplete_response = SimpleNamespace(
            incomplete_details=SimpleNamespace(reason="max_output_tokens")
        )
        fake_client = _FakeClient(
            [
                SimpleNamespace(type="response.output_text.delta", delta="일부 내용"),
                SimpleNamespace(type="response.incomplete", response=incomplete_response),
            ]
        )

        with patch("core.report_generator.OpenAI", return_value=fake_client):
            with self.assertRaisesRegex(RuntimeError, "최대 출력 토큰 한도"):
                list(
                    generate_vf_report_streaming(
                        api_key="test-key",
                        patent_documents=[("a.pdf", "특허 본문")],
                        section_name="Ⅰ. 기술 개요",
                        prev_sections="",
                    )
                )


if __name__ == "__main__":
    unittest.main()
