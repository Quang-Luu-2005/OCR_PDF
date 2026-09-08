import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from src.processing.scientific_translator import (
    GlossaryTerm,
    TranslationResult,
    TranslationUsage,
)


class DummyOCR:
    image_path = None

    def __init__(self, *args, **kwargs):
        pass

    def process_pdf(self, pdf_path):
        return {
            "markdown": "# Results\n\n[IMAGE_PLACEHOLDER_1]\n",
            "images": [{"file_path": str(self.image_path)}],
        }


class DummyDigitalParser:
    def is_digital_pdf(self, pdf_path):
        return True

    def convert(self, pdf_path, output_path):
        Path(output_path).write_bytes(b"digital-docx")


class DummyExporter:
    calls = []

    def markdown_to_word(self, markdown, output_path, images=None):
        self.calls.append((markdown, Path(output_path), images or []))
        Path(output_path).write_bytes(b"word-with-image")
        return output_path


class DummyTranslator:
    model = "gemini-3.8-flash"
    base_url = "https://generativelanguage.googleapis.com/v1beta/openai/"
    cache_dir = None

    def __init__(self, fail=False):
        self.calls = 0
        self.fail = fail

    def process(self, markdown, document_title=""):
        self.calls += 1
        if self.fail:
            raise RuntimeError("translation failed")
        return TranslationResult(
            corrected_markdown=markdown.replace("Results", "Corrected Results"),
            vi_markdown=markdown.replace("Results", "Kết quả"),
            glossary=[GlossaryTerm("Results", "Kết quả")],
            usage=TranslationUsage(
                input_tokens=10,
                output_tokens=8,
                total_tokens=18,
                chunks=1,
                retries=1,
            ),
        )


class PipelineTranslationTests(unittest.TestCase):
    def setUp(self):
        DummyExporter.calls = []

    def make_pipeline(self, root, translator=None, enabled=True):
        from src.core.pipeline import OCRPipeline

        with (
            patch("src.core.pipeline.OCREngine", DummyOCR),
            patch("src.core.pipeline.DigitalParser", DummyDigitalParser),
            patch("src.core.pipeline.WordExporter", DummyExporter),
        ):
            return OCRPipeline(
                output_dir=Path(root) / "output",
                temp_dir=Path(root) / "temp",
                enable_vi_translation=enabled,
                translator=translator,
            )

    def prepare_input(self, root):
        pdf_path = Path(root) / "paper.pdf"
        pdf_path.write_bytes(b"%PDF-test")
        image_path = Path(root) / "source.png"
        image_path.write_bytes(b"png")
        DummyOCR.image_path = image_path
        return pdf_path

    def test_scan_writes_corrected_and_vietnamese_outputs_atomically(self):
        with tempfile.TemporaryDirectory() as root:
            translator = DummyTranslator()
            pipeline = self.make_pipeline(root, translator)
            pdf_path = self.prepare_input(root)
            output_path = pipeline.process_pdf(pdf_path, mode="scan")

            self.assertTrue(output_path.exists())
            self.assertIn("Corrected Results", pipeline.output_artifacts["corrected_markdown"].read_text(encoding="utf-8"))
            self.assertIn("Kết quả", pipeline.output_artifacts["vi_markdown"].read_text(encoding="utf-8"))
            self.assertTrue(pipeline.output_artifacts["vi_docx"].exists())
            self.assertEqual(2, len(DummyExporter.calls))
            self.assertTrue(all(call[2] for call in DummyExporter.calls))
            self.assertEqual(18, pipeline.metrics.metrics["translation"]["total_tokens"])

    def test_digital_keeps_pdf2docx_and_also_builds_translated_word(self):
        with tempfile.TemporaryDirectory() as root:
            translator = DummyTranslator()
            pipeline = self.make_pipeline(root, translator)
            pdf_path = self.prepare_input(root)
            output_path = pipeline.process_pdf(pdf_path, mode="digital")

            self.assertEqual(b"digital-docx", output_path.read_bytes())
            self.assertTrue(pipeline.output_artifacts["vi_docx"].exists())
            self.assertEqual(1, len(DummyExporter.calls))
            self.assertEqual(1, translator.calls)

    def test_no_translate_digital_never_calls_marker_or_api(self):
        with tempfile.TemporaryDirectory() as root:
            pipeline = self.make_pipeline(root, translator=None, enabled=False)
            pdf_path = self.prepare_input(root)
            with patch.object(pipeline.ocr_engine, "process_pdf", side_effect=AssertionError("Marker called")):
                output_path = pipeline.process_pdf(pdf_path, mode="digital")

            self.assertEqual(b"digital-docx", output_path.read_bytes())
            self.assertFalse(pipeline.metrics.metrics["translation"]["enabled"])

    def test_failed_translation_does_not_leave_vi_outputs(self):
        with tempfile.TemporaryDirectory() as root:
            pipeline = self.make_pipeline(root, DummyTranslator(fail=True))
            pdf_path = self.prepare_input(root)
            with self.assertRaises(RuntimeError):
                pipeline.process_pdf(pdf_path, mode="scan")

            output_dir = Path(root) / "output"
            self.assertFalse((output_dir / "paper_ocr_results_vi.md").exists())
            self.assertFalse((output_dir / "paper_ocr_results_vi.docx").exists())


if __name__ == "__main__":
    unittest.main()
