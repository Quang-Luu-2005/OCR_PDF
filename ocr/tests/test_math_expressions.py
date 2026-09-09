import base64
import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
import zipfile

import fitz

from src.export.exporter import WordExporter
from src.processing.math_expression_processor import (
    EquationRecord,
    MathExpressionProcessor,
)


ONE_PIXEL_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9ZlKsAAAAASUVORK5CYII="
)


def response(payload):
    return SimpleNamespace(
        choices=[SimpleNamespace(message=SimpleNamespace(content=json.dumps(payload)))],
        usage=SimpleNamespace(prompt_tokens=20, completion_tokens=10, total_tokens=30),
    )


class FakeCompletions:
    def __init__(self, payload, failures=0):
        self.payload = payload
        self.failures = failures
        self.calls = []

    def create(self, **kwargs):
        self.calls.append(kwargs)
        if self.failures:
            self.failures -= 1
            raise RuntimeError("429")
        return response(self.payload)


class FakePage:
    rect = fitz.Rect(0, 0, 500, 700)

    class Pixmap:
        @staticmethod
        def save(path):
            Path(path).write_bytes(ONE_PIXEL_PNG)

    def get_pixmap(self, **kwargs):
        return self.Pixmap()


class MathExpressionTests(unittest.TestCase):
    def test_detects_numbered_equation_and_calls_vision_with_image(self):
        payload = {
            "latex": r"Z=\frac{X-\min(X)}{\max(X)-\min(X)}",
            "equation_number": "1",
            "plain_text": "Z equals normalized X",
            "symbols": ["Z", "X", "min", "max"],
            "confidence": 0.98,
        }
        completions = FakeCompletions(payload)
        client = SimpleNamespace(chat=SimpleNamespace(completions=completions))
        blocks = [
            (250.0, 200.0, 490.0, 225.0, "(1) Z = X - min(X)", 0, 0),
            (330.0, 224.0, 340.0, 232.0, "N", 0, 0),
            (420.0, 210.0, 425.0, 218.0, ",", 0, 0),
        ]

        with tempfile.TemporaryDirectory() as root:
            processor = MathExpressionProcessor(
                client=client,
                output_dir=Path(root) / "equations",
                cache_dir=Path(root) / "cache",
                sleep_fn=lambda _: None,
            )
            records = processor.extract_page(FakePage(), 3, blocks, "Paper")

            self.assertEqual(1, len(records))
            self.assertTrue(records[0].validated)
            self.assertEqual([0, 1, 2], records[0].consumed_block_indexes)
            self.assertEqual("omml", records[0].output_mode)
            self.assertIn("$$Z=", records[0].markdown)
            content = completions.calls[0]["messages"][-1]["content"]
            self.assertEqual("image_url", content[1]["type"])
            self.assertTrue(content[1]["image_url"]["url"].startswith("data:image/png;base64,"))

    def test_low_confidence_and_api_failure_use_image_fallback(self):
        payload = {
            "latex": r"x=1",
            "equation_number": "2",
            "plain_text": "x equals one",
            "symbols": ["x", "1"],
            "confidence": 0.2,
        }
        completions = FakeCompletions(payload)
        client = SimpleNamespace(chat=SimpleNamespace(completions=completions))
        blocks = [(250.0, 200.0, 490.0, 225.0, "(2) x = 1", 0, 0)]
        with tempfile.TemporaryDirectory() as root:
            processor = MathExpressionProcessor(
                client=client, output_dir=root, min_confidence=0.9
            )
            record = processor.extract_page(FakePage(), 1, blocks, "Paper")[0]
            self.assertFalse(record.validated)
            self.assertIn("![equation:", record.markdown)

    def test_cache_avoids_duplicate_vision_calls(self):
        payload = {
            "latex": r"x=1",
            "equation_number": "1",
            "plain_text": "x equals one",
            "symbols": ["x", "1"],
            "confidence": 0.99,
        }
        completions = FakeCompletions(payload)
        client = SimpleNamespace(chat=SimpleNamespace(completions=completions))
        blocks = [(250.0, 200.0, 490.0, 225.0, "(1) x = 1", 0, 0)]
        with tempfile.TemporaryDirectory() as root:
            processor = MathExpressionProcessor(
                client=client,
                output_dir=Path(root) / "equations",
                cache_dir=Path(root) / "cache",
            )
            processor.extract_page(FakePage(), 1, blocks, "Paper")
            processor.extract_page(FakePage(), 1, blocks, "Paper")
            self.assertEqual(1, len(completions.calls))
            self.assertEqual(1, processor.usage.cache_hits)

    def test_word_export_contains_native_omml_and_image_fallback(self):
        with tempfile.TemporaryDirectory() as root:
            root_path = Path(root)
            image_path = root_path / "equation.png"
            image_path.write_bytes(ONE_PIXEL_PNG)
            native = EquationRecord(
                "eq_native", 1, (0, 0, 100, 20), "1", r"x=\frac{1}{2}",
                "", [], "", str(image_path), 0.99, True, "omml",
            )
            fallback = EquationRecord(
                "eq_image", 1, (0, 0, 100, 20), "2", None,
                "", [], "", str(image_path), 0.0, False, "image",
            )
            markdown = (
                "$$x=\\frac{1}{2}$$ <!-- equation:id=eq_native;number=1 -->\n"
                "![equation: eq_image](equation.png) <!-- equation:number=2 -->\n"
                "Inline $y^2$ value.\n"
            )
            output = root_path / "math.docx"
            WordExporter().markdown_to_word(
                markdown, output, equations=[native, fallback]
            )
            with zipfile.ZipFile(output) as package:
                document_xml = package.read("word/document.xml").decode("utf-8")
                package_names = package.namelist()
            self.assertGreaterEqual(document_xml.count("<m:oMath"), 2)
            self.assertIn("(1)", document_xml)
            self.assertTrue(any(name.startswith("word/media/") for name in package_names))

    def test_large_operator_uses_portable_image_fallback(self):
        self.assertFalse(WordExporter._native_equation_is_portable(r"\sum_{i=1}^N x_i"))
        self.assertTrue(WordExporter._native_equation_is_portable(r"x=\frac{1}{2}"))


if __name__ == "__main__":
    unittest.main()
