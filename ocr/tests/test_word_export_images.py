import base64
import tempfile
import unittest
from pathlib import Path

from docx import Document

from src.export.exporter import WordExporter
from src.processing.markdown_processor import MarkdownProcessor


ONE_PIXEL_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9ZlKsAAAAASUVORK5CYII="
)


class WordExporterImageTests(unittest.TestCase):
    def test_marker_wrapped_placeholder_becomes_one_valid_image_reference(self):
        processor = MarkdownProcessor(use_llm_correction=False)
        result = processor.process(
            "Before\n\n![]([IMAGE_PLACEHOLDER_1])\n\nAfter",
            images=[{"image_id": "paper_img_001"}],
        )
        self.assertIn("![id: paper_img_001](paper_img_001.png)", result)
        self.assertNotIn("![](![", result)

    def test_source_and_vietnamese_word_keep_same_embedded_image_count(self):
        with tempfile.TemporaryDirectory() as root:
            root_path = Path(root)
            image_path = root_path / "figure.png"
            image_path.write_bytes(ONE_PIXEL_PNG)
            images = [{"image_id": "paper_img_001", "output_path": str(image_path)}]
            exporter = WordExporter()

            source_docx = root_path / "source.docx"
            vi_docx = root_path / "vi.docx"
            exporter.markdown_to_word(
                "# Results\n\n![id: paper_img_001](paper_img_001.png)\n",
                str(source_docx),
                images=images,
            )
            exporter.markdown_to_word(
                "# Kết quả\n\n![id: paper_img_001](paper_img_001.png)\n",
                str(vi_docx),
                images=images,
            )

            self.assertEqual(1, len(Document(source_docx).inline_shapes))
            self.assertEqual(1, len(Document(vi_docx).inline_shapes))


if __name__ == "__main__":
    unittest.main()
