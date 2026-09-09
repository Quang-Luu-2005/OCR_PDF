import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace

from src.processing.scientific_translator import ScientificTranslator, TranslationError


def _response(payload, prompt_tokens=10, completion_tokens=5):
    return SimpleNamespace(
        choices=[SimpleNamespace(message=SimpleNamespace(content=json.dumps(payload)))],
        usage=SimpleNamespace(
            prompt_tokens=prompt_tokens,
            completion_tokens=completion_tokens,
            total_tokens=prompt_tokens + completion_tokens,
        ),
    )


class FakeCompletions:
    def __init__(self, malformed_chunks=0, remove_protected=False):
        self.calls = []
        self.malformed_chunks = malformed_chunks
        self.remove_protected = remove_protected

    def create(self, **kwargs):
        self.calls.append(kwargs)
        user_text = kwargs["messages"][-1]["content"]
        payload = json.loads(user_text.split("\n", 1)[1])
        if "document" in payload:
            return _response({
                "terms": [{
                    "source": "bounding box",
                    "vi": "hộp giới hạn",
                    "keep_english": True,
                }]
            })

        if self.malformed_chunks:
            self.malformed_chunks -= 1
            return _response({"items": []})

        items = []
        for item in payload["items"]:
            corrected = item["source"].replace("teh", "the")
            translated = corrected.replace("Results", "Kết quả")
            if self.remove_protected:
                corrected = corrected.replace("`KEEP_TOKEN_0000`", "")
                translated = translated.replace("`KEEP_TOKEN_0000`", "")
            items.append({
                "id": item["id"],
                "corrected_source": corrected,
                "vi": translated,
            })
        return _response({"items": items})


class FakeClient:
    def __init__(self, completions):
        self.chat = SimpleNamespace(completions=completions)


class ScientificTranslatorTests(unittest.TestCase):
    def make_translator(self, completions, cache_dir=None, max_retries=2):
        return ScientificTranslator(
            client=FakeClient(completions),
            cache_dir=cache_dir,
            chunk_chars=1000,
            max_retries=max_retries,
            sleep_fn=lambda _: None,
        )

    def test_preserves_markdown_scientific_content_and_image_positions(self):
        source = (
            "## Results\n\n"
            "The teh cohort included 10 patients (Nagy et al., 2022), with 25%. "
            "See https://example.org/a and doi:10.1000/xyz123. The score was $x=2$.\n\n"
            "![id: paper_img_001](paper_img_001.png)\n\n"
            "| Metric | Value |\n"
            "|---|---|\n"
            "| Accuracy | 95% |\n\n"
            "```python\nvalue = 10\n```\n"
            "\n$$Z=\\frac{X-\\min(X)}{\\max(X)-\\min(X)}$$ "
            "<!-- equation:id=eq_p003_001;number=1 -->\n"
        )
        completions = FakeCompletions()
        result = self.make_translator(completions).process(source, "Paper")

        self.assertIn("The the cohort", result.corrected_markdown)
        self.assertIn("## Kết quả", result.vi_markdown)
        for protected in (
            "https://example.org/a",
            "doi:10.1000/xyz123",
            "$x=2$",
            "![id: paper_img_001](paper_img_001.png)",
            "```python\nvalue = 10\n```",
            "$$Z=\\frac{X-\\min(X)}{\\max(X)-\\min(X)}$$ "
            "<!-- equation:id=eq_p003_001;number=1 -->",
        ):
            self.assertEqual(source.count(protected), result.vi_markdown.count(protected))
        self.assertEqual(source.count("|"), result.vi_markdown.count("|"))
        self.assertEqual(1, len(result.glossary))
        self.assertGreater(result.usage.total_tokens, 0)
        chunk_call = completions.calls[-1]
        self.assertEqual({"type": "json_object"}, chunk_call["response_format"])
        self.assertEqual("low", chunk_call["reasoning_effort"])
        chunk_payload = json.loads(chunk_call["messages"][-1]["content"].split("\n", 1)[1])
        cell_sources = [item["source"].strip() for item in chunk_payload["items"]]
        self.assertIn("Metric", cell_sources)
        self.assertIn("Value", cell_sources)

    def test_retries_empty_or_incomplete_json(self):
        completions = FakeCompletions(malformed_chunks=1)
        result = self.make_translator(completions).process("Results included 10 cases.")
        self.assertEqual(1, result.usage.retries)
        self.assertEqual(3, len(completions.calls))  # glossary + failed chunk + retry

    def test_retries_network_or_rate_limit_errors(self):
        completions = FakeCompletions()
        original_create = completions.create
        state = {"failed": False}

        def flaky_create(**kwargs):
            user_text = kwargs["messages"][-1]["content"]
            if "Return JSON for this batch" in user_text and not state["failed"]:
                state["failed"] = True
                completions.calls.append(kwargs)
                raise RuntimeError("429 rate limit")
            return original_create(**kwargs)

        completions.create = flaky_create
        result = self.make_translator(completions).process("Results included 10 cases.")
        self.assertEqual(1, result.usage.retries)

    def test_checkpoint_avoids_duplicate_api_calls(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            completions = FakeCompletions()
            translator = self.make_translator(completions, Path(temp_dir))
            first = translator.process("Results included 10 cases.")
            first_call_count = len(completions.calls)
            second = translator.process("Results included 10 cases.")

            self.assertEqual(2, first_call_count)
            self.assertEqual(first_call_count, len(completions.calls))
            self.assertEqual(2, second.usage.cache_hits)
            self.assertEqual(first.vi_markdown, second.vi_markdown)

    def test_rejects_changed_image_placeholder(self):
        source = "Figure ![id: paper_img_001](paper_img_001.png) shows 10 cases."
        completions = FakeCompletions(remove_protected=True)
        translator = self.make_translator(completions, max_retries=1)
        with self.assertRaises(TranslationError):
            translator.process(source)
        self.assertEqual(3, len(completions.calls))  # glossary + two chunk attempts


if __name__ == "__main__":
    unittest.main()
