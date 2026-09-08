"""Scientific English-to-Vietnamese post-processing for OCR Markdown."""

from __future__ import annotations

from collections import Counter
from dataclasses import asdict, dataclass, field
from hashlib import sha256
import json
import logging
import os
from pathlib import Path
import re
import time
from typing import Any, Callable, Dict, Iterable, List, Optional, Tuple


logger = logging.getLogger(__name__)


class TranslationError(RuntimeError):
    """Raised when a complete, structurally valid translation cannot be produced."""


class TranslationValidationError(TranslationError):
    """Raised when a model response changes protected document content."""


@dataclass(frozen=True)
class GlossaryTerm:
    source: str
    vi: str
    keep_english: bool = False


@dataclass
class TranslationUsage:
    input_tokens: int = 0
    output_tokens: int = 0
    total_tokens: int = 0
    chunks: int = 0
    cache_hits: int = 0
    retries: int = 0


@dataclass(frozen=True)
class TranslationResult:
    corrected_markdown: str
    vi_markdown: str
    glossary: List[GlossaryTerm]
    usage: TranslationUsage


@dataclass(frozen=True)
class _MarkdownBlock:
    block_id: str
    text: str
    translatable: bool


@dataclass(frozen=True)
class _ProtectedText:
    text: str
    replacements: Dict[str, str] = field(default_factory=dict)


class ScientificTranslator:
    """Correct OCR English and translate it to formal scientific Vietnamese."""

    PROMPT_VERSION = "scientific-vi-v2-gemini"
    DEFAULT_BASE_URL = "https://generativelanguage.googleapis.com/v1beta/openai/"
    DEFAULT_MODEL = "gemini-3.8-flash"

    _IMAGE_RE = re.compile(
        r"!\[[^\]]*\]\([^\)]+\)|\[IMAGE_PLACEHOLDER_\d+\]",
        re.IGNORECASE,
    )
    _NUMBER_RE = re.compile(
        r"(?<![\w])[-+]?\d+(?:[.,]\d+)*(?:\s*[-–]\s*\d+(?:[.,]\d+)*)?%?"
    )
    _PROTECTED_RE = re.compile(
        r"!\[[^\]]*\]\([^\)]+\)"
        r"|\[IMAGE_PLACEHOLDER_\d+\]"
        r"|\$\$[\s\S]*?\$\$"
        r"|\$[^$\n]+\$"
        r"|`[^`\n]+`"
        r"|https?://[^\s<>)\]]+"
        r"|(?:doi:\s*)?10\.\d{4,9}/[-._;()/:%A-Z0-9]+"
        r"|\[(?:\d+[a-z]?\s*[,;–-]?\s*)+\]"
        r"|\([A-Z][A-Za-z'’-]+(?:\s+et\s+al\.)?,?\s+\d{4}[a-z]?\)"
        r"|(?<![\w])[-+]?\d+(?:[.,]\d+)*\s*(?:%|mm|cm|km|m|kg|mg|g|mL|L|s|ms|Hz|MHz|GHz|px|dpi)\b"
        r"|(?-i:\b[A-Z][A-Z0-9-]{1,}\b)"
        r"|</?[^>]+>",
        re.IGNORECASE,
    )
    _TABLE_SEPARATOR_RE = re.compile(
        r"^\s*\|?\s*:?-{3,}:?\s*(?:\|\s*:?-{3,}:?\s*)+\|?\s*$"
    )

    def __init__(
        self,
        *,
        api_key: Optional[str] = None,
        base_url: Optional[str] = None,
        model: Optional[str] = None,
        timeout_seconds: float = 120.0,
        chunk_chars: int = 8000,
        max_retries: int = 3,
        max_tokens: int = 12000,
        cache_dir: Optional[str | Path] = None,
        client: Any = None,
        sleep_fn: Callable[[float], None] = time.sleep,
    ) -> None:
        self.base_url = (base_url or self.DEFAULT_BASE_URL).rstrip("/")
        self.model = model or self.DEFAULT_MODEL
        self.timeout_seconds = float(timeout_seconds)
        self.chunk_chars = max(1000, int(chunk_chars))
        self.max_retries = max(0, int(max_retries))
        self.max_tokens = max(1000, int(max_tokens))
        self.cache_dir = Path(cache_dir) if cache_dir else None
        self._sleep = sleep_fn

        if self.cache_dir:
            self.cache_dir.mkdir(parents=True, exist_ok=True)

        if client is not None:
            self.client = client
            return

        resolved_key = (
            api_key
            or os.getenv("TRANSLATION_API_KEY")
            or os.getenv("GEMINI_API_KEY")
        )
        if not resolved_key:
            raise TranslationError(
                "Vietnamese translation is enabled but no API key was found. "
                "Set GEMINI_API_KEY (or TRANSLATION_API_KEY) in ocr/.env, "
                "or run with --no-translate-vi."
            )

        try:
            from openai import OpenAI
        except ImportError as exc:
            raise TranslationError(
                "The 'openai' package is required. Run: "
                "python -m pip install -r ocr/requirements.txt"
            ) from exc

        self.client = OpenAI(
            api_key=resolved_key,
            base_url=self.base_url,
            timeout=self.timeout_seconds,
        )

    @classmethod
    def from_env(
        cls,
        *,
        model: Optional[str] = None,
        base_url: Optional[str] = None,
        cache_dir: Optional[str | Path] = None,
        client: Any = None,
    ) -> "ScientificTranslator":
        """Build a translator from ``ocr/.env`` and process environment values."""
        env_path = Path(__file__).resolve().parents[2] / ".env"
        try:
            from dotenv import load_dotenv

            load_dotenv(env_path)
        except ImportError:
            pass

        return cls(
            api_key=os.getenv("TRANSLATION_API_KEY") or os.getenv("GEMINI_API_KEY"),
            base_url=base_url or os.getenv("TRANSLATION_API_BASE") or cls.DEFAULT_BASE_URL,
            model=model or os.getenv("TRANSLATION_MODEL") or cls.DEFAULT_MODEL,
            timeout_seconds=float(os.getenv("TRANSLATION_TIMEOUT_SECONDS", "120")),
            chunk_chars=int(os.getenv("TRANSLATION_CHUNK_CHARS", "8000")),
            max_retries=int(os.getenv("TRANSLATION_MAX_RETRIES", "3")),
            max_tokens=int(os.getenv("TRANSLATION_MAX_TOKENS", "12000")),
            cache_dir=cache_dir,
            client=client,
        )

    def process(self, markdown: str, document_title: str = "") -> TranslationResult:
        if not markdown.strip():
            return TranslationResult(markdown, markdown, [], TranslationUsage())

        usage = TranslationUsage()
        glossary = self._build_glossary(markdown, document_title, usage)
        blocks = self._parse_blocks(markdown)
        batches = list(self._chunk_blocks(block for block in blocks if block.translatable))

        corrected_by_id: Dict[str, str] = {}
        vi_by_id: Dict[str, str] = {}
        glossary_payload = [asdict(term) for term in glossary]
        seen_first_use_terms = set()

        for batch_index, batch in enumerate(batches, start=1):
            protected_items = {
                block.block_id: self._protect(block.text) for block in batch
            }
            request_items = [
                {"id": block.block_id, "source": protected_items[block.block_id].text}
                for block in batch
            ]
            batch_source = "\n".join(block.text for block in batch).casefold()
            first_use_terms = []
            for term in glossary:
                term_key = term.source.casefold()
                if (
                    term.keep_english
                    and term_key not in seen_first_use_terms
                    and term_key in batch_source
                ):
                    first_use_terms.append(term.source)
                    seen_first_use_terms.add(term_key)
            payload = {
                "document_title": document_title,
                "glossary": glossary_payload,
                "include_english_on_first_use": first_use_terms,
                "items": request_items,
            }
            messages = [
                {"role": "system", "content": self._translation_system_prompt()},
                {
                    "role": "user",
                    "content": "Return JSON for this batch:\n" + json.dumps(payload, ensure_ascii=False),
                },
            ]

            def validate(data: Dict[str, Any]) -> None:
                self._validate_translation_response(data, batch, protected_items)

            data, request_usage, cache_hit, retries = self._request_json(
                messages,
                purpose=f"chunk-{batch_index:04d}",
                validator=validate,
                max_tokens=self.max_tokens,
            )
            self._accumulate_usage(usage, request_usage, cache_hit, retries)
            usage.chunks += 1

            for item in data["items"]:
                block_id = item["id"]
                protected = protected_items[block_id]
                corrected_by_id[block_id] = self._restore(
                    item["corrected_source"], protected.replacements
                )
                vi_by_id[block_id] = self._restore(item["vi"], protected.replacements)

        corrected = "".join(
            corrected_by_id.get(block.block_id, block.text) for block in blocks
        )
        translated = "".join(vi_by_id.get(block.block_id, block.text) for block in blocks)

        expected_images = self._IMAGE_RE.findall(markdown)
        if self._IMAGE_RE.findall(corrected) != expected_images:
            raise TranslationValidationError("Corrected Markdown changed image placeholders.")
        if self._IMAGE_RE.findall(translated) != expected_images:
            raise TranslationValidationError("Vietnamese Markdown changed image placeholders.")
        if corrected.count("|") != markdown.count("|"):
            raise TranslationValidationError("Corrected Markdown changed table structure.")
        if translated.count("|") != markdown.count("|"):
            raise TranslationValidationError("Vietnamese Markdown changed table structure.")

        return TranslationResult(corrected, translated, glossary, usage)

    def _build_glossary(
        self,
        markdown: str,
        document_title: str,
        usage: TranslationUsage,
    ) -> List[GlossaryTerm]:
        translatable_text = "\n\n".join(
            self._PROTECTED_RE.sub(" ", block.text)
            for block in self._parse_blocks(markdown)
            if block.translatable
        )
        payload = {
            "document_title": document_title,
            "source_language": "English",
            "target_language": "Vietnamese",
            "document": translatable_text,
        }
        messages = [
            {
                "role": "system",
                "content": (
                    "You are a scientific terminology editor. Identify only important, repeated domain terms "
                    "needed for a consistent English-to-Vietnamese translation. Return valid JSON exactly as "
                    '{"terms":[{"source":"...","vi":"...","keep_english":false}]}. '
                    "Use established Vietnamese scientific terminology. Preserve dataset, software, model, "
                    "organization and proper names. Use keep_english=true for specialized terms that should "
                    "also appear in English on first use. Do not include ordinary words."
                ),
            },
            {
                "role": "user",
                "content": "Build the JSON glossary for:\n" + json.dumps(payload, ensure_ascii=False),
            },
        ]

        def validate(data: Dict[str, Any]) -> None:
            terms = data.get("terms")
            if not isinstance(terms, list):
                raise TranslationValidationError("Glossary response is missing a terms array.")
            for term in terms:
                if not isinstance(term, dict):
                    raise TranslationValidationError("Glossary term must be an object.")
                if not isinstance(term.get("source"), str) or not term["source"].strip():
                    raise TranslationValidationError("Glossary term has no source text.")
                if not isinstance(term.get("vi"), str) or not term["vi"].strip():
                    raise TranslationValidationError("Glossary term has no Vietnamese text.")

        data, request_usage, cache_hit, retries = self._request_json(
            messages,
            purpose="glossary",
            validator=validate,
            max_tokens=min(4000, self.max_tokens),
        )
        self._accumulate_usage(usage, request_usage, cache_hit, retries)

        seen = set()
        glossary: List[GlossaryTerm] = []
        for item in data["terms"]:
            source = item["source"].strip()
            key = source.casefold()
            if key in seen:
                continue
            seen.add(key)
            glossary.append(
                GlossaryTerm(
                    source=source,
                    vi=item["vi"].strip(),
                    keep_english=bool(item.get("keep_english", False)),
                )
            )
        return glossary

    def _request_json(
        self,
        messages: List[Dict[str, str]],
        *,
        purpose: str,
        validator: Callable[[Dict[str, Any]], None],
        max_tokens: int,
    ) -> Tuple[Dict[str, Any], Dict[str, int], bool, int]:
        cache_key = sha256(
            json.dumps(
                {
                    "prompt_version": self.PROMPT_VERSION,
                    "base_url": self.base_url,
                    "model": self.model,
                    "messages": messages,
                },
                ensure_ascii=False,
                sort_keys=True,
            ).encode("utf-8")
        ).hexdigest()
        cache_path = self.cache_dir / f"{purpose}-{cache_key}.json" if self.cache_dir else None

        if cache_path and cache_path.exists():
            try:
                cached = json.loads(cache_path.read_text(encoding="utf-8"))
                data = cached["data"]
                validator(data)
                empty_usage = {"input_tokens": 0, "output_tokens": 0, "total_tokens": 0}
                return data, empty_usage, True, 0
            except Exception:
                logger.warning("Ignoring invalid translation cache: %s", cache_path)

        last_error: Optional[BaseException] = None
        for attempt in range(self.max_retries + 1):
            try:
                kwargs: Dict[str, Any] = {
                    "model": self.model,
                    "messages": messages,
                    "response_format": {"type": "json_object"},
                    "max_tokens": max_tokens,
                    "stream": False,
                    "temperature": 0.1,
                }
                if "generativelanguage.googleapis.com" in self.base_url.lower():
                    kwargs["reasoning_effort"] = "low"

                response = self.client.chat.completions.create(**kwargs)
                content = response.choices[0].message.content
                if not isinstance(content, str) or not content.strip():
                    raise TranslationValidationError("The translation API returned empty content.")
                data = json.loads(content)
                if not isinstance(data, dict):
                    raise TranslationValidationError("The translation response must be a JSON object.")
                validator(data)
                request_usage = self._read_usage(response)

                if cache_path:
                    self._atomic_write_json(cache_path, {"data": data, "usage": request_usage})
                return data, request_usage, False, attempt
            except Exception as exc:
                last_error = exc
                if attempt >= self.max_retries:
                    break
                delay = min(8.0, 2.0 ** attempt)
                logger.warning(
                    "Translation request %s failed (attempt %d/%d): %s. Retrying in %.1fs.",
                    purpose,
                    attempt + 1,
                    self.max_retries + 1,
                    exc,
                    delay,
                )
                self._sleep(delay)

        raise TranslationError(
            f"Translation request '{purpose}' failed after {self.max_retries + 1} attempts: {last_error}"
        ) from last_error

    @staticmethod
    def _read_usage(response: Any) -> Dict[str, int]:
        usage = getattr(response, "usage", None)
        input_tokens = int(
            getattr(usage, "prompt_tokens", getattr(usage, "input_tokens", 0)) or 0
        )
        output_tokens = int(
            getattr(usage, "completion_tokens", getattr(usage, "output_tokens", 0)) or 0
        )
        total_tokens = int(getattr(usage, "total_tokens", input_tokens + output_tokens) or 0)
        return {
            "input_tokens": input_tokens,
            "output_tokens": output_tokens,
            "total_tokens": total_tokens,
        }

    @staticmethod
    def _accumulate_usage(
        usage: TranslationUsage,
        request_usage: Dict[str, int],
        cache_hit: bool,
        retries: int,
    ) -> None:
        usage.input_tokens += request_usage.get("input_tokens", 0)
        usage.output_tokens += request_usage.get("output_tokens", 0)
        usage.total_tokens += request_usage.get("total_tokens", 0)
        usage.cache_hits += int(cache_hit)
        usage.retries += retries

    def _validate_translation_response(
        self,
        data: Dict[str, Any],
        blocks: List[_MarkdownBlock],
        protected_items: Dict[str, _ProtectedText],
    ) -> None:
        items = data.get("items")
        if not isinstance(items, list):
            raise TranslationValidationError("Translation response is missing an items array.")

        expected_ids = [block.block_id for block in blocks]
        actual_ids = [item.get("id") for item in items if isinstance(item, dict)]
        if actual_ids != expected_ids:
            raise TranslationValidationError(
                f"Translation IDs changed or were reordered: expected {expected_ids}, got {actual_ids}."
            )

        source_by_id = {block.block_id: block.text for block in blocks}
        for item in items:
            block_id = item["id"]
            corrected = item.get("corrected_source")
            vi = item.get("vi")
            if not isinstance(corrected, str) or not corrected.strip():
                raise TranslationValidationError(f"Block {block_id} has no corrected source text.")
            if not isinstance(vi, str) or not vi.strip():
                raise TranslationValidationError(f"Block {block_id} has no Vietnamese translation.")

            protected = protected_items[block_id]
            for placeholder in protected.replacements:
                if corrected.count(placeholder) != 1 or vi.count(placeholder) != 1:
                    raise TranslationValidationError(
                        f"Block {block_id} changed protected token {placeholder}."
                    )

            source = source_by_id[block_id]
            restored_corrected = self._restore(corrected, protected.replacements)
            restored_vi = self._restore(vi, protected.replacements)
            source_numbers = Counter(self._NUMBER_RE.findall(source))
            if Counter(self._NUMBER_RE.findall(restored_corrected)) != source_numbers:
                raise TranslationValidationError(f"Block {block_id} changed source numerical values.")
            if Counter(self._NUMBER_RE.findall(restored_vi)) != source_numbers:
                raise TranslationValidationError(f"Block {block_id} changed translated numerical values.")

            source_heading = re.match(r"^\s*(#{1,6})\s+", source)
            if source_heading:
                corrected_heading = re.match(r"^\s*(#{1,6})\s+", restored_corrected)
                vi_heading = re.match(r"^\s*(#{1,6})\s+", restored_vi)
                expected = source_heading.group(1)
                if not corrected_heading or corrected_heading.group(1) != expected:
                    raise TranslationValidationError(f"Block {block_id} changed heading level.")
                if not vi_heading or vi_heading.group(1) != expected:
                    raise TranslationValidationError(f"Block {block_id} changed translated heading level.")

            if "|" in source:
                if restored_corrected.count("|") != source.count("|"):
                    raise TranslationValidationError(f"Block {block_id} changed table columns.")
                if restored_vi.count("|") != source.count("|"):
                    raise TranslationValidationError(f"Block {block_id} changed translated table columns.")

    @staticmethod
    def _translation_system_prompt() -> str:
        return (
            "You are a scientific English editor and English-to-Vietnamese translator. Return valid JSON only "
            'with exactly this shape: {"items":[{"id":"...","corrected_source":"...","vi":"..."}]}. '
            "Keep every input item exactly once and in the same order. First, minimally correct OCR, spelling, "
            "broken line-wrap hyphenation and obvious grammar errors in corrected_source without rewriting or "
            "changing meaning. Then translate that corrected text into formal, precise Vietnamese suitable for a "
            "peer-reviewed scientific paper. Never summarize, omit, expand, explain or invent. Follow the supplied "
            "glossary consistently. Preserve Markdown syntax, paragraph role, table pipes, datasets, software, model "
            "names, proper nouns, acronyms, citations, values and units. Tokens like __KEEP_0001__ are immutable and "
            "must occur exactly once in both outputs. For specialized or ambiguous terms marked keep_english, include "
            "the English term in parentheses only when that source term appears in include_english_on_first_use; "
            "otherwise use the glossary's Vietnamese term without adding English again."
        )

    @classmethod
    def _protect(cls, text: str) -> _ProtectedText:
        replacements: Dict[str, str] = {}
        next_index = 0

        def repl(match: re.Match[str]) -> str:
            nonlocal next_index
            placeholder = f"__KEEP_{next_index:04d}__"
            while placeholder in text or placeholder in replacements:
                next_index += 1
                placeholder = f"__KEEP_{next_index:04d}__"
            replacements[placeholder] = match.group(0)
            next_index += 1
            return placeholder

        return _ProtectedText(cls._PROTECTED_RE.sub(repl, text), replacements)

    @staticmethod
    def _restore(text: str, replacements: Dict[str, str]) -> str:
        restored = text
        for placeholder, original in replacements.items():
            restored = restored.replace(placeholder, original)
        return restored

    @classmethod
    def _parse_blocks(cls, markdown: str) -> List[_MarkdownBlock]:
        blocks: List[_MarkdownBlock] = []
        paragraph: List[str] = []
        code: List[str] = []
        in_code = False
        fence = ""

        def add(text: str, translatable: bool) -> None:
            if text:
                blocks.append(_MarkdownBlock(f"b{len(blocks):06d}", text, translatable))

        def flush_paragraph() -> None:
            if not paragraph:
                return
            joined = "".join(paragraph)
            paragraph.clear()
            match = re.search(r"(?:\r?\n)+$", joined)
            trailing = match.group(0) if match else ""
            body = joined[:-len(trailing)] if trailing else joined
            add(body, not cls._is_nontranslatable(body))
            add(trailing, False)

        for raw_line in markdown.splitlines(keepends=True):
            stripped = raw_line.strip()
            if in_code:
                code.append(raw_line)
                if stripped.startswith(fence):
                    add("".join(code), False)
                    code.clear()
                    in_code = False
                continue

            if stripped.startswith("```") or stripped.startswith("~~~"):
                flush_paragraph()
                fence = stripped[:3]
                code = [raw_line]
                in_code = True
                continue

            if not stripped:
                flush_paragraph()
                add(raw_line, False)
                continue

            if stripped.startswith("|"):
                flush_paragraph()
                line = raw_line.rstrip("\r\n")
                ending = raw_line[len(line):]
                if cls._TABLE_SEPARATOR_RE.fullmatch(line):
                    add(line, False)
                else:
                    # Treat each table cell as its own item while retaining every
                    # delimiter verbatim, so cells cannot be reordered or merged.
                    for part in re.split(r"(\|)", line):
                        add(part, part != "|" and not cls._is_nontranslatable(part))
                add(ending, False)
                continue

            paragraph.append(raw_line)

        flush_paragraph()
        if code:
            add("".join(code), False)
        return blocks

    @classmethod
    def _is_nontranslatable(cls, text: str) -> bool:
        stripped = text.strip()
        if not stripped or stripped == "</break>":
            return True
        if cls._IMAGE_RE.fullmatch(stripped):
            return True
        if cls._TABLE_SEPARATOR_RE.fullmatch(stripped):
            return True
        return not bool(re.search(r"[A-Za-z]", stripped))

    def _chunk_blocks(self, blocks: Iterable[_MarkdownBlock]) -> Iterable[List[_MarkdownBlock]]:
        current: List[_MarkdownBlock] = []
        current_size = 0
        for block in blocks:
            block_size = len(block.text)
            if current and current_size + block_size > self.chunk_chars:
                yield current
                current = []
                current_size = 0
            current.append(block)
            current_size += block_size
        if current:
            yield current

    @staticmethod
    def _atomic_write_json(path: Path, payload: Dict[str, Any]) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        temp_path = path.with_suffix(path.suffix + ".tmp")
        temp_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        temp_path.replace(path)
