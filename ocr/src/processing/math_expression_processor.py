"""Detection and Gemini Vision recognition for displayed PDF equations."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
import base64
from hashlib import sha256
import json
import logging
from pathlib import Path
import re
import time
from typing import Any, Dict, List, Optional, Sequence, Tuple


logger = logging.getLogger(__name__)


@dataclass
class EquationUsage:
    api_calls: int = 0
    cache_hits: int = 0
    retries: int = 0
    input_tokens: int = 0
    output_tokens: int = 0
    total_tokens: int = 0


@dataclass
class EquationRecord:
    equation_id: str
    page_num: int
    bbox: Tuple[float, float, float, float]
    equation_number: Optional[str]
    latex: Optional[str]
    plain_text: str
    symbols: List[str]
    transcript: str
    image_path: str
    confidence: float
    validated: bool
    output_mode: str
    consumed_block_indexes: List[int] = field(default_factory=list)
    validation_error: Optional[str] = None

    @property
    def markdown(self) -> str:
        number = self.equation_number or ""
        if self.validated and self.latex:
            latex = " ".join(self.latex.splitlines())
            return (
                f'$${latex}$$ '
                f'<!-- equation:id={self.equation_id};number={number} -->'
            )
        filename = Path(self.image_path).name
        return (
            f'![equation: {self.equation_id}]({filename}) '
            f'<!-- equation:number={number} -->'
        )

    def to_dict(self) -> Dict[str, Any]:
        data = asdict(self)
        data["bbox"] = list(self.bbox)
        data["markdown"] = self.markdown
        return data


class MathExpressionProcessor:
    """Find numbered display equations and recognize them with Gemini Vision."""

    PROMPT_VERSION = "equation-vision-v1"
    _NUMBER_RE = re.compile(r"\((\d{1,3})\)")
    _MATH_FONT_RE = re.compile(
        r"(?:math|symbol|mtmi|mtsyn|cmex|cmsy|cmmi|cambria)", re.IGNORECASE
    )

    def __init__(
        self,
        *,
        client: Any = None,
        model: str = "gemini-3.5-flash",
        output_dir: str | Path = "./output/extracted_equations",
        cache_dir: str | Path | None = None,
        enabled: bool = True,
        vision_enabled: bool = True,
        render_dpi: int = 300,
        min_confidence: float = 0.90,
        max_retries: int = 3,
        timeout_seconds: float = 120.0,
        sleep_fn=time.sleep,
    ) -> None:
        self.client = client
        self.model = model
        self.output_dir = Path(output_dir)
        self.cache_dir = Path(cache_dir) if cache_dir else None
        self.enabled = bool(enabled)
        self.vision_enabled = bool(vision_enabled)
        self.render_dpi = max(150, int(render_dpi))
        self.min_confidence = min(1.0, max(0.0, float(min_confidence)))
        self.max_retries = max(0, int(max_retries))
        self.timeout_seconds = float(timeout_seconds)
        self._sleep = sleep_fn
        self.usage = EquationUsage()
        self.records: List[EquationRecord] = []
        self.output_dir.mkdir(parents=True, exist_ok=True)
        if self.cache_dir:
            self.cache_dir.mkdir(parents=True, exist_ok=True)

    def reset(self) -> None:
        self.usage = EquationUsage()
        self.records = []

    def extract_page(
        self,
        page: Any,
        page_number: int,
        blocks: Sequence[tuple],
        document_title: str,
    ) -> List[EquationRecord]:
        if not self.enabled:
            return []

        anchors = self._find_numbered_anchors(page, blocks)
        page_records: List[EquationRecord] = []
        used_indexes: set[int] = set()
        for sequence, (anchor_index, equation_number) in enumerate(anchors, start=1):
            indexes = self._collect_equation_blocks(
                page, blocks, anchor_index, used_indexes
            )
            if not indexes:
                continue
            used_indexes.update(indexes)
            bbox = self._union_bbox([blocks[index][:4] for index in indexes])
            clip = self._padded_clip(page, bbox)
            equation_id = f"eq_p{page_number:03d}_{sequence:03d}"
            image_path = self.output_dir / f"{equation_id}.png"
            scale = self.render_dpi / 72.0
            page.get_pixmap(
                matrix=__import__("fitz").Matrix(scale, scale),
                clip=clip,
                alpha=False,
            ).save(str(image_path))

            transcript = " ".join(
                blocks[index][4].strip()
                for index in sorted(indexes, key=lambda i: (blocks[i][1], blocks[i][0]))
                if len(blocks[index]) >= 5 and blocks[index][4].strip()
            )
            context = self._nearby_context(blocks, indexes)
            recognized: Dict[str, Any] = {}
            error: Optional[str] = None
            try:
                recognized = self._recognize(
                    image_path=image_path,
                    document_title=document_title,
                    page_number=page_number,
                    equation_number=equation_number,
                    transcript=transcript,
                    context=context,
                )
                valid, error = self._validate_recognition(
                    recognized, transcript, equation_number
                )
            except Exception as exc:
                valid = False
                error = str(exc)
                logger.warning("Equation %s uses image fallback: %s", equation_id, exc)

            latex = self._clean_latex(recognized.get("latex"))
            confidence = self._coerce_confidence(recognized.get("confidence"))
            record = EquationRecord(
                equation_id=equation_id,
                page_num=page_number,
                bbox=tuple(float(value) for value in clip),
                equation_number=equation_number,
                latex=latex,
                plain_text=str(recognized.get("plain_text", "") or ""),
                symbols=[str(value) for value in recognized.get("symbols", [])]
                if isinstance(recognized.get("symbols", []), list)
                else [],
                transcript=transcript,
                image_path=str(image_path),
                confidence=confidence,
                validated=valid,
                output_mode="omml" if valid else "image",
                consumed_block_indexes=sorted(indexes),
                validation_error=error,
            )
            page_records.append(record)
            self.records.append(record)
        return page_records

    def save_index(self, path: str | Path, source_pdf: str | Path) -> Path:
        path = Path(path)
        payload = {
            "source_pdf": str(source_pdf),
            "total_equations": len(self.records),
            "omml_candidates": sum(record.validated for record in self.records),
            "image_fallbacks": sum(not record.validated for record in self.records),
            "usage": asdict(self.usage),
            "equations": [record.to_dict() for record in self.records],
        }
        path.parent.mkdir(parents=True, exist_ok=True)
        temp_path = path.with_suffix(path.suffix + ".tmp")
        temp_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        temp_path.replace(path)
        return path

    def _find_numbered_anchors(
        self, page: Any, blocks: Sequence[tuple]
    ) -> List[Tuple[int, str]]:
        anchors: List[Tuple[int, str]] = []
        for index, block in enumerate(blocks):
            if len(block) < 5 or block[2] < page.rect.width * 0.88:
                continue
            matches = list(self._NUMBER_RE.finditer(block[4]))
            if not matches:
                continue
            # Equation numbers in journal layouts sit at the right margin and
            # share a short horizontal band with mathematical glyphs.
            if block[3] - block[1] > 35:
                continue
            nearby_math = any(
                self._looks_mathematical(candidate)
                and self._vertical_gap(block, candidate) <= 18
                for candidate in blocks
            )
            if nearby_math:
                anchors.append((index, matches[-1].group(1)))
        return anchors

    def _collect_equation_blocks(
        self,
        page: Any,
        blocks: Sequence[tuple],
        anchor_index: int,
        used_indexes: set[int],
    ) -> List[int]:
        anchor = blocks[anchor_index]
        min_x = page.rect.width * 0.30
        indexes = []
        for index, block in enumerate(blocks):
            if index in used_indexes or len(block) < 5 or not block[4].strip():
                continue
            if block[0] < min_x or self._vertical_gap(anchor, block) > 18:
                continue
            short_math_fragment = (
                len(block[4].strip()) <= 3
                and not re.search(r"[.!?;:]", block[4].strip())
            )
            if (
                index == anchor_index
                or self._looks_mathematical(block)
                or short_math_fragment
            ):
                indexes.append(index)
        return indexes

    def _looks_mathematical(self, block: tuple) -> bool:
        text = block[4] if len(block) >= 5 else ""
        if any(ord(char) < 32 and char not in "\r\n\t" for char in text):
            return True
        if re.search(r"[=\u2212+\u00d7\u00f7\u2211\u222b]", text):
            return True
        return bool(self._NUMBER_RE.search(text) and len(text.strip()) < 80)

    @staticmethod
    def _vertical_gap(first: tuple, second: tuple) -> float:
        if first[3] < second[1]:
            return float(second[1] - first[3])
        if second[3] < first[1]:
            return float(first[1] - second[3])
        return 0.0

    @staticmethod
    def _union_bbox(boxes: Sequence[Sequence[float]]) -> Tuple[float, float, float, float]:
        return (
            min(box[0] for box in boxes),
            min(box[1] for box in boxes),
            max(box[2] for box in boxes),
            max(box[3] for box in boxes),
        )

    @staticmethod
    def _padded_clip(page: Any, bbox: Sequence[float]) -> Any:
        fitz = __import__("fitz")
        return fitz.Rect(
            max(page.rect.x0, bbox[0] - 8),
            max(page.rect.y0, bbox[1] - 6),
            min(page.rect.x1, bbox[2] + 8),
            min(page.rect.y1, bbox[3] + 6),
        )

    @staticmethod
    def _nearby_context(blocks: Sequence[tuple], indexes: Sequence[int]) -> str:
        first, last = min(indexes), max(indexes)
        context_parts = []
        for index in (first - 1, last + 1):
            if 0 <= index < len(blocks) and len(blocks[index]) >= 5:
                context_parts.append(blocks[index][4].strip()[-500:])
        return "\n".join(part for part in context_parts if part)

    def _recognize(
        self,
        *,
        image_path: Path,
        document_title: str,
        page_number: int,
        equation_number: Optional[str],
        transcript: str,
        context: str,
    ) -> Dict[str, Any]:
        if not self.vision_enabled:
            raise RuntimeError("Gemini Vision equation recognition is disabled")
        if self.client is None:
            raise RuntimeError("Gemini client is unavailable")

        image_bytes = image_path.read_bytes()
        request_key = sha256(
            self.PROMPT_VERSION.encode("utf-8")
            + self.model.encode("utf-8")
            + image_bytes
            + transcript.encode("utf-8", errors="replace")
        ).hexdigest()
        cache_path = self.cache_dir / f"equation-{request_key}.json" if self.cache_dir else None
        if cache_path and cache_path.exists():
            data = json.loads(cache_path.read_text(encoding="utf-8"))
            if isinstance(data, dict) and data.get("latex"):
                self.usage.cache_hits += 1
                return data

        encoded = base64.b64encode(image_bytes).decode("ascii")
        payload = {
            "document_title": document_title,
            "page": page_number,
            "expected_equation_number": equation_number,
            "pdf_glyph_transcript": transcript,
            "nearby_text": context,
        }
        messages = [
            {
                "role": "system",
                "content": (
                    "You transcribe a single cropped scientific equation. Return JSON only with keys "
                    "latex, equation_number, plain_text, symbols, confidence. latex must contain only the "
                    "equation body: no dollar signs, markdown, equation number, prose label, or explanation. "
                    "Use standard LaTeX and preserve every variable, index, operator, delimiter and numerical "
                    "value visible in the image. The PDF transcript may contain broken control characters; "
                    "use it only as supporting evidence. confidence must be between 0 and 1."
                ),
            },
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": json.dumps(payload, ensure_ascii=False)},
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:image/png;base64,{encoded}"},
                    },
                ],
            },
        ]

        last_error: Optional[BaseException] = None
        for attempt in range(self.max_retries + 1):
            try:
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=messages,
                    response_format={"type": "json_object"},
                    max_tokens=1600,
                    stream=False,
                    temperature=0.0,
                    reasoning_effort="low",
                )
                self.usage.api_calls += 1
                content = response.choices[0].message.content
                data = json.loads(content)
                if not isinstance(data, dict) or not data.get("latex"):
                    raise ValueError("Vision response does not contain LaTeX")
                self._add_response_usage(response)
                if cache_path:
                    temp_path = cache_path.with_suffix(".tmp")
                    temp_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
                    temp_path.replace(cache_path)
                return data
            except Exception as exc:
                last_error = exc
                if attempt >= self.max_retries:
                    break
                self.usage.retries += 1
                self._sleep(min(8.0, 2.0 ** attempt))
        raise RuntimeError(
            f"Gemini Vision failed after {self.max_retries + 1} attempts: {last_error}"
        ) from last_error

    def _validate_recognition(
        self,
        data: Dict[str, Any],
        transcript: str,
        expected_number: Optional[str],
    ) -> Tuple[bool, Optional[str]]:
        latex = self._clean_latex(data.get("latex"))
        if not latex:
            return False, "empty LaTeX"
        if "$" in latex or "```" in latex:
            return False, "LaTeX contains Markdown delimiters"
        if not self._balanced_latex(latex):
            return False, "unbalanced LaTeX delimiters"
        confidence = self._coerce_confidence(data.get("confidence"))
        if confidence < self.min_confidence:
            return False, f"confidence {confidence:.2f} is below {self.min_confidence:.2f}"
        actual_number = str(data.get("equation_number", "") or "").strip("() ")
        if expected_number and actual_number and actual_number != expected_number:
            return False, "equation number does not match the PDF"

        source_numbers = re.findall(r"(?<![A-Za-z])\d+(?:\.\d+)?", transcript)
        if expected_number in source_numbers:
            source_numbers.remove(expected_number)
        latex_numbers = re.findall(r"(?<![A-Za-z])\d+(?:\.\d+)?", latex)
        missing = [number for number in source_numbers if number not in latex_numbers]
        if missing:
            return False, f"LaTeX is missing numerical tokens: {missing}"
        return True, None

    @staticmethod
    def _clean_latex(value: Any) -> Optional[str]:
        if not isinstance(value, str):
            return None
        latex = value.strip()
        latex = re.sub(r"^\$\$?|\$\$?$", "", latex).strip()
        latex = re.sub(r"\\tag\s*\{[^}]*\}\s*$", "", latex).strip()
        return latex or None

    @staticmethod
    def _balanced_latex(latex: str) -> bool:
        pairs = {"{": "}", "[": "]", "(": ")"}
        stack: List[str] = []
        escaped = False
        for char in latex:
            if escaped:
                escaped = False
                continue
            if char == "\\":
                escaped = True
                continue
            if char in pairs:
                stack.append(pairs[char])
            elif char in pairs.values():
                if not stack or stack.pop() != char:
                    return False
        return not stack

    @staticmethod
    def _coerce_confidence(value: Any) -> float:
        try:
            return min(1.0, max(0.0, float(value)))
        except (TypeError, ValueError):
            return 0.0

    def _add_response_usage(self, response: Any) -> None:
        usage = getattr(response, "usage", None)
        input_tokens = int(getattr(usage, "prompt_tokens", 0) or 0)
        output_tokens = int(getattr(usage, "completion_tokens", 0) or 0)
        total_tokens = int(getattr(usage, "total_tokens", input_tokens + output_tokens) or 0)
        self.usage.input_tokens += input_tokens
        self.usage.output_tokens += output_tokens
        self.usage.total_tokens += total_tokens
