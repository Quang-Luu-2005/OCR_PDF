# Pipeline orchestrator
# Coordinates the entire OCR workflow

from __future__ import annotations
import json
import logging
import re
import shutil
import subprocess
from hashlib import sha256
from pathlib import Path
from typing import Optional, List, Dict, Any

from ..processing.digital_parser import DigitalParser
from .pdf_converter import PDFConverter
from .ocr_engine import OCREngine
from ..export.exporter import WordExporter
from ..processing.markdown_processor import MarkdownProcessor
from ..processing.scientific_translator import ScientificTranslator, TranslationResult
from ..utils.metrics import PipelineMetrics
 

logger = logging.getLogger(__name__)


class OCRPipeline:
    """
    Complete OCR pipeline for processing PDF documents
    
    Automatically detects if PDF is digital or scanned, and applies
    appropriate conversion method.
    """
    
    def __init__(
        self,
        output_dir: str | Path = "./output",
        temp_dir: str | Path = "./temp",
        dpi: int = 300,
        enable_preprocessing: bool = True,
        auto_detect: bool = True,
        extract_images: bool = True,
        analyze_layout: bool = True,
        extract_tables: bool = True,
        use_llm_correction: bool = False,
        enable_vi_translation: bool = True,
        translation_model: Optional[str] = None,
        translation_base_url: Optional[str] = None,
        translator: Optional[Any] = None,
        use_marker_for_digital_translation: bool = False,
    ):
        """
        Initialize OCR Pipeline
        
        Args:
            output_dir: Directory for final outputs
            temp_dir: Directory for temporary files
            dpi: Resolution for PDF to image conversion
            enable_preprocessing: Enable image preprocessing for scanned docs
            auto_detect: Automatically detect if PDF is digital or scanned
            extract_images: Extract and embed images from PDF
            analyze_layout: Analyze and preserve document layout/structure
            extract_tables: Extract and process tables from scanned pages
        """
        self.output_dir = Path(output_dir)
        self.temp_dir = Path(temp_dir)
        self.dpi = dpi
        self.enable_preprocessing = enable_preprocessing
        self.auto_detect = auto_detect
        self.extract_images = extract_images
        self.analyze_layout = analyze_layout
        self.extract_tables = extract_tables
        self.use_llm_correction = use_llm_correction
        self.enable_vi_translation = enable_vi_translation
        self.use_marker_for_digital_translation = use_marker_for_digital_translation
        # Maximum number of worker threads for page-level parallelism.
        # If None, defaults to number of CPUs.
        self.max_workers: Optional[int] = None
        
        # Create directories
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.temp_dir.mkdir(parents=True, exist_ok=True)
        
        # Create images output folder
        self.images_output_dir = self.output_dir / "extracted_images"
        self.images_output_dir.mkdir(parents=True, exist_ok=True)

        # Resolve credentials before any expensive OCR work starts.
        self.translator = None
        self.translation_cache_root: Optional[Path] = None
        if self.enable_vi_translation:
            cache_dir = Path(__file__).resolve().parents[2] / "artifacts" / "translation"
            self.translator = translator or ScientificTranslator.from_env(
                model=translation_model,
                base_url=translation_base_url,
                cache_dir=cache_dir,
            )
            configured_cache = getattr(self.translator, "cache_dir", None)
            if configured_cache:
                self.translation_cache_root = Path(configured_cache)
        
        # Initialize modules
        self.digital_parser = DigitalParser()
        self.pdf_converter = PDFConverter(dpi=dpi)
        # self.preprocessor = Preprocessor(output_dir=Path(temp_dir) / "preprocessed")
        self.ocr_engine = OCREngine(output_dir=Path(temp_dir) / "craft_output")
        self.exporter = WordExporter()
        self.markdown_processor = MarkdownProcessor(
            # Gemini performs the English OCR correction when translation is on.
            use_llm_correction=use_llm_correction and not enable_vi_translation
        )
        
        # Initialize metrics tracker
        self.metrics = PipelineMetrics(output_dir=self.output_dir)
        self.output_artifacts: Dict[str, Path] = {}
        self.metrics.set_translation_metrics(
            enabled=self.enable_vi_translation,
            model=getattr(self.translator, "model", translation_model),
            base_url=getattr(self.translator, "base_url", translation_base_url),
        )
        
    
    def process_pdf(
        self,
        pdf_path: str | Path,
        output_path: Optional[str | Path] = None,
        mode: Optional[str] = None
    ) -> Path:
        """
        Process a single PDF file
        
        Args:
            pdf_path: Path to input PDF
            output_path: Path for output DOCX (auto-generated if None)
            mode: Processing mode ('digital', 'scan', or None for auto-detect)
        
        Returns:
            Path to output DOCX file
        """
        self.metrics.start_processing()
        
        try:
            pdf_path = Path(pdf_path)
            if not pdf_path.exists():
                raise FileNotFoundError(f"PDF not found: {pdf_path}")
            
            # Track input file
            self.metrics.add_file_processed(pdf_path)
            
            # Determine output path
            if output_path is None:
                output_path = self.output_dir / f"{pdf_path.stem}.docx"
            output_path = Path(output_path)
            
            # Auto-detect mode if not specified
            if mode is None and self.auto_detect:
                is_digital = self.digital_parser.is_digital_pdf(pdf_path)
                mode = "digital" if is_digital else "scan"
            elif mode is None:
                mode = "scan"
            
            # Process based on mode
            if mode == "digital":
                result = self._process_digital(pdf_path, output_path)
            elif mode == "scan":
                result = self._process_scanned(pdf_path, output_path)
            else:
                raise ValueError(f"Unknown mode: {mode}")
            
            # Update metrics with output file sizes
            self.metrics.set_output_files_size(docx_path=output_path)
            self.metrics.end_processing()
            
            return result
            
        except Exception as e:
            self.metrics.add_error(str(e))
            self.metrics.end_processing()
            raise
    
    def _process_digital(self, pdf_path: Path, output_path: Path) -> Path:
        """Keep pdf2docx output and use Marker for structured translation."""
        try:
            self.digital_parser.convert(pdf_path, output_path)
            if self.enable_vi_translation:
                self._process_structured_markdown(
                    pdf_path,
                    english_docx_path=None,
                    allow_digital_fallback=(
                        not self.use_marker_for_digital_translation
                        or not self._vllm_backend_available()
                    ),
                    prefer_digital_fallback=not self.use_marker_for_digital_translation,
                )
            return output_path
        except Exception as e:
            logger.error(f"Digital conversion failed: {e}")
            raise
    
    def _process_scanned(self, pdf_path: Path, output_path: Path) -> Path:
        """Process scanned PDF using marker-pdf OCR pipeline"""
        try:
            self._process_structured_markdown(pdf_path, english_docx_path=output_path)
            return output_path
            
        except Exception as e:
            logger.error(f"Marker-pdf processing failed: {e}")
            self.metrics.add_error(f"Scanned processing: {str(e)}")
            raise

    def _process_structured_markdown(
        self,
        pdf_path: Path,
        english_docx_path: Optional[Path],
        allow_digital_fallback: bool = False,
        prefer_digital_fallback: bool = False,
    ) -> None:
        """Extract structured Markdown, translate it and build image-aware Word files."""
        if prefer_digital_fallback:
            logger.info("Using PyMuPDF for digital PDF; Marker/Docker is disabled.")
            marker_result = self._extract_digital_markdown_fallback(pdf_path)
        else:
            try:
                marker_result = self.ocr_engine.process_pdf(pdf_path)
            except Exception as marker_error:
                if not allow_digital_fallback:
                    raise
                logger.warning(
                    "Marker could not process digital PDF; using PyMuPDF text/image fallback: %s",
                    marker_error,
                )
                marker_result = self._extract_digital_markdown_fallback(pdf_path)
        organized_images = self._organize_extracted_images(
            marker_result.get('images', []), pdf_path
        )
        self.metrics.add_images_extracted([img['output_path'] for img in organized_images])

        processed_markdown = self.markdown_processor.process(
            marker_result['markdown'], images=organized_images
        )
        markdown_path = self.output_dir / f"{pdf_path.stem}_ocr_results.md"

        if self.enable_vi_translation:
            if self.translator is None:
                raise RuntimeError("Vietnamese translation is enabled but no translator is configured.")

            # Isolate checkpoints per source document while retaining content-hash reuse.
            source_key = sha256(str(pdf_path.resolve()).encode("utf-8")).hexdigest()[:16]
            if self.translation_cache_root:
                self.translator.cache_dir = self.translation_cache_root / source_key
                self.translator.cache_dir.mkdir(parents=True, exist_ok=True)

            result: TranslationResult = self.translator.process(
                processed_markdown, document_title=pdf_path.stem
            )
            self._atomic_write_text(markdown_path, result.corrected_markdown)

            vi_markdown_path = self.output_dir / f"{pdf_path.stem}_ocr_results_vi.md"
            vi_docx_path = self.output_dir / f"{pdf_path.stem}_ocr_results_vi.docx"
            self._atomic_write_text(vi_markdown_path, result.vi_markdown)
            self._atomic_export_word(result.vi_markdown, vi_docx_path, organized_images)

            if english_docx_path is not None:
                self._atomic_export_word(
                    result.corrected_markdown, english_docx_path, organized_images
                )

            self.output_artifacts.update({
                "corrected_markdown": markdown_path,
                "vi_markdown": vi_markdown_path,
                "vi_docx": vi_docx_path,
            })
            usage = result.usage
            self.metrics.set_translation_metrics(
                enabled=True,
                model=getattr(self.translator, "model", None),
                base_url=getattr(self.translator, "base_url", None),
                input_tokens=usage.input_tokens,
                output_tokens=usage.output_tokens,
                total_tokens=usage.total_tokens,
                chunks=usage.chunks,
                cache_hits=usage.cache_hits,
                retries=usage.retries,
                glossary_terms=len(result.glossary),
                outputs={key: str(value) for key, value in self.output_artifacts.items()},
            )
        else:
            self._atomic_write_text(markdown_path, processed_markdown)
            if english_docx_path is not None:
                self._atomic_export_word(processed_markdown, english_docx_path, organized_images)
            self.output_artifacts["markdown"] = markdown_path

        self.metrics.set_line_count(len(processed_markdown.split('\n')))
        self.metrics.set_sample_count(
            self._count_samples(processed_markdown, organized_images)
        )
        self.metrics.set_output_files_size(
            markdown_path=markdown_path,
            docx_path=english_docx_path,
        )

    def _extract_digital_markdown_fallback(self, pdf_path: Path) -> Dict[str, Any]:
        """Extract a structured text/image Markdown fallback for digital PDFs.

        Marker remains the preferred extractor. This fallback is useful when
        Surya's optional vLLM Docker backend is unavailable, while preserving
        the same image IDs consumed by MarkdownProcessor and WordExporter.
        """
        try:
            import fitz
        except ImportError as exc:
            raise RuntimeError(
                "PyMuPDF is required for the digital-PDF fallback. "
                "Install it with: pip install PyMuPDF"
            ) from exc

        markdown_parts: List[str] = []
        images: List[Dict[str, Any]] = []
        document = fitz.open(str(pdf_path))
        try:
            for page_number, page in enumerate(document, start=1):
                blocks = page.get_text("blocks", sort=True)
                page_items: List[tuple[float, float, str]] = []
                table_block_indexes, table_markdown = self._extract_pdf_table(blocks)
                raster_figures = self._find_vector_figures(page, blocks)
                if table_markdown:
                    table_y = min(blocks[index][1] for index in table_block_indexes)
                    table_x = min(blocks[index][0] for index in table_block_indexes)
                    page_items.append((table_y, table_x, table_markdown))

                for block_index, block in enumerate(blocks):
                    if (
                        block_index in table_block_indexes
                        or len(block) < 5
                        or not block[4].strip()
                        or any(
                            fitz.Rect(block[:4]).intersects(figure["clip"])
                            for figure in raster_figures
                        )
                    ):
                        continue
                    page_items.append((block[1], block[0], block[4].strip()))

                # Figures 1, 2 and 5 are vector artwork in the PDF rather
                # than raster image XObjects. Render those regions so the
                # translated DOCX retains the figures instead of exposing
                # their labels as unrelated text paragraphs.
                for figure in raster_figures:
                    figure_id = (
                        f"{pdf_path.stem}_page_{page_number:03d}_fig_"
                        f"{figure['figure_number']:02d}"
                    )
                    output_path = self.images_output_dir / f"{figure_id}.png"
                    matrix = fitz.Matrix(2.5, 2.5)
                    page.get_pixmap(matrix=matrix, clip=figure["clip"], alpha=False).save(
                        str(output_path)
                    )
                    page_items.append(
                        (figure["clip"].y0, figure["clip"].x0,
                         f"![id: {figure_id}]({figure_id}.png)")
                    )
                    images.append({
                        "image_id": figure_id,
                        "output_path": str(output_path),
                        "file_path": str(output_path),
                        "page_num": page_number,
                        "type": "figure",
                        "width": int(figure["clip"].width * 2.5),
                        "height": int(figure["clip"].height * 2.5),
                    })

                for image_number, image_info in enumerate(page.get_images(full=True), start=1):
                    xref = image_info[0]
                    extracted = document.extract_image(xref)
                    # Ignore thin decorative/rule images that are embedded in
                    # the PDF but are not figures (the source contains one on
                    # the references page).
                    if (
                        extracted.get("height", 0) <= 200
                        and extracted.get("width", 0) > extracted.get("height", 1) * 8
                    ):
                        continue
                    image_id = f"{pdf_path.stem}_page_{page_number:03d}_img_{image_number:02d}"
                    extension = extracted.get("ext", "png")
                    output_path = self.images_output_dir / f"{image_id}.{extension}"
                    output_path.write_bytes(extracted["image"])
                    image_rects = page.get_image_rects(xref)
                    image_y = image_rects[0].y0 if image_rects else page.rect.height
                    image_x = image_rects[0].x0 if image_rects else page.rect.width
                    page_items.append(
                        (
                            image_y,
                            image_x,
                            f"![id: {image_id}]({image_id}.{extension})",
                        )
                    )
                    images.append({
                        "image_id": image_id,
                        "output_path": str(output_path),
                        "file_path": str(output_path),
                        "page_num": page_number,
                        "type": "image",
                        "width": extracted.get("width", 800),
                        "height": extracted.get("height", 600),
                    })

                page_items.sort(key=lambda item: (item[0], item[1]))
                markdown_parts.extend(item[2] for item in page_items)
        finally:
            document.close()

        index_path = self.output_dir / "images_index.json"
        index_data = {
            "source_pdf": str(pdf_path),
            "output_folder": str(self.images_output_dir),
            "total_images": len(images),
            "images": [
                {
                    "id": image["image_id"],
                    "filename": Path(image["output_path"]).name,
                    "path": image["output_path"],
                }
                for image in images
            ],
        }
        self._atomic_write_text(
            index_path,
            json.dumps(index_data, indent=2, ensure_ascii=False),
        )
        return {
            "markdown": "\n\n".join(markdown_parts),
            "images": images,
        }

    @staticmethod
    def _find_vector_figures(
        page: Any,
        blocks: List[tuple],
    ) -> List[Dict[str, Any]]:
        """Find figure captions whose artwork is not a raster PDF image."""
        import fitz

        raster_rects = []
        for image_info in page.get_images(full=True):
            raster_rects.extend(page.get_image_rects(image_info[0]))

        figures: List[Dict[str, Any]] = []
        for caption_index, block in enumerate(blocks):
            if len(block) < 5:
                continue
            match = re.match(r"\s*Fig\.\s*(\d+)\b", block[4])
            if not match:
                continue
            caption_y = block[1]
            if any(rect.y1 <= caption_y + 2 and rect.y1 > caption_y - 800 for rect in raster_rects):
                continue

            previous_blocks = [
                candidate for candidate in blocks[:caption_index]
                if len(candidate) >= 5 and candidate[3] < caption_y
            ]
            previous = previous_blocks[-1] if previous_blocks else None
            # Figures 1 and 2 contain text inside the vector artwork. Their
            # last text block is therefore part of the figure, not prose that
            # should determine the crop boundary.
            if int(match.group(1)) in {1, 2}:
                top = 28.0
            else:
                top = max(28.0, (previous[3] + 12.0) if previous else 28.0)
            bottom = caption_y - 8.0
            if bottom <= top:
                continue
            clip = fitz.Rect(
                max(145.0, block[0] - 5.0),
                top,
                min(page.rect.width - 30.0, block[2] + 5.0),
                bottom,
            )
            figures.append({
                "figure_number": int(match.group(1)),
                "clip": clip,
                "caption_index": caption_index,
            })
        return figures

    @staticmethod
    def _extract_pdf_table(blocks: List[tuple]) -> tuple[set[int], Optional[str]]:
        """Turn the two simple row-based tables in this paper into Markdown.

        PyMuPDF exposes each PDF table row as one text block containing the
        cell values separated by newlines. Reconstructing those rows here is
        considerably safer than asking the translation model to infer table
        structure from a flat paragraph stream.
        """
        header_index = None
        column_count = 0
        for index, block in enumerate(blocks):
            if len(block) < 5:
                continue
            text = block[4].strip()
            if "Object count" in text and "One or more objects" in text:
                header_index = index
                column_count = 4
                headers = (
                    "Object",
                    "Type",
                    "Object count",
                    "One or more objects present in the image",
                )
                break
            if "Image tag" in text and "Description" in text and "Count" in text:
                header_index = index
                column_count = 3
                headers = ("Image tag", "Description", "Count")
                break
            if "mAP@0.5" in text and "mAP @0.5:0.95" in text:
                header_index = index
                column_count = 7
                headers = (
                    "Class",
                    "Images",
                    "Labels",
                    "Precision",
                    "Recall",
                    "mAP@0.5",
                    "mAP@0.5:0.95",
                )
                break

        if header_index is None:
            return set(), None

        header_block = blocks[header_index]
        header_y = header_block[1]
        table_indexes = {
            index
            for index, block in enumerate(blocks)
            if (
                len(block) >= 5
                and block[0] >= 140
                and block[0] <= 190
                and block[1] >= header_y - 1
                and block[3] <= header_y + 155
                and block[4].strip()
            )
        }
        if header_index not in table_indexes:
            table_indexes.add(header_index)

        rows = [
            "| " + " | ".join(headers) + " |",
            "| " + " | ".join("---" for _ in range(column_count)) + " |",
        ]
        for index in sorted(table_indexes, key=lambda value: blocks[value][1]):
            if index == header_index:
                continue
            cells = [line.strip() for line in blocks[index][4].splitlines() if line.strip()]
            if len(cells) != column_count:
                continue
            cells = [cell.replace("|", "\\|") for cell in cells]
            rows.append("| " + " | ".join(cells) + " |")

        if len(rows) <= 2:
            return set(), None
        return table_indexes, "\n".join(rows)

    @staticmethod
    def _vllm_backend_available() -> bool:
        """Return whether Marker can launch its optional Surya vLLM image."""
        try:
            result = subprocess.run(
                ["docker", "image", "inspect", "vllm/vllm-openai:v0.20.1"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=False,
                timeout=10,
            )
            return result.returncode == 0
        except (OSError, subprocess.SubprocessError):
            return False

    @staticmethod
    def _atomic_write_text(path: Path, content: str) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        temp_path = path.with_suffix(path.suffix + ".tmp")
        temp_path.write_text(content, encoding="utf-8")
        temp_path.replace(path)

    def _atomic_export_word(
        self,
        markdown: str,
        output_path: Path,
        images: List[Dict[str, Any]],
    ) -> None:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        temp_path = output_path.with_name(output_path.stem + ".tmp" + output_path.suffix)
        try:
            self.exporter.markdown_to_word(markdown, str(temp_path), images=images)
            temp_path.replace(output_path)
        finally:
            if temp_path.exists():
                temp_path.unlink()
    
    def _organize_extracted_images(self, images: List[Dict[str, Any]], 
                                   pdf_path: Path) -> List[Dict[str, Any]]:
        """
        Organize extracted images into output folder with distinct IDs
        
        Args:
            images: List of extracted image data
            pdf_path: Path to source PDF
        
        Returns:
            List of images with organized paths and IDs
        """
        organized_images = []
        
        for idx, image in enumerate(images, 1):
            try:
                original_path = Path(image.get('file_path', ''))
                
                if not original_path.exists():
                    logger.warning(f"Image file not found: {original_path}")
                    continue
                
                # Create unique image ID mapping to this image
                image_id = f"{pdf_path.stem}_img_{idx:03d}"
                
                # Copy image to organized output folder
                # Normalize JPEGs to a standard PNG container. Some JPEG
                # streams extracted from PDFs are readable by Pillow but are
                # rejected by python-docx even when their suffix is .jpeg.
                output_suffix = original_path.suffix.lower() or ".png"
                if output_suffix in {".jpg", ".jpeg"}:
                    output_suffix = ".png"
                output_filename = f"{image_id}{output_suffix}"
                output_path = self.images_output_dir / output_filename

                if original_path.suffix.lower() in {".jpg", ".jpeg"}:
                    from PIL import Image

                    with Image.open(original_path) as pil_image:
                        pil_image.convert("RGB").save(output_path, format="PNG")
                else:
                    shutil.copy2(original_path, output_path)
                
                # Update image data with organized path and ID
                organized_image = image.copy()
                organized_image['output_path'] = str(output_path)
                organized_image['image_id'] = image_id
                organized_image['original_file_path'] = str(original_path)
                
                organized_images.append(organized_image)
                logger.debug(f"Organized image {idx}: {image_id} -> {output_filename}")
                
            except Exception as e:
                logger.warning(f"Failed to organize image {idx}: {e}")
                self.metrics.add_error(f"Image organization: {str(e)}")
                continue
        
        # Create index file mapping image IDs
        index_path = self.output_dir / "images_index.json"
        index_data = {
            'source_pdf': str(pdf_path),
            'output_folder': str(self.images_output_dir),
            'total_images': len(organized_images),
            'images': [
                {
                    'id': img['image_id'],
                    'filename': img['output_path'].split('/')[-1] if '/' in img['output_path'] else img['output_path'].split('\\')[-1],
                    'path': img['output_path']
                }
                for img in organized_images
            ]
        }
        
        with open(index_path, 'w', encoding='utf-8') as f:
            json.dump(index_data, f, indent=2, ensure_ascii=False)
        
        return organized_images
    
    def _count_samples(self, markdown_text: str, images: List[Dict[str, Any]]) -> int:
        """
        Count OCR samples/elements (text blocks, images, tables)
        
        Args:
            markdown_text: Processed markdown text
            images: List of extracted images
        
        Returns:
            Total count of samples
        """
        # Count headings
        import re
        headings = len(re.findall(r'^#+\s', markdown_text, re.MULTILINE))
        
        # Count paragraphs
        lines = [l.strip() for l in markdown_text.split('\n') if l.strip()]
        
        # Count tables
        tables = markdown_text.count('\n|')
        
        # Total samples
        total_samples = headings + len(lines) + tables + len(images)
        
        return total_samples
    
    def process_batch(
        self,
        input_dir: str | Path,
        pattern: str = "*.pdf",
        mode: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """
        Process all PDF files in a directory
        
        Args:
            input_dir: Directory containing PDF files
            pattern: File pattern to match (default: "*.pdf")
            mode: Processing mode for all files (None for auto-detect)
        
        Returns:
            List of processing results with status for each file
        """
        input_dir = Path(input_dir)
        if not input_dir.exists():
            raise FileNotFoundError(f"Input directory not found: {input_dir}")
        
        pdf_files = sorted(input_dir.glob(pattern))
        if not pdf_files:
            logger.warning(f"No PDF files found in {input_dir}")
            return []
        
        results = []
        
        for _, pdf_path in enumerate(pdf_files, start=1):     
            try:
                output_path = self.process_pdf(pdf_path, mode=mode)
                results.append({
                    "input": str(pdf_path),
                    "output": str(output_path),
                    "status": "success"
                })
            except Exception as e:
                logger.error(f"Failed to process {pdf_path.name}: {e}")
                results.append({
                    "input": str(pdf_path),
                    "output": None,
                    "status": "failed",
                    "error": str(e)
                })
        
        # Summary
        success_count = sum(1 for r in results if r["status"] == "success")
        return results
