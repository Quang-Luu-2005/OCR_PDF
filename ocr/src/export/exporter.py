# Contains python-docx & tagging logic
import json
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import parse_xml
import re
import logging
import os
from pathlib import Path

logger = logging.getLogger(__name__)


class WordExporter:
    def __init__(
        self,
        base_spacing=1.0,
        font_size=12,
        max_image_width=6.0,
        mml2omml_xsl_path=None,
    ):
        self.base_spacing = base_spacing
        self.font_size = font_size
        self.max_image_width = max_image_width  # Maximum width in inches for embedded images
        self.mml2omml_xsl_path = self._resolve_mml2omml_xsl(mml2omml_xsl_path)
        self._mml2omml_transform = None

    @staticmethod
    def _resolve_mml2omml_xsl(configured_path=None):
        candidates = [
            configured_path,
            os.getenv("MML2OMML_XSL_PATH"),
            r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL",
            r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL",
        ]
        for candidate in candidates:
            if candidate and Path(candidate).is_file():
                return Path(candidate)
        return None

    def _latex_to_omml(self, latex):
        if not self.mml2omml_xsl_path:
            raise RuntimeError("MML2OMML.XSL was not found")
        try:
            from latex2mathml.converter import convert
            from lxml import etree
        except ImportError as exc:
            raise RuntimeError("latex2mathml and lxml are required for Word equations") from exc

        if self._mml2omml_transform is None:
            self._mml2omml_transform = etree.XSLT(
                etree.parse(str(self.mml2omml_xsl_path))
            )
        mathml = convert(latex)
        transformed = self._mml2omml_transform(etree.fromstring(mathml.encode("utf-8")))
        xml = etree.tostring(transformed, encoding="unicode")
        if "oMath" not in xml:
            raise RuntimeError("MathML conversion did not produce OMML")
        return parse_xml(xml)

    def _append_native_equation(self, paragraph, latex):
        try:
            paragraph._p.append(self._latex_to_omml(latex))
            return True
        except Exception as exc:
            logger.warning("Could not create native Word equation for %r: %s", latex, exc)
            return False

    @staticmethod
    def _native_equation_is_portable(latex):
        """Use OMML only for shapes verified to render in Word and LibreOffice.

        LibreOffice currently fragments transformed OMML containing large
        operators across lines/cells. The source crop is the lossless fallback
        for those expressions; ordinary fractions, powers and inline math stay
        editable.
        """
        return not re.search(
            r"\\(?:sum|prod|int|iint|iiint)(?=[_{\s])"
            r"|\\begin\s*\{(?:matrix|cases|aligned)\}",
            latex or "",
        )

    @staticmethod
    def _equation_value(equation, name, default=None):
        if isinstance(equation, dict):
            return equation.get(name, default)
        return getattr(equation, name, default)

    def _add_equation_image(self, paragraph, equation):
        image_path = Path(self._equation_value(equation, "image_path", ""))
        if not image_path.is_file():
            return False
        width_inches = 6.2
        try:
            from PIL import Image

            with Image.open(image_path) as image:
                width_inches = min(6.2, max(1.0, image.width / 300.0))
        except Exception:
            pass
        paragraph.add_run().add_picture(str(image_path), width=Inches(width_inches))
        return True

    def _add_display_equation(self, doc, latex, equation=None, number=None):
        # A source crop already includes its equation number. Put it in a
        # full-width paragraph so long equations are never clipped by a table
        # cell in LibreOffice.
        if equation is not None and (
            not latex or not self._native_equation_is_portable(latex)
        ):
            image_paragraph = doc.add_paragraph()
            image_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            image_paragraph.paragraph_format.space_before = Pt(3)
            image_paragraph.paragraph_format.space_after = Pt(3)
            if self._add_equation_image(image_paragraph, equation):
                if not isinstance(equation, dict):
                    equation.output_mode = "image"
                return True

        table = doc.add_table(rows=1, cols=3)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        widths = (Inches(0.55), Inches(5.25), Inches(0.65))
        for cell, width in zip(table.rows[0].cells, widths):
            cell.width = width
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            cell.paragraphs[0].paragraph_format.space_before = Pt(2)
            cell.paragraphs[0].paragraph_format.space_after = Pt(2)

        formula_paragraph = table.cell(0, 1).paragraphs[0]
        formula_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rendered = False
        if latex and self._native_equation_is_portable(latex):
            rendered = self._append_native_equation(formula_paragraph, latex)
        if not rendered and equation is not None:
            rendered = self._add_equation_image(formula_paragraph, equation)
            if rendered and not isinstance(equation, dict):
                equation.output_mode = "image"
                if not equation.validation_error:
                    equation.validation_error = "OMML conversion failed"
        if not rendered:
            run = formula_paragraph.add_run("[Equation unavailable]")
            run.font.italic = True
            run.font.color.rgb = RGBColor(128, 128, 128)

        if number:
            number_paragraph = table.cell(0, 2).paragraphs[0]
            number_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            number_paragraph.add_run(f"({str(number).strip('() ')})")
        return rendered

    @staticmethod
    def center_y(box):
        return sum(p[1] for p in box) / len(box)
    
    def add_image_to_document(self, doc, image_data, show_id_caption=True):
        """
        Add an image to the Word document with proper formatting
        
        Args:
            doc: Document object
            image_data: Dictionary with 'file_path', 'image_id', 'width', 'height'
        
        Returns:
            bool: True if image was added successfully, False otherwise
        """
        try:
            image_path_value = image_data.get('output_path') or image_data.get('file_path')
            if not image_path_value:
                logger.debug(
                    f"Skipping image {image_data.get('image_id', 'unknown')}: no image path"
                )
                return False
            image_path = Path(image_path_value)
            
            # Skip if file doesn't exist (filtered out during extraction)
            if not image_path.exists():
                logger.debug(f"Skipping image {image_data.get('image_id', 'unknown')}: file not found (likely filtered)")
                return False
            
            # Add image with caption
            image_id = image_data.get('image_id', 'unknown')
            
            # Add spacing before image
            spacing_p = doc.add_paragraph()
            spacing_p.paragraph_format.space_before = Pt(6)
            spacing_p.paragraph_format.space_after = Pt(3)
            
            # Add paragraph for image
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(6)
            
            # Calculate image dimensions (maintain aspect ratio)
            original_width = image_data.get('width', 800)
            original_height = image_data.get('height', 600)

            # Validate dimensions and set defaults if invalid
            if original_width <= 0:
                logger.warning(f"Invalid width for image {image_id}, using default")
                original_width = 800
            if original_height <= 0:
                logger.warning(f"Invalid height for image {image_id}, using default")
                original_height = 600

            aspect_ratio = original_height / original_width
            
            # Adjust max width based on aspect ratio (taller images get less width)
            if aspect_ratio > 1.5:  # Very tall image
                max_width = min(4.0, self.max_image_width)
            elif aspect_ratio > 1.2:  # Tall image  
                max_width = min(5.0, self.max_image_width)
            else:  # Normal or wide image
                max_width = self.max_image_width
            
            # Calculate final dimensions
            width_inches = min(max_width, original_width / 100)
            height_inches = width_inches * aspect_ratio
            
            # Limit height to avoid overly tall images
            max_height = 8.0  # Maximum 8 inches tall
            if height_inches > max_height:
                height_inches = max_height
                width_inches = height_inches / aspect_ratio

            # Add the image
            run = p.add_run()
            run.add_picture(str(image_path), width=Inches(width_inches))
            
            # Internal IDs are useful for debug/native exports, but should not
            # appear in the reader-facing Markdown translation document.
            if show_id_caption:
                caption_p = doc.add_paragraph()
                caption_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                caption_p.paragraph_format.space_before = Pt(3)
                caption_p.paragraph_format.space_after = Pt(6)
                caption_run = caption_p.add_run(f"[{image_id}]")
                caption_run.font.size = Pt(9)
                caption_run.font.italic = True
                caption_run.font.color.rgb = RGBColor(100, 100, 100)
            
            logger.debug(f"Added image {image_id} to document ({width_inches:.1f}x{height_inches:.1f} inches)")
            return True
            
        except FileNotFoundError:
            logger.debug(f"Image file not found: {image_data.get('image_id', 'unknown')}")
            return False
        except Exception as e:
            logger.warning(f"Failed to add image {image_data.get('image_id', 'unknown')}: {e}")
            return False

    def add_table_to_document(self, doc, table_data):
        """
        Add a table to the Word document with improved formatting
        
        Args:
            doc: Document object
            table_data: Dictionary with table structure and cell data
        
        Returns:
            bool: True if table was added successfully
        """
        try:
            table_id = table_data.get('table_id', 'unknown')
            structure = table_data.get('structure', {})
            cells = structure.get('cells', [])
            
            if not cells:
                # If no structure, add as image
                return self.add_image_to_document(doc, table_data)
            
            rows = len(cells)
            cols = len(cells[0]) if cells else 0
            
            if rows == 0 or cols == 0:
                return False
            
            # Add spacing before table
            spacing_p = doc.add_paragraph()
            spacing_p.paragraph_format.space_before = Pt(6)
            spacing_p.paragraph_format.space_after = Pt(3)
            
            # Add table title
            title_p = doc.add_paragraph()
            title_run = title_p.add_run(f"[{table_id}]")
            title_run.font.bold = True
            title_run.font.size = Pt(10)
            title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            title_p.paragraph_format.space_before = Pt(6)
            title_p.paragraph_format.space_after = Pt(3)
            
            # Create table
            table = doc.add_table(rows=rows, cols=cols)
            table.style = 'Table Grid'
            
            # Apply table formatting
            table.autofit = False
            table.allow_autofit = False
            
            # Fill cells
            for i, row_data in enumerate(cells):
                for j, cell_data in enumerate(row_data):
                    cell = table.rows[i].cells[j]
                    cell.text = cell_data.get('text', '')
                    
                    # Format cell text
                    for paragraph in cell.paragraphs:
                        paragraph.paragraph_format.space_before = Pt(2)
                        paragraph.paragraph_format.space_after = Pt(2)
                        for run in paragraph.runs:
                            run.font.size = Pt(9)
                            
                    # Add cell padding
                    cell.vertical_alignment = 1  # Center vertically
            
            # Add spacing after table
            after_spacing_p = doc.add_paragraph()
            after_spacing_p.paragraph_format.space_before = Pt(3)
            after_spacing_p.paragraph_format.space_after = Pt(6)
            
            logger.debug(f"Added table {table_id} ({rows}x{cols}) to document")
            return True
            
        except Exception as e:
            logger.warning(f"Failed to add table {table_data.get('table_id', 'unknown')}: {e}")
            # Fallback: try to add as image
            try:
                return self.add_image_to_document(doc, table_data)
            except:
                return False

    @staticmethod
    def inject_break_tag(text: str) -> str:
        """Insert </break> tag after level 2 headings (e.g., 2.1, 2.2, 3.1, etc.)"""
        # Pattern matches level 2 headings: X.Y where X and Y are digits
        pattern = r"^\s*\d+\.\d+\s"
        if re.match(pattern, text.strip()):
            # Check if it's exactly level 2 (one dot, not 2.1.1)
            heading_part = text.strip().split()[0] if text.strip().split() else ""
            if heading_part.count('.') == 1:
                return text.rstrip() + " </break>"
        return text

    def write_to_word(self, data, output_path="output.docx", images=None, tables=None):
        """
        Export OCR results to Word document
        
        Args:
            data: Can be:
                - str: path to JSON file
                - list: OCR results (assumes single page)
                - dict with 'results' key: OCR results
                - dict with 'pages' key: Multi-page OCR results
            output_path: str, path to save docx
            images: Optional list of image dictionaries to embed
            tables: Optional list of table dictionaries to embed
        
        Returns:
            Path to saved document
        """
        if isinstance(data, str):
            # Assume it's a path to JSON file
            with open(data, "r", encoding="utf-8") as f:
                data = json.load(f)

        if isinstance(data, dict) and "results" in data:
            items = data["results"]
        elif isinstance(data, dict) and "pages" in data:
            # Multi-page format: {"pages": [{page_num: 1, results: [...]}, ...]}
            return self._write_multipage_to_word(data["pages"], output_path, images, tables)
        elif isinstance(data, list):
            items = data
        elif isinstance(data, dict):
            # For layout export: expects keys 'text', 'images', 'tables'
            doc = Document()
            if data.get("text"):
                doc.add_paragraph(data["text"])
            # Images and tables can be handled here if needed
            doc.save(output_path)
            print("Saved:", output_path)
            return output_path
        else:
            raise ValueError("Unsupported data format for export")

        # OCR results export - handle both with and without page_id
        pages = {}
        for it in items:
            page_id = it.get("page_id", 1)
            pages.setdefault(page_id, []).append(it)

        doc = Document()

        for page_index, (page_id, page_items) in enumerate(pages.items()):
            # Separate text items, image items, and table items
            text_items = [it for it in page_items if it.get("element_type") not in ["image", "table"]]
            image_items = [it for it in page_items if it.get("element_type") == "image"]
            table_items = [it for it in page_items if it.get("element_type") == "table"]
            
            # Sort text items by position
            text_items = sorted(
                text_items,
                key=lambda t: (self.center_y(t.get("box", [[0,0]])), t.get("box", [[0,0]])[0][0])
            )

            last_y = None
            last_height = 12

            for it in text_items:
                # Skip headers and footers if marked
                if it.get("skip", False):
                    continue
                
                text = it.get("text", "")
                box = it.get("box", [[0, 0], [0, 0], [0, 0], [0, 0]])
                element_type = it.get("element_type", "paragraph")

                # Apply break tag for level 2 headings
                text = self.inject_break_tag(text)

                y = self.center_y(box)
                height = abs(box[2][1] - box[0][1])

                if last_y is None:
                    spacing = self.base_spacing
                else:
                    gap = y - last_y
                    spacing = max(self.base_spacing, gap / last_height)

                p = doc.add_paragraph()
                run = p.add_run(text)
                run.font.size = Pt(self.font_size)
                
                # Apply heading formatting
                if element_type == "heading":
                    run.font.bold = True
                    heading_level = it.get("heading_level", 0)
                    if heading_level == 1:
                        run.font.size = Pt(16)
                    elif heading_level == 2:
                        run.font.size = Pt(14)
                    elif heading_level >= 3:
                        run.font.size = Pt(13)

                p.paragraph_format.line_spacing = spacing

                last_y = y
                last_height = max(height, 8)
            
            # Add images for this page
            if image_items:
                added_count = 0
                for img_item in image_items:
                    if self.add_image_to_document(doc, img_item):
                        added_count += 1
                # if added_count > 0:
                #     logger.info(f"Added {added_count}/{len(image_items)} images for page {page_id}")
            
            # Add tables for this page
            if table_items:
                added_count = 0
                for tbl_item in table_items:
                    if self.add_table_to_document(doc, tbl_item):
                        added_count += 1
                # if added_count > 0:
                #     logger.info(f"Added {added_count}/{len(table_items)} tables for page {page_id}")

            if page_index < len(pages) - 1:
                doc.add_page_break()

        doc.save(output_path)
        print("Saved:", output_path)
        return output_path
    
    def _write_multipage_to_word(self, pages_data, output_path="output.docx", images=None, tables=None):
        """
        Export multi-page OCR results to Word document
        
        Args:
            pages_data: List of dicts with 'page_num' and 'results' keys
            output_path: str, path to save docx
            images: Optional list of image dictionaries to embed
            tables: Optional list of table dictionaries to embed
        
        Returns:
            Path to saved document
        """
        doc = Document()
        
        # Sort pages by page number
        pages_data = sorted(pages_data, key=lambda p: p.get("page_num", 1))
        
        # Group images and tables by page
        images_by_page = {}
        tables_by_page = {}
        
        if images:
            for img in images:
                page_num = img.get("page_num", 1)
                images_by_page.setdefault(page_num, []).append(img)
        
        if tables:
            for tbl in tables:
                page_num = tbl.get("page_num", 1)
                tables_by_page.setdefault(page_num, []).append(tbl)
        
        for page_idx, page_data in enumerate(pages_data):
            page_num = page_data.get("page_num", page_idx + 1)
            results = page_data.get("results", [])
            
            if not results:
                continue
            
            # Separate text, image, and table elements
            text_results = [r for r in results if r.get("element_type") not in ["image", "table"]]
            image_results = [r for r in results if r.get("element_type") == "image"]
            table_results = [r for r in results if r.get("element_type") == "table"]
            
            # Sort text results by vertical position, then horizontal
            text_results = sorted(
                text_results,
                key=lambda t: (self.center_y(t.get("box", [[0,0]])), t.get("box", [[0,0]])[0][0])
            )
            
            last_y = None
            last_height = 12
            
            for it in text_results:
                # Skip headers and footers
                if it.get("skip", False):
                    continue
                
                text = it.get("text", "")
                box = it.get("box", [[0, 0], [0, 0], [0, 0], [0, 0]])
                element_type = it.get("element_type", "paragraph")
                
                # Apply break tag
                text = self.inject_break_tag(text)
                
                y = self.center_y(box)
                height = abs(box[2][1] - box[0][1])
                
                if last_y is None:
                    spacing = self.base_spacing
                else:
                    gap = y - last_y
                    spacing = max(self.base_spacing, gap / last_height)
                
                p = doc.add_paragraph()
                run = p.add_run(text)
                run.font.size = Pt(self.font_size)
                
                # Apply heading formatting
                if element_type == "heading":
                    run.font.bold = True
                    heading_level = it.get("heading_level", 0)
                    if heading_level == 1:
                        run.font.size = Pt(16)
                    elif heading_level == 2:
                        run.font.size = Pt(14)
                    elif heading_level >= 3:
                        run.font.size = Pt(13)
                
                p.paragraph_format.line_spacing = spacing
                
                last_y = y
                last_height = max(height, 8)
            
            # Add images from results
            added_from_results = 0
            for img_item in image_results:
                if self.add_image_to_document(doc, img_item):
                    added_from_results += 1
            
            # Add images from separate image list
            added_from_list = 0
            if page_num in images_by_page:
                for img in images_by_page[page_num]:
                    if self.add_image_to_document(doc, img):
                        added_from_list += 1
            
            # Add tables from results
            added_tables_from_results = 0
            for tbl_item in table_results:
                if self.add_table_to_document(doc, tbl_item):
                    added_tables_from_results += 1
            
            # Add tables from separate table list
            added_tables_from_list = 0
            if page_num in tables_by_page:
                for tbl in tables_by_page[page_num]:
                    if self.add_table_to_document(doc, tbl):
                        added_tables_from_list += 1
            
            # Log summary
            total_images = added_from_results + added_from_list
            total_tables = added_tables_from_results + added_tables_from_list
            
            # if total_images > 0:
            #     logger.info(f"Page {page_num}: Added {total_images} images")
            # if total_tables > 0:
            #     logger.info(f"Page {page_num}: Added {total_tables} tables")
            
            # Add page break between pages (except last page)
            if page_idx < len(pages_data) - 1:
                doc.add_page_break()
        
        doc.save(output_path)
        print(f"Saved {len(pages_data)} pages to: {output_path}")
        return output_path

    

    def markdown_to_word(
        self,
        markdown_text,
        output_path="output.docx",
        images=None,
        equations=None,
    ):
        """
        Convert markdown text to Word document with native tables and formatting.
        Properly handles markdown syntax including bold, italic, code, links, and more.
        """
        doc = Document()
        
        # --- Helper: Image Lookup ---
        image_map = {}
        for img in images or []:
            aliases = {
                img.get('image_id', ''),
                img.get('id', ''),
                Path(img.get('output_path', '')).stem if img.get('output_path') else '',
                Path(img.get('file_path', '')).stem if img.get('file_path') else '',
                Path(img.get('path', '')).stem if img.get('path') else '',
                Path(img.get('filename', '')).stem if img.get('filename') else '',
                Path(img.get('original_file_path', '')).stem
                if img.get('original_file_path')
                else '',
            }
            for alias in aliases - {''}:
                image_map[alias] = img

        equation_map = {
            self._equation_value(equation, "equation_id"): equation
            for equation in equations or []
            if self._equation_value(equation, "equation_id")
        }

        # --- Helper: Apply formatting (Bold/Italic/Code/Links/Subscript/Superscript/Math) ---
        def add_formatted_text(paragraph, text):
            """
            Parse and apply markdown formatting to text:
            - **bold** or __bold__
            - *italic* or _italic_
            - ***bold+italic***
            - `code`
            - [link text](url)
            - <sub>subscript</sub>
            - <sup>superscript</sup>
            - $math$ (inline LaTeX)
            """
            # Process text sequentially, handling nested formatting
            
            # First, protect and convert LaTeX math and HTML tags
            parts_to_process = [text]
            final_parts = []
            
            # Pattern to match: LaTeX math, HTML sub/sup, or markdown formatting
            # Priority order: LaTeX > HTML tags > Markdown
            combined_pattern = r'(\$[^$\n]+\$|<sub>.*?</sub>|<sup>.*?</sup>|<[^>]+>|\*\*\*.*?\*\*\*|\*\*.*?\*\*|__.*?__|_.*?_|\*.*?\*|`.*?`|\[.*?\]\(.*?\))'
            
            parts = re.split(combined_pattern, text)
            
            for part in parts:
                if not part:
                    continue
                
                run = paragraph.add_run()
                
                # LaTeX inline math $...$
                if part.startswith('$') and part.endswith('$') and len(part) > 2:
                    math_content = part[1:-1]
                    paragraph._p.remove(run._element)
                    if not self._append_native_equation(paragraph, math_content):
                        fallback_run = paragraph.add_run("[Equation unavailable]")
                        fallback_run.font.italic = True
                        fallback_run.font.color.rgb = RGBColor(128, 128, 128)
                # Subscript <sub>...</sub>
                elif part.startswith('<sub>') and part.endswith('</sub>'):
                    subscript_text = part[5:-6]  # Remove <sub></sub>
                    run.text = subscript_text
                    run.font.subscript = True
                    run.font.size = Pt(8)
                # Superscript <sup>...</sup>
                elif part.startswith('<sup>') and part.endswith('</sup>'):
                    superscript_text = part[5:-6]  # Remove <sup></sup>
                    run.text = superscript_text
                    run.font.superscript = True
                    run.font.size = Pt(8)
                # Bold + Italic (*** or ___)
                elif (part.startswith('***') and part.endswith('***') and len(part) > 6):
                    run.text = part[3:-3]
                    run.font.bold = True
                    run.font.italic = True
                # Bold (** or __)
                elif (part.startswith('**') and part.endswith('**') and len(part) > 4):
                    run.text = part[2:-2]
                    run.font.bold = True
                elif (part.startswith('__') and part.endswith('__') and len(part) > 4):
                    run.text = part[2:-2]
                    run.font.bold = True
                # Italic (* or _)
                elif (part.startswith('*') and part.endswith('*') and len(part) > 2 and not part.startswith('**')):
                    run.text = part[1:-1]
                    run.font.italic = True
                elif (part.startswith('_') and part.endswith('_') and len(part) > 2 and not part.startswith('__')):
                    run.text = part[1:-1]
                    run.font.italic = True
                # Code (`)
                elif (part.startswith('`') and part.endswith('`') and len(part) > 2):
                    run.text = part[1:-1]
                    run.font.name = 'Courier New'
                    run.font.size = Pt(10)
                    run.font.color.rgb = RGBColor(199, 37, 78)
                # Link ([text](url))
                elif re.match(r'\[.*?\]\(.*?\)', part):
                    link_match = re.match(r'\[(.*?)\]\((.*?)\)', part)
                    if link_match:
                        link_text = link_match.group(1)
                        link_url = link_match.group(2)
                        run.text = link_text
                        run.font.color.rgb = RGBColor(0, 0, 255)
                        run.font.underline = True
                        # Note: python-docx doesn't support clickable hyperlinks easily
                        # This just styles it like a link
                else:
                    run.text = part
                    
        # --- Helper: Process Image Placeholder ---
        def process_image_placeholder(line):
            """Handle both [IMAGE_PLACEHOLDER_X] and ![id: img_X](img_X.png) formats"""
            # New format: ![id: img_X](img_X.png)
            match = re.search(r'!\[id:\s*([^\]]+)\]\([^\)]+\)', line)
            if match:
                img_id = match.group(1).strip()
                if img_id in image_map:
                    self.add_image_to_document(doc, image_map[img_id], show_id_caption=False)
                    return True
                else:
                    # Try old format fallback
                    p = doc.add_paragraph()
                    run = p.add_run(f"[{img_id}]")
                    run.font.italic = True
                    run.font.color.rgb = RGBColor(128, 128, 128)
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    return True
            
            # Old format: [IMAGE_PLACEHOLDER_X]
            match = re.search(r'\[IMAGE_PLACEHOLDER_(\d+)\]', line)
            if match:
                img_num = match.group(1)
                img_id = f'img_{img_num}'
                if img_id in image_map:
                    self.add_image_to_document(doc, image_map[img_id], show_id_caption=False)
                else:
                    p = doc.add_paragraph()
                    run = p.add_run(f"[IMAGE {img_num}]")
                    run.font.italic = True
                    run.font.color.rgb = RGBColor(128, 128, 128)
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                return True
            return False

        # --- Main Parsing Loop ---
        lines = markdown_text.split('\n')
        table_buffer = [] # To store table rows temporarily
        in_table = False
        in_code_block = False
        code_block_buffer = []
        
        i = 0
        while i < len(lines):
            line = lines[i]
            clean_line = line.strip()

            # Display equation with stable metadata.
            equation_match = re.fullmatch(
                r'\s*\$\$(.*?)\$\$\s*<!--\s*equation:id=([^;]+);number=([^\s>]*)\s*-->\s*',
                line,
            )
            if equation_match:
                latex = equation_match.group(1).strip()
                equation_id = equation_match.group(2).strip()
                number = equation_match.group(3).strip() or None
                self._add_display_equation(
                    doc,
                    latex,
                    equation=equation_map.get(equation_id),
                    number=number,
                )
                i += 1
                continue

            # Marker and user-authored Markdown may contain display math
            # without pipeline metadata. Convert it through the same OMML path.
            plain_display_match = re.fullmatch(r'\s*\$\$(.*?)\$\$\s*', line)
            if plain_display_match:
                self._add_display_equation(
                    doc, plain_display_match.group(1).strip()
                )
                i += 1
                continue
            if clean_line == "$$":
                display_lines = []
                i += 1
                while i < len(lines) and lines[i].strip() != "$$":
                    display_lines.append(lines[i])
                    i += 1
                self._add_display_equation(doc, "\n".join(display_lines).strip())
                i += int(i < len(lines))
                continue

            equation_image_match = re.fullmatch(
                r'\s*!\[equation:\s*([^\]]+)\]\([^\)]+\)\s*'
                r'<!--\s*equation:number=([^\s>]*)\s*-->\s*',
                line,
            )
            if equation_image_match:
                equation_id = equation_image_match.group(1).strip()
                number = equation_image_match.group(2).strip() or None
                self._add_display_equation(
                    doc,
                    None,
                    equation=equation_map.get(equation_id),
                    number=number,
                )
                i += 1
                continue
                        
            # 1. Handle Table Logic
            if clean_line.startswith('|'):
                # It's a table row
                if not in_table:
                    # Check if the NEXT line is a separator (|---|) to confirm it's a header
                    if i + 1 < len(lines) and '---' in lines[i+1]:
                        in_table = True
                        table_buffer = [line] # Start buffer
                        i += 1
                        continue # Skip processing this line as normal text
                else:
                    # If we are already in a table, just add the line
                    # Ignore the separator line (e.g. |---|---|)
                    if '---' in line:
                        i += 1
                        continue 
                    table_buffer.append(line)
                    i += 1
                    continue
            
            # If we were in a table, but this line is NOT a table line (or empty)
            if in_table:
                # Render the buffered table now
                if table_buffer:
                    # Calculate rows and columns
                    rows_data = [[cell.strip() for cell in row.strip('|').split('|')] for row in table_buffer]
                    num_rows = len(rows_data)
                    num_cols = len(rows_data[0]) if rows_data else 0
                    
                    if num_rows > 0 and num_cols > 0:
                        table = doc.add_table(rows=num_rows, cols=num_cols)
                        table.style = 'Table Grid'
                        
                        for r_idx, row_data in enumerate(rows_data):
                            row_cells = table.rows[r_idx].cells
                            for c_idx, cell_text in enumerate(row_data):
                                if c_idx < len(row_cells):
                                    # Convert HTML <br> tags to proper line breaks
                                    cell_text = cell_text.replace('<br>', '\n').replace('<br/>', '\n').replace('<br />', '\n')
                                    
                                    # Split by newlines and add as separate paragraphs for proper line breaks
                                    lines_in_cell = cell_text.split('\n')
                                    p = row_cells[c_idx].paragraphs[0]
                                    add_formatted_text(p, lines_in_cell[0])
                                    
                                    # Format header row (first row) as bold
                                    if r_idx == 0:
                                        for run in p.runs:
                                            run.font.bold = True
                                    
                                    # Add additional lines as new paragraphs
                                    for cell_line in lines_in_cell[1:]:
                                        new_p = row_cells[c_idx].add_paragraph()
                                        add_formatted_text(new_p, cell_line)
                                        if r_idx == 0:
                                            for run in new_p.runs:
                                                run.font.bold = True
                
                # Reset table state
                in_table = False
                table_buffer = []
                
                # If this current line was empty, we are done with the table. 
                if not clean_line:
                    i += 1
                    continue

            # 2. Handle Images
            if '[IMAGE_PLACEHOLDER_' in line or '![id:' in line:
                if process_image_placeholder(line):
                    i += 1
                    continue

            # 3. Skip empty lines (outside of tables)
            if not clean_line:
                i += 1
                continue

            # 4. Handle Headings
            if line.startswith('#'):
                # Count the number of # characters
                hash_count = 0
                for char in line:
                    if char == '#':
                        hash_count += 1
                    else:
                        break
                
                if hash_count <= 6 and hash_count > 0:
                    level = hash_count
                    text = line[level:].strip()
                    p = doc.add_paragraph()
                    add_formatted_text(p, text)
                    
                    # Apply heading styles
                    if level <= 4:
                        p.style = f'Heading {level}'
                    else:
                        # For h5 and h6, apply manual formatting
                        for run in p.runs:
                            run.font.bold = True
                            run.font.size = Pt(11 if level == 5 else 10)
                    i += 1
                    continue

            # 5. Handle Lists (Unordered)
            if line.strip().startswith(('- ', '* ', '+ ')):
                p = doc.add_paragraph(style='List Bullet')
                text = line.strip()[2:]
                add_formatted_text(p, text)
                p.paragraph_format.left_indent = Pt(20)
                i += 1
                continue
            
            # 6. Handle Lists (Ordered)
            elif re.match(r'^\s*\d+\.\s', line):
                # Numbered list
                text = re.sub(r'^\s*\d+\.\s', '', line)
                p = doc.add_paragraph(style='List Number')
                add_formatted_text(p, text)
                p.paragraph_format.left_indent = Pt(20)
                i += 1
                continue

            # 7. Handle horizontal rules (---, ***, ___)
            if re.match(r'^\s*(-{3,}|\*{3,}|_{3,})\s*$', line):
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(6)
                p.add_run('_' * 50)  # Horizontal line
                i += 1
                continue

            # 8. Handle break tags (skip them as they're document structure markers)
            if clean_line == '</break>':
                # Could add a page break or section break here if needed
                # doc.add_page_break()
                i += 1
                continue
            
            # 9. Handle blockquotes (> text)
            if line.strip().startswith('>'):
                text = line.strip()[1:].strip()
                p = doc.add_paragraph()
                add_formatted_text(p, text)
                p.paragraph_format.left_indent = Pt(36)
                p.paragraph_format.space_before = Pt(3)
                p.paragraph_format.space_after = Pt(3)
                # Add quote styling
                for run in p.runs:
                    run.font.italic = True
                    run.font.color.rgb = RGBColor(96, 96, 96)
                i += 1
                continue
            
            # 10. Handle HTML <br> tags in regular paragraphs
            if '<br>' in line or '<br/>' in line or '<br />' in line:
                line = line.replace('<br>', '\n').replace('<br/>', '\n').replace('<br />', '\n')
                # Split by newlines and add as paragraph with internal line breaks
                p = doc.add_paragraph()
                lines_in_para = line.split('\n')
                for idx, para_line in enumerate(lines_in_para):
                    if idx > 0:
                        p.add_run('\n')  # Line break within paragraph
                    add_formatted_text(p, para_line)
            else:
                # Regular Paragraph
                p = doc.add_paragraph()
                add_formatted_text(p, line)
            
            i += 1

        doc.save(output_path)
        return output_path

