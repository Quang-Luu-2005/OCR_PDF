import re
import logging
from pathlib import Path
from typing import Optional, Dict, Any

logger = logging.getLogger(__name__)


class MarkdownProcessor:
    """
    Post-processes markdown output to:
    - Insert section breaks after level-2 headings (e.g., 2.1, 2.2)
    - Fix spelling errors using pattern-based correction
    - Preserve images, tables, and structure
    """
    
    def __init__(self, use_llm_correction: bool = True):
        """
        Initialize Markdown Processor
        
        Args:
            use_llm_correction: Enable pattern-based spelling correction
        """
        self.use_llm_correction = use_llm_correction
        
        # Regex pattern for level-2 headings: 2.1, 2.2, 3.1, 3.2, etc.
        # Matches patterns like "## 2.1 Title" or "2.1 Title" or "2.1. Title"
        self.level2_heading_pattern = re.compile(r'^(#{1,6}\s*)?\d+\.\d+\.?\s+.+$', re.MULTILINE)
        
    def insert_section_breaks(self, markdown_text: str) -> str:
        """
        Insert </break> tags at the end of level-2 heading sections (e.g., after content of 2.1, 2.2, etc.)
        Handles any numbering/lettering scheme for level-2 headings.
        
        Args:
            markdown_text: Input markdown text
            
        Returns:
            Markdown text with section breaks inserted at the end of level-2 sections
        """
        lines = markdown_text.split('\n')
        processed_lines = []
        last_heading_level = 0
        last_level2_index = -1  # Track where the last level-2 heading content ends
        
        for i, line in enumerate(lines):
            current_heading_level = self._get_heading_level(line)
            
            # If we encounter a new level-2 heading and we had a previous level-2 section,
            # insert </break> at the end of the previous section (before this new heading)
            if current_heading_level == 2 and last_heading_level == 2:
                processed_lines.append('</break>')
                logger.debug(f"Inserted section break after previous level-2 section")
            
            # If we encounter a level-1 heading after a level-2, also insert break
            elif current_heading_level == 1 and last_heading_level == 2:
                processed_lines.append('</break>')
                logger.debug(f"Inserted section break before level-1 heading")
            
            processed_lines.append(line)
            
            # Update last heading level if this line is a heading
            if current_heading_level > 0:
                last_heading_level = current_heading_level
                if current_heading_level == 2:
                    last_level2_index = len(processed_lines) - 1
        
        # Insert </break> at the end of the last level-2 section if document ends with one
        if last_heading_level == 2:
            processed_lines.append('</break>')
            logger.debug(f"Inserted section break at end of document")
        
        return '\n'.join(processed_lines)
    
    def _get_heading_level(self, line: str) -> int:
        """
        Detect markdown heading level from a line.
        Returns the heading level (1-6) based on # count, or 0 if not a heading.
        
        Handles:
        - Markdown headings: # (level 1), ## (level 2), ### (level 3), etc.
        - Also detects numbered/lettered headings as level 2 if they match X.Y format
        
        Args:
            line: Line to check
            
        Returns:
            Heading level (1-6) or 0 if not a heading
        """
        line = line.strip()
        
        # Check for markdown heading syntax (# ## ### etc.)
        hash_match = re.match(r'^(#{1,6})\s+', line)
        if hash_match:
            hash_count = len(hash_match.group(1))
            return hash_count
        
        # Check for numbered level-2 headings (X.Y format: 2.1, 3.2, etc.)
        # These are considered level-2 since they're sub-sections
        numbered_match = re.match(r'^(\d+)\.(\d+)\s+', line)
        if numbered_match:
            return 2
        
        # Check for lettered level-2 headings (a) b) c) etc.)
        lettered_match = re.match(r'^[a-z]\)\s+', line, re.IGNORECASE)
        if lettered_match:
            return 2
        
        # Check for roman numeral level-2 headings (i) ii) iii) etc.)
        roman_match = re.match(r'^[ivx]+\)\s+', line, re.IGNORECASE)
        if roman_match:
            return 2
        
        return 0
    
    def fix_spelling_with_llm(self, markdown_text: str, images: list = None) -> str:
        """
        Fix spelling errors in Vietnamese markdown text using a hybrid, transformer-free approach.
        Preserves image placeholders, tables, markdown structure, LaTeX math, and HTML tags.
        """
        # If correction disabled, return original
        if not self.use_llm_correction:
            return markdown_text

        try:
            # Protect special elements before correction
            protected_text, protection_map = self._protect_special_elements(markdown_text)
            
            # Try hybrid correction methods in order of preference (no transformers)
            corrected_text = self._fix_with_hybrid_correction(protected_text)
            
            # Restore protected elements
            final_text = self._restore_protected_elements(corrected_text, protection_map)
            
            return final_text

        except Exception as e:
            logger.error(f"Spelling correction failed: {e}")
            logger.warning("Continuing with uncorrected text...")
            return markdown_text
    
    def _protect_special_elements(self, text: str) -> tuple:
        """
        Protect special elements from being modified during spell correction.
        Returns: (protected_text, protection_map)
        """
        protection_map = {}
        protected_text = text
        counter = 0
        
        # 1. Protect LaTeX math expressions: $...$, $$...$$
        # Match both inline ($x$) and display ($$x$$) math
        math_patterns = [
            (r'\$\$[^$]+\$\$', 'MATH_DISPLAY'),  # Display math $$...$$
            (r'\$[^$\n]+\$', 'MATH_INLINE'),      # Inline math $...$
        ]
        
        for pattern, prefix in math_patterns:
            matches = re.finditer(pattern, protected_text)
            for match in reversed(list(matches)):  # Reverse to maintain positions
                placeholder = f"__{prefix}_{counter}__"
                protection_map[placeholder] = match.group(0)
                protected_text = protected_text[:match.start()] + placeholder + protected_text[match.end():]
                counter += 1
        
        # 2. Protect HTML tags: <sub>, <sup>, <br>, and all other HTML tags
        html_tag_pattern = r'<[^>]+>.*?</[^>]+>|<[^>]+/>|<[^>]+>'
        matches = re.finditer(html_tag_pattern, protected_text)
        for match in reversed(list(matches)):
            placeholder = f"__HTML_TAG_{counter}__"
            protection_map[placeholder] = match.group(0)
            protected_text = protected_text[:match.start()] + placeholder + protected_text[match.end():]
            counter += 1
        
        # 3. Protect image placeholders and references
        image_patterns = [
            r'!\[id:\s*[^\]]+\]\([^\)]+\)',  # New format: ![id: img_1](img_1.png)
            r'\[IMAGE_PLACEHOLDER_\d+\]',      # Old format: [IMAGE_PLACEHOLDER_1]
        ]
        
        for pattern in image_patterns:
            matches = re.finditer(pattern, protected_text)
            for match in reversed(list(matches)):
                placeholder = f"__IMAGE_{counter}__"
                protection_map[placeholder] = match.group(0)
                protected_text = protected_text[:match.start()] + placeholder + protected_text[match.end():]
                counter += 1
        
        # 4. Protect markdown headers
        header_pattern = r'^#{1,6}\s+.+$'
        lines = protected_text.split('\n')
        for i, line in enumerate(lines):
            if re.match(header_pattern, line.strip()):
                placeholder = f"__HEADER_{counter}__"
                protection_map[placeholder] = line
                lines[i] = placeholder
                counter += 1
        protected_text = '\n'.join(lines)
        
        # 5. Protect table rows
        table_pattern = r'^\|.+\|$'
        lines = protected_text.split('\n')
        for i, line in enumerate(lines):
            if re.match(table_pattern, line.strip()):
                placeholder = f"__TABLE_ROW_{counter}__"
                protection_map[placeholder] = line
                lines[i] = placeholder
                counter += 1
        protected_text = '\n'.join(lines)
        
        return protected_text, protection_map
    
    def _restore_protected_elements(self, text: str, protection_map: dict) -> str:
        """
        Restore protected elements after spell correction.
        """
        result = text
        # Restore in reverse order to handle nested protections
        for placeholder, original in protection_map.items():
            result = result.replace(placeholder, original)
        return result
    
    def _fix_with_hybrid_correction(self, text: str) -> str:
        """
        Pattern-based correction for Vietnamese text.
        """
        logger.info("Using enhanced pattern-based correction...")
        return self._enhanced_pattern_correction(text)
    
    def _enhanced_pattern_correction(self, text: str) -> str:
        """
        Enhanced pattern-based correction with context awareness.
        """
        lines = text.split('\n')
        corrected_lines = []
        
        for line in lines:
            if line.strip():
                corrected_line = self._apply_enhanced_corrections(line)
                corrected_lines.append(corrected_line)
            else:
                corrected_lines.append(line)
        
        return '\n'.join(corrected_lines)
    

    def _apply_enhanced_corrections(self, line: str) -> str:
        """
        Enhanced OCR corrections with context-aware pattern matching.
        """
        # First apply basic corrections
        line = self._apply_basic_corrections(line)
        
        # Additional context-aware corrections
        # Fix common Vietnamese phrases
        context_corrections = {
            r'\bc h ứ a\b': 'chữa',
            r'\bt h u ố c\b': 'thuốc',
            r'\bb ệ n h\b': 'bệnh',
            r'\bt á c\s+d ụ n g\b': 'tác dụng',
            r'\bc h ấ t\b': 'chất',
            r'\bc h ứ n g\b': 'chứng',
            r'\bt h ể\b': 'thể',
            r'\bc ơ\s+t h ể\b': 'cơ thể',
        }
        
        for pattern, replacement in context_corrections.items():
            line = re.sub(pattern, replacement, line)
        
        return line
    
    def _apply_basic_corrections(self, line: str) -> str:
        """Apply targeted OCR corrections to a single line without breaking formatting"""
        corrections = {
            # Common Vietnamese word splits (OCR breaking words incorrectly)
            r'\bl\s+à\b': 'là',  # "l à" -> "là"
            r'\bc\s+ó\b': 'có',  # "c ó" -> "có"
            r'\bt\s+ừ\b': 'từ',  # "t ừ" -> "từ"
            r'\bđ\s+ược\b': 'được',  # "đ ược" -> "được"
            r'\bn\s+hư\b': 'như',  # "n hư" -> "như"
            r'\bk\s+hi\b': 'khi',  # "k hi" -> "khi"
            r'\bv\s+ới\b': 'với',  # "v ới" -> "với"
            r'\bs\s+ô\b': 'số',  # "s ô" -> "số"
            r'\bm\s+ủ\b': 'mủ',  # "m ủ" -> "mủ"
            r'\bv\s+ự\b': 'vự',  # "v ự" -> "vự"
            r'\bt\s+ại\b': 'tại',  # "t ại" -> "tại"
            r'\bn\s+ên\b': 'nên',  # "n ên" -> "nên"
            r'\bth\s+ì\b': 'thì',  # "th ì" -> "thì"
            r'\btr\s+ong\b': 'trong',  # "tr ong" -> "trong"
            r'\bch\s+o\b': 'cho',  # "ch o" -> "cho"
            r'\bc\s+ủa\b': 'của',  # "c ủa" -> "của"
            r'\bh\s+ay\b': 'hay',  # "h ay" -> "hay"
            r'\bv\s+à\b': 'và',  # "v à" -> "và"
            r'\bn\s+ày\b': 'này',  # "n ày" -> "này"
            r'\bđ\s+ó\b': 'đó',  # "đ ó" -> "đó"
            r'\bk\s+ể\b': 'kể',  # "k ể" -> "kể"
            r'\bt\s+ên\b': 'tên',  # "t ên" -> "tên"
            r'\bch\s+ưa\b': 'chưa',  # "ch ưa" -> "chưa"
            r'\bđ\s+ã\b': 'đã',  # "đ ã" -> "đã"
            r'\bm\s+ột\b': 'một',  # "m ột" -> "một"
            r'\bh\s+ơn\b': 'hơn',  # "h ơn" -> "hơn"
            r'\bcũ\s+ng\b': 'cũng',  # "cũ ng" -> "cũng"
            r'\bnh\s+ững\b': 'những',  # "nh ững" -> "những"
            
            # Common Vietnamese misspellings
            r'\bđươc\b': 'được',
            r'\bduợc\b': 'được',
            r'\bđuợc\b': 'được',
            r'\bnhư\s+ng\b': 'nhưng',
            r'\btr\s+ên\b': 'trên',
            r'\bdu\s+ới\b': 'dưới',
            r'\bsa\s+u\b': 'sau',
            r'\btr\s+ước\b': 'trước',
            r'\bgi\s+ữa\b': 'giữa',
            r'\bhi\s+ện\b': 'hiện',
            r'\bth\s+ực\b': 'thực',
            r'\bch\s+ính\b': 'chính',
            
            # Character confusions
            r'\bΝιάς\b': 'Nước',
            r'\b0\s+\b': 'ở',  # "0" often confused for "ở"
            r'\b1\s+à\b': 'là',  # "1" confused for "l"
            r'\bII\b': 'II',  # Roman numerals
            
            # Number formatting
            r'\.\s+\b1\s+9': '.1.9',  # ". 1,9" -> ".1,9"
            r'(\d+)\s+\.\s+(\d+)': r'\1.\2',  # "5 . 2" -> "5.2"
            r'(\d+)\s+,\s+(\d+)': r'\1,\2',  # "5 , 2" -> "5,2"
            
            # Common punctuation issues
            r'\s+,': ',',  # Space before comma
            r'\s+\.': '.',  # Space before period
            r'\(\s+': '(',  # Space after opening paren
            r'\s+\)': ')',  # Space before closing paren
            
            # Vietnamese tone mark errors
            r'\bhoạ\s+t\b': 'hoạt',
            r'\bkhô\s+ng\b': 'không',
            r'\bthế\s+o\b': 'theo',
            r'\bcá\s+c\b': 'các',
        }
        
        result = line
        for pattern, replacement in corrections.items():
            result = re.sub(pattern, replacement, result)
        
        return result
    
    def _basic_vietnamese_correction(self, text: str) -> str:
        """Basic pattern-based Vietnamese text correction"""
        result = text
        lines = result.split('\n')
        corrected_lines = []
        
        for line in lines:
            if line.strip():
                corrected_line = self._apply_enhanced_corrections(line)
                corrected_lines.append(corrected_line)
            else:
                corrected_lines.append(line)
        
        return '\n'.join(corrected_lines)
    
    def _split_into_chunks(self, text: str, chunk_size: int) -> list:
        """
        Split text into chunks while preserving paragraph boundaries
        
        Args:
            text: Text to split
            chunk_size: Maximum characters per chunk
            
        Returns:
            List of text chunks
        """
        if len(text) <= chunk_size:
            return [text]
        
        chunks = []
        current_chunk = []
        current_length = 0
        
        # Split by paragraphs (double newline)
        paragraphs = text.split('\n\n')
        
        for para in paragraphs:
            para_length = len(para) + 2  # +2 for \n\n
            
            if current_length + para_length > chunk_size and current_chunk:
                # Save current chunk and start new one
                chunks.append('\n\n'.join(current_chunk))
                current_chunk = [para]
                current_length = para_length
            else:
                current_chunk.append(para)
                current_length += para_length
        
        # Add remaining chunk
        if current_chunk:
            chunks.append('\n\n'.join(current_chunk))
        
        return chunks
    
    def process(self, markdown_text: str, images: list = None) -> str:
        """
        Full post-processing pipeline:
        1. Fix markdown syntax errors
        2. Insert image IDs for mapping
        3. Insert section breaks
        4. Fix spelling with LLM
        
        Args:
            markdown_text: Input markdown text
            images: List of image references with IDs
            
        Returns:
            Processed markdown text with image IDs
        """
        # Step 0: Fix markdown syntax errors
        markdown_text = self._fix_markdown_syntax(markdown_text)
        
        # Step 1: Inject image IDs into markdown
        if images:
            markdown_text = self._inject_image_ids(markdown_text, images)
        
        # Step 2: Insert section breaks
        markdown_text = self.insert_section_breaks(markdown_text)
        
        # Step 3: Fix spelling with LLM
        if self.use_llm_correction:
            markdown_text = self.fix_spelling_with_llm(markdown_text, images)
        
        return markdown_text
    
    
    def _fix_markdown_syntax(self, text: str) -> str:
        """
        Fix common markdown syntax errors from OCR output.
        
        Handles:
        - Missing spaces after markdown markers (#, >, -, *)
        - Broken bold/italic markers (**text, *text without closing)
        - Malformed links and images
        - Broken list formatting
        - Broken table structures
        
        Args:
            text: Input markdown text
            
        Returns:
            Fixed markdown text
        """
        lines = text.split('\n')
        fixed_lines = []
        
        for line in lines:
            fixed_line = self._fix_line_markdown_syntax(line)
            fixed_lines.append(fixed_line)
        
        result = '\n'.join(fixed_lines)
        
        # Fix multi-line issues
        result = self._fix_formatting_markers(result)
        result = self._fix_broken_links(result)
        result = self._fix_broken_tables(result)
        
        return result
    
    def _fix_line_markdown_syntax(self, line: str) -> str:
        """Fix markdown syntax issues on a single line"""
        # Fix heading syntax: #Heading -> # Heading
        line = re.sub(r'^(#{1,6})([^\s#])', r'\1 \2', line)
        
        # Fix blockquote syntax: >text -> > text
        line = re.sub(r'^(>+)([^\s>])', r'\1 \2', line)
        
        # Fix list syntax: -text -> - text
        line = re.sub(r'^(\s*)([-*+])([^\s])', r'\1\2 \3', line)
        
        # Fix numbered list: 1.text -> 1. text
        line = re.sub(r'^(\s*)(\d+)\.([^\s.])', r'\1\2. \3', line)
        
        return line
    
    def _fix_formatting_markers(self, text: str) -> str:
        """
        Fix broken bold, italic, and code formatting markers.
        """
        # Find lines with unpaired ** or * markers
        lines = text.split('\n')
        fixed_lines = []
        
        for line in lines:
            # Count unescaped markers
            bold_count = len(re.findall(r'(?<!\\)\*\*', line))
            italic_count = len(re.findall(r'(?<!\\)[*_]', line)) - (bold_count * 2)
            
            # If we have odd number of bold markers, it's likely malformed
            if bold_count % 2 == 1:
                # Try to fix by adding closing marker at line end if text looks like should be bold
                if '**' in line and not line.rstrip().endswith('**'):
                    line = line.rstrip() + '**'
            
            # Same for italic
            if italic_count % 2 == 1:
                # Check if line has unclosed italic
                if re.search(r'[*_][^*_\s]', line) and not line.rstrip().endswith(('*', '_')):
                    # Detect which marker was used
                    if '*' in line:
                        line = line.rstrip() + '*'
            
            fixed_lines.append(line)
        
        return '\n'.join(fixed_lines)
    
    def _fix_broken_links(self, text: str) -> str:
        """
        Fix broken markdown links and image references.
        """
        # Fix incomplete link syntax
        text = re.sub(r'\[([^\]]+)$', r'[\1]', text, flags=re.MULTILINE)
        
        # Fix broken image syntax
        text = re.sub(r'!\[([^\]]+)$', r'![\1]', text, flags=re.MULTILINE)
        
        # Fix broken URL in parentheses
        text = re.sub(r'\(([^\)]+)$', r'(\1)', text, flags=re.MULTILINE)
        
        return text
    
    def _fix_broken_tables(self, text: str) -> str:
        """
        Fix broken table structures.
        """
        lines = text.split('\n')
        fixed_lines = []
        i = 0
        
        while i < len(lines):
            line = lines[i]
            
            # Check if this looks like a table row
            if '|' in line:
                # Ensure line starts and ends with |
                if not line.strip().startswith('|'):
                    line = '|' + line
                if not line.strip().endswith('|'):
                    line = line + '|'
                
                # Check if next line is a separator
                if i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    if re.match(r'^\|[\s\-|:]+\|?$', next_line):
                        # This is a table header, ensure separator is properly formatted
                        if not next_line.startswith('|'):
                            lines[i + 1] = '|' + next_line
                        if not next_line.endswith('|'):
                            lines[i + 1] = next_line + '|'
            
            fixed_lines.append(line)
            i += 1
        
        return '\n'.join(fixed_lines)
    
    def _inject_image_ids(self, markdown_text: str, images: list = None) -> str:
        """
        Inject unique image IDs into markdown for mapping to extracted images
        
        Args:
            markdown_text: Input markdown text
            images: List of image data with 'image_id' field
            
        Returns:
            Markdown text with image IDs injected as references
        """
        if not images:
            return markdown_text
        
        def replacement(match):
            image_number = int(match.group(1))
            image_index = image_number - 1
            if not 0 <= image_index < len(images):
                return match.group(0)
            image_id = images[image_index].get('image_id', f'img_{image_number}')
            logger.debug(f"Injected image ID: {image_id} at placeholder {match.group(0)}")
            return f"![id: {image_id}]({image_id}.png)"

        # Marker commonly wraps its image key in ![](...). Replace the whole
        # wrapper first to avoid producing nested, malformed Markdown images.
        result = re.sub(
            r'!\[[^\]]*\]\(\s*\[IMAGE_PLACEHOLDER_(\d+)\]\s*\)',
            replacement,
            markdown_text,
        )
        return re.sub(r'\[IMAGE_PLACEHOLDER_(\d+)\]', replacement, result)
