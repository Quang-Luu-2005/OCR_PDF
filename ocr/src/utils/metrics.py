# Pipeline metrics monitoring and tracking
import json
import logging
from pathlib import Path
from typing import Dict, Any, Optional
from datetime import datetime
import os

logger = logging.getLogger(__name__)


class PipelineMetrics:
    """
    Tracks and monitors OCR pipeline metrics including:
    - số dòng (number of lines)
    - số file (number of files)
    - số ảnh (number of images)
    - số mẫu (number of samples/elements)
    - dung lượng (file size/storage)
    """
    
    def __init__(self, output_dir: str | Path = "./output"):
        """
        Initialize Pipeline Metrics tracker
        
        Args:
            output_dir: Directory for saving metrics
        """
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        self.reset()
    
    def reset(self):
        """Reset all metrics to zero"""
        self.metrics = {
            'start_time': None,
            'end_time': None,
            'num_lines': 0,                    # số dòng
            'num_files': 0,                    # số file
            'num_images': 0,                   # số ảnh
            'num_samples': 0,                  # số mẫu (OCR elements)
            'total_size_bytes': 0,             # dung lượng (total bytes)
            'output_size_bytes': 0,            # dung lượng output
            'image_size_bytes': 0,             # dung lượng images
            'markdown_size_bytes': 0,          # dung lượng markdown
            'docx_size_bytes': 0,              # dung lượng docx
            'processing_time_seconds': 0,      # thời gian xử lý
            'files_processed': [],             # danh sách files
            'images_extracted': [],            # danh sách images
            'translation': {
                'enabled': False,
                'model': None,
                'base_url': None,
                'input_tokens': 0,
                'output_tokens': 0,
                'total_tokens': 0,
                'chunks': 0,
                'cache_hits': 0,
                'retries': 0,
                'glossary_terms': 0,
                'outputs': {}
            },
            'errors': []                       # danh sách lỗi
        }
    
    def start_processing(self):
        """Mark the start of processing"""
        self.metrics['start_time'] = datetime.now()
        # logger.info("Pipeline metrics tracking started")
    
    def end_processing(self):
        """Mark the end of processing and calculate duration"""
        self.metrics['end_time'] = datetime.now()
        if self.metrics['start_time']:
            delta = self.metrics['end_time'] - self.metrics['start_time']
            self.metrics['processing_time_seconds'] = delta.total_seconds()
        logger.info(f"Pipeline processing completed in {self.metrics['processing_time_seconds']:.2f} seconds")
    
    def add_file_processed(self, file_path: str | Path, file_size_bytes: Optional[int] = None):
        """
        Record a file that was processed
        
        Args:
            file_path: Path to the processed file
            file_size_bytes: Size of the file in bytes
        """
        file_path = Path(file_path)
        self.metrics['num_files'] += 1
        self.metrics['files_processed'].append({
            'path': str(file_path),
            'name': file_path.name,
            'size_bytes': file_size_bytes or file_path.stat().st_size if file_path.exists() else 0
        })
        if file_size_bytes is None and file_path.exists():
            file_size_bytes = file_path.stat().st_size
        self.metrics['total_size_bytes'] += file_size_bytes or 0
    
    def add_images_extracted(self, image_paths: list | str | Path):
        """
        Record extracted images
        
        Args:
            image_paths: List of image paths or single path
        """
        if isinstance(image_paths, (str, Path)):
            image_paths = [image_paths]
        
        for img_path in image_paths:
            img_path = Path(img_path)
            self.metrics['num_images'] += 1
            
            size_bytes = img_path.stat().st_size if img_path.exists() else 0
            self.metrics['image_size_bytes'] += size_bytes
            
            self.metrics['images_extracted'].append({
                'id': img_path.stem,
                'path': str(img_path),
                'name': img_path.name,
                'size_bytes': size_bytes
            })
    
    def set_line_count(self, num_lines: int):
        """Set the number of lines in extracted text"""
        self.metrics['num_lines'] = num_lines
    
    def set_sample_count(self, num_samples: int):
        """
        Set the number of samples (OCR elements like text blocks, tables, etc.)
        
        Args:
            num_samples: Number of samples/elements extracted
        """
        self.metrics['num_samples'] = num_samples
    
    def set_output_files_size(self, markdown_path: Optional[str | Path] = None, 
                            docx_path: Optional[str | Path] = None):
        """
        Update output file sizes
        
        Args:
            markdown_path: Path to output markdown file
            docx_path: Path to output DOCX file
        """
        if markdown_path:
            markdown_path = Path(markdown_path)
            if markdown_path.exists():
                self.metrics['markdown_size_bytes'] = markdown_path.stat().st_size
        
        if docx_path:
            docx_path = Path(docx_path)
            if docx_path.exists():
                self.metrics['docx_size_bytes'] = docx_path.stat().st_size
        
        self.metrics['output_size_bytes'] = (
            self.metrics['markdown_size_bytes'] + 
            self.metrics['docx_size_bytes'] + 
            self.metrics['image_size_bytes']
        )
    
    def add_error(self, error_msg: str):
        """Record an error during processing"""
        self.metrics['errors'].append(error_msg)
        logger.warning(f"Pipeline error recorded: {error_msg}")

    def set_translation_metrics(
        self,
        *,
        enabled: bool,
        model: Optional[str] = None,
        base_url: Optional[str] = None,
        input_tokens: int = 0,
        output_tokens: int = 0,
        total_tokens: int = 0,
        chunks: int = 0,
        cache_hits: int = 0,
        retries: int = 0,
        glossary_terms: int = 0,
        outputs: Optional[Dict[str, str]] = None,
    ):
        """Record scientific-translation configuration, usage and outputs."""
        self.metrics['translation'] = {
            'enabled': bool(enabled),
            'model': model,
            'base_url': base_url,
            'input_tokens': int(input_tokens),
            'output_tokens': int(output_tokens),
            'total_tokens': int(total_tokens),
            'chunks': int(chunks),
            'cache_hits': int(cache_hits),
            'retries': int(retries),
            'glossary_terms': int(glossary_terms),
            'outputs': outputs or {},
        }
    
    def get_metrics_summary(self) -> Dict[str, Any]:
        """Get the current metrics summary"""
        return self.metrics.copy()
    
    def get_formatted_summary(self) -> str:
        """Get a formatted string summary of metrics"""
        metrics = self.metrics
        
        lines = []
        lines.append("\n" + "="*70)
        lines.append("PIPELINE METRICS SUMMARY")
        lines.append("="*70)
        
        # Processing time
        lines.append(f"\nProcessing Time: {metrics['processing_time_seconds']:.2f} seconds")
        
        # File counts
        lines.append(f"\nFiles:")
        lines.append(f"   • Số file (Number of files): {metrics['num_files']}")
        
        # Text metrics
        lines.append(f"\nText:")
        lines.append(f"   • Số dòng (Number of lines): {metrics['num_lines']:,}")
        lines.append(f"   • Số mẫu (Number of samples): {metrics['num_samples']:,}")
        
        # Image metrics
        lines.append(f"\nImages:")
        lines.append(f"   • Số ảnh (Number of images): {metrics['num_images']}")
        
        # Storage metrics
        lines.append(f"\nStorage:")
        lines.append(f"   • Input: {self._format_bytes(metrics['total_size_bytes'])}")
        lines.append(f"   • Output total: {self._format_bytes(metrics['output_size_bytes'])}")
        if metrics['markdown_size_bytes']:
            lines.append(f"   • Markdown: {self._format_bytes(metrics['markdown_size_bytes'])}")
        if metrics['docx_size_bytes']:
            lines.append(f"   • DOCX: {self._format_bytes(metrics['docx_size_bytes'])}")
        if metrics['image_size_bytes']:
            lines.append(f"   • Images: {self._format_bytes(metrics['image_size_bytes'])}")

        translation = metrics.get('translation', {})
        lines.append("\nVietnamese scientific translation:")
        lines.append(f"   • Enabled: {translation.get('enabled', False)}")
        if translation.get('enabled'):
            lines.append(f"   • Model: {translation.get('model')}")
            lines.append(f"   • Chunks: {translation.get('chunks', 0)}")
            lines.append(f"   • Cache hits: {translation.get('cache_hits', 0)}")
            lines.append(f"   • Retries: {translation.get('retries', 0)}")
            lines.append(f"   • Tokens: {translation.get('total_tokens', 0):,}")
            for label, path in translation.get('outputs', {}).items():
                lines.append(f"   • {label}: {path}")
        
        # Files processed
        if metrics['files_processed']:
            lines.append(f"\nFiles Processed:")
            for f in metrics['files_processed']:
                lines.append(f"   • {f['name']}: {self._format_bytes(f['size_bytes'])}")
        
        # Images extracted
        if metrics['images_extracted']:
            lines.append(f"\nImages Extracted:")
            for img in metrics['images_extracted'][:5]:  # Show first 5
                lines.append(f"   • {img['name']} ({img['id']}): {self._format_bytes(img['size_bytes'])}")
            if len(metrics['images_extracted']) > 5:
                lines.append(f"   ... and {len(metrics['images_extracted']) - 5} more images")
        
        # Errors
        if metrics['errors']:
            lines.append(f"\nErrors ({len(metrics['errors'])}):")
            for err in metrics['errors'][:3]:
                lines.append(f"   • {err}")
            if len(metrics['errors']) > 3:
                lines.append(f"   ... and {len(metrics['errors']) - 3} more errors")
        
        lines.append("\n" + "="*70 + "\n")
        
        return '\n'.join(lines)
    
    @staticmethod
    def _format_bytes(bytes_size: int) -> str:
        """Format bytes to human readable size"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if bytes_size < 1024:
                return f"{bytes_size:.1f} {unit}"
            bytes_size /= 1024
        return f"{bytes_size:.1f} TB"
    
    def save_metrics_json(self, output_path: Optional[str | Path] = None) -> Path:
        """
        Save metrics to JSON file
        
        Args:
            output_path: Path to save metrics JSON (auto-generated if None)
        
        Returns:
            Path to saved metrics file
        """
        if output_path is None:
            output_path = self.output_dir / "pipeline_metrics.json"
        
        output_path = Path(output_path)
        
        # Convert datetime objects to strings for JSON serialization
        metrics_serializable = self.metrics.copy()
        if metrics_serializable['start_time']:
            metrics_serializable['start_time'] = metrics_serializable['start_time'].isoformat()
        if metrics_serializable['end_time']:
            metrics_serializable['end_time'] = metrics_serializable['end_time'].isoformat()
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(metrics_serializable, f, indent=2, ensure_ascii=False)
        
        return output_path
    
    def load_metrics_json(self, input_path: str | Path) -> Dict[str, Any]:
        """Load metrics from JSON file"""
        input_path = Path(input_path)
        with open(input_path, 'r', encoding='utf-8') as f:
            self.metrics = json.load(f)
        return self.metrics
