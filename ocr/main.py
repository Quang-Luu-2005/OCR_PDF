# Entry point for OCR phase
# Parses arguments and coordinates OCR modules

import argparse
import sys
import logging
import os
from pathlib import Path

from src.core.pipeline import OCRPipeline

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def check_ocr_dependencies(translation_enabled=False):
    """Check if all required libraries for marker-pdf OCR are available"""
    missing = []
    try:
        from marker.converters.pdf import PdfConverter
    except ImportError:
        missing.append("marker-pdf")
    if translation_enabled:
        try:
            import openai  # noqa: F401
            import dotenv  # noqa: F401
        except ImportError:
            missing.append("openai and python-dotenv")
    if missing:
        print("Missing required libraries:")
        for lib in missing:
            print(f"   - {lib}")
        print("\nInstall missing libraries with:")
        print("   pip install -r requirements.txt")
        return False
    return True

def main():
    # Windows consoles may use a legacy code page that cannot print valid
    # scientific Unicode, which must not abort an otherwise completed run
    # before metrics are persisted.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(
        description="Document OCR & Conversion Pipeline - Convert PDFs to Word documents"
    )

    parser.add_argument(
        "--input",
        required=False,
        default=None,
        help="Input PDF file or directory containing PDFs (defaults to config.INPUT_DIR if present)"
    )
    
    parser.add_argument(
        "--output",
        default="./output",
        help="Output directory for converted files (default: ./output)"
    )

    parser.add_argument(
        "--mode",
        choices=["auto", "scan", "digital"],
        default="auto",
        help="Processing mode: auto (detect) | scan (OCR) | digital (direct conversion)"
    )

    parser.add_argument(
        "--batch",
        action="store_true",
        help="Process all PDFs in input directory"
    )

    parser.add_argument(
        "--workers",
        type=int,
        default=None,
        help="Number of worker threads for page-level parallelism (default: CPU count)"
    )

    parser.add_argument(
        "--no-preprocess",
        action="store_true",
        help="Disable preprocessing for scanned PDFs"
    )

    parser.add_argument(
        "--dpi",
        type=int,
        default=300,
        help="DPI for PDF to image conversion (default: 300)"
    )
    
    parser.add_argument(
        "--no-images",
        action="store_true",
        help="Disable image extraction and embedding"
    )
    
    parser.add_argument(
        "--no-layout",
        action="store_true",
        help="Disable layout analysis (headers/footers/structure preservation)"
    )
    
    parser.add_argument(
        "--no-tables",
        action="store_true",
        help="Disable table extraction and processing"
    )
    
    parser.add_argument(
        "--easydataset",
        action="store_true",
        help="Generate EasyDataset format output for post-processing"
    )
    
    parser.add_argument(
        "--llm-correction",
        action="store_true",
        help="Enable FREE spelling correction for OCR output (no API costs!)"
    )

    parser.add_argument(
        "--no-translate-vi",
        action="store_true",
        help="Disable the default scientific English-to-Vietnamese translation"
    )

    parser.add_argument(
        "--translation-model",
        default=None,
        help="OpenAI-compatible translation model (default: config/env gemini-3.5-flash)"
    )

    parser.add_argument(
        "--translation-base-url",
        default=None,
        help="OpenAI-compatible API base URL (API key is read only from the environment)"
    )
    
    args = parser.parse_args()

    # CLI arguments always take precedence over config values.
    try:
        import config as cfg
        cfg_loaded = True
        print("Loaded configuration from 'config.py'")
    except Exception:
        cfg = None
        cfg_loaded = False

    try:
        from dotenv import load_dotenv

        load_dotenv(Path(__file__).resolve().parent / ".env")
    except ImportError:
        pass

    if cfg_loaded:
        for setting in (
            "TRANSLATION_API_BASE",
            "TRANSLATION_MODEL",
            "TRANSLATION_TIMEOUT_SECONDS",
            "TRANSLATION_CHUNK_CHARS",
            "TRANSLATION_MAX_RETRIES",
            "TRANSLATION_MAX_TOKENS",
        ):
            if hasattr(cfg, setting):
                os.environ.setdefault(setting, str(getattr(cfg, setting)))

    config_translation_enabled = (
        bool(getattr(cfg, "ENABLE_VI_TRANSLATION", True)) if cfg_loaded else True
    )
    translation_enabled = config_translation_enabled and not args.no_translate_vi

    if translation_enabled and not (
        os.getenv("GEMINI_API_KEY") or os.getenv("TRANSLATION_API_KEY")
    ):
        print("Vietnamese translation is enabled, but no API key was found.")
        print("Copy ocr/.env.example to ocr/.env and set GEMINI_API_KEY,")
        print("or run with --no-translate-vi.")
        sys.exit(1)

    # Check dependencies only after the cheap credential preflight.
    if not check_ocr_dependencies(translation_enabled):
        sys.exit(1)

    # Resolve input path: CLI arg > config.INPUT_DIR > error
    input_arg = args.input
    if input_arg is None and cfg_loaded and hasattr(cfg, "INPUT_DIR"):
        input_arg = cfg.INPUT_DIR

    if input_arg is None:
        print("No input path provided. Pass --input or set INPUT_DIR in config.py")
        sys.exit(1)

    input_path = Path(input_arg)
    if not input_path.exists():
        print(f"Input path does not exist: {input_path}")
        sys.exit(1)


    # Apply config defaults where CLI uses defaults
    if cfg_loaded:
        # Override output if user did not supply a custom value
        if args.output == "./output" and hasattr(cfg, "OUTPUT_DIR"):
            args.output = cfg.OUTPUT_DIR

        # Override dpi if user did not supply a custom value
        if args.dpi == 300 and hasattr(cfg, "DPI"):
            args.dpi = cfg.DPI

    translation_model = args.translation_model or os.getenv("TRANSLATION_MODEL") or (
        getattr(cfg, "TRANSLATION_MODEL", None) if cfg_loaded else None
    )
    translation_base_url = (
        args.translation_base_url
        or os.getenv("TRANSLATION_API_BASE")
        or (getattr(cfg, "TRANSLATION_API_BASE", None) if cfg_loaded else None)
    )
    use_marker_for_digital = bool(
        getattr(cfg, "USE_MARKER_FOR_DIGITAL_TRANSLATION", False)
        if cfg_loaded else False
    )

    # Determine enable_preprocessing
    if args.no_preprocess:
        enable_preprocessing = False
    else:
        if cfg_loaded and hasattr(cfg, "ENABLE_PREPROCESSING"):
            enable_preprocessing = bool(cfg.ENABLE_PREPROCESSING)
        else:
            enable_preprocessing = True

    # Determine auto-detect behavior for pipeline
    if args.mode == "auto":
        auto_detect = cfg.AUTO_DETECT if (cfg_loaded and hasattr(cfg, "AUTO_DETECT")) else True
        mode_arg = None
    else:
        auto_detect = False
        mode_arg = args.mode

    # Initialize pipeline
    pipeline = OCRPipeline(
        output_dir=args.output,
        temp_dir="./temp",
        dpi=args.dpi,
        enable_preprocessing=enable_preprocessing,
        auto_detect=auto_detect,
        extract_images=not args.no_images,
        analyze_layout=not args.no_layout,
        extract_tables=not args.no_tables,
        use_llm_correction=args.llm_correction,
        enable_vi_translation=translation_enabled,
        translation_model=translation_model,
        translation_base_url=translation_base_url,
        use_marker_for_digital_translation=use_marker_for_digital,
    )

    mode = None if args.mode == "auto" else args.mode

    try:
        if args.batch or input_path.is_dir():
            print("Batch processing is disabled. Provide a single PDF file as --input.")
            print("   To process multiple files, run the script separately for each PDF.")
            sys.exit(1)

        # Initialize pipeline with worker setting
        pipeline.max_workers = args.workers

        # Single file processing
        print(f"\nProcessing: {input_path.name}")
        output_path = pipeline.process_pdf(input_path, mode=mode)
        print(f"\nSuccess! Output saved to: {output_path}")
        for label, artifact_path in pipeline.output_artifacts.items():
            print(f"   {label}: {artifact_path}")
        
        # Display metrics summary
        print(pipeline.metrics.get_formatted_summary())
        
        # Save metrics to JSON
        metrics_path = pipeline.metrics.save_metrics_json()
        print(f"Metrics saved to: {metrics_path}")
        
        # Display organized images info
        images_dir = pipeline.images_output_dir
        if images_dir.exists():
            image_files = list(images_dir.glob("*.png"))
            print(f"\nExtracted images organized in: {images_dir}/")
            print(f"   Total images: {len(image_files)}")
            
            # Show image index
            index_path = Path(args.output) / "images_index.json"
            if index_path.exists():
                print(f"   Image mapping: {index_path}")
        
        # Generate EasyDataset format if requested
        if args.easydataset:
            print("\nGenerating EasyDataset format...")
            from src.processing.easydataset_processor import EasyDatasetProcessor
            
            processor = EasyDatasetProcessor(chunk_size=512, overlap=50)
            
            # Find the OCR results JSON
            json_path = Path(args.output) / f"{input_path.stem}_ocr_results.json"
            
            if json_path.exists():
                # Process to EasyDataset format
                easydataset_path = Path(args.output) / f"{input_path.stem}_easydataset.json"
                dataset = processor.process_ocr_results(json_path, easydataset_path)
                
                # Export for Q&A generation
                qa_path = Path(args.output) / f"{input_path.stem}_qa.json"
                processor.export_for_qa_generation(dataset, qa_path)
                
                # Export for retrieval
                retrieval_path = Path(args.output) / f"{input_path.stem}_retrieval.json"
                processor.export_for_retrieval(dataset, retrieval_path)
                
                print(f"EasyDataset format: {easydataset_path}")
                print(f"Q&A format: {qa_path}")
                print(f"Retrieval format: {retrieval_path}")
            else:
                print(f"OCR results JSON not found: {json_path}")

    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()

