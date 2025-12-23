"""
File Data Extraction - Main Entry Point
Extracts text from various document formats (PDF, HWP, HWPX, Excel, Word, ZIP)
"""

import argparse
import logging
import sys
from pathlib import Path

from text_extraction import TextExtractor

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


def extract_single_file(file_path: Path, output_path: Path = None) -> dict:
    """Extract text from a single file and optionally save to output."""
    extractor = TextExtractor(data_lake_root=file_path.parent)

    logger.info(f"Extracting text from: {file_path.name}")
    result = extractor.extract_text(file_path)

    if result.get('error'):
        logger.error(f"Extraction failed: {result['error']}")
        return result

    logger.info(f"Extracted {result.get('char_count', 0)} characters, {result.get('word_count', 0)} words")

    # Save to output file if specified
    if output_path and result.get('text'):
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(result['text'], encoding='utf-8')
        logger.info(f"Saved extracted text to: {output_path}")

    return result


def extract_directory(input_dir: Path, output_dir: Path = None) -> dict:
    """Extract text from all supported files in a directory."""
    supported_extensions = {'.pdf', '.hwp', '.hwpx', '.xlsx', '.xls', '.xlsm', '.docx', '.zip'}

    extractor = TextExtractor(data_lake_root=input_dir)

    # Find all supported files
    files = [f for f in input_dir.rglob('*') if f.suffix.lower() in supported_extensions]

    if not files:
        logger.warning(f"No supported files found in {input_dir}")
        return {'processed': 0, 'errors': 0, 'files': []}

    logger.info(f"Found {len(files)} supported files to process")

    results = {
        'processed': 0,
        'errors': 0,
        'files': []
    }

    for idx, file_path in enumerate(files, 1):
        logger.info(f"[{idx}/{len(files)}] Processing: {file_path.name}")

        result = extractor.extract_text(file_path)

        file_result = {
            'file': str(file_path),
            'success': not result.get('error'),
            'chars': result.get('char_count', 0),
            'words': result.get('word_count', 0)
        }

        if result.get('error'):
            file_result['error'] = result['error']
            results['errors'] += 1
            logger.warning(f"  Failed: {result['error']}")
        else:
            results['processed'] += 1
            logger.info(f"  Extracted {result.get('char_count', 0)} chars")

            # Save to output directory if specified
            if output_dir and result.get('text'):
                relative_path = file_path.relative_to(input_dir)
                output_file = output_dir / relative_path.with_suffix('.txt')
                output_file.parent.mkdir(parents=True, exist_ok=True)
                output_file.write_text(result['text'], encoding='utf-8')

        results['files'].append(file_result)

    logger.info("=" * 60)
    logger.info(f"Extraction complete: {results['processed']} succeeded, {results['errors']} failed")

    return results


def main():
    parser = argparse.ArgumentParser(
        description='Extract text from documents (PDF, HWP, HWPX, Excel, Word, ZIP)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python main.py document.pdf                    # Extract and print to console
  python main.py document.pdf -o output.txt     # Extract and save to file
  python main.py ./docs/ -o ./output/           # Process entire directory
  python main.py document.hwp --verbose         # Extract with verbose logging
        """
    )

    parser.add_argument(
        'input',
        type=str,
        help='Input file or directory to process'
    )

    parser.add_argument(
        '-o', '--output',
        type=str,
        default=None,
        help='Output file (for single file) or directory (for directory input)'
    )

    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Enable verbose logging'
    )

    parser.add_argument(
        '-q', '--quiet',
        action='store_true',
        help='Suppress output except errors'
    )

    parser.add_argument(
        '--print-text',
        action='store_true',
        help='Print extracted text to console'
    )

    args = parser.parse_args()

    # Set logging level
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    elif args.quiet:
        logging.getLogger().setLevel(logging.ERROR)

    input_path = Path(args.input)
    output_path = Path(args.output) if args.output else None

    if not input_path.exists():
        logger.error(f"Input path does not exist: {input_path}")
        sys.exit(1)

    if input_path.is_file():
        result = extract_single_file(input_path, output_path)

        if args.print_text and result.get('text'):
            print("\n" + "=" * 60)
            print("EXTRACTED TEXT:")
            print("=" * 60)
            print(result['text'])

        sys.exit(0 if not result.get('error') else 1)

    elif input_path.is_dir():
        results = extract_directory(input_path, output_path)
        sys.exit(0 if results['errors'] == 0 else 1)

    else:
        logger.error(f"Invalid input path: {input_path}")
        sys.exit(1)


if __name__ == '__main__':
    main()
