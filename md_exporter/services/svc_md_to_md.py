#!/usr/bin/env python3
"""
MdToMd service
"""

from pathlib import Path

from ..utils import get_logger
from ..utils.markdown_utils import get_md_text

logger = get_logger(__name__)


def convert_md_to_md(md_text: str, output_path: Path, is_strip_wrapper: bool = False) -> Path:
    """
    Convert Markdown text to .md file
    Args:
        md_text: Markdown text to convert
        output_path: Path to save the output MD file
        is_strip_wrapper: Whether to remove code block wrapper if present
    Returns:
        Path to the created MD file
    Raises:
        ValueError: If input processing fails
        Exception: If conversion fails
    """
    # Process Markdown text
    processed_md = get_md_text(md_text, is_strip_wrapper=is_strip_wrapper)

    logger.info(f"Processing Markdown file: {output_path}")

    # Write to output file
    try:
        output_path.write_text(processed_md, encoding="utf-8")
        logger.info(f"Successfully created MD file: {output_path}")
        return output_path
    except Exception as e:
        logger.error(f"Failed to save MD file: {e}")
        raise Exception(f"Failed to save MD file: {e}")
