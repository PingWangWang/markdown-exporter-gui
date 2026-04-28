#!/usr/bin/env python3
"""
MdToLatex service
"""

from pathlib import Path

from ..utils import get_logger
from ..utils.markdown_utils import get_md_text
from ..utils.table_utils import parse_md_to_tables

logger = get_logger(__name__)


def convert_md_to_latex(md_text: str, output_path: Path, is_strip_wrapper: bool = False) -> list[Path]:
    """
    Convert Markdown tables to LaTeX format
    Args:
        md_text: Markdown text to convert
        output_path: Path to save the output LaTeX file
        is_strip_wrapper: Whether to remove code block wrapper if present
    Returns:
        List of paths to the created LaTeX files
    Raises:
        ValueError: If input processing fails
        Exception: If conversion fails
    """
    # Process Markdown text
    processed_md = get_md_text(md_text, is_strip_wrapper=is_strip_wrapper)

    logger.info(f"Converting Markdown tables to LaTeX")

    # Parse Markdown tables
    tables = parse_md_to_tables(processed_md)

    # Convert to LaTeX
    created_files = []
    for i, table in enumerate(tables):
        table_latex_str = table.to_latex(index=False, bold_rows=True)
        doc_latex_str = (
            "\\documentclass[]{article}\n"
            + "\\usepackage{booktabs}\n"
            + "\\begin{document}\n"
            + "\n"
            + table_latex_str
            + "\n"
            + "\\end{document}\n"
        )
        result_file_bytes = doc_latex_str.encode("utf-8")

        # Determine output file name
        if len(tables) > 1:
            output_file = output_path.parent / f"{output_path.stem}_{i + 1}.tex"
        else:
            output_file = output_path

        # Write to file
        output_file.write_bytes(result_file_bytes)
        created_files.append(output_file)

    logger.info(f"Successfully created {len(created_files)} LaTeX file(s): {[str(f) for f in created_files]}")
    return created_files
