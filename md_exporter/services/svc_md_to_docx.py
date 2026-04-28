#!/usr/bin/env python3
"""
Markdown to DOCX conversion service
Provides common functionality for converting Markdown to DOCX format
"""

import os
import re
from pathlib import Path
from tempfile import NamedTemporaryFile

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt, RGBColor

from ..utils import get_logger
from ..utils.markdown_utils import get_md_text
from ..utils.pandoc_utils import pandoc_convert_file

logger = get_logger(__name__)

# ---------------------------------------------------------------------------
# Style configuration table
# ---------------------------------------------------------------------------
STYLE_CONFIGS: list[dict] = [
    {
        "name": "Heading 1",
        "style_keywords": ["Heading 1"],
        "font_name": "宋体",
        "font_size": Pt(20),
        "bold": True,
        "color": RGBColor(0, 0, 0),
        "first_line_indent": Pt(0),
        "left_indent": Pt(0),
        "line_spacing": 1.3,
        "space_before": Pt(6),
        "space_after": Pt(6),
    },
    {
        "name": "Heading 2",
        "style_keywords": ["Heading 2"],
        "font_name": "宋体",
        "font_size": Pt(18),
        "bold": True,
        "color": RGBColor(0, 0, 0),
        "first_line_indent": Pt(0),
        "left_indent": Pt(0),
        "line_spacing": 1.3,
        "space_before": Pt(6),
        "space_after": Pt(6),
    },
    {
        "name": "Heading 3",
        "style_keywords": ["Heading 3"],
        "font_name": "宋体",
        "font_size": Pt(16),
        "bold": True,
        "color": RGBColor(0, 0, 0),
        "first_line_indent": Pt(0),
        "left_indent": Pt(0),
        "line_spacing": 1.3,
        "space_before": Pt(6),
        "space_after": Pt(6),
    },
    {
        "name": "Heading 4",
        "style_keywords": ["Heading 4"],
        "font_name": "宋体",
        "font_size": Pt(14),
        "bold": True,
        "color": RGBColor(0, 0, 0),
        "first_line_indent": Pt(0),
        "left_indent": Pt(0),
        "line_spacing": 1.3,
        "space_before": Pt(6),
        "space_after": Pt(6),
    },
    {
        "name": "Heading 5",
        "style_keywords": ["Heading 5"],
        "font_name": "宋体",
        "font_size": Pt(12),
        "bold": True,
        "color": RGBColor(0, 0, 0),
        "first_line_indent": Pt(0),
        "left_indent": Pt(0),
        "line_spacing": 1.3,
        "space_before": Pt(6),
        "space_after": Pt(6),
    },
    {
        "name": "Heading 6",
        "style_keywords": ["Heading 6"],
        "font_name": "宋体",
        "font_size": Pt(12),
        "bold": True,
        "color": RGBColor(0, 0, 0),
        "first_line_indent": Pt(0),
        "left_indent": Pt(0),
        "line_spacing": 1.3,
        "space_before": Pt(6),
        "space_after": Pt(6),
    },
    {
        "name": "Normal",
        "style_keywords": ["Normal"],
        "font_name": "宋体",
        "font_size": Pt(12),
        "bold": False,
        "color": RGBColor(0, 0, 0),
        "first_line_indent": Pt(24),
        "left_indent": Pt(0),
        "line_spacing": 1.3,
        "space_before": Pt(0),
        "space_after": Pt(0),
    },
    {
        "name": "List Paragraph",
        "style_keywords": ["List Paragraph", "List"],
        "font_name": "宋体",
        "font_size": Pt(12),
        "bold": False,
        "color": RGBColor(0, 0, 0),
        "first_line_indent": Pt(0),
        "left_indent": Pt(0),
        "line_spacing": 1.3,
        "space_before": Pt(0),
        "space_after": Pt(0),
        "is_list": True,
    },
    {
        "name": "Table Text",
        "style_keywords": [],
        "font_name": "宋体",
        "font_size": Pt(12),
        "bold": False,
        "color": RGBColor(0, 0, 0),
        "first_line_indent": Pt(0),
        "left_indent": Pt(0),
        "line_spacing": 1.0,
        "space_before": Pt(0),
        "space_after": Pt(0),
        "alignment": 1,
        "is_table": True,
    },
]

REQUIRED_PARAGRAPH_STYLES: set[str] = {
    "Normal",
    "Heading 1",
    "Heading 2",
    "Heading 3",
    "Heading 4",
    "Heading 5",
    "Heading 6",
    "Table Text",
    "List Paragraph",
}

REQUIRED_CHARACTER_STYLES: set[str] = {
    "Default Paragraph Font",
    "Hyperlink",
    "Strong",
    "Emphasis",
}

_NORMAL_CONFIG: dict = next(c for c in STYLE_CONFIGS if c["name"] == "Normal")
_TABLE_CONFIG: dict = next(c for c in STYLE_CONFIGS if c.get("is_table"))


# ---------------------------------------------------------------------------
# Module-level helpers
# ---------------------------------------------------------------------------


def _get_config_for_style(style_name: str) -> dict:
    for config in STYLE_CONFIGS:
        if config.get("is_table"):
            continue
        for keyword in config["style_keywords"]:
            if keyword in style_name:
                return config
    return _NORMAL_CONFIG


def _set_rfonts(rpr_elem, font_name: str) -> None:
    rFonts = rpr_elem.get_or_add_rFonts()
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:eastAsia"), font_name)
    rFonts.set(qn("w:cs"), font_name)
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:themeEastAsia", "w:cstheme"):
        rFonts.attrib.pop(qn(attr), None)


# Matches paragraphs that start with a number/letter list marker and should not be indented
_NO_INDENT_PATTERN = re.compile(r"^\s*(\d+[.、）)）]|[（(]\d+[）)]|[一二三四五六七八九十百]+[、.]|[a-zA-Z][.)])")


def _needs_no_indent(paragraph) -> bool:
    """Return True when a Normal paragraph should NOT receive first-line indent.

    Only paragraphs that start with a numbered / lettered list marker are
    suppressed; bold-lead paragraphs follow normal indentation rules.
    """
    return bool(_NO_INDENT_PATTERN.match(paragraph.text))


def _has_num_pr(paragraph) -> bool:
    """Return True if the paragraph contains a w:numPr element (pandoc list item)."""
    pPr = paragraph._p.find(qn("w:pPr"))
    return pPr is not None and pPr.find(qn("w:numPr")) is not None


def _has_image(paragraph) -> bool:
    """Return True if the paragraph contains an image."""
    # Check for drawing elements (inline images)
    for child in paragraph._element.iter():
        if child.tag.endswith("drawing") or child.tag.endswith("pic"):
            return True
    return False


def _apply_para_formatting(paragraph, config: dict, is_table: bool = False) -> None:
    pf = paragraph.paragraph_format
    pf.line_spacing = config["line_spacing"]
    pf.space_before = config["space_before"]
    pf.space_after = config["space_after"]
    # List items (w:numPr) manage their own indentation via numbering definition;
    # overriding first_line_indent / left_indent would break bullet alignment.
    if not _has_num_pr(paragraph):
        # Image paragraphs should not have first-line indent and should be centered
        if _has_image(paragraph):
            pf.first_line_indent = Pt(0)
            pf.left_indent = Pt(0)
            pf.alignment = 1  # Center alignment
        elif config["first_line_indent"] and not is_table and _needs_no_indent(paragraph):
            pf.first_line_indent = Pt(0)
        else:
            pf.first_line_indent = config["first_line_indent"]
            pf.left_indent = config["left_indent"]
    if is_table and "alignment" in config:
        pf.alignment = config["alignment"]

    for run in paragraph.runs:
        run.font.color.rgb = config["color"]
        run.font.name = config["font_name"]
        run.font.size = config["font_size"]
        # Preserve explicit bold set by pandoc (e.g. from **..** markdown);
        # only apply the style's bold value when the run has no explicit bold.
        if config["bold"]:
            run.font.bold = True
        elif run.font.bold is not True:
            run.font.bold = config["bold"]
        _set_rfonts(run._element.get_or_add_rPr(), config["font_name"])


# ---------------------------------------------------------------------------
# Three-step formatting pipeline
# ---------------------------------------------------------------------------


def _step1_delete_styles(doc) -> None:
    normal_style = doc.styles["Normal"]
    existing_names = {s.name for s in doc.styles}
    # List Paragraph may not exist yet before _step2 creates it; create a minimal
    # placeholder now so that list sub-styles can be reassigned to it.
    if "List Paragraph" not in existing_names:
        list_style = doc.styles.add_style("List Paragraph", 1)
        list_style.base_style = normal_style
    else:
        list_style = doc.styles["List Paragraph"]

    styles_to_delete = [
        s
        for s in doc.styles
        if (s.type == 1 and s.name not in REQUIRED_PARAGRAPH_STYLES)
        or (s.type == 2 and s.name not in REQUIRED_CHARACTER_STYLES)
    ]

    for style_obj in styles_to_delete:
        try:
            if style_obj.type == 1:
                # Pandoc generates "List Paragraph 1/2/..." for nested lists;
                # reassign those to List Paragraph so they keep bullet formatting.
                is_list_sub = "List" in style_obj.name
                fallback_style = list_style if is_list_sub else normal_style
                for para in doc.paragraphs:
                    if para.style.name == style_obj.name:
                        para.style = fallback_style
                for tbl in doc.tables:
                    for row in tbl.rows:
                        for cell in row.cells:
                            for para in cell.paragraphs:
                                if para.style.name == style_obj.name:
                                    para.style = fallback_style
            else:
                style_id = style_obj._element.get(qn("w:styleId"))
                all_paras = list(doc.paragraphs)
                for tbl in doc.tables:
                    for row in tbl.rows:
                        for cell in row.cells:
                            all_paras.extend(cell.paragraphs)
                for para in all_paras:
                    for run in para.runs:
                        rPr = run._element.find(qn("w:rPr"))
                        if rPr is not None:
                            rStyle = rPr.find(qn("w:rStyle"))
                            if rStyle is not None and rStyle.get(qn("w:val")) == style_id:
                                rPr.remove(rStyle)

            style_obj._element.getparent().remove(style_obj._element)
            logger.info(f"Deleted style: {style_obj.name}")
        except Exception as exc:
            logger.warning(f"Failed to delete style '{style_obj.name}': {exc}")


def _step2_create_styles(doc) -> None:
    for config in STYLE_CONFIGS:
        name = config["name"]
        try:
            existing_names = {s.name for s in doc.styles}
            style = doc.styles[name] if name in existing_names else doc.styles.add_style(name, 1)

            style.font.name = config["font_name"]
            style.font.size = config["font_size"]
            style.font.bold = config["bold"]
            style.font.color.rgb = config["color"]

            # Use style.element.get_or_add_rPr() to get CT_RPr — the only
            # reliable path; style.font._element may return CT_Style instead.
            rPr = style.element.get_or_add_rPr()
            _set_rfonts(rPr, config["font_name"])

            # Strip theme-color overrides
            color_elem = rPr.find(qn("w:color"))
            if color_elem is not None:
                for attr in ("w:themeColor", "w:themeShade", "w:themeTint"):
                    color_elem.attrib.pop(qn(attr), None)

            pf = style.paragraph_format
            pf.line_spacing = config["line_spacing"]
            pf.first_line_indent = config["first_line_indent"]
            if config["left_indent"] is not None:
                pf.left_indent = config["left_indent"]
            pf.space_before = config["space_before"]
            pf.space_after = config["space_after"]
        except Exception as exc:
            logger.warning(f"Failed to create/update style '{name}': {exc}")


def _step3_apply_to_content(doc) -> None:
    for para in doc.paragraphs:
        try:
            style_name = para.style.name if para.style else "Normal"
            _apply_para_formatting(para, _get_config_for_style(style_name))
        except Exception as exc:
            logger.warning(f"Failed to format paragraph: {exc}")

    table_text_style = doc.styles["Table Text"]
    for tbl in doc.tables:
        # Auto-fit table to window width (100% of page width)
        # Set table width to 100% using w:tblW element
        tblPr = tbl._element.tblPr
        if tblPr is None:
            tblPr = parse_xml(f"<w:tblPr {nsdecls('w')}/>")
            tbl._element.insert(0, tblPr)

        # Remove existing tblW if present
        existing_tblW = tblPr.find(qn("w:tblW"))
        if existing_tblW is not None:
            tblPr.remove(existing_tblW)

        # Create new tblW element for 100% width
        tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:type="pct" w:w="5000"/>')
        tblPr.append(tblW)

        for row in tbl.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    try:
                        para.style = table_text_style
                        _apply_para_formatting(para, _TABLE_CONFIG, is_table=True)
                    except Exception as exc:
                        logger.warning(f"Failed to format table cell: {exc}")


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def _apply_formatting(docx_path: Path) -> None:
    """
    Three-step formatting pipeline:
      1. Delete all styles outside the required sets.
      2. Create / refresh styles from STYLE_CONFIGS with explicit fonts.
      3. Apply styles and run-level formatting to every paragraph.
    """
    doc = Document(docx_path)
    _step1_delete_styles(doc)
    _step2_create_styles(doc)
    _step3_apply_to_content(doc)
    doc.save(docx_path)
    logger.info(f"Applied formatting to {docx_path}")


def convert_md_to_docx(
    md_text: str,
    output_path: Path,
    template_path: Path | None = None,
    is_strip_wrapper: bool = False,
    is_enable_toc: bool = False,
) -> None:
    """
    Convert Markdown text to DOCX format.

    Args:
        md_text: Markdown text to convert
        output_path: Path to save the output DOCX file
        template_path: Optional path to DOCX template file
        is_strip_wrapper: Whether to remove code block wrapper if present
        is_enable_toc: Whether to include table of contents in the output

    Raises:
        ValueError: If input processing fails
        Exception: If conversion fails
    """
    processed_md = get_md_text(md_text, is_strip_wrapper=is_strip_wrapper)

    final_template_path = template_path or get_default_template()

    extra_args: list[str] = []
    if final_template_path and final_template_path.exists():
        extra_args.append(f"--reference-doc={final_template_path}")
    if is_enable_toc:
        extra_args.append("--toc")

    with NamedTemporaryFile(suffix=".md", delete=False, mode="w", encoding="utf-8") as tmp:
        tmp.write(processed_md)
        tmp_path = tmp.name

    try:
        pandoc_convert_file(
            source_file=tmp_path,
            input_format="markdown",
            dest_format="docx",
            outputfile=str(output_path),
            extra_args=extra_args,
        )
        _apply_formatting(output_path)
    finally:
        os.unlink(tmp_path)


def get_default_template() -> Path | None:
    """Return the built-in DOCX template path, or None if absent."""
    template = Path(__file__).resolve().parent.parent / "assets" / "template" / "docx_template.docx"
    if template.exists():
        return template
    logger.warning(f"Default DOCX template not found at {template}")
    return None
