#!/usr/bin/env python3
"""
Markdown to DOCX conversion service
Provides common functionality for converting Markdown to DOCX format
"""

import os
import re
from pathlib import Path
from tempfile import NamedTemporaryFile, TemporaryDirectory

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt, RGBColor

from ..utils import get_logger
from ..utils.markdown_utils import get_md_text
from ..utils.mermaid_utils import replace_mermaid_with_images, cleanup_temp_images, extract_mermaid_blocks
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
        "font_name_latin": "Times New Roman",
        "font_size": Pt(20),
        "bold": True,
        "color": RGBColor(0, 0, 0),
        "first_line_indent": Pt(0),
        "left_indent": Pt(0),
        "line_spacing": 1.3,
        "space_before": Pt(3),
        "space_after": Pt(3),
    },
    {
        "name": "Heading 2",
        "style_keywords": ["Heading 2"],
        "font_name": "宋体",
        "font_name_latin": "Times New Roman",
        "font_size": Pt(18),
        "bold": True,
        "color": RGBColor(0, 0, 0),
        "first_line_indent": Pt(0),
        "left_indent": Pt(0),
        "line_spacing": 1.3,
        "space_before": Pt(3),
        "space_after": Pt(3),
    },
    {
        "name": "Heading 3",
        "style_keywords": ["Heading 3"],
        "font_name": "宋体",
        "font_name_latin": "Times New Roman",
        "font_size": Pt(16),
        "bold": True,
        "color": RGBColor(0, 0, 0),
        "first_line_indent": Pt(0),
        "left_indent": Pt(0),
        "line_spacing": 1.3,
        "space_before": Pt(3),
        "space_after": Pt(3),
    },
    {
        "name": "Heading 4",
        "style_keywords": ["Heading 4"],
        "font_name": "宋体",
        "font_name_latin": "Times New Roman",
        "font_size": Pt(14),
        "bold": True,
        "color": RGBColor(0, 0, 0),
        "first_line_indent": Pt(0),
        "left_indent": Pt(0),
        "line_spacing": 1.3,
        "space_before": Pt(3),
        "space_after": Pt(3),
    },
    {
        "name": "Heading 5",
        "style_keywords": ["Heading 5"],
        "font_name": "宋体",
        "font_name_latin": "Times New Roman",
        "font_size": Pt(12),
        "bold": True,
        "color": RGBColor(0, 0, 0),
        "first_line_indent": Pt(0),
        "left_indent": Pt(0),
        "line_spacing": 1.3,
        "space_before": Pt(3),
        "space_after": Pt(3),
    },
    {
        "name": "Heading 6",
        "style_keywords": ["Heading 6"],
        "font_name": "宋体",
        "font_name_latin": "Times New Roman",
        "font_size": Pt(12),
        "bold": True,
        "color": RGBColor(0, 0, 0),
        "first_line_indent": Pt(0),
        "left_indent": Pt(0),
        "line_spacing": 1.3,
        "space_before": Pt(3),
        "space_after": Pt(3),
    },
    {
        "name": "Normal",
        "style_keywords": ["Normal"],
        "font_name": "宋体",
        "font_name_latin": "Times New Roman",
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
        "font_name_latin": "Times New Roman",
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
        "name": "Custom List",
        "style_keywords": [],
        "font_name": "宋体",
        "font_name_latin": "Times New Roman",
        "font_size": Pt(12),
        "bold": False,
        "color": RGBColor(0, 0, 0),
        "first_line_indent": Pt(24),
        "left_indent": Pt(0),
        "line_spacing": 1.3,
        "space_before": Pt(0),
        "space_after": Pt(0),
        "is_custom_list": True,
    },
    {
        "name": "Table Text",
        "style_keywords": [],
        "font_name": "宋体",
        "font_name_latin": "Times New Roman",
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
    {
        "name": "Image Paragraph",
        "style_keywords": [],
        "font_name": "宋体",
        "font_name_latin": "Times New Roman",
        "font_size": Pt(12),
        "bold": False,
        "color": RGBColor(0, 0, 0),
        "first_line_indent": Pt(0),
        "left_indent": Pt(0),
        "line_spacing": 1.3,
        "space_before": Pt(0),
        "space_after": Pt(0),
        "alignment": 1,
        "is_image": True,
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
    "Image Paragraph",
    "Custom List",
}

REQUIRED_CHARACTER_STYLES: set[str] = {
    "Default Paragraph Font",
    "Hyperlink",
    "Strong",
    "Emphasis",
}

_NORMAL_CONFIG: dict = next(c for c in STYLE_CONFIGS if c["name"] == "Normal")
_TABLE_CONFIG: dict = next(c for c in STYLE_CONFIGS if c.get("is_table"))
_IMAGE_CONFIG: dict = next(c for c in STYLE_CONFIGS if c.get("is_image"))
_CUSTOM_LIST_CONFIG: dict = next(c for c in STYLE_CONFIGS if c.get("is_custom_list"))


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


def _set_rfonts(rpr_elem, font_name: str, font_name_latin: str | None = None) -> None:
    rFonts = rpr_elem.get_or_add_rFonts()
    # Set Latin fonts (ascii and hAnsi)
    latin_font = font_name_latin if font_name_latin else font_name
    rFonts.set(qn("w:ascii"), latin_font)
    rFonts.set(qn("w:hAnsi"), latin_font)
    # Set East Asian fonts
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
    
    # Custom list style: set explicit indent to override pandoc defaults
    if config.get("is_custom_list"):
        pf.left_indent = config["left_indent"]
        pf.first_line_indent = config["first_line_indent"]
    # List items (w:numPr) manage their own indentation via numbering definition;
    # but we need to override left_indent to control the overall indent level.
    elif config.get("is_list"):
        # For list items, only set left_indent, let numbering handle the rest
        pf.left_indent = config["left_indent"]
        pf.first_line_indent = Pt(0)
    elif not _has_num_pr(paragraph):
        if config["first_line_indent"] and not is_table and _needs_no_indent(paragraph):
            pf.first_line_indent = Pt(0)
        else:
            pf.first_line_indent = config["first_line_indent"]
            pf.left_indent = config["left_indent"]
    if is_table and "alignment" in config:
        pf.alignment = config["alignment"]
    # Image paragraph alignment is handled by style, but ensure it's set
    if config.get("is_image"):
        pf.alignment = 1  # Center alignment

    for run in paragraph.runs:
        run.font.color.rgb = config["color"]
        run.font.name = config.get("font_name_latin", config["font_name"])
        run.font.size = config["font_size"]
        # Preserve explicit bold set by pandoc (e.g. from **..** markdown);
        # only apply the style's bold value when the run has no explicit bold.
        if config["bold"]:
            run.font.bold = True
        elif run.font.bold is not True:
            run.font.bold = config["bold"]
        _set_rfonts(
            run._element.get_or_add_rPr(),
            config["font_name"],
            config.get("font_name_latin"),
        )


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

            style.font.name = config.get("font_name_latin", config["font_name"])
            style.font.size = config["font_size"]
            style.font.bold = config["bold"]
            style.font.color.rgb = config["color"]

            # Use style.element.get_or_add_rPr() to get CT_RPr — the only
            # reliable path; style.font._element may return CT_Style instead.
            rPr = style.element.get_or_add_rPr()
            _set_rfonts(rPr, config["font_name"], config.get("font_name_latin"))

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
    table_text_style = doc.styles["Table Text"]
    image_para_style = doc.styles["Image Paragraph"]
    custom_list_style = doc.styles["Custom List"]
    
    for para in doc.paragraphs:
        try:
            # 检查是否为图片段落，如果是则应用图片样式
            if _has_image(para):
                para.style = image_para_style
                _apply_para_formatting(para, _IMAGE_CONFIG)
            # 检查是否为列表项，如果是则应用自定义列表样式
            elif _has_num_pr(para):
                para.style = custom_list_style
                _apply_para_formatting(para, _CUSTOM_LIST_CONFIG)
            else:
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

        # 设置表格框线为实线（单线）
        # 构建 tblBorders 元素：所有边框使用 single 样式
        tblBorders_xml = f'''<w:tblBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        </w:tblBorders>'''
        tblBorders = parse_xml(tblBorders_xml)
        
        # 移除现有的 tblBorders
        existing_tblBorders = tblPr.find(qn("w:tblBorders"))
        if existing_tblBorders is not None:
            tblPr.remove(existing_tblBorders)
        
        tblPr.append(tblBorders)

        for row in tbl.rows:
            for cell in row.cells:
                # 设置单元格垂直居中对齐
                tc = cell._element
                tcPr = tc.tcPr if tc.tcPr is not None else parse_xml(f'<w:tcPr {nsdecls("w")}/>')
                if tc.tcPr is None:
                    tc.insert(0, tcPr)
                
                # 移除现有的 vAlign 元素
                existing_vAlign = tcPr.find(qn("w:vAlign"))
                if existing_vAlign is not None:
                    tcPr.remove(existing_vAlign)
                
                # 设置垂直居中 (center)
                vAlign = parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>')
                tcPr.append(vAlign)
                
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
    save_mermaid_images: bool = False,
    output_dir: Path | None = None,
) -> None:
    """
    Convert Markdown text to DOCX format.

    Args:
        md_text: Markdown text to convert
        output_path: Path to save the output DOCX file
        template_path: Optional path to DOCX template file
        is_strip_wrapper: Whether to remove code block wrapper if present
        is_enable_toc: Whether to include table of contents in the output
        save_mermaid_images: Whether to save Mermaid images to output directory
        output_dir: Output directory for saving Mermaid images (required if save_mermaid_images is True)

    Raises:
        ValueError: If input processing fails
        Exception: If conversion fails
    """
    processed_md = get_md_text(md_text, is_strip_wrapper=is_strip_wrapper)
    
    # 检查是否有 Mermaid 代码块需要转换
    mermaid_blocks = extract_mermaid_blocks(processed_md)
    
    if mermaid_blocks:
        logger.info(f"检测到 {len(mermaid_blocks)} 个 Mermaid 图表，开始转换...")
        
        # 根据是否保存图片决定使用临时目录还是输出目录
        if save_mermaid_images and output_dir:
            # 创建图片保存目录
            images_dir = output_dir / "mermaid_images"
            images_dir.mkdir(exist_ok=True)
            save_path = images_dir
            logger.info(f"Mermaid 图片将保存到: {images_dir}")
        else:
            # 使用临时目录，转换后删除
            save_path = None
        
        # 创建临时目录存放图片和 Markdown 文件
        with TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            
            # 如果不需要保存图片，图片保存在临时目录
            image_save_path = save_path if save_mermaid_images else temp_path
            
            # 替换 Mermaid 代码块为图片引用（使用 PNG 格式，通过 scale 参数提高清晰度）
            modified_md, generated_images = replace_mermaid_with_images(
                processed_md,
                image_save_path,
                image_format="png",
                timeout=15,  # 增加超时时间，因为大图片需要更长时间
                max_retries=3,
                retry_delay=2,
                scale=3  # 3倍缩放提高清晰度
            )
            
            # 使用修改后的 Markdown（包含图片引用）进行转换
            final_template_path = template_path or get_default_template()
            
            extra_args: list[str] = []
            if final_template_path and final_template_path.exists():
                extra_args.append(f"--reference-doc={final_template_path}")
            if is_enable_toc:
                extra_args.append("--toc")
            
            # 将修改后的 Markdown 写入临时目录中的文件
            temp_md_file = temp_path / "temp.md"
            temp_md_file.write_text(modified_md, encoding="utf-8")
            
            # 添加资源路径，让 Pandoc 能找到图片
            # 如果图片保存在输出目录，需要添加输出目录到资源路径
            resource_paths = [str(temp_path)]
            if save_mermaid_images and save_path:
                resource_paths.append(str(save_path))
            extra_args.append(f"--resource-path={';'.join(resource_paths)}")
            
            try:
                pandoc_convert_file(
                    source_file=str(temp_md_file),
                    input_format="markdown",
                    dest_format="docx",
                    outputfile=str(output_path),
                    extra_args=extra_args,
                )
                # 仅在没有提供自定义模板时应用格式化
                if not template_path:
                    _apply_formatting(output_path)
                else:
                    logger.info(f"Using custom template, skipping formatting: {template_path}")
            finally:
                # 如果不保存图片，清理临时图片
                if not save_mermaid_images:
                    cleanup_temp_images(generated_images)
                    logger.info("已清理临时图片文件")
                else:
                    logger.info(f"已保存 {len(generated_images)} 个 Mermaid 图片到: {save_path}")
    else:
        # 没有 Mermaid 图表，使用原有逻辑
        logger.info("未检测到 Mermaid 图表，使用标准转换流程")
        
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
            # 仅在没有提供自定义模板时应用格式化
            if not template_path:
                _apply_formatting(output_path)
            else:
                logger.info(f"Using custom template, skipping formatting: {template_path}")
        finally:
            os.unlink(tmp_path)


def get_default_template() -> Path | None:
    """Return the built-in DOCX template path, or None if absent."""
    template = Path(__file__).resolve().parent.parent / "assets" / "template" / "docx_template.docx"
    if template.exists():
        return template
    logger.warning(f"Default DOCX template not found at {template}")
    return None
