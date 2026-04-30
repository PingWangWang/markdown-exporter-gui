#!/usr/bin/env python3
"""
MdToPptx service
"""

import os
from pathlib import Path
from tempfile import NamedTemporaryFile, TemporaryDirectory

from ..utils import get_logger
from ..utils.markdown_utils import get_md_text
from ..utils.mermaid_utils import replace_mermaid_with_images, cleanup_temp_images, extract_mermaid_blocks
from ..utils.pandoc_utils import pandoc_convert_file

logger = get_logger(__name__)


def get_default_template() -> Path | None:
    """
    Get the default PPTX template path

    Returns:
        Optional[Path]: Path to default template if it exists, None otherwise
    """
    script_dir = Path(__file__).resolve().parent.parent
    default_template = script_dir / "assets" / "template" / "pptx_template.pptx"
    if default_template.exists():
        return default_template
    else:
        logger.warning(f"Default PPTX template not found at {default_template}")
        return None


def convert_md_to_pptx(
    md_text: str, output_path: Path, template_path: Path | None = None, is_strip_wrapper: bool = False
) -> Path:
    """
    Convert Markdown text to PPTX format using pandoc
    Args:
        md_text: Markdown text to convert
        output_path: Path to save the output PPTX file
        template_path: Path to PPTX template file (optional)
        is_strip_wrapper: Whether to remove code block wrapper if present
    Returns:
        Path to the created PPTX file
    Raises:
        ValueError: If input processing fails
        Exception: If conversion fails
    """
    # Process Markdown text
    processed_md = get_md_text(md_text, is_strip_wrapper=is_strip_wrapper)
    
    # 检查是否有 Mermaid 代码块需要转换
    mermaid_blocks = extract_mermaid_blocks(processed_md)
    
    if mermaid_blocks:
        logger.info(f"检测到 {len(mermaid_blocks)} 个 Mermaid 图表，开始转换...")
        
        # 创建临时目录存放图片和 Markdown 文件
        with TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            
            # 替换 Mermaid 代码块为图片引用（图片会保存在 temp_path）
            modified_md, generated_images = replace_mermaid_with_images(
                processed_md,
                temp_path,
                image_format="png",
                timeout=10,
                max_retries=3,
                retry_delay=2
            )
            
            # 使用修改后的 Markdown（包含图片引用）进行转换
            final_template_path = template_path
            if not final_template_path:
                final_template_path = get_default_template()
            
            extra_args = []
            if final_template_path and final_template_path.exists():
                extra_args.append(f"--reference-doc={final_template_path}")
            
            # 将修改后的 Markdown 写入临时目录中的文件
            temp_md_file = temp_path / "temp.md"
            temp_md_file.write_text(modified_md, encoding="utf-8")
            
            # 添加资源路径，让 Pandoc 能找到图片
            extra_args.append(f"--resource-path={temp_path}")
            
            try:
                logger.info(f"Converting Markdown to PPTX: {output_path}")
                pandoc_convert_file(
                    source_file=str(temp_md_file),
                    input_format="markdown",
                    dest_format="pptx",
                    outputfile=str(output_path),
                    extra_args=extra_args,
                )
                logger.info(f"Successfully created PPTX: {output_path}")
                return output_path
            finally:
                # 清理临时图片
                cleanup_temp_images(generated_images)
                logger.info("已清理临时图片文件")
    else:
        # 没有 Mermaid 图表，使用原有逻辑
        logger.info("未检测到 Mermaid 图表，使用标准转换流程")
        
        # Determine template file
        final_template_path = template_path
        if not final_template_path:
            final_template_path = get_default_template()
        
        extra_args = []
        if final_template_path and final_template_path.exists():
            extra_args.append(f"--reference-doc={final_template_path}")
        
        with NamedTemporaryFile(suffix=".md", delete=False, mode="w", encoding="utf-8") as temp_md_file:
            temp_md_file.write(processed_md)
            temp_md_file_path = temp_md_file.name
        
        try:
            logger.info(f"Converting Markdown to PPTX: {output_path}")
            pandoc_convert_file(
                source_file=temp_md_file_path,
                input_format="markdown",
                dest_format="pptx",
                outputfile=str(output_path),
                extra_args=extra_args,
            )
            logger.info(f"Successfully created PPTX: {output_path}")
            return output_path
        finally:
            if os.path.exists(temp_md_file_path):
                os.unlink(temp_md_file_path)
