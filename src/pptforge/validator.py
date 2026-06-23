import os
import zipfile
from lxml import etree

from pptforge.models import ProposalConfig, SlideSource
from pptforge.constants import REL_TYPES


class ValidationError(Exception):
    def __init__(self, errors: list[str]):
        self.errors = errors

    def __str__(self):
        return "\n".join(f"✗ {e}" for e in self.errors)


def validate_static(proposal: ProposalConfig, force: bool = False) -> None:
    errors = []

    output_dir = os.path.dirname(proposal.output_path) or "."
    if not os.path.isdir(output_dir):
        try:
            os.makedirs(output_dir, exist_ok=True)
        except OSError:
            errors.append(f"输出目录无法创建：{output_dir}")

    if os.path.exists(proposal.output_path) and not force:
        errors.append(
            f"输出文件已存在：{proposal.output_path}（使用 --force 覆盖）"
        )

    for src in proposal.sources:
        if not os.path.exists(src.pptx_path):
            errors.append(f"文件不存在：{src.pptx_path}")
        elif not src.pptx_path.lower().endswith(".pptx"):
            errors.append(f"文件不是 .pptx 格式：{src.pptx_path}")

        for page in src.pages:
            if page < 1:
                errors.append(
                    f"页码必须为正整数：{src.pptx_path} 请求了第 {page} 页"
                )

    if errors:
        raise ValidationError(errors)


def _get_slide_count(src_zip: zipfile.ZipFile) -> int:
    rels_xml = src_zip.read("ppt/_rels/presentation.xml.rels")
    root = etree.fromstring(rels_xml)
    slide_type = REL_TYPES["slide"]
    return sum(1 for rel in root if rel.get("Type") == slide_type)


def validate_content(proposal: ProposalConfig) -> list[str]:
    errors = []
    warnings = []

    for src in proposal.sources:
        try:
            with zipfile.ZipFile(src.pptx_path, "r") as z:
                if "ppt/presentation.xml" not in z.namelist():
                    errors.append(f"不是有效的 PPTX 文件：{src.pptx_path}")
                    continue

                slide_count = _get_slide_count(z)
                for page in src.pages:
                    if page > slide_count:
                        errors.append(
                            f"页码越界：{os.path.basename(src.pptx_path)} "
                            f"共 {slide_count} 页，请求了第 {page} 页"
                        )
        except (zipfile.BadZipFile, Exception) as e:
            errors.append(f"无法打开文件 {src.pptx_path}：{e}")

    if errors:
        raise ValidationError(errors)

    return warnings
