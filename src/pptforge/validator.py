import os
import zipfile

from lxml import etree

from pptforge.config import RESERVED_TAG_CHARS, _get_tag_filtered_pages
from pptforge.constants import REL_TYPES
from pptforge.extractor import _collect_notes_metadata, _compute_tags, extract_index
from pptforge.models import ProposalConfig


class ValidationError(Exception):
    def __init__(self, errors: list[str]):
        self.errors = errors

    def __str__(self):
        return "\n".join(f"✗ {e}" for e in self.errors)


def _get_slide_count(src_zip: zipfile.ZipFile) -> int:
    rels_xml = src_zip.read("ppt/_rels/presentation.xml.rels")
    root = etree.fromstring(rels_xml)
    slide_type = REL_TYPES["slide"]
    return sum(1 for rel in root if rel.get("Type") == slide_type)


def validate_static(proposal: ProposalConfig, force: bool = False) -> None:
    errors = []

    output_path = os.path.realpath(proposal.output_path)
    output_dir = os.path.dirname(output_path) or "."
    if not os.path.isdir(output_dir):
        try:
            os.makedirs(output_dir, exist_ok=True)
        except OSError:
            errors.append(f"输出目录无法创建：{output_dir}")

    if os.path.exists(output_path) and not force:
        errors.append(
            f"输出文件已存在：{proposal.output_path}（使用 --force 覆盖）"
        )

    for src in proposal.sources:
        if not os.path.exists(src.pptx_path):
            errors.append(f"文件不存在：\"{src.pptx_path}\"")
        elif not src.pptx_path.lower().endswith(".pptx"):
            errors.append(f"文件不是 .pptx 格式：\"{src.pptx_path}\"")

        if src.tags:
            for tag in src.tags:
                if not tag:
                    errors.append(f"tag 名为空：\"{src.pptx_path}\"")
                for char in RESERVED_TAG_CHARS:
                    if char in tag:
                        errors.append(
                            f'tag 名包含保留字符 "{char}"：{tag}'
                        )
                        break

        src_path = os.path.realpath(src.pptx_path)
        if src_path == output_path:
            errors.append(
                f"输出文件路径与源文件冲突：\"{proposal.output_path}\""
            )

    if errors:
        raise ValidationError(errors)


def validate_content(proposal: ProposalConfig) -> list[str]:
    errors = []
    warnings = []

    for src in proposal.sources:
        try:
            with zipfile.ZipFile(src.pptx_path, "r") as z:
                if "ppt/presentation.xml" not in z.namelist():
                    errors.append(f"不是有效的 PPTX 文件：\"{src.pptx_path}\"")
                    continue

                slide_count = _get_slide_count(z)
                tag_errors = validate_tags_in_pptx(src.pptx_path)
                if tag_errors:
                    for err in tag_errors:
                        errors.append(f"\"{src.pptx_path}\"：{err}")
                    continue

                index = None
                if src.tags:
                    try:
                        index = extract_index(src.pptx_path)
                    except Exception as e:
                        errors.append(
                            f"无法读取 \"{src.pptx_path}\"：{e}"
                        )
                        continue
                    for tag in src.tags:
                        if tag not in index.tags:
                            errors.append(
                                f"tag \"{tag}\" 不在 "
                                f"\"{src.pptx_path}\" 中"
                            )

                if errors:
                    continue

                if src.tags:
                    base_count = len(_get_tag_filtered_pages(index, src.tag_groups)) if index else 0
                else:
                    base_count = slide_count

                if src.pages is not None:
                    for spec in src.pages:
                        if spec > 0:
                            if spec > base_count:
                                errors.append(
                                    f"页码越界：\"{src.pptx_path}\" "
                                    f"共 {base_count} 页，请求了第 {spec} 页"
                                )
                        else:
                            if abs(spec) > base_count:
                                errors.append(
                                    f"页码越界：\"{src.pptx_path}\" "
                                    f"共 {base_count} 页，请求了第 {spec} 页"
                                )
        except (zipfile.BadZipFile, Exception) as e:
            errors.append(f"无法打开文件 \"{src.pptx_path}\"：{e}")

    if errors:
        raise ValidationError(errors)

    return warnings


def validate_tags_in_pptx(pptx_path: str) -> list[str]:
    try:
        per_page_notes = _collect_notes_metadata(pptx_path)
    except Exception as e:
        return [f"无法读取 \"{pptx_path}\"：{e}"]

    tag_name_errors = _validate_tag_names(per_page_notes)
    _, _, compute_errors = _compute_tags(per_page_notes)
    return tag_name_errors + compute_errors


def _validate_tag_names(per_page_notes: dict[int, dict]) -> list[str]:
    errors: list[str] = []
    seen: set[tuple[int, str]] = set()

    for page_num, notes in per_page_notes.items():
        for field in ("tags", "tag-start", "tag-end"):
            for tag in notes.get(field, []):
                key = (page_num, tag)
                if key in seen:
                    continue
                seen.add(key)
                for char in RESERVED_TAG_CHARS:
                    if char in tag:
                        errors.append(
                            f'第 {page_num} 页：tag 名包含保留字符 "{char}"：{tag}'
                        )
                        break

    return errors
