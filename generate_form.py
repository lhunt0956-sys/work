"""Generate a 工作任务联系单 (Work Task Contact Form) as a Word document.

This is the reference application for this repository: it builds a structured
DOCX form from a small data dictionary using python-docx. Run it directly to
produce an example document under ``output/``.
"""

from __future__ import annotations

import argparse
import datetime as dt
from pathlib import Path

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

# Default field values for the example document. Override any of these on the
# command line (see ``main``) to generate a different form.
DEFAULT_FIELDS: dict[str, str] = {
    "编号": "RW-2024-001",
    "项目名称": "示例项目",
    "提出部门": "工程部",
    "承办部门": "技术部",
    "联系人": "张三",
    "联系电话": "138-0000-0000",
    "日期": dt.date.today().isoformat(),
    "任务描述": "请协助完成本期开发环境的搭建与验证工作，并反馈处理结果。",
    "处理意见": "",
    "备注": "",
}


def _set_cell_text(cell, text: str, *, bold: bool = False) -> None:
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.size = Pt(11)


def build_form(fields: dict[str, str]) -> Document:
    """Build a Word document for the given field mapping."""
    document = Document()

    title = document.add_heading("工作任务联系单", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Header rows: label/value pairs laid out two columns wide.
    header_pairs = [
        ("编号", "项目名称"),
        ("提出部门", "承办部门"),
        ("联系人", "联系电话"),
    ]

    table = document.add_table(rows=0, cols=4)
    table.style = "Table Grid"

    for left_key, right_key in header_pairs:
        row = table.add_row().cells
        _set_cell_text(row[0], left_key, bold=True)
        _set_cell_text(row[1], fields.get(left_key, ""))
        _set_cell_text(row[2], right_key, bold=True)
        _set_cell_text(row[3], fields.get(right_key, ""))

    # Full-width date row.
    date_row = table.add_row().cells
    _set_cell_text(date_row[0], "日期", bold=True)
    date_cell = date_row[1]
    date_cell.merge(date_row[3])
    _set_cell_text(date_cell, fields.get("日期", ""))

    # Full-width free-text blocks.
    for key in ("任务描述", "处理意见", "备注"):
        label_row = table.add_row().cells
        _set_cell_text(label_row[0], key, bold=True)
        value_cell = label_row[1]
        value_cell.merge(label_row[3])
        _set_cell_text(value_cell, fields.get(key, ""))

    return document


def generate(fields: dict[str, str], output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document = build_form(fields)
    document.save(output_path)
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("output") / "工作任务联系单.docx",
        help="Path to write the generated .docx file.",
    )
    for key in DEFAULT_FIELDS:
        parser.add_argument(f"--{key}", dest=key, help=f"Value for the 「{key}」 field.")

    args = parser.parse_args()

    fields = dict(DEFAULT_FIELDS)
    for key in DEFAULT_FIELDS:
        value = getattr(args, key)
        if value is not None:
            fields[key] = value

    saved = generate(fields, args.output)
    print(f"Wrote {saved.resolve()}")


if __name__ == "__main__":
    main()
