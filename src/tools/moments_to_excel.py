"""Convert an exported moments.json into an Excel (or CSV) table.

Usage:
    python src/tools/moments_to_excel.py <output_dir> [--output moments.xlsx]

The script reads ``moments.json`` produced by wxMoments and writes a table
with one row per post. Columns: 时间, UID, 手机号, 朋友圈正文, 图片路径, 商品分类ID.


商品分类ID is derived from the remark prefix: 服装→23, 百货→21, 鞋子→26,
家纺→27. Each output file holds at most 500 rows; extra rows spill into
moments_2.xlsx, moments_3.xlsx, ...

It prefers ``openpyxl`` for .xlsx output and falls back to CSV when no Excel
library is available.
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from pathlib import Path

_PHONE_RE = re.compile(r"1[3-9]\d{9}")


def _strip_phone(text: str) -> str:
    """去除文本中的手机号"""
    return _PHONE_RE.sub("", str(text or "")).strip()

MAX_ROWS_PER_FILE = 500

CATEGORY_ID_MAP = {
    "服装": "23",
    "百货": "21",
    "鞋子": "26",
    "家纺": "27",
}

# 商品分类统一复用 clean_moments 的 classify_product（含复合词消歧与备注优先规则）
try:
    from tools.clean_moments import classify_product as _classify_product
except ImportError:
    import importlib.util as _ilu
    _spec = _ilu.spec_from_file_location(
        "clean_moments", Path(__file__).parent / "clean_moments.py"
    )
    _mod = _ilu.module_from_spec(_spec)
    _spec.loader.exec_module(_mod)
    _classify_product = _mod.classify_product

COLUMNS = [
    ("时间", "time_text"),
    ("UID", "uid"),
    ("手机号", "phone"),
    ("朋友圈正文", "content"),
    ("图片路径", "images"),
    ("商品分类ID", "category_id"),
]


def category_id_for(record: dict) -> str:
    """清洗后的记录直接用 _category_id；未清洗的复用 classify_product 统一分类
    （内容唯一命中直接用；多分类歧义时商家备注优先；无命中按备注兜底）。"""
    # 清洗后的记录可能已带 _category_id
    if record.get("_category_id"):
        return str(record["_category_id"])
    category = _classify_product(record)
    return CATEGORY_ID_MAP.get(category, "")


def load_records(moments_json: Path) -> list[dict]:
    if not moments_json.exists():
        raise FileNotFoundError(
            f"未找到 {moments_json.name}。请先使用 wxMoments 导出朋友圈（会生成 moments.json）。"
        )
    data = json.loads(moments_json.read_text(encoding="utf-8-sig"))
    if not isinstance(data, list):
        raise ValueError("moments.json 结构异常：顶层应为数组。")
    return data


def row_values(record: dict, index: int) -> list[str]:
    values: list[str] = []
    for _, key in COLUMNS:
        if key == "images":
            images = record.get("images") or []
            values.append("|".join(images) if isinstance(images, list) else str(images))
        elif key == "category_id":
            values.append(category_id_for(record))
        elif key == "content":
            values.append(_strip_phone(record.get(key) or ""))
        else:
            values.append(str(record.get(key) or ""))
    return values


def write_csv(records: list[dict], out_path: Path) -> None:
    with out_path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow([label for label, _ in COLUMNS])
        for idx, record in enumerate(records, 1):
            writer.writerow(row_values(record, idx))


def write_xlsx(records: list[dict], out_path: Path) -> None:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill
    except ImportError as exc:
        raise RuntimeError("未安装 openpyxl") from exc

    wb = Workbook()
    ws = wb.active
    ws.title = "朋友圈"

    header_fill = PatternFill("solid", fgColor="D9E1F2")
    header_font = Font(bold=True)
    top_align = Alignment(vertical="top", wrap_text=True)

    ws.append([label for label, _ in COLUMNS])
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(vertical="center", wrap_text=True)

    for idx, record in enumerate(records, 1):
        ws.append(row_values(record, idx))

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = top_align

    widths = {
        "时间": 20,
        "UID": 28,
        "手机号": 16,
        "朋友圈正文": 60,
        "图片路径": 50,
        "商品分类ID": 10,
    }
    for col_idx, (label, _) in enumerate(COLUMNS, 1):
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = widths.get(label, 15)

    ws.freeze_panes = "A2"
    wb.save(out_path)


def chunk_records(records: list[dict], size: int) -> list[list[dict]]:
    return [records[i : i + size] for i in range(0, len(records), size)]


def chunked_path(out_path: Path, index: int) -> Path:
    """First file keeps the original name; later ones get _2, _3, ... suffix."""
    if index == 0:
        return out_path
    return out_path.with_name(f"{out_path.stem}_{index + 1}{out_path.suffix}")


def write_xlsx_all(records: list[dict], out_path: Path) -> list[Path]:
    """Write records to one or more xlsx files, splitting at MAX_ROWS_PER_FILE.

    Returns the list of generated file paths.
    """
    chunks = chunk_records(records, MAX_ROWS_PER_FILE)
    paths: list[Path] = []
    for index, chunk in enumerate(chunks):
        path = chunked_path(out_path, index)
        write_xlsx(chunk, path)
        paths.append(path)
    return paths


def main() -> int:
    parser = argparse.ArgumentParser(description="将 wxMoments 导出的 moments.json 转为 Excel/CSV 表格")
    parser.add_argument("output_dir", help="wxMoments 导出目录（包含 moments.json）")
    parser.add_argument("--output", default="", help="输出文件路径，默认在导出目录下生成 moments.xlsx")
    args = parser.parse_args()

    output_dir = Path(args.output_dir).expanduser().resolve()
    moments_json = output_dir / "moments.json"

    try:
        records = load_records(moments_json)
    except (FileNotFoundError, ValueError, json.JSONDecodeError) as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1

    if not records:
        print("moments.json 为空，没有可导出的记录。", file=sys.stderr)
        return 0

    out_path = Path(args.output).expanduser().resolve() if args.output else output_dir / "moments.xlsx"
    out_path.parent.mkdir(parents=True, exist_ok=True)

    chunks = chunk_records(records, MAX_ROWS_PER_FILE)

    try:
        for index, chunk in enumerate(chunks):
            write_xlsx(chunk, chunked_path(out_path, index))
        paths = [chunked_path(out_path, i) for i in range(len(chunks))]
        if len(paths) == 1:
            print(f"已生成 Excel：{paths[0]}（共 {len(records)} 条）")
        else:
            print(f"共 {len(records)} 条，按每个文件 {MAX_ROWS_PER_FILE} 条拆分为 {len(paths)} 个 Excel：")
            for path, chunk in zip(paths, chunks):
                print(f"  {path}（{len(chunk)} 条）")
        return 0
    except RuntimeError:
        csv_base = out_path.with_suffix(".csv")
        for index, chunk in enumerate(chunks):
            write_csv(chunk, chunked_path(csv_base, index))
        paths = [chunked_path(csv_base, i) for i in range(len(chunks))]
        listing = "\n".join(f"  {p}" for p in paths)
        print(
            f"未安装 openpyxl，已生成 CSV：\n{listing}\n"
            f"如需 xlsx 格式，请执行：pip install openpyxl",
            file=sys.stderr,
        )
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
