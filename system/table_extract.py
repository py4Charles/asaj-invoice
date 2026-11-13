from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Sequence

SYSTEM_DIR = Path(__file__).resolve().parent
OCR_SCRIPT = SYSTEM_DIR / "ocr_windows.ps1"


@dataclass(frozen=True)
class Word:
    text: str
    x: float
    y: float
    w: float
    h: float


@dataclass
class Column:
    left: float
    right: float

    @property
    def center(self) -> float:
        return (self.left + self.right) / 2


def run_ocr(image_path: Path) -> List[Word]:
    if not image_path.exists():
        raise FileNotFoundError(f"Image not found: {image_path}")
    if not OCR_SCRIPT.exists():
        raise FileNotFoundError(f"OCR helper not found: {OCR_SCRIPT}")

    command = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(OCR_SCRIPT),
        "-Path",
        str(image_path),
    ]
    completed = subprocess.run(
        command,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if completed.returncode != 0:
        raise RuntimeError(f"OCR failed:\n{completed.stderr.strip()}")

    payload = completed.stdout.strip()
    if not payload:
        raise RuntimeError("OCR returned no output.")

    start = min(i for i in (payload.find("["), payload.find("{")) if i != -1)
    data = json.loads(payload[start:])
    if isinstance(data, dict):
        data = [data]

    return [
        Word(
            text=str(item.get("text", "")),
            x=float(item.get("x", 0)),
            y=float(item.get("y", 0)),
            w=float(item.get("w", 0)),
            h=float(item.get("h", 0)),
        )
        for item in data
        if str(item.get("text", "")).strip()
    ]


def group_rows(words: Sequence[Word], tolerance: float) -> List[List[Word]]:
    rows: List[List[Word]] = []
    anchors: List[float] = []
    for word in sorted(words, key=lambda w: (w.y, w.x)):
        if not rows or abs(word.y - anchors[-1]) > tolerance:
            rows.append([word])
            anchors.append(word.y)
        else:
            rows[-1].append(word)
    for row in rows:
        row.sort(key=lambda w: w.x)
    return rows


def _cluster_indices(values: Sequence[float], gap: float) -> List[int]:
    order = sorted(range(len(values)), key=lambda i: values[i])
    groups: List[List[int]] = []
    for index in order:
        if not groups or values[index] - values[groups[-1][-1]] > gap:
            groups.append([index])
        else:
            groups[-1].append(index)
    return groups


def _kmeans_indices(values: Sequence[float], count: int) -> List[int]:
    ordered = sorted(values)
    if count >= len(ordered):
        return [[i] for i in range(len(values))]
    step = (len(ordered) - 1) / max(count - 1, 1)
    centers = [ordered[round(i * step)] for i in range(count)]

    assignments = [0] * len(values)
    for _ in range(50):
        changed = False
        for i, value in enumerate(values):
            best = min(range(count), key=lambda c: abs(value - centers[c]))
            if assignments[i] != best:
                assignments[i] = best
                changed = True
        for c in range(count):
            members = [values[i] for i in range(len(values)) if assignments[i] == c]
            if members:
                centers[c] = sum(members) / len(members)
        if not changed:
            break

    groups: List[List[int]] = [[] for _ in range(count)]
    for i, c in enumerate(assignments):
        groups[c].append(i)
    return [group for group in groups if group]


def detect_columns(
    rows: Sequence[Sequence[Word]],
    gap: float,
    column_count: int | None = None,
) -> List[Column]:
    positions: List[float] = []
    for row in rows:
        for word in row:
            positions.append(word.x)
    if not positions:
        return []

    if column_count:
        groups = _kmeans_indices(positions, column_count)
    else:
        groups = _cluster_indices(positions, gap)

    columns: List[Column] = []
    for group in groups:
        xs = [positions[i] for i in group]
        columns.append(Column(left=min(xs), right=max(xs)))
    columns.sort(key=lambda c: c.center)
    return columns


def build_table(rows: Sequence[Sequence[Word]], columns: Sequence[Column]) -> List[List[str]]:
    table: List[List[str]] = []
    for row in rows:
        cells: List[List[str]] = [[] for _ in columns]
        for word in row:
            nearest = min(range(len(columns)), key=lambda i: abs(word.x - columns[i].center))
            cells[nearest].append(word.text)
        table.append([" ".join(parts).strip() for parts in cells])
    return table


def _clean(value: str) -> str:
    return " ".join(value.replace("\u00a0", " ").split())


def _parse_number(value: str) -> int | float | None:
    text = value.replace(",", "").strip()
    if not text:
        return None
    if text.isdigit():
        return int(text)
    try:
        return float(text)
    except ValueError:
        return None


def post_process_cell(header: str, value: str) -> object:
    text = _clean(value)
    if not text:
        return ""

    name = header.lower()

    if name == "time" and text.isdigit():
        digits = text.zfill(6)
        if len(digits) == 6:
            return f"{digits[:2]}:{digits[2:4]}:{digits[4:]}"
        return text

    if "weight" in name:
        number = _parse_number(text)
        return number if number is not None else text

    if name == "row_no" or name.endswith("_no"):
        number = _parse_number(text)
        return number if number is not None else text

    if name == "asa_id" or text.upper().startswith("ASA-"):
        return text.upper().rstrip("-")

    if "product" in name or name in {"code", "product_code"}:
        return text.replace(" ", "").upper().rstrip("-")

    number = _parse_number(text)
    if number is not None and not text.startswith("0") and text.count(".") <= 1:
        return number
    return text


WEIGHT_FLAG_RE = re.compile(r"^(?P<value>\d+(?:[.,]\d+)?)\s*(?P<flag>DB|US)$", re.IGNORECASE)


def post_process_table(
    headers: Sequence[str],
    table: Sequence[Sequence[str]],
    raw: bool = False,
) -> List[List[object]]:
    lowered = [header.lower() for header in headers]
    status_index = next((i for i, h in enumerate(lowered) if h == "status"), None)
    unit_index = next((i for i, h in enumerate(lowered) if h == "unit"), None)
    weight_indexes = [i for i, h in enumerate(lowered) if "weight" in h]

    rows: List[List[object]] = []
    for row in table:
        cells = list(row)

        for index in weight_indexes:
            match = WEIGHT_FLAG_RE.match(_clean(cells[index]))
            if not match:
                continue
            flag = match.group("flag").upper()
            target = status_index if flag == "DB" else unit_index
            if target is None:
                continue
            if not _clean(cells[target]):
                cells[target] = flag
            cells[index] = match.group("value")

        if raw:
            rows.append(cells)
        else:
            rows.append(
                [post_process_cell(headers[i], cells[i]) for i in range(len(cells))]
            )
    return rows


def default_names(count: int) -> List[str]:
    return [f"column_{i + 1}" for i in range(count)]


def reorder(
    headers: List[str],
    rows: List[List[object]],
    order: Sequence[str],
) -> tuple[List[str], List[List[object]]]:
    index = {name: i for i, name in enumerate(headers)}
    valid = [name for name in order if name in index]
    rest = [name for name in headers if name not in valid]
    final = valid + rest
    positions = [index[name] for name in final]
    return final, [[row[i] for i in positions] for row in rows]


def write_xlsx(
    path: Path,
    headers: Sequence[str],
    rows: Sequence[Sequence[object]],
    sheet_name: str = "Data",
) -> None:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError as exc:
        raise SystemExit(
            "openpyxl is required for .xlsx output. Install it with: "
            "python -m pip install openpyxl"
        ) from exc

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_name
    sheet.append(list(headers))

    header_fill = PatternFill("solid", fgColor="1A1A2E")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = header_font

    for row in rows:
        sheet.append(list(row))

    sheet.freeze_panes = "A2"
    if headers:
        sheet.auto_filter.ref = sheet.dimensions

    for index, header in enumerate(headers, start=1):
        width = max(
            [len(str(header))]
            + [len(str(row[index - 1])) if index - 1 < len(row) else 0 for row in rows]
        )
        sheet.column_dimensions[get_column_letter(index)].width = min(max(width + 2, 10), 40)

    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def parse_args(argv: Iterable[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract a table from an image using Windows OCR and export it to Excel."
    )
    parser.add_argument("image", type=Path, help="Path to the table image.")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Output .xlsx path (default: <image name>.xlsx).",
    )
    parser.add_argument(
        "--columns",
        type=int,
        default=None,
        help="Force the number of columns (default: auto-detect).",
    )
    parser.add_argument(
        "--names",
        default=None,
        help="Comma-separated column names, in detection order.",
    )
    parser.add_argument(
        "--order",
        default=None,
        help="Comma-separated column names to move to the front, in order.",
    )
    parser.add_argument(
        "--row-tolerance",
        type=float,
        default=8.0,
        help="Max y difference (px) for words on the same row (default: 8).",
    )
    parser.add_argument(
        "--column-gap",
        type=float,
        default=30.0,
        help="Min x gap (px) that starts a new column when auto-detecting (default: 30).",
    )
    parser.add_argument(
        "--sheet",
        default="Data",
        help="Worksheet name (default: Data).",
    )
    parser.add_argument(
        "--raw",
        action="store_true",
        help="Keep raw OCR text without value normalization.",
    )
    return parser.parse_args(argv)


def main(argv: Iterable[str] | None = None) -> int:
    args = parse_args(argv)

    words = run_ocr(args.image)
    if not words:
        raise SystemExit("No text was detected in the image.")

    rows = group_rows(words, args.row_tolerance)
    columns = detect_columns(rows, args.column_gap, args.columns)
    if not columns:
        raise SystemExit("Could not detect any columns.")

    table = build_table(rows, columns)

    if args.names:
        names = [part.strip() for part in args.names.split(",") if part.strip()]
        if len(names) < len(columns):
            names += [f"column_{i + 1}" for i in range(len(names), len(columns))]
        headers = names[: len(columns)]
    else:
        headers = default_names(len(columns))

    data = post_process_table(headers, table, raw=args.raw)

    if args.order:
        order = [part.strip() for part in args.order.split(",") if part.strip()]
        headers, data = reorder(headers, data, order)

    output = args.output or args.image.with_suffix(".xlsx")
    write_xlsx(output, headers, data, args.sheet)

    print(f"Detected {len(columns)} columns and {len(data)} rows.")
    print(f"Wrote {output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
