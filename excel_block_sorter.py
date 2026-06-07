from __future__ import annotations

import argparse
from copy import copy
from dataclasses import dataclass
from pathlib import Path
import re
from typing import Iterable

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


DEFAULT_OUTPUT_DIR = "siralanmis"


@dataclass(frozen=True)
class SortConfig:
    sort_column: int
    first_row: int | None
    last_column: int | None = None
    descending: bool = True
    sheet_name: str | None = None
    include_negative: bool = True
    keep_blank_rows_at_bottom: bool = True
    sort_columns: tuple[int, ...] = ()
    sort_total_blocks: bool = True

    def active_sort_columns(self) -> tuple[int, ...]:
        return self.sort_columns or (self.sort_column,)


@dataclass(frozen=True)
class ProcessResult:
    input_path: Path
    output_path: Path | None
    changed_blocks: int
    total_rows: int
    detail_rows: int
    skipped: bool = False
    reason: str = ""


def column_to_index(value: str) -> int:
    text = str(value).strip()
    if not text:
        raise ValueError("Kolon bos olamaz.")
    if text.isdigit():
        number = int(text)
        if number < 1:
            raise ValueError("Kolon numarasi 1 veya daha buyuk olmali.")
        return number
    return column_index_from_string(text.upper())


def columns_to_indices(value: str) -> tuple[int, ...]:
    parts = [part for part in re.split(r"[,;\s]+", value.strip()) if part]
    if not parts:
        raise ValueError("En az bir siralama kolonu girilmeli.")
    columns = tuple(column_to_index(part) for part in parts)
    if len(set(columns)) != len(columns):
        raise ValueError("Ayni siralama kolonu birden fazla yazilamaz.")
    return columns


def parse_number(value) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return None

    text = text.lower()
    multiplier = 1.0
    if "milyar" in text:
        multiplier = 1_000_000_000.0
    elif "milyon" in text:
        multiplier = 1_000_000.0
    elif "bin" in text:
        multiplier = 1_000.0

    text = (
        text.replace("milyar", "")
        .replace("milyon", "")
        .replace("bin", "")
        .replace("tl", "")
        .replace("try", "")
        .replace("\xa0", " ")
        .strip()
    )
    if re.search(r"[^\W\d_]", text, flags=re.UNICODE):
        return None

    if "," in text and "." in text:
        last_comma = text.rfind(",")
        last_dot = text.rfind(".")
        if last_comma > last_dot:
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    else:
        parts = text.split(".")
        if len(parts) > 2 and all(len(part) == 3 for part in parts[1:]):
            text = "".join(parts)
        elif len(parts) == 2 and len(parts[1]) == 3 and parts[0].replace("-", "").isdigit():
            text = "".join(parts)

    text = re.sub(r"[^0-9.\-]", "", text)
    if text in {"", "-", ".", "-."}:
        return None

    try:
        return float(text) * multiplier
    except ValueError:
        return None


def cell_has_total_style(cell: Cell) -> bool:
    return (
        cell.fill.fill_type == "solid"
        or cell.border.top.style == "double"
        or cell.border.bottom.style == "double"
    )


def is_total_row(
    ws: Worksheet,
    row_number: int,
    sort_col: int,
    last_col: int | None = None,
) -> bool:
    cell = ws.cell(row_number, sort_col)
    if parse_number(cell.value) is None:
        return False
    if cell_has_total_style(cell):
        return True
    if last_col is None:
        return False
    return any(cell_has_total_style(ws.cell(row_number, col)) for col in range(1, last_col + 1))


def has_merged_cell_in_range(ws: Worksheet, start_row: int, end_row: int, max_col: int) -> bool:
    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_row <= end_row and merged_range.max_row >= start_row:
            if merged_range.min_col <= max_col and merged_range.max_col >= 1:
                return True
    return False


def detect_last_column(ws: Worksheet, sort_cols: int | Iterable[int]) -> int:
    if isinstance(sort_cols, int):
        sort_columns = (sort_cols,)
    else:
        sort_columns = tuple(sort_cols)
    max_col = max((ws.max_column, *sort_columns))
    for row in ws.iter_rows():
        for cell in row:
            if cell.value not in (None, ""):
                max_col = max(max_col, cell.column)
    return max_col


def row_snapshot(ws: Worksheet, row_number: int, max_col: int) -> list[dict]:
    cells = []
    for col in range(1, max_col + 1):
        cell = ws.cell(row_number, col)
        if isinstance(cell, MergedCell):
            raise ValueError(
                f"{ws.title}!{get_column_letter(col)}{row_number} birlesik hucre icinde."
            )
        cells.append(snapshot_cell(cell))
    return cells


def rows_snapshot(ws: Worksheet, start_row: int, end_row: int, max_col: int) -> list[list[dict]]:
    return [row_snapshot(ws, row, max_col) for row in range(start_row, end_row + 1)]


def write_rows(ws: Worksheet, start_row: int, snapshots: list[list[dict]]) -> None:
    for offset, snapshot in enumerate(snapshots):
        write_row(ws, start_row + offset, snapshot)


def snapshot_cell(cell: Cell) -> dict:
    return {
        "value": cell.value,
        "style": copy(cell._style),
        "number_format": cell.number_format,
        "font": copy(cell.font),
        "fill": copy(cell.fill),
        "border": copy(cell.border),
        "alignment": copy(cell.alignment),
        "protection": copy(cell.protection),
        "comment": copy(cell.comment),
        "hyperlink": copy(cell.hyperlink),
    }


def write_row(ws: Worksheet, row_number: int, snapshot: list[dict]) -> None:
    for col, data in enumerate(snapshot, start=1):
        cell = ws.cell(row_number, col)
        cell.value = data["value"]
        cell._style = copy(data["style"])
        cell.number_format = data["number_format"]
        cell.font = copy(data["font"])
        cell.fill = copy(data["fill"])
        cell.border = copy(data["border"])
        cell.alignment = copy(data["alignment"])
        cell.protection = copy(data["protection"])
        cell.comment = copy(data["comment"])
        cell.hyperlink = copy(data["hyperlink"])


def iter_target_sheets(wb, sheet_name: str | None):
    if sheet_name:
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sayfa bulunamadi: {sheet_name}")
        yield wb[sheet_name]
        return
    yield from wb.worksheets


def total_row_key(ws: Worksheet, row_number: int, sort_columns: tuple[int, ...]) -> tuple[float, ...]:
    values = [parse_number(ws.cell(row_number, col).value) for col in sort_columns]
    return tuple(value if value is not None else float("-inf") for value in values)


def is_fixed_total_row(ws: Worksheet, row_number: int, last_column: int) -> bool:
    labels = []
    for col in range(1, last_column + 1):
        value = ws.cell(row_number, col).value
        if value not in (None, ""):
            labels.append(str(value).strip().lower())
    text = " ".join(labels)
    return "genel toplam" in text or "grand total" in text


def sort_total_blocks(
    ws: Worksheet,
    total_rows: list[int],
    last_column: int,
    sort_columns: tuple[int, ...],
    descending: bool,
) -> None:
    if len(total_rows) < 2:
        return

    blocks = []
    for index, total_row in enumerate(total_rows):
        end = (total_rows[index + 1] - 1) if index + 1 < len(total_rows) else ws.max_row
        blocks.append(
            {
                "start": total_row,
                "end": end,
                "key": total_row_key(ws, total_row, sort_columns),
                "fixed": is_fixed_total_row(ws, total_row, last_column),
                "rows": rows_snapshot(ws, total_row, end, last_column),
            }
        )

    output_blocks = []
    segment = []
    for block in blocks:
        if block["fixed"]:
            output_blocks.extend(sorted(segment, key=lambda item: item["key"], reverse=descending))
            segment = []
            output_blocks.append(block)
        else:
            segment.append(block)
    output_blocks.extend(sorted(segment, key=lambda item: item["key"], reverse=descending))

    row_pointer = total_rows[0]
    for block in output_blocks:
        write_rows(ws, row_pointer, block["rows"])
        row_pointer += len(block["rows"])


def sort_sheet(ws: Worksheet, config: SortConfig) -> tuple[int, int, int]:
    first_row = config.first_row or 1
    sort_columns = config.active_sort_columns()
    primary_sort_column = sort_columns[0]
    last_column = config.last_column or detect_last_column(ws, sort_columns)
    total_rows = [
        row
        for row in range(first_row, ws.max_row + 1)
        if is_total_row(ws, row, primary_sort_column, last_column)
    ]

    changed_blocks = 0
    moved_detail_rows = 0

    for index, total_row in enumerate(total_rows):
        start = total_row + 1
        end = (total_rows[index + 1] - 1) if index + 1 < len(total_rows) else ws.max_row
        if start > end:
            continue

        if has_merged_cell_in_range(ws, start, end, last_column):
            raise ValueError(
                f"{ws.title}!A{start}:{get_column_letter(last_column)}{end} "
                "araliginda birlesik hucre var; bu blok guvenli siralanamaz."
            )

        numeric_rows = []
        other_rows = []
        for row in range(start, end + 1):
            key_values = tuple(parse_number(ws.cell(row, col).value) for col in sort_columns)
            snapshot = row_snapshot(ws, row, last_column)
            primary_key = key_values[0]
            if primary_key is None or (primary_key < 0 and not config.include_negative):
                other_rows.append((key_values, snapshot))
            else:
                key = tuple(value if value is not None else float("-inf") for value in key_values)
                numeric_rows.append((key, snapshot))

        if len(numeric_rows) < 2:
            continue

        sorted_numeric_rows = sorted(
            numeric_rows,
            key=lambda item: item[0],
            reverse=config.descending,
        )
        ordered_rows = (
            sorted_numeric_rows + other_rows
            if config.keep_blank_rows_at_bottom
            else sorted_numeric_rows
        )

        for offset, (_, snapshot) in enumerate(ordered_rows):
            write_row(ws, start + offset, snapshot)

        changed_blocks += 1
        moved_detail_rows += len(ordered_rows)

    if config.sort_total_blocks:
        sort_total_blocks(ws, total_rows, last_column, sort_columns, config.descending)

    return changed_blocks, len(total_rows), moved_detail_rows


def collect_input_files(input_path: Path, output_dir: Path) -> list[Path]:
    if input_path.is_file():
        return [input_path]

    output_dir_resolved = output_dir.resolve()
    files = []
    for path in sorted(input_path.glob("*.xlsx")):
        if path.name.startswith("~$"):
            continue
        if output_dir_resolved in path.resolve().parents:
            continue
        files.append(path)
    return files


def build_output_path(
    input_file: Path,
    input_path: Path,
    output_dir: Path,
    suffix: str,
    output_name: str | None = None,
) -> Path:
    if output_name:
        name = Path(output_name)
        if name.suffix.lower() != ".xlsx":
            name = name.with_suffix(".xlsx")
        return output_dir / name.name
    if input_path.is_file():
        return output_dir / f"{input_file.stem}{suffix}{input_file.suffix}"
    relative_parent = input_file.parent.relative_to(input_path)
    return output_dir / relative_parent / f"{input_file.stem}{suffix}{input_file.suffix}"


def process_workbook(
    input_file: Path,
    input_path: Path,
    output_dir: Path,
    config: SortConfig,
    suffix: str,
    dry_run: bool = False,
    output_name: str | None = None,
) -> ProcessResult:
    if input_file.suffix.lower() != ".xlsx":
        return ProcessResult(input_file, None, 0, 0, 0, skipped=True, reason="xlsx degil")

    wb = load_workbook(input_file)
    changed_blocks = 0
    total_rows = 0
    detail_rows = 0

    for ws in iter_target_sheets(wb, config.sheet_name):
        sheet_changed, sheet_totals, sheet_details = sort_sheet(ws, config)
        changed_blocks += sheet_changed
        total_rows += sheet_totals
        detail_rows += sheet_details

    output_path = build_output_path(input_file, input_path, output_dir, suffix, output_name)
    if not dry_run:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(output_path)

    return ProcessResult(input_file, output_path, changed_blocks, total_rows, detail_rows)


def process_files(
    input_path: Path,
    output_dir: Path,
    config: SortConfig,
    suffix: str = "_sirali",
    dry_run: bool = False,
    output_name: str | None = None,
) -> list[ProcessResult]:
    files = collect_input_files(input_path, output_dir)
    results = []
    for input_file in files:
        try:
            results.append(
                process_workbook(
                    input_file,
                    input_path,
                    output_dir,
                    config,
                    suffix,
                    dry_run,
                    output_name,
                )
            )
        except Exception as exc:
            results.append(
                ProcessResult(
                    input_file,
                    None,
                    0,
                    0,
                    0,
                    skipped=True,
                    reason=str(exc),
                )
            )
    return results


def optional_positive_int(value: str) -> int | None:
    text = value.strip().lower()
    if text in {"", "auto", "otomatik"}:
        return None
    number = int(text)
    if number < 1:
        raise argparse.ArgumentTypeError("1 veya daha buyuk olmali.")
    return number


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Excel dosyalarinda toplam satirlarini sabit tutup alt detay satirlarini "
            "secili kolona gore siralar."
        )
    )
    parser.add_argument("--input", default=".", help="Excel dosyasi veya klasoru. Varsayilan: .")
    parser.add_argument("--output", default=DEFAULT_OUTPUT_DIR, help="Cikti klasoru.")
    parser.add_argument("--sort-column", default="J", help="Siralanacak kolon(lar). Ornek: J veya J,K")
    parser.add_argument(
        "--first-row",
        type=optional_positive_int,
        default=None,
        help="Ilk toplam satiri arama satiri. Varsayilan: auto",
    )
    parser.add_argument(
        "--last-column",
        default=None,
        help="Satirla birlikte tasinacak son kolon. Bos ise otomatik bulunur.",
    )
    parser.add_argument("--sheet", default=None, help="Sadece belirli sayfayi isle. Bos ise tum sayfalar.")
    parser.add_argument("--ascending", action="store_true", help="Kucukten buyuge sirala.")
    parser.add_argument("--suffix", default="_sirali", help="Coklu dosyada cikti dosyasi ek adi.")
    parser.add_argument("--output-name", default=None, help="Tek dosya icin yeni cikti dosyasi adi.")
    parser.add_argument("--dry-run", action="store_true", help="Dosya yazmadan kac blok bulunacagini goster.")
    return parser


def result_line(result: ProcessResult, dry_run: bool) -> str:
    if result.skipped:
        return f"SKIP {result.input_path.name}: {result.reason}"

    action = "kontrol edildi" if dry_run else f"yazildi -> {result.output_path}"
    return (
        f"OK {result.input_path.name}: {result.changed_blocks} blok siralandi, "
        f"{result.total_rows} toplam satiri bulundu, {result.detail_rows} detay satiri islendi, "
        f"{action}"
    )


def main(argv: Iterable[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    input_path = Path(args.input)
    output_dir = Path(args.output)
    if not input_path.exists():
        parser.error(f"Girdi bulunamadi: {input_path}")

    sort_columns = columns_to_indices(args.sort_column)
    config = SortConfig(
        sort_column=sort_columns[0],
        first_row=args.first_row,
        last_column=column_to_index(args.last_column) if args.last_column else None,
        descending=not args.ascending,
        sheet_name=args.sheet,
        sort_columns=sort_columns,
    )

    if config.last_column is not None and config.last_column < max(config.active_sort_columns()):
        parser.error("--last-column, --sort-column kolonunu kapsamali.")

    if args.output_name and input_path.is_dir():
        parser.error("--output-name sadece tek Excel dosyasi secildiginde kullanilir.")

    results = process_files(
        input_path,
        output_dir,
        config,
        args.suffix,
        args.dry_run,
        args.output_name,
    )
    if not results:
        print("Islenecek .xlsx dosyasi bulunamadi.")
        return 1

    for result in results:
        print(result_line(result, args.dry_run))
    return 2 if any(result.skipped for result in results) else 0


if __name__ == "__main__":
    raise SystemExit(main())
