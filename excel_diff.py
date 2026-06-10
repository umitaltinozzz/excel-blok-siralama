from __future__ import annotations

from copy import copy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import MergedCell

# ---------------------------------------------------------------------------
# Renk sabitleri
# ---------------------------------------------------------------------------

FILL_CHANGED = PatternFill("solid", fgColor="FFFF00")   # Sarı  – değişen hücre
FILL_ADDED = PatternFill("solid", fgColor="C6EFCE")     # Yeşil – yalnızca yeni dosyada
FILL_REMOVED = PatternFill("solid", fgColor="FFC7CE")   # Kırmızı – yalnızca eski dosyada


# ---------------------------------------------------------------------------
# Veri sınıfları
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class CellDiff:
    """Tek bir hücre farkını temsil eder."""

    row: int
    col: int
    old_value: Any
    new_value: Any
    diff_type: str  # 'changed', 'added', 'removed'
    comment: str = ""


@dataclass(frozen=True)
class SheetDiffResult:
    """Bir sayfa için karşılaştırma sonucu."""

    sheet_name: str
    changed_cells: int
    added_rows: int
    removed_rows: int
    diffs: list[CellDiff] = field(default_factory=list)


@dataclass(frozen=True)
class DiffResult:
    """Tüm çalışma kitabı karşılaştırma sonucu."""

    old_path: Path
    new_path: Path
    output_path: Path | None
    sheet_results: list[SheetDiffResult] = field(default_factory=list)
    added_sheets: list[str] = field(default_factory=list)
    removed_sheets: list[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Yardımcı fonksiyonlar
# ---------------------------------------------------------------------------

def _cell_value(ws: Worksheet, row: int, col: int) -> Any:
    """Hücre değerini güvenli şekilde döndürür."""
    return ws.cell(row, col).value


def _row_values(ws: Worksheet, row: int, max_col: int) -> list[Any]:
    """Bir satırdaki tüm hücre değerlerini liste olarak döndürür."""
    return [_cell_value(ws, row, col) for col in range(1, max_col + 1)]


def _is_row_empty(ws: Worksheet, row: int, max_col: int) -> bool:
    """Satırın tamamen boş olup olmadığını kontrol eder."""
    return all(
        _cell_value(ws, row, col) in (None, "")
        for col in range(1, max_col + 1)
    )


def _effective_max_col(ws_old: Worksheet, ws_new: Worksheet) -> int:
    """İki sayfanın ortak maksimum kolon sayısını hesaplar."""
    return max(ws_old.max_column or 1, ws_new.max_column or 1)


def _build_key_map(
    ws: Worksheet,
    max_col: int,
    key_column: int,
) -> dict[Any, int]:
    """Anahtar kolondaki değere göre satır numarası eşlemesi oluşturur.

    Aynı anahtar birden fazla satırda varsa son satır kullanılır.
    """
    key_map: dict[Any, int] = {}
    for row in range(1, (ws.max_row or 0) + 1):
        key = _cell_value(ws, row, key_column)
        if key not in (None, ""):
            key_map[key] = row
    return key_map


def _compare_cells(
    old_val: Any, old_form: Any, old_num_fmt: Any,
    new_val: Any, new_form: Any, new_num_fmt: Any
) -> tuple[bool, str]:
    """İki hücreyi değer, formül ve biçim açısından karşılaştırır.
    
    Fark varsa True ve açıklayıcı yorum metni döndürür.
    """
    # Her iki hücre de boş ise fark yoktur
    if old_val in (None, "") and new_val in (None, "") and old_form in (None, "") and new_form in (None, ""):
        return False, ""

    old_is_formula = str(old_form).startswith("=") if old_form is not None else False
    new_is_formula = str(new_form).startswith("=") if new_form is not None else False

    # 1. Formül -> Formül Değişimi
    if old_is_formula and new_is_formula:
        if old_form != new_form:
            return True, f"Formül Değişti:\nEski: {old_form}\nYeni: {new_form}"
        if old_val != new_val:
            return True, f"Formül Sonucu Değişti:\nEski Değer: {old_val} → Yeni Değer: {new_val}"
            
    # 2. Formül -> Sabit Değer (Formülün iptal edilmesi)
    elif old_is_formula and not new_is_formula:
        return True, f"Formül İptal Edildi (Sabit Değer):\nEski Formül: {old_form}\nYeni Değer: {new_val}"
        
    # 3. Sabit Değer -> Formül (Yeni formül eklenmesi)
    elif not old_is_formula and new_is_formula:
        return True, f"Formül Eklendi:\nEski Değer: {old_val}\nYeni Formül: {new_form}"
        
    # 4. Sabit Değer -> Sabit Değer Değişimi
    else:
        if old_val != new_val:
            if old_val in (None, ""):
                return True, f"Değer Eklendi: {new_val}"
            if new_val in (None, ""):
                return True, f"Değer Silindi (Eski: {old_val})"
            return True, f"Değer Değişti:\nEski: {old_val} → Yeni: {new_val}"

    # 5. Değerler aynı iken Sayı Biçimlendirmesi (Number Format) Değişimi
    if old_num_fmt != new_num_fmt and old_num_fmt is not None and new_num_fmt is not None:
        if old_num_fmt != "General" or new_num_fmt != "General":
            return True, f"Biçimlendirme Değişti:\nEski Format: {old_num_fmt} → Yeni Format: {new_num_fmt}"

    return False, ""


# ---------------------------------------------------------------------------
# Sayfa karşılaştırma
# ---------------------------------------------------------------------------

def diff_sheets(
    ws_old_val: Worksheet,
    ws_old_form: Worksheet,
    ws_new_val: Worksheet,
    ws_new_form: Worksheet,
    max_col: int | None = None,
    key_column: int | None = None,
) -> SheetDiffResult:
    """İki sayfayı hücre hücre karşılaştırır (formül ve değer farkları dahil).

    Args:
        ws_old_val: Eski sayfa (data_only=True).
        ws_old_form: Eski sayfa (data_only=False).
        ws_new_val: Yeni sayfa (data_only=True).
        ws_new_form: Yeni sayfa (data_only=False).
        max_col: Karşılaştırılacak maksimum kolon.
        key_column: Satır eşleştirmede kullanılacak anahtar kolon.

    Returns:
        Sayfa bazında fark sonucu.
    """
    if max_col is None:
        max_col = _effective_max_col(ws_old_val, ws_new_val)

    diffs: list[CellDiff] = []
    changed_cells = 0
    added_rows_count = 0
    removed_rows_count = 0

    if key_column is not None:
        # Anahtar kolon tabanlı eşleştirme
        old_map = _build_key_map(ws_old_val, max_col, key_column)
        new_map = _build_key_map(ws_new_val, max_col, key_column)

        all_keys = list(dict.fromkeys(list(old_map.keys()) + list(new_map.keys())))

        for key in all_keys:
            old_row = old_map.get(key)
            new_row = new_map.get(key)

            if old_row is not None and new_row is None:
                # Satır yalnızca eski dosyada var
                removed_rows_count += 1
                for col in range(1, max_col + 1):
                    val = ws_old_val.cell(old_row, col).value
                    form = ws_old_form.cell(old_row, col).value
                    is_formula = str(form).startswith("=") if form is not None else False
                    diffs.append(CellDiff(
                        row=old_row,
                        col=col,
                        old_value=form if is_formula else val,
                        new_value=None,
                        diff_type="removed",
                    ))
            elif old_row is None and new_row is not None:
                # Satır yalnızca yeni dosyada var
                added_rows_count += 1
                for col in range(1, max_col + 1):
                    val = ws_new_val.cell(new_row, col).value
                    form = ws_new_form.cell(new_row, col).value
                    is_formula = str(form).startswith("=") if form is not None else False
                    diffs.append(CellDiff(
                        row=new_row,
                        col=col,
                        old_value=None,
                        new_value=form if is_formula else val,
                        diff_type="added",
                    ))
            else:
                # Her iki dosyada da var – hücre bazında karşılaştır
                assert old_row is not None and new_row is not None
                for col in range(1, max_col + 1):
                    old_val = ws_old_val.cell(old_row, col).value
                    old_form = ws_old_form.cell(old_row, col).value
                    old_num_fmt = ws_old_form.cell(old_row, col).number_format

                    new_val = ws_new_val.cell(new_row, col).value
                    new_form = ws_new_form.cell(new_row, col).value
                    new_num_fmt = ws_new_form.cell(new_row, col).number_format

                    is_changed, comment = _compare_cells(
                        old_val, old_form, old_num_fmt,
                        new_val, new_form, new_num_fmt
                    )
                    if is_changed:
                        changed_cells += 1
                        old_is_formula = str(old_form).startswith("=") if old_form is not None else False
                        new_is_formula = str(new_form).startswith("=") if new_form is not None else False
                        diffs.append(CellDiff(
                            row=new_row,
                            col=col,
                            old_value=old_form if old_is_formula else old_val,
                            new_value=new_form if new_is_formula else new_val,
                            diff_type="changed",
                            comment=comment,
                        ))
    else:
        # Satır numarasına göre eşleştirme
        old_max_row = ws_old_val.max_row or 0
        new_max_row = ws_new_val.max_row or 0
        common_rows = min(old_max_row, new_max_row)

        # Ortak satırları karşılaştır
        for row in range(1, common_rows + 1):
            for col in range(1, max_col + 1):
                old_val = ws_old_val.cell(row, col).value
                old_form = ws_old_form.cell(row, col).value
                old_num_fmt = ws_old_form.cell(row, col).number_format

                new_val = ws_new_val.cell(row, col).value
                new_form = ws_new_form.cell(row, col).value
                new_num_fmt = ws_new_form.cell(row, col).number_format

                is_changed, comment = _compare_cells(
                    old_val, old_form, old_num_fmt,
                    new_val, new_form, new_num_fmt
                )
                if is_changed:
                    changed_cells += 1
                    old_is_formula = str(old_form).startswith("=") if old_form is not None else False
                    new_is_formula = str(new_form).startswith("=") if new_form is not None else False
                    diffs.append(CellDiff(
                        row=row,
                        col=col,
                        old_value=old_form if old_is_formula else old_val,
                        new_value=new_form if new_is_formula else new_val,
                        diff_type="changed",
                        comment=comment,
                    ))

        # Yeni dosyada fazla satırlar
        for row in range(common_rows + 1, new_max_row + 1):
            if not _is_row_empty(ws_new_val, row, max_col):
                added_rows_count += 1
                for col in range(1, max_col + 1):
                    val = ws_new_val.cell(row, col).value
                    form = ws_new_form.cell(row, col).value
                    is_formula = str(form).startswith("=") if form is not None else False
                    diffs.append(CellDiff(
                        row=row,
                        col=col,
                        old_value=None,
                        new_value=form if is_formula else val,
                        diff_type="added",
                    ))

        # Eski dosyada fazla satırlar
        for row in range(common_rows + 1, old_max_row + 1):
            if not _is_row_empty(ws_old_val, row, max_col):
                removed_rows_count += 1
                for col in range(1, max_col + 1):
                    val = ws_old_val.cell(row, col).value
                    form = ws_old_form.cell(row, col).value
                    is_formula = str(form).startswith("=") if form is not None else False
                    diffs.append(CellDiff(
                        row=row,
                        col=col,
                        old_value=form if is_formula else val,
                        new_value=None,
                        diff_type="removed",
                    ))

    return SheetDiffResult(
        sheet_name=ws_new_val.title,
        changed_cells=changed_cells,
        added_rows=added_rows_count,
        removed_rows=removed_rows_count,
        diffs=diffs,
    )


# ---------------------------------------------------------------------------
# Fark çıktısı yazma
# ---------------------------------------------------------------------------

def _apply_changed_fill(cell, diff: CellDiff) -> None:
    """Değişen hücreye sarı dolgu ve açıklayıcı yorum ekler."""
    if isinstance(cell, MergedCell):
        return
    cell.fill = FILL_CHANGED
    comment_text = diff.comment or f"Eski: {diff.old_value} → Yeni: {diff.new_value}"
    cell.comment = Comment(comment_text, "ExcelDiff")


def _apply_added_fill(ws: Worksheet, row: int, max_col: int) -> None:
    """Eklenen satırın tüm hücrelerini yeşil renkle boyar."""
    for col in range(1, max_col + 1):
        cell = ws.cell(row, col)
        if not isinstance(cell, MergedCell):
            cell.fill = FILL_ADDED


def _apply_removed_row(ws: Worksheet, row: int, diff_cells: list[CellDiff]) -> None:
    """Silinen satırı kırmızı arka planla çıktıya yazar."""
    for diff in diff_cells:
        cell = ws.cell(row, diff.col)
        if not isinstance(cell, MergedCell):
            cell.value = diff.old_value
            cell.fill = FILL_REMOVED


def _add_summary_sheet(wb: Workbook, diff_result: DiffResult) -> None:
    """Çalışma kitabına 'Fark Özeti' sayfası ekler."""
    ws = wb.create_sheet("Fark Özeti")

    # Başlık satırı
    headers = ["Sayfa Adı", "Değişen Hücre/Formül", "Eklenen Satır", "Silinen Satır"]
    header_font = Font(bold=True)
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(1, col)
        cell.value = header
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    # Sayfa sonuçları
    for row_idx, result in enumerate(diff_result.sheet_results, start=2):
        ws.cell(row_idx, 1).value = result.sheet_name
        ws.cell(row_idx, 2).value = result.changed_cells
        ws.cell(row_idx, 3).value = result.added_rows
        ws.cell(row_idx, 4).value = result.removed_rows

    # Eklenen ve silinen sayfalar
    info_row = len(diff_result.sheet_results) + 3
    if diff_result.added_sheets:
        ws.cell(info_row, 1).value = "Eklenen Sayfalar"
        ws.cell(info_row, 1).font = Font(bold=True)
        ws.cell(info_row, 2).value = ", ".join(diff_result.added_sheets)
        info_row += 1
    if diff_result.removed_sheets:
        ws.cell(info_row, 1).value = "Silinen Sayfalar"
        ws.cell(info_row, 1).font = Font(bold=True)
        ws.cell(info_row, 2).value = ", ".join(diff_result.removed_sheets)

    # Kolon genişlikleri
    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 18


def write_diff_workbook(
    wb_new: Workbook,
    diff_result: DiffResult,
    output_path: Path,
) -> None:
    """Yeni çalışma kitabının kopyasına farkları ve yorumları uygulayıp kaydeder."""
    for sheet_result in diff_result.sheet_results:
        if sheet_result.sheet_name not in wb_new.sheetnames:
            continue

        ws = wb_new[sheet_result.sheet_name]
        max_col = ws.max_column or 1

        # Eklenen satır numaralarını belirle
        added_rows: set[int] = set()
        removed_diffs_by_row: dict[int, list[CellDiff]] = {}

        for diff in sheet_result.diffs:
            if diff.diff_type == "added":
                added_rows.add(diff.row)
            elif diff.diff_type == "removed":
                removed_diffs_by_row.setdefault(diff.row, []).append(diff)

        # Değişen hücreleri/formülleri boyar ve yorum ekler
        for diff in sheet_result.diffs:
            if diff.diff_type == "changed":
                cell = ws.cell(diff.row, diff.col)
                _apply_changed_fill(cell, diff)

        # Eklenen satırları yeşile boyar
        for row in added_rows:
            _apply_added_fill(ws, row, max_col)

        # Silinen satırları tablonun sonuna ekler
        if removed_diffs_by_row:
            insert_row = (ws.max_row or 0) + 2
            for removed_row in sorted(removed_diffs_by_row.keys()):
                _apply_removed_row(ws, insert_row, removed_diffs_by_row[removed_row])
                insert_row += 1

    # Özet sayfası ekle
    _add_summary_sheet(wb_new, diff_result)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb_new.save(output_path)


# ---------------------------------------------------------------------------
# Ana karşılaştırma fonksiyonu
# ---------------------------------------------------------------------------

def diff_workbooks(
    old_path: Path | str,
    new_path: Path | str,
    output_dir: Path | str | None = None,
    sheet_name: str | None = None,
    key_column: int | None = None,
    sheet_name_old: str | None = None,
    sheet_name_new: str | None = None,
) -> DiffResult:
    """İki Excel dosyasını formül ve değer bazında karşılaştırıp fark raporu üretir.

    Args:
        old_path: Eski Excel dosyasının yolu.
        new_path: Yeni Excel dosyasının yolu.
        output_dir: Çıktı dizini. ``None`` ise fark dosyası oluşturulmaz.
        sheet_name: Sadece belirli bir sayfayı karşılaştır (iki dosyada da aynı isimli).
        key_column: Satır eşleştirmede kullanılacak anahtar kolon (1-tabanlı).
        sheet_name_old: Eski dosyada karşılaştırılacak sayfa adı.
        sheet_name_new: Yeni dosyada karşılaştırılacak sayfa adı.

    Returns:
        Karşılaştırma sonucu.
    """
    old_path = Path(old_path)
    new_path = Path(new_path)

    # 1. Çift Yükleme: Değerler ve formüller için ayrı ayrı yükle
    wb_old_val = load_workbook(old_path, data_only=True)
    wb_old_form = load_workbook(old_path, data_only=False)

    wb_new_val = load_workbook(new_path, data_only=True)
    wb_new_form = load_workbook(new_path, data_only=False)

    old_sheets = set(wb_old_val.sheetnames)
    new_sheets = set(wb_new_form.sheetnames)

    added_sheets = sorted(new_sheets - old_sheets)
    removed_sheets = sorted(old_sheets - new_sheets)

    # Karşılaştırılacak sayfaları belirle
    if sheet_name_old is not None and sheet_name_new is not None:
        if sheet_name_old not in old_sheets:
            raise ValueError(f"Eski dosyada sayfa bulunamadı: {sheet_name_old}")
        if sheet_name_new not in new_sheets:
            raise ValueError(f"Yeni dosyada sayfa bulunamadı: {sheet_name_new}")
        target_sheets = [(sheet_name_old, sheet_name_new)]
    elif sheet_name is not None:
        if sheet_name not in old_sheets:
            raise ValueError(f"Eski dosyada sayfa bulunamadı: {sheet_name}")
        if sheet_name not in new_sheets:
            raise ValueError(f"Yeni dosyada sayfa bulunamadı: {sheet_name}")
        target_sheets = [(sheet_name, sheet_name)]
    else:
        target_sheets = [(name, name) for name in wb_new_form.sheetnames if name in old_sheets]

    sheet_results: list[SheetDiffResult] = []
    for name_old, name_new in target_sheets:
        ws_old_val = wb_old_val[name_old]
        ws_old_form = wb_old_form[name_old]
        ws_new_val = wb_new_val[name_new]
        ws_new_form = wb_new_form[name_new]
        
        max_col = _effective_max_col(ws_old_val, ws_new_val)
        result = diff_sheets(
            ws_old_val, ws_old_form,
            ws_new_val, ws_new_form,
            max_col, key_column
        )
        sheet_results.append(result)

    # Çıktı dosyasını oluştur
    output_path: Path | None = None
    if output_dir is not None:
        output_dir = Path(output_dir)
        output_path = output_dir / f"{new_path.stem}_fark.xlsx"
        
        diff_result = DiffResult(
            old_path=old_path,
            new_path=new_path,
            output_path=output_path,
            sheet_results=sheet_results,
            added_sheets=added_sheets,
            removed_sheets=removed_sheets,
        )
        # wb_new_form kopyasına farkları çiz ve kaydet
        write_diff_workbook(wb_new_form, diff_result, output_path)
        return diff_result

    return DiffResult(
        old_path=old_path,
        new_path=new_path,
        output_path=None,
        sheet_results=sheet_results,
        added_sheets=added_sheets,
        removed_sheets=removed_sheets,
    )
