"""跨工作簿复制页签、行块复制、页签重排（06 / 09 Mac 模式共用）。"""

from __future__ import annotations

from copy import copy

from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


def _copy_style(source, target) -> None:
    for attr in ("font", "fill", "border", "alignment", "number_format", "protection"):
        setattr(target, attr, copy(getattr(source, attr)))


def _copy_cell(source_cell, target_cell) -> None:
    target_cell._value = source_cell._value
    target_cell.data_type = source_cell.data_type
    if source_cell.has_style:
        if source_cell.parent.parent is target_cell.parent.parent:
            target_cell._style = copy(source_cell._style)
        else:
            _copy_style(source_cell, target_cell)
    if source_cell.hyperlink:
        target_cell._hyperlink = copy(source_cell.hyperlink)
    if source_cell.comment:
        target_cell.comment = copy(source_cell.comment)


def _copy_cells(source: Worksheet, target: Worksheet) -> None:
    for (row, col), source_cell in source._cells.items():
        _copy_cell(source_cell, target.cell(column=col, row=row))


def _copy_dimensions(source: Worksheet, target: Worksheet) -> None:
    for attr in ("row_dimensions", "column_dimensions"):
        src_dims = getattr(source, attr)
        dst_dims = getattr(target, attr)
        for key, dim in src_dims.items():
            target_dim = copy(dim)
            target_dim.parent = target
            if dim.has_style:
                target_dim._style = None
                _copy_style(dim, target_dim)
            dst_dims[key] = target_dim


def copy_sheet_between_workbooks(source: Worksheet, target: Worksheet) -> None:
    """复制整页签（跨工作簿，较慢；仅在没有整文件可复制时使用）。"""
    _copy_cells(source, target)
    _copy_dimensions(source, target)
    target.sheet_format = copy(source.sheet_format)
    target.sheet_properties = copy(source.sheet_properties)
    target.merged_cells = copy(source.merged_cells)
    target.page_margins = copy(source.page_margins)
    target.page_setup = copy(source.page_setup)
    target.print_options = copy(source.print_options)


def copy_sheet_within_workbook(workbook, source: Worksheet) -> Worksheet:
    return workbook.copy_worksheet(source)


def reorder_workbook_sheets(wb: Workbook, desired_order: list[str]) -> None:
    """将已有页签重排为指定顺序（不复制单元格，很快）。"""
    for target_idx, name in enumerate(desired_order):
        sheet = wb[name]
        offset = target_idx - wb.index(sheet)
        if offset:
            wb.move_sheet(sheet, offset=offset)


def copy_row_range(
    source: Worksheet,
    target: Worksheet,
    src_start: int,
    src_end: int,
    dest_start: int,
    last_col: int,
) -> int:
    """复制连续行（含样式）；返回下一条目标行号。"""
    dest_row = dest_start
    for src_row in range(src_start, src_end + 1):
        for col in range(1, last_col + 1):
            source_cell = source._cells.get((src_row, col))
            if source_cell is not None:
                _copy_cell(source_cell, target.cell(row=dest_row, column=col))
        src_height = source.row_dimensions[src_row].height
        if src_height is not None:
            target.row_dimensions[dest_row].height = src_height
        dest_row += 1
    return dest_row


def copy_header_layout(
    source: Worksheet,
    target: Worksheet,
    header_rows: int,
    last_col: int,
) -> None:
    """复制前 N 行表头：列宽、行高、单元格内容与样式。"""
    for col in range(1, last_col + 1):
        letter = get_column_letter(col)
        width = source.column_dimensions[letter].width
        if width is not None:
            target.column_dimensions[letter].width = width
    copy_row_range(source, target, 1, header_rows, 1, last_col)


def copy_header_merged_cells(source: Worksheet, target: Worksheet, header_rows: int) -> None:
    for mrange in list(target.merged_cells.ranges):
        if mrange.max_row <= header_rows:
            target.unmerge_cells(str(mrange))
    for mrange in source.merged_cells.ranges:
        if mrange.max_row <= header_rows:
            target.merge_cells(str(mrange))


def restore_header_region(
    source: Worksheet,
    target: Worksheet,
    header_rows: int,
    last_col: int,
) -> None:
    """用源页签前 N 行覆盖目标页签表头区（含公式、样式、列宽、行高、合并单元格）。"""
    copy_header_layout(source, target, header_rows, last_col)
    copy_header_merged_cells(source, target, header_rows)
