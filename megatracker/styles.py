"""Shared styles: Calibri 11, grey headers with white text, light-blue inputs, yellow totals."""
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo


def font(size=11, bold=False, italic=False, color="000000"):
    return Font(name="Calibri", size=size, bold=bold, italic=italic, color=color)


FONT_BODY, FONT_BOLD = font(), font(bold=True)
FONT_HEADER = font(bold=True, color="FFFFFF")
FONT_NOTE = font(size=10, italic=True, color="808080")
FILL_HEADER = PatternFill("solid", fgColor="7F7F7F")   # grey header, white text
FILL_SECTION = PatternFill("solid", fgColor="404040")  # dashboard section bars
FILL_INPUT = PatternFill("solid", fgColor="D6E4F0")    # light blue = you type here
FILL_CALC = PatternFill("solid", fgColor="FFFFFF")     # white = formula
FILL_TOTAL = PatternFill("solid", fgColor="FFF2CC")    # light yellow = key result
CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)
_S = Side(style="thin", color="B4B4B4")
BORDER = Border(left=_S, right=_S, top=_S, bottom=_S)


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


def title(ws, text, span, color="1F4E79"):
    ws.merge_cells(span)
    c = ws[span.split(":")[0]]
    c.value, c.font, c.alignment = text, font(16, True, color=color), LEFT


def section(ws, row, col, text, color="1F4E79"):
    ws.cell(row=row, column=col, value=text).font = font(12, True, color=color)


def headers(ws, row, col, names):
    for i, name in enumerate(names):
        c = ws.cell(row=row, column=col + i, value=name)
        c.font, c.fill, c.alignment, c.border = FONT_HEADER, FILL_HEADER, CENTER, BORDER


def put(ws, row, col, value=None, kind="calc", fmt=None, bold=False, align=CENTER):
    """kind: input (blue), calc (white), total (yellow bold), label (bold text)."""
    c = ws.cell(row=row, column=col, value=value)
    c.border = BORDER
    c.alignment = align
    if kind == "input":
        c.fill, c.font = FILL_INPUT, FONT_BOLD if bold else FONT_BODY
    elif kind == "total":
        c.fill, c.font = FILL_TOTAL, FONT_BOLD
    elif kind == "label":
        c.font, c.alignment = FONT_BOLD, LEFT
    elif kind == "note":
        c.font, c.alignment = FONT_NOTE, LEFT
    else:
        c.fill, c.font = FILL_CALC, FONT_BOLD if bold else FONT_BODY
    if fmt:
        c.number_format = fmt
    return c


def table(ws, name, ref, style="TableStyleLight1"):
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=False)
    ws.add_table(t)


def dropdown(ws, formula, sqref):
    dv = DataValidation(type="list", formula1=formula, allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(sqref)


def landscape(ws, fit_height=0):
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, fit_height
