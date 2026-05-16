from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
import re
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape


OUT = Path("staff_yearly_leave_2026_management_ready.xlsx")
BACKUP = Path(f"staff_yearly_leave_2026.before_leave_type_cleanup_{datetime.now():%Y%m%d_%H%M%S}.xlsx")

NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
Q = '"'
LAST_COL = "ND"
CODES = ["AL", "SL", "TOR", "TR"]
TYPES = ["Annual Leave", "Sick Leave", "Time of Request", "Training"]
STATUSES = ["Pending", "Approved", "Rejected", "Cancelled", "Completed"]
MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]


def col_name(n: int) -> str:
    s = ""
    while n:
        n -= 1
        s = chr(65 + n % 26) + s
        n //= 26
    return s


def cell(row: int, col: int, value: str | float | int | None = None, style: int = 0, formula: bool = False) -> str:
    ref = f"{col_name(col)}{row}"
    if formula:
        cached = None
        if isinstance(value, tuple):
            value, cached = value
        cached_xml = "" if cached is None else f"<v>{escape(str(cached))}</v>"
        return f'<c r="{ref}" s="{style}"><f>{escape(str(value or ""))}</f>{cached_xml}</c>'
    if value is None:
        return f'<c r="{ref}" s="{style}"/>'
    if isinstance(value, (int, float)):
        return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t>{escape(value)}</t></is></c>'


def row(r: int, cells: list[str], attrs: str = "") -> str:
    return f'<row r="{r}"{attrs}>{"".join(cells)}</row>'


def content_types() -> str:
    sheets = "".join(
        f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for i in range(1, 6)
    )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>{sheets}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>'''


def root_rels() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'''


def workbook() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Setup" sheetId="1" r:id="rId1"/><sheet name="Yearly Calendar" sheetId="2" r:id="rId2"/><sheet name="Requests" sheetId="3" r:id="rId3"/><sheet name="Summary" sheetId="4" r:id="rId4"/><sheet name="Management Dashboard" sheetId="5" r:id="rId5"/></sheets><definedNames><definedName name="PlannerYear">Setup!$B$3</definedName><definedName name="StaffNames">Setup!$A$6:$A$25</definedName><definedName name="LeaveCodes">Setup!$D$6:$D$9</definedName><definedName name="LeaveTypes">Setup!$E$6:$E$9</definedName><definedName name="RequestStatuses">Setup!$H$6:$H$10</definedName></definedNames><calcPr calcId="191029" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>'''


def workbook_rels() -> str:
    rels = "".join(
        f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
        for i in range(1, 6)
    )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>'''


def styles() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="3"><numFmt numFmtId="164" formatCode="dd-mmm-yyyy"/><numFmt numFmtId="165" formatCode="h:mm AM/PM"/><numFmt numFmtId="166" formatCode="0.00"/></numFmts><fonts count="7"><font><sz val="11"/><color rgb="FF111827"/><name val="Calibri"/></font><font><b/><sz val="18"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FF111827"/><name val="Calibri"/></font><font><i/><sz val="10"/><color rgb="FF475569"/><name val="Calibri"/></font><font><b/><sz val="16"/><color rgb="FF0F172A"/><name val="Calibri"/></font><font><b/><sz val="20"/><color rgb="FF0F172A"/><name val="Calibri"/></font></fonts><fills count="14"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F2937"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF2563EB"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF8FAFC"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE5E7EB"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDCFCE7"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFEE2E2"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDBEAFE"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFDE68A"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE0F2FE"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFECFDF5"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF7ED"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF1F5F9"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFE2E8F0"/></left><right style="thin"><color rgb="FFE2E8F0"/></right><top style="thin"><color rgb="FFE2E8F0"/></top><bottom style="thin"><color rgb="FFE2E8F0"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="22"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="3" fillId="12" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf><xf numFmtId="0" fontId="3" fillId="10" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf><xf numFmtId="0" fontId="3" fillId="11" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="5" fillId="11" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="6" fillId="11" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="3" fillId="13" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="166" fontId="5" fillId="11" borderId="1" xfId="0" applyFont="1" applyFill="1" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="5"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFBBF7D0"/></patternFill></fill></dxf><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFECACA"/></patternFill></fill></dxf><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFDBEAFE"/></patternFill></fill></dxf><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFDE68A"/></patternFill></fill></dxf><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFF1F5F9"/></patternFill></fill></dxf></dxfs><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>'''


def theme() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="263238"/></a:dk2><a:lt2><a:srgbClr val="F8FAFC"/></a:lt2><a:accent1><a:srgbClr val="2563EB"/></a:accent1><a:accent2><a:srgbClr val="059669"/></a:accent2><a:accent3><a:srgbClr val="F59E0B"/></a:accent3><a:accent4><a:srgbClr val="DC2626"/></a:accent4><a:accent5><a:srgbClr val="7C3AED"/></a:accent5><a:accent6><a:srgbClr val="0891B2"/></a:accent6><a:hlink><a:srgbClr val="2563EB"/></a:hlink><a:folHlink><a:srgbClr val="7C3AED"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle/></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>'''


def setup_sheet() -> str:
    rows = [
        row(1, [cell(1, 1, "Staff Leave Planner Setup", 1)], ' ht="30" customHeight="1"'),
        row(3, [cell(3, 1, "Planner Year", 12), cell(3, 2, 2026, 3), cell(3, 4, "Change B3 to update calendar, dashboard, and request year.", 11)]),
        row(5, [cell(5, 1, "Staff", 9), cell(5, 2, "Department", 9), cell(5, 4, "Code", 9), cell(5, 5, "Leave Type", 9), cell(5, 8, "Request Status", 9)]),
    ]
    for r in range(6, 26):
        cells = [cell(r, 1, f"Staff {r-5:02d}", 6), cell(r, 2, "", 7)]
        if r <= 9:
            i = r - 6
            cells += [cell(r, 4, CODES[i], 7), cell(r, 5, TYPES[i], 7)]
        if r <= 10:
            cells.append(cell(r, 8, STATUSES[r - 6], 7))
        rows.append(row(r, cells))
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet {NS}><dimension ref="A1:I25"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="18"/><cols><col min="1" max="1" width="20" customWidth="1"/><col min="2" max="2" width="18" customWidth="1"/><col min="4" max="5" width="22" customWidth="1"/><col min="8" max="8" width="18" customWidth="1"/><col min="9" max="9" width="36" customWidth="1"/></cols><sheetData>{"".join(rows)}</sheetData><mergeCells count="1"><mergeCell ref="A1:I1"/></mergeCells><dataValidations count="1"><dataValidation type="whole" operator="between" allowBlank="0" showErrorMessage="1" sqref="B3"><formula1>1900</formula1><formula2>9999</formula2></dataValidation></dataValidations><pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.3" footer="0.3"/></worksheet>'''


def calendar_sheet() -> str:
    rows = [row(1, [cell(1, 1, "Yearly Leave Calendar", 1), cell(1, 3, "Setup!$B$3", 1, True)], ' ht="30" customHeight="1"')]
    rows.append(row(2, [cell(2, 1, "Staff", 5), cell(2, 2, "Department", 5)] + [
        cell(2, c, f'IF(YEAR(DATE(Setup!$B$3,1,{c-2}))=Setup!$B$3,IF(DAY(DATE(Setup!$B$3,1,{c-2}))=1,TEXT(DATE(Setup!$B$3,1,{c-2}),"mmm"),""),"")', 2, True)
        for c in range(3, 369)
    ]))
    for rr, fmt in [(3, "DAY({col}$5)"), (4, 'TEXT({col}$5,"ddd")')]:
        rows.append(row(rr, [cell(rr, 1, "", 5), cell(rr, 2, "", 5)] + [
            cell(rr, c, f'IF({col_name(c)}$5="","",{fmt.format(col=col_name(c))})', 3, True)
            for c in range(3, 369)
        ]))
    rows.append(row(5, [cell(5, 1, "", 0), cell(5, 2, "", 0)] + [
        cell(5, c, f'IF(YEAR(DATE(Setup!$B$3,1,{c-2}))=Setup!$B$3,DATE(Setup!$B$3,1,{c-2}),"")', 13, True)
        for c in range(3, 369)
    ], ' hidden="1"'))
    for r in range(6, 26):
        rows.append(row(r, [cell(r, 1, f"Staff {r-5:02d}", 6), cell(r, 2, "", 6)] + [cell(r, c, None, 7) for c in range(3, 369)]))
    cf = f'<conditionalFormatting sqref="C2:{LAST_COL}25"><cfRule type="expression" priority="1" dxfId="4"><formula>AND(C$5&lt;&gt;"",WEEKDAY(C$5,2)&gt;5)</formula></cfRule></conditionalFormatting>'
    leave_cf = '<conditionalFormatting sqref="C6:ND25">' + ''.join(
        f'<cfRule type="cellIs" priority="{i+2}" operator="equal" dxfId="{i}"><formula>"{code}"</formula></cfRule>'
        for i, code in enumerate(CODES)
    ) + '</conditionalFormatting>'
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet {NS}><dimension ref="A1:{LAST_COL}25"/><sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="5" topLeftCell="C6" activePane="bottomRight" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="18"/><cols><col min="1" max="1" width="18" customWidth="1"/><col min="2" max="2" width="14" customWidth="1"/><col min="3" max="368" width="4.2" customWidth="1"/></cols><sheetData>{"".join(rows)}</sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>{cf}{leave_cf}<dataValidations count="1"><dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="C6:{LAST_COL}25"><formula1>LeaveCodes</formula1></dataValidation></dataValidations><pageMargins left="0.25" right="0.25" top="0.5" bottom="0.5" header="0.3" footer="0.3"/></worksheet>'''


def requests_sheet() -> str:
    rows = [
        row(1, [cell(1, 1, "Leave Requests", 1), cell(1, 4, "Setup!$B$3", 1, True)], ' ht="30" customHeight="1"'),
        row(2, [cell(2, 1, "Record requests here. Dashboard and Summary use this sheet automatically.", 11)]),
        row(3, [cell(3, i + 1, h, 9) for i, h in enumerate(["Request Date", "Staff", "Department", "Leave Type", "Leave Date", "Start Time", "End Time", "Hours", "Status", "Approved By", "Notes"])], ' ht="22" customHeight="1"'),
    ]
    for r in range(4, 504):
        rows.append(row(r, [
            cell(r, 1, None, 13), cell(r, 2, None, 7),
            cell(r, 3, f'IFERROR(VLOOKUP(B{r},Setup!$A$6:$B$25,2,FALSE),"")', 7, True),
            cell(r, 4, None, 7), cell(r, 5, None, 13), cell(r, 6, None, 14), cell(r, 7, None, 14),
            cell(r, 8, f'IF(AND(F{r}<>"",G{r}<>""),MOD(G{r}-F{r},1)*24,"")', 15, True),
            cell(r, 9, None, 7), cell(r, 10, None, 7), cell(r, 11, None, 16),
        ]))
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet {NS}><dimension ref="A1:K503"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="18"/><cols><col min="1" max="1" width="14" customWidth="1"/><col min="2" max="2" width="18" customWidth="1"/><col min="3" max="3" width="16" customWidth="1"/><col min="4" max="4" width="18" customWidth="1"/><col min="5" max="5" width="14" customWidth="1"/><col min="6" max="7" width="12" customWidth="1"/><col min="8" max="8" width="10" customWidth="1"/><col min="9" max="9" width="13" customWidth="1"/><col min="10" max="10" width="16" customWidth="1"/><col min="11" max="11" width="36" customWidth="1"/></cols><sheetData>{"".join(rows)}</sheetData><mergeCells count="2"><mergeCell ref="A1:C1"/><mergeCell ref="A2:K2"/></mergeCells><autoFilter ref="A3:K503"/><dataValidations count="4"><dataValidation type="date" operator="between" allowBlank="1" showErrorMessage="1" sqref="A4:A503 E4:E503"><formula1>DATE(Setup!$B$3,1,1)</formula1><formula2>DATE(Setup!$B$3,12,31)</formula2></dataValidation><dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="B4:B503"><formula1>StaffNames</formula1></dataValidation><dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="D4:D503"><formula1>LeaveTypes</formula1></dataValidation><dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="I4:I503"><formula1>RequestStatuses</formula1></dataValidation></dataValidations><pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.3" footer="0.3"/></worksheet>'''


def summary_sheet() -> str:
    headers = ["Staff", "Annual Leave", "Sick Leave", "Time of Request", "Training", "Total Calendar Days", "Request Hours", "Pending", "Approved", "Rejected", "Completed"]
    rows = [
        row(1, [cell(1, 1, "Staff Summary", 1), cell(1, 4, ("Setup!$B$3", 2026), 1, True)], ' ht="30" customHeight="1"'),
        row(2, [cell(2, i + 1, h, 9) for i, h in enumerate(headers)], ' ht="22" customHeight="1"'),
    ]
    for r in range(3, 23):
        cal_row = r + 3
        setup_row = r + 3
        cells = [cell(r, 1, f"Staff {setup_row-5:02d}", 6)]
        cells += [cell(r, 2 + i, f'COUNTIF(\'Yearly Calendar\'!C{cal_row}:{LAST_COL}{cal_row},"{code}")', 10, True) for i, code in enumerate(CODES)]
        cells.append(cell(r, 6, f"SUM(B{r}:E{r})", 10, True))
        cells.append(cell(r, 7, f'SUMIFS(Requests!$H$4:$H$503,Requests!$B$4:$B$503,A{r},Requests!$E$4:$E$503,">="&DATE(Setup!$B$3,1,1),Requests!$E$4:$E$503,"<"&DATE(Setup!$B$3+1,1,1))', 15, True))
        for c, status in [(8, "Pending"), (9, "Approved"), (10, "Rejected"), (11, "Completed")]:
            cells.append(cell(r, c, f'COUNTIFS(Requests!$B$4:$B$503,A{r},Requests!$E$4:$E$503,">="&DATE(Setup!$B$3,1,1),Requests!$E$4:$E$503,"<"&DATE(Setup!$B$3+1,1,1),Requests!$I$4:$I$503,"{status}")', 10, True))
        rows.append(row(r, cells))
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet {NS}><dimension ref="A1:K23"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="18"/><cols><col min="1" max="1" width="18" customWidth="1"/><col min="2" max="10" width="15" customWidth="1"/><col min="11" max="11" width="18" customWidth="1"/></cols><sheetData>{"".join(rows)}</sheetData><mergeCells count="1"><mergeCell ref="A1:C1"/></mergeCells><autoFilter ref="A2:K22"/><pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.3" footer="0.3"/></worksheet>'''


def dashboard_sheet() -> str:
    rows = [
        row(1, [cell(1, 1, "Management Leave Dashboard", 1)], ' ht="32" customHeight="1"'),
        row(3, [cell(3, 1, "Year", 12), cell(3, 2, ("Setup!$B$3", 2026), 18, True), cell(3, 4, "Staff Focus", 12), cell(3, 5, "All", 18), cell(3, 14, "All", 7)]),
    ]
    # Hidden dropdown helper list plus visible dashboard rows.
    for r in range(4, 24):
        cells = []
        if r == 4:
            def kpi(status: str) -> str:
                return f'COUNTIFS(Requests!$E$4:$E$503,">="&DATE($B$3,1,1),Requests!$E$4:$E$503,"<"&DATE($B$3+1,1,1),Requests!$I$4:$I$503,"{status}",Requests!$B$4:$B$503,IF($E$3="All","*",$E$3))'
            cells += [cell(r, 1, "Pending", 17), cell(r, 2, kpi("Pending"), 19, True), cell(r, 4, "Approved", 17), cell(r, 5, kpi("Approved"), 19, True), cell(r, 7, "Hours", 17), cell(r, 8, 'SUMIFS(Requests!$H$4:$H$503,Requests!$E$4:$E$503,">="&DATE($B$3,1,1),Requests!$E$4:$E$503,"<"&DATE($B$3+1,1,1),Requests!$B$4:$B$503,IF($E$3="All","*",$E$3))', 21, True)]
        elif r == 6:
            cells += [cell(r, 1, "Monthly Request Status", 12), cell(r, 10, "Leave Type Mix", 12)]
        elif r == 7:
            cells += [cell(r, 1, "Month", 9)] + [cell(r, i + 2, s, 9) for i, s in enumerate(STATUSES)] + [cell(r, 7, "Total", 9), cell(r, 8, "Hours", 9), cell(r, 10, "Leave Type", 9), cell(r, 11, "Pending", 9), cell(r, 12, "Approved", 9), cell(r, 13, "Total", 9)]
        elif 8 <= r <= 19:
            m = r - 7
            cells.append(cell(r, 1, MONTHS[m - 1], 6))
            for c in range(2, 7):
                cells.append(cell(r, c, f'COUNTIFS(Requests!$E$4:$E$503,">="&DATE($B$3,{m},1),Requests!$E$4:$E$503,"<"&EDATE(DATE($B$3,{m},1),1),Requests!$I$4:$I$503,{col_name(c)}$7,Requests!$B$4:$B$503,IF($E$3="All","*",$E$3))', 10, True))
            cells += [cell(r, 7, f"SUM(B{r}:F{r})", 10, True), cell(r, 8, f'SUMIFS(Requests!$H$4:$H$503,Requests!$E$4:$E$503,">="&DATE($B$3,{m},1),Requests!$E$4:$E$503,"<"&EDATE(DATE($B$3,{m},1),1),Requests!$B$4:$B$503,IF($E$3="All","*",$E$3))', 15, True)]
            if r <= 11:
                typ = TYPES[r - 8]
                cells += [cell(r, 10, typ, 6), cell(r, 11, f'COUNTIFS(Requests!$E$4:$E$503,">="&DATE($B$3,1,1),Requests!$E$4:$E$503,"<"&DATE($B$3+1,1,1),Requests!$D$4:$D$503,J{r},Requests!$I$4:$I$503,"Pending",Requests!$B$4:$B$503,IF($E$3="All","*",$E$3))', 10, True), cell(r, 12, f'COUNTIFS(Requests!$E$4:$E$503,">="&DATE($B$3,1,1),Requests!$E$4:$E$503,"<"&DATE($B$3+1,1,1),Requests!$D$4:$D$503,J{r},Requests!$I$4:$I$503,"Approved",Requests!$B$4:$B$503,IF($E$3="All","*",$E$3))', 10, True), cell(r, 13, f"SUM(K{r}:L{r})", 10, True)]
        elif r == 22:
            cells.append(cell(r, 1, "Staff Follow-up", 12))
        elif r == 23:
            cells += [cell(r, i + 1, h, 9) for i, h in enumerate(["Staff", "Pending", "Approved", "Rejected", "Cancelled", "Completed", "Total", "Hours"])]
        cells.append(cell(r, 14, f"Staff {r-3:02d}", 7))
        rows.append(row(r, cells))
    for r in range(24, 44):
        cells = [cell(r, 1, f"Staff {r-23:02d}", 6)]
        for c in range(2, 7):
            cells.append(cell(r, c, f'COUNTIFS(Requests!$E$4:$E$503,">="&DATE($B$3,1,1),Requests!$E$4:$E$503,"<"&DATE($B$3+1,1,1),Requests!$B$4:$B$503,A{r},Requests!$I$4:$I$503,{col_name(c)}$23)', 10, True))
        cells += [cell(r, 7, f"SUM(B{r}:F{r})", 10, True), cell(r, 8, f'SUMIFS(Requests!$H$4:$H$503,Requests!$E$4:$E$503,">="&DATE($B$3,1,1),Requests!$E$4:$E$503,"<"&DATE($B$3+1,1,1),Requests!$B$4:$B$503,A{r})', 15, True)]
        rows.append(row(r, cells))
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet {NS}><dimension ref="A1:N46"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="7" topLeftCell="A8" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="18"/><cols><col min="1" max="1" width="18" customWidth="1"/><col min="2" max="8" width="13" customWidth="1"/><col min="10" max="13" width="17" customWidth="1"/><col min="14" max="14" width="18" hidden="1" customWidth="1"/></cols><sheetData>{"".join(rows)}</sheetData><mergeCells count="3"><mergeCell ref="A1:M1"/><mergeCell ref="A6:H6"/><mergeCell ref="J6:M6"/></mergeCells><autoFilter ref="A23:H43"/><dataValidations count="1"><dataValidation type="list" allowBlank="0" showErrorMessage="1" sqref="E3"><formula1>$N$3:$N$23</formula1></dataValidation></dataValidations><pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/></worksheet>'''


def core() -> str:
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Staff Leave Planner</dc:title><dc:creator>Codex</dc:creator><cp:lastModifiedBy>Codex</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified></cp:coreProperties>'''


def app() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Codex</Application></Properties>'''


parts = {
    "[Content_Types].xml": content_types(),
    "_rels/.rels": root_rels(),
    "docProps/app.xml": app(),
    "docProps/core.xml": core(),
    "xl/workbook.xml": workbook(),
    "xl/_rels/workbook.xml.rels": workbook_rels(),
    "xl/styles.xml": styles(),
    "xl/theme/theme1.xml": theme(),
    "xl/worksheets/sheet1.xml": setup_sheet(),
    "xl/worksheets/sheet2.xml": calendar_sheet(),
    "xl/worksheets/sheet3.xml": requests_sheet(),
    "xl/worksheets/sheet4.xml": summary_sheet(),
    "xl/worksheets/sheet5.xml": dashboard_sheet(),
}

def cache_zero_formulas(xml: str) -> str:
    return re.sub(r'(<c r="[^"]+" s="\d+"><f>.*?</f>)</c>', r'\1<v>0</v></c>', xml)


parts["xl/worksheets/sheet4.xml"] = cache_zero_formulas(parts["xl/worksheets/sheet4.xml"])
parts["xl/worksheets/sheet5.xml"] = cache_zero_formulas(parts["xl/worksheets/sheet5.xml"])

if OUT.exists():
    OUT.replace(BACKUP)

with ZipFile(OUT, "w", ZIP_DEFLATED) as zf:
    for name, data in parts.items():
        zf.writestr(name, data)

print(f"Created {OUT.resolve()}")
print(f"Backup {BACKUP.resolve()}")
