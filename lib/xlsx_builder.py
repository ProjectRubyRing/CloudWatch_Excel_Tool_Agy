"""
xlsx_builder.py - Python standard library only XLSX generator.
Designed for RHEL 9.8 environments with zero external dependencies (no openpyxl, no pandas).
Produces beautiful, professional Excel reports with 'Meiryo UI' font, executive colors,
proper cell borders, auto-adjusted column widths, text wrapping, and freeze panes.
"""

import zipfile
import re
import datetime

# Stylesheet XML defining fonts, fills, borders, and cell formatting
STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="8">
    <font><name val="Meiryo UI"/><family val="2"/><sz val="9.5"/><color rgb="FF1E293B"/></font>
    <font><name val="Meiryo UI"/><b val="1"/><sz val="14"/><color rgb="FF1E3A8A"/></font>
    <font><name val="Meiryo UI"/><b val="1"/><sz val="10"/><color rgb="00FFFFFF"/></font>
    <font><name val="Meiryo UI"/><b val="1"/><sz val="10.5"/><color rgb="00FFFFFF"/></font>
    <font><name val="Meiryo UI"/><b val="1"/><sz val="9.5"/><color rgb="FF991B1B"/></font>
    <font><name val="Meiryo UI"/><b val="1"/><sz val="9"/><color rgb="FF334155"/></font>
    <font><name val="Meiryo UI"/><sz val="9"/><color rgb="FF0F172A"/></font>
    <font><name val="Meiryo UI"/><b val="1"/><sz val="11"/><color rgb="FF1E293B"/></font>
  </fonts>
  <fills count="8">
    <fill><patternFill/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1E3A8A"/><bgColor rgb="FF1E3A8A"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF334155"/><bgColor rgb="FF334155"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF8FAFC"/><bgColor rgb="FFF8FAFC"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFEE2E2"/><bgColor rgb="FFFEE2E2"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF1F5F9"/><bgColor rgb="FFF1F5F9"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF3730A3"/><bgColor rgb="FF3730A3"/></patternFill></fill>
  </fills>
  <borders count="4">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFCBD5E1"/></left>
      <right style="thin"><color rgb="FFCBD5E1"/></right>
      <top style="thin"><color rgb="FFCBD5E1"/></top>
      <bottom style="thin"><color rgb="FFCBD5E1"/></bottom>
    </border>
    <border>
      <left style="thin"><color rgb="FF94A3B8"/></left>
      <right style="thin"><color rgb="FF94A3B8"/></right>
      <top style="thin"><color rgb="FF94A3B8"/></top>
      <bottom style="medium"><color rgb="FF0F172A"/></bottom>
    </border>
    <border>
      <left style="thin"><color rgb="FFFCA5A5"/></left>
      <right style="thin"><color rgb="FFFCA5A5"/></right>
      <top style="thin"><color rgb="FFFCA5A5"/></top>
      <bottom style="thin"><color rgb="FFFCA5A5"/></bottom>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="19">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="5" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="6" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="4" fillId="5" borderId="3" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="4" fillId="5" borderId="3" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="4" fillId="5" borderId="3" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="2" fillId="7" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="7" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="7" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>"""

def sanitize_for_xml(text):
    """Safely cleans and escapes arbitrary text for OpenXML 1.0."""
    if text is None:
        return ""
    text = str(text)
    # Remove ANSI terminal escape sequences
    text = re.sub(r'\x1b\[[0-9;]*[a-zA-Z]', '', text)
    # Remove invalid XML 1.0 control characters (keep \t=0x09, \n=0x0A, \r=0x0D)
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', text)
    # Escape standard XML entities
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
    return text

def col_letter(col_idx):
    """Converts 1-based column index to Excel column letters (1 -> A, 27 -> AA)."""
    result = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result

def calc_visual_width(text):
    """Calculates approximate visual width of text (full-width vs half-width)."""
    if not text:
        return 0
    width = 0.0
    for ch in str(text):
        if ord(ch) > 127:
            width += 1.85  # Full-width / CJK
        else:
            width += 1.05  # Half-width ASCII
    return width

class XlsxWorksheet:
    """Represents an individual worksheet in the workbook."""
    def __init__(self, name, sheet_id, freeze_row=0):
        self.name = name
        self.sheet_id = sheet_id
        self.freeze_row = freeze_row
        self.rows = {}  # row_idx -> {col_idx: (value, style_id)}
        self.row_heights = {}  # row_idx -> height
        self.merges = []  # list of "A1:E1" strings
        self.col_widths = {}  # col_idx -> max width
        self.max_row = 0
        self.max_col = 0

    def set_cell(self, row, col, value, style_id=0):
        """Sets a cell value and style (1-based row and col)."""
        if row not in self.rows:
            self.rows[row] = {}
        self.rows[row][col] = (value, style_id)
        if row > self.max_row:
            self.max_row = row
        if col > self.max_col:
            self.max_col = col

        # Calculate column width tracking
        if value is not None:
            # If multi-line, take max line width
            lines = str(value).split('\n')
            max_line_w = max(calc_visual_width(line) for line in lines) if lines else 0
            cur_w = self.col_widths.get(col, 0)
            if max_line_w > cur_w:
                self.col_widths[col] = max_line_w

    def set_row_height(self, row, height):
        self.row_heights[row] = height

    def merge_cells(self, range_str):
        self.merges.append(range_str)

    def to_xml(self):
        """Generates the sheetN.xml document."""
        cols_xml = []
        if self.max_col > 0:
            cols_xml.append("  <cols>")
            for c in range(1, self.max_col + 1):
                raw_w = self.col_widths.get(c, 10)
                # Pad slightly for margins and readability
                calc_w = max(8.0, min(105.0, raw_w + 3.5))
                cols_xml.append(f'    <col min="{c}" max="{c}" width="{calc_w:.1f}" customWidth="1"/>')
            cols_xml.append("  </cols>")

        sheet_views = ['  <sheetViews>', '    <sheetView tabSelected="1" workbookViewId="0">']
        if self.freeze_row > 0:
            split_cell = f"A{self.freeze_row + 1}"
            sheet_views.append(
                f'      <pane ySplit="{self.freeze_row}" topLeftCell="{split_cell}" activePane="bottomLeft" state="frozen"/>'
            )
        sheet_views.append('    </sheetView>')
        sheet_views.append('  </sheetViews>')

        data_xml = ['  <sheetData>']
        sorted_row_indices = sorted(self.rows.keys())
        for r in sorted_row_indices:
            row_data = self.rows[r]
            ht_attr = ""
            if r in self.row_heights:
                ht_attr = f' ht="{self.row_heights[r]}" customHeight="1"'
            data_xml.append(f'    <row r="{r}"{ht_attr}>')
            sorted_cols = sorted(row_data.keys())
            for c in sorted_cols:
                val, style_id = row_data[c]
                ref = f"{col_letter(c)}{r}"
                s_attr = f' s="{style_id}"' if style_id != 0 else ''
                if val is None or val == "":
                    data_xml.append(f'      <c r="{ref}"{s_attr}/>')
                elif isinstance(val, (int, float)) and not isinstance(val, bool):
                    data_xml.append(f'      <c r="{ref}"{s_attr}><v>{val}</v></c>')
                else:
                    escaped_val = sanitize_for_xml(val)
                    data_xml.append(
                        f'      <c r="{ref}"{s_attr} t="inlineStr"><is><t xml:space="preserve">{escaped_val}</t></is></c>'
                    )
            data_xml.append('    </row>')
        data_xml.append('  </sheetData>')

        merges_xml = []
        if self.merges:
            merges_xml.append(f'  <mergeCells count="{len(self.merges)}">')
            for m in self.merges:
                merges_xml.append(f'    <mergeCell ref="{m}"/>')
            merges_xml.append('  </mergeCells>')

        xml_lines = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
            '\n'.join(sheet_views),
            '\n'.join(cols_xml) if cols_xml else '',
            '\n'.join(data_xml),
            '\n'.join(merges_xml) if merges_xml else '',
            '</worksheet>'
        ]
        return '\n'.join([line for line in xml_lines if line])

class XlsxWorkbook:
    """Manages sheets, packaging, and writing the final .xlsx zip file."""
    def __init__(self):
        self.sheets = []
        self.sheet_names_seen = set()

    def add_worksheet(self, title, freeze_row=0):
        # Sanitize sheet title (Excel forbids: \ / ? * [ ] : and max 31 chars)
        clean_title = re.sub(r'[\\/*?\[\]:]+', '_', title).strip('_').strip()
        if not clean_title:
            clean_title = "Sheet"
        # Truncate to 31 chars
        clean_title = clean_title[:31].rstrip('_')

        # Ensure uniqueness
        base_title = clean_title
        counter = 1
        while clean_title.lower() in self.sheet_names_seen:
            suffix = f"_{counter}"
            clean_title = base_title[:31 - len(suffix)] + suffix
            counter += 1

        self.sheet_names_seen.add(clean_title.lower())
        sheet_id = len(self.sheets) + 1
        ws = XlsxWorksheet(clean_title, sheet_id, freeze_row=freeze_row)
        self.sheets.append(ws)
        return ws

    def save(self, filepath):
        """Builds all OpenXML files and writes to a valid ZIP (.xlsx)."""
        content_types = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
            '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
            '  <Default Extension="xml" ContentType="application/xml"/>',
            '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
            '  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        ]
        for s in self.sheets:
            content_types.append(
                f'  <Override PartName="/xl/worksheets/sheet{s.sheet_id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            )
        content_types.append('</Types>')

        rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

        wb_rels = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
            '  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        ]
        for s in self.sheets:
            wb_rels.append(
                f'  <Relationship Id="rId{s.sheet_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{s.sheet_id}.xml"/>'
            )
        wb_rels.append('</Relationships>')

        wb_xml = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
            '  <bookViews><workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="12000"/></bookViews>',
            '  <sheets>'
        ]
        for s in self.sheets:
            wb_xml.append(
                f'    <sheet name="{sanitize_for_xml(s.name)}" sheetId="{s.sheet_id}" r:id="rId{s.sheet_id}"/>'
            )
        wb_xml.append('  </sheets>')
        wb_xml.append('</workbook>')

        with zipfile.ZipFile(filepath, 'w', compression=zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', '\n'.join(content_types))
            z.writestr('_rels/.rels', rels)
            z.writestr('xl/_rels/workbook.xml.rels', '\n'.join(wb_rels))
            z.writestr('xl/workbook.xml', '\n'.join(wb_xml))
            z.writestr('xl/styles.xml', STYLES_XML)
            for s in self.sheets:
                z.writestr(f'xl/worksheets/sheet{s.sheet_id}.xml', s.to_xml())

# Style ID Constants for convenient reference
STYLE_NORMAL = 0
STYLE_TITLE = 1
STYLE_META_LABEL = 2
STYLE_META_VALUE = 3
STYLE_HEADER_CENTER = 4
STYLE_HEADER_LEFT = 5
STYLE_STREAM_BANNER = 6
STYLE_DATA_CENTER = 7
STYLE_DATA_CENTER_ZEBRA = 8
STYLE_DATA_LEFT = 9
STYLE_DATA_LEFT_ZEBRA = 10
STYLE_DATA_MSG = 11
STYLE_DATA_MSG_ZEBRA = 12
STYLE_HIGHLIGHT_CENTER = 13
STYLE_HIGHLIGHT_LEFT = 14
STYLE_HIGHLIGHT_MSG = 15
STYLE_SUMMARY_HEADER_CENTER = 16
STYLE_SUMMARY_HEADER_LEFT = 17
STYLE_SECTION_HEADER = 18
