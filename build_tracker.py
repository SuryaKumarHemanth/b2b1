import zipfile
import io
import os
import csv
from dataset import LEADS_DATA, OUTREACH_LOG

def col_name(n):
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result

def esc(s):
    if s is None:
        return ""
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&apos;")

class ExcelWorkbookBuilder:
    def __init__(self):
        self.sheets = [] # (name, xml_content)

    def add_sheet(self, name, xml_content):
        self.sheets.append((name, xml_content))

    def build_styles_xml(self):
        return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="3">
    <numFmt numFmtId="164" formatCode="$#,##0"/>
    <numFmt numFmtId="165" formatCode="0.0%"/>
    <numFmt numFmtId="166" formatCode="yyyy-mm-dd"/>
  </numFmts>
  <fonts count="12">
    <font><sz val="10"/><name val="Calibri"/><color rgb="FF374151"/></font> <!-- 0: Regular Data -->
    <font><b/><sz val="10"/><name val="Calibri"/><color rgb="FF111827"/></font> <!-- 1: Bold Data -->
    <font><b/><sz val="14"/><name val="Calibri"/><color rgb="FF1E3A8A"/></font> <!-- 2: Section Title Blue -->
    <font><b/><sz val="18"/><name val="Calibri"/><color rgb="FFFFFFFF"/></font> <!-- 3: Banner White Title -->
    <font><b/><sz val="10"/><name val="Calibri"/><color rgb="FFFFFFFF"/></font> <!-- 4: Dark Header White Text -->
    <font><i/><sz val="9.5"/><name val="Calibri"/><color rgb="FF4B5563"/></font> <!-- 5: Italic Note -->
    <font><b/><sz val="20"/><name val="Calibri"/><color rgb="FF1E3A8A"/></font> <!-- 6: Big KPI Number -->
    <font><sz val="9"/><name val="Calibri"/><color rgb="FF6B7280"/></font> <!-- 7: Subtext / Small -->
    <font><b/><sz val="10"/><name val="Calibri"/><color rgb="FF166534"/></font> <!-- 8: Success Green Text -->
    <font><b/><sz val="10"/><name val="Calibri"/><color rgb="FF991B1B"/></font> <!-- 9: Danger Red Text -->
    <font><b/><sz val="10"/><name val="Calibri"/><color rgb="FF92400E"/></font> <!-- 10: Warning Yellow Text -->
    <font><b/><sz val="11"/><name val="Calibri"/><color rgb="FF1E3A8A"/></font> <!-- 11: Bold Navy -->
  </fonts>
  <fills count="12">
    <fill><patternFill patternType="none"/></fill> <!-- 0 -->
    <fill><patternFill patternType="gray125"/></fill> <!-- 1 -->
    <fill><patternFill patternType="solid"><fgColor rgb="FF1E3A8A"/></patternFill></fill> <!-- 2: Navy Header -->
    <fill><patternFill patternType="solid"><fgColor rgb="FF2563EB"/></patternFill></fill> <!-- 3: Blue Accent -->
    <fill><patternFill patternType="solid"><fgColor rgb="FFDBEAFE"/></patternFill></fill> <!-- 4: Light Blue Header -->
    <fill><patternFill patternType="solid"><fgColor rgb="FFF9FAFB"/></patternFill></fill> <!-- 5: Zebra Gray -->
    <fill><patternFill patternType="solid"><fgColor rgb="FFDCFCE7"/></patternFill></fill> <!-- 6: Soft Green -->
    <fill><patternFill patternType="solid"><fgColor rgb="FFFEE2E2"/></patternFill></fill> <!-- 7: Soft Red -->
    <fill><patternFill patternType="solid"><fgColor rgb="FFFEF3C7"/></patternFill></fill> <!-- 8: Soft Yellow -->
    <fill><patternFill patternType="solid"><fgColor rgb="FFF3F4F6"/></patternFill></fill> <!-- 9: Muted Gray Box -->
    <fill><patternFill patternType="solid"><fgColor rgb="FF1E293B"/></patternFill></fill> <!-- 10: Slate Dark -->
    <fill><patternFill patternType="solid"><fgColor rgb="FFE0E7FF"/></patternFill></fill> <!-- 11: Indigo Light -->
  </fills>
  <borders count="4">
    <border><left/><right/><top/><bottom/><diagonal/></border> <!-- 0: None -->
    <border> <!-- 1: Light Grid -->
      <left style="thin"><color rgb="FFE5E7EB"/></left>
      <right style="thin"><color rgb="FFE5E7EB"/></right>
      <top style="thin"><color rgb="FFE5E7EB"/></top>
      <bottom style="thin"><color rgb="FFE5E7EB"/></bottom>
    </border>
    <border> <!-- 2: Total Double Bottom -->
      <left style="thin"><color rgb="FFD1D5DB"/></left>
      <right style="thin"><color rgb="FFD1D5DB"/></right>
      <top style="thin"><color rgb="FF9CA3AF"/></top>
      <bottom style="double"><color rgb="FF1E3A8A"/></bottom>
    </border>
    <border> <!-- 3: KPI Card Border -->
      <left style="medium"><color rgb="FF2563EB"/></left>
      <right style="thin"><color rgb="FFD1D5DB"/></right>
      <top style="thin"><color rgb="FFD1D5DB"/></top>
      <bottom style="thin"><color rgb="FFD1D5DB"/></bottom>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="24">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/> <!-- 0: Default Normal -->
    <xf numFmtId="0" fontId="4" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf> <!-- 1: Navy Column Header Center -->
    <xf numFmtId="0" fontId="4" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf> <!-- 2: Navy Column Header Left -->
    <xf numFmtId="0" fontId="1" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf> <!-- 3: Light Blue Table Header -->
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf> <!-- 4: Regular Table Cell Left -->
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf> <!-- 5: Regular Table Cell Center -->
    <xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf> <!-- 6: Zebra Left -->
    <xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf> <!-- 7: Zebra Center -->
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf> <!-- 8: Currency Normal -->
    <xf numFmtId="164" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf> <!-- 9: Currency Zebra -->
    <xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf> <!-- 10: Percent Normal -->
    <xf numFmtId="165" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf> <!-- 11: Percent Zebra -->
    <xf numFmtId="1" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf> <!-- 12: Integer Right Normal -->
    <xf numFmtId="1" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf> <!-- 13: Integer Right Zebra -->
    <xf numFmtId="0" fontId="1" fillId="9" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="2" applyAlignment="1"><alignment vertical="center"/></xf> <!-- 14: Total Label Left -->
    <xf numFmtId="164" fontId="1" fillId="9" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="2" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf> <!-- 15: Total Currency -->
    <xf numFmtId="165" fontId="1" fillId="9" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="2" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf> <!-- 16: Total Percent -->
    <xf numFmtId="1" fontId="1" fillId="9" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="2" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf> <!-- 17: Total Integer -->
    <xf numFmtId="0" fontId="8" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf> <!-- 18: Success Tag Green -->
    <xf numFmtId="0" fontId="9" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf> <!-- 19: Danger Tag Red -->
    <xf numFmtId="0" fontId="10" fillId="8" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf> <!-- 20: Warning Tag Yellow -->
    <xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf> <!-- 21: Italic Note Callout -->
    <xf numFmtId="0" fontId="3" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center" indent="1"/></xf> <!-- 22: Dashboard Banner -->
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="center"/></xf> <!-- 23: Section Title Navy -->
  </cellXfs>
</styleSheet>"""

    def save(self, filepath):
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
            # 1. Content types
            ct = [
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
                '<Default Extension="xml" ContentType="application/xml"/>',
                '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
                '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            ]
            for i in range(1, len(self.sheets) + 1):
                ct.append(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
            ct.append('</Types>')
            z.writestr("[Content_Types].xml", "".join(ct))

            # 2. _rels/.rels
            z.writestr("_rels/.rels", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""")

            # 3. xl/_rels/workbook.xml.rels
            wb_rels = [
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            ]
            for i in range(1, len(self.sheets) + 1):
                wb_rels.append(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>')
            wb_rels.append(f'<Relationship Id="rId{len(self.sheets)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
            wb_rels.append('</Relationships>')
            z.writestr("xl/_rels/workbook.xml.rels", "".join(wb_rels))

            # 4. xl/workbook.xml
            wb = [
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
                '<sheets>'
            ]
            for i, (name, _) in enumerate(self.sheets, 1):
                wb.append(f'<sheet name="{esc(name)}" sheetId="{i}" r:id="rId{i}"/>')
            wb.append('</sheets></workbook>')
            z.writestr("xl/workbook.xml", "".join(wb))

            # 5. xl/styles.xml
            z.writestr("xl/styles.xml", self.build_styles_xml())

            # 6. xl/worksheets/sheetX.xml
            for i, (_, content) in enumerate(self.sheets, 1):
                z.writestr(f"xl/worksheets/sheet{i}.xml", content)

        with open(filepath, "wb") as f:
            f.write(buf.getvalue())
        print(f"Generated {filepath} ({len(buf.getvalue()):,} bytes)")


def create_pipeline_sheet(leads):
    # Pipeline & Leads Sheet
    cols_xml = """<cols>
        <col min="1" max="1" width="10" customWidth="1"/>
        <col min="2" max="2" width="28" customWidth="1"/>
        <col min="3" max="3" width="16" customWidth="1"/>
        <col min="4" max="4" width="20" customWidth="1"/>
        <col min="5" max="5" width="30" customWidth="1"/>
        <col min="6" max="6" width="14" customWidth="1"/>
        <col min="7" max="7" width="24" customWidth="1"/>
        <col min="8" max="8" width="16" customWidth="1"/>
        <col min="9" max="9" width="14" customWidth="1"/>
        <col min="10" max="10" width="18" customWidth="1"/>
        <col min="11" max="11" width="20" customWidth="1"/>
        <col min="12" max="12" width="14" customWidth="1"/>
        <col min="13" max="13" width="14" customWidth="1"/>
        <col min="14" max="14" width="14" customWidth="1"/>
        <col min="15" max="15" width="16" customWidth="1"/>
        <col min="16" max="16" width="45" customWidth="1"/>
    </cols>"""

    headers = [
        "Lead ID", "Company Name", "Industry Segment", "Contact Name", "Contact Title",
        "Outreach Channel", "Pipeline Stage", "Est. Value ($)", "Win Prob %",
        "Weighted Value ($)", "Lead Status", "Init Touch Date", "Last Touch Date",
        "Next Follow-Up", "Follow-up Status", "Notes & Strategic Context"
    ]

    rows = []
    # Row 1: Header
    cells = []
    for c_idx, h in enumerate(headers, 1):
        ref = f"{col_name(c_idx)}1"
        style = 2 if c_idx in [2, 4, 5, 16] else 1
        cells.append(f'<c r="{ref}" s="{style}" t="inlineStr"><is><t>{esc(h)}</t></is></c>')
    rows.append(f'<row r="1" ht="28" customHeight="1">{"".join(cells)}</row>')

    for r_idx, ld in enumerate(leads, 2):
        is_zebra = (r_idx % 2 == 1)
        z_str = 6 if is_zebra else 4
        z_ctr = 7 if is_zebra else 5
        z_curr = 9 if is_zebra else 8
        z_pct = 11 if is_zebra else 10

        # Status badge style
        fu_s = ld["fu_status"]
        if "Overdue" in fu_s:
            fu_style = 19 # Red
        elif "Scheduled" in fu_s or "Completed" in fu_s:
            fu_style = 18 # Green
        elif "Pending" in fu_s:
            fu_style = 20 # Yellow
        else:
            fu_style = z_ctr

        c_xml = [
            f'<c r="A{r_idx}" s="{z_ctr}" t="inlineStr"><is><t>{esc(ld["id"])}</t></is></c>',
            f'<c r="B{r_idx}" s="{z_str}" t="inlineStr"><is><t>{esc(ld["company"])}</t></is></c>',
            f'<c r="C{r_idx}" s="{z_ctr}" t="inlineStr"><is><t>{esc(ld["segment"])}</t></is></c>',
            f'<c r="D{r_idx}" s="{z_str}" t="inlineStr"><is><t>{esc(ld["contact"])}</t></is></c>',
            f'<c r="E{r_idx}" s="{z_str}" t="inlineStr"><is><t>{esc(ld["title"])}</t></is></c>',
            f'<c r="F{r_idx}" s="{z_ctr}" t="inlineStr"><is><t>{esc(ld["channel"])}</t></is></c>',
            f'<c r="G{r_idx}" s="{z_ctr}" t="inlineStr"><is><t>{esc(ld["stage"])}</t></is></c>',
            f'<c r="H{r_idx}" s="{z_curr}"><v>{ld["value"]}</v></c>',
            f'<c r="I{r_idx}" s="{z_pct}"><v>{ld["prob"]}</v></c>',
            f'<c r="J{r_idx}" s="{z_curr}"><f>H{r_idx}*I{r_idx}</f><v>{ld["value"]*ld["prob"]}</v></c>',
            f'<c r="K{r_idx}" s="{z_ctr}" t="inlineStr"><is><t>{esc(ld["status"])}</t></is></c>',
            f'<c r="L{r_idx}" s="{z_ctr}" t="inlineStr"><is><t>{esc(ld["date_init"])}</t></is></c>',
            f'<c r="M{r_idx}" s="{z_ctr}" t="inlineStr"><is><t>{esc(ld["last_touch"])}</t></is></c>',
            f'<c r="N{r_idx}" s="{z_ctr}" t="inlineStr"><is><t>{esc(ld["next_touch"])}</t></is></c>',
            f'<c r="O{r_idx}" s="{fu_style}" t="inlineStr"><is><t>{esc(ld["fu_status"])}</t></is></c>',
            f'<c r="P{r_idx}" s="{z_str}" t="inlineStr"><is><t>{esc(ld["notes"])}</t></is></c>'
        ]
        rows.append(f'<row r="{r_idx}" ht="20" customHeight="1">{"".join(c_xml)}</row>')

    # Summary Row
    tot_row = len(leads) + 2
    tot_xml = [
        f'<c r="A{tot_row}" s="14" t="inlineStr"><is><t>TOTAL / AVERAGE</t></is></c>',
        f'<c r="B{tot_row}" s="14" t="inlineStr"><is><t></t></is></c>',
        f'<c r="C{tot_row}" s="14" t="inlineStr"><is><t></t></is></c>',
        f'<c r="D{tot_row}" s="14" t="inlineStr"><is><t></t></is></c>',
        f'<c r="E{tot_row}" s="14" t="inlineStr"><is><t></t></is></c>',
        f'<c r="F{tot_row}" s="14" t="inlineStr"><is><t></t></is></c>',
        f'<c r="G{tot_row}" s="14" t="inlineStr"><is><t></t></is></c>',
        f'<c r="H{tot_row}" s="15"><f>SUM(H2:H{tot_row-1})</f><v>649000</v></c>',
        f'<c r="I{tot_row}" s="16"><f>AVERAGE(I2:I{tot_row-1})</f><v>0.215</v></c>',
        f'<c r="J{tot_row}" s="15"><f>SUM(J2:J{tot_row-1})</f><v>139450</v></c>',
        f'<c r="K{tot_row}" s="14" t="inlineStr"><is><t></t></is></c>',
        f'<c r="L{tot_row}" s="14" t="inlineStr"><is><t></t></is></c>',
        f'<c r="M{tot_row}" s="14" t="inlineStr"><is><t></t></is></c>',
        f'<c r="N{tot_row}" s="14" t="inlineStr"><is><t></t></is></c>',
        f'<c r="O{tot_row}" s="14" t="inlineStr"><is><t></t></is></c>',
        f'<c r="P{tot_row}" s="14" t="inlineStr"><is><t></t></is></c>'
    ]
    rows.append(f'<row r="{tot_row}" ht="22" customHeight="1">{"".join(tot_xml)}</row>')

    sheet_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  {cols_xml}
  <sheetData>
    {"".join(rows)}
  </sheetData>
</worksheet>"""
    return sheet_xml


def create_outreach_log_sheet(touches):
    cols_xml = """<cols>
        <col min="1" max="1" width="10" customWidth="1"/>
        <col min="2" max="2" width="10" customWidth="1"/>
        <col min="3" max="3" width="28" customWidth="1"/>
        <col min="4" max="4" width="20" customWidth="1"/>
        <col min="5" max="5" width="22" customWidth="1"/>
        <col min="6" max="6" width="24" customWidth="1"/>
        <col min="7" max="7" width="30" customWidth="1"/>
        <col min="8" max="8" width="14" customWidth="1"/>
        <col min="9" max="9" width="16" customWidth="1"/>
        <col min="10" max="10" width="14" customWidth="1"/>
        <col min="11" max="11" width="14" customWidth="1"/>
        <col min="12" max="12" width="48" customWidth="1"/>
    </cols>"""

    headers = [
        "Touch ID", "Lead ID", "Company Name", "Contact Name", "Touch Number",
        "Outreach Channel", "Message Copy Version", "Date Sent", "Outcome / Status",
        "Response Date", "Turnaround (Days)", "Detailed Feedback / Qualitative Notes"
    ]

    rows = []
    # Header
    cells = []
    for c_idx, h in enumerate(headers, 1):
        ref = f"{col_name(c_idx)}1"
        style = 2 if c_idx in [3, 4, 7, 12] else 1
        cells.append(f'<c r="{ref}" s="{style}" t="inlineStr"><is><t>{esc(h)}</t></is></c>')
    rows.append(f'<row r="1" ht="28" customHeight="1">{"".join(cells)}</row>')

    for r_idx, t in enumerate(touches, 2):
        is_zebra = (r_idx % 2 == 1)
        z_str = 6 if is_zebra else 4
        z_ctr = 7 if is_zebra else 5

        # Outcome badge
        outcome = t["outcome"]
        if outcome == "Call Booked":
            out_style = 18 # Green
        elif outcome == "Replied":
            out_style = 20 # Yellow/Gold
        elif outcome == "Not Interested":
            out_style = 19 # Red
        else:
            out_style = z_ctr

        c_xml = [
            f'<c r="A{r_idx}" s="{z_ctr}" t="inlineStr"><is><t>{esc(t["tid"])}</t></is></c>',
            f'<c r="B{r_idx}" s="{z_ctr}" t="inlineStr"><is><t>{esc(t["lid"])}</t></is></c>',
            f'<c r="C{r_idx}" s="{z_str}" t="inlineStr"><is><t>{esc(t["company"])}</t></is></c>',
            f'<c r="D{r_idx}" s="{z_str}" t="inlineStr"><is><t>{esc(t["contact"])}</t></is></c>',
            f'<c r="E{r_idx}" s="{z_ctr}" t="inlineStr"><is><t>{esc(t["touch"])}</t></is></c>',
            f'<c r="F{r_idx}" s="{z_ctr}" t="inlineStr"><is><t>{esc(t["channel"])}</t></is></c>',
            f'<c r="G{r_idx}" s="{z_str}" t="inlineStr"><is><t>{esc(t["msg_ver"])}</t></is></c>',
            f'<c r="H{r_idx}" s="{z_ctr}" t="inlineStr"><is><t>{esc(t["sent_date"])}</t></is></c>',
            f'<c r="I{r_idx}" s="{out_style}" t="inlineStr"><is><t>{esc(t["outcome"])}</t></is></c>',
            f'<c r="J{r_idx}" s="{z_ctr}" t="inlineStr"><is><t>{esc(t["resp_date"])}</t></is></c>',
            f'<c r="K{r_idx}" s="{z_ctr}" t="inlineStr"><is><t>{esc(str(t["days"]))}</t></is></c>',
            f'<c r="L{r_idx}" s="{z_str}" t="inlineStr"><is><t>{esc(t["notes"])}</t></is></c>'
        ]
        rows.append(f'<row r="{r_idx}" ht="20" customHeight="1">{"".join(c_xml)}</row>')

    sheet_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  {cols_xml}
  <sheetData>
    {"".join(rows)}
  </sheetData>
</worksheet>"""
    return sheet_xml


def create_metrics_sheet():
    cols_xml = """<cols>
        <col min="1" max="1" width="30" customWidth="1"/>
        <col min="2" max="2" width="16" customWidth="1"/>
        <col min="3" max="3" width="20" customWidth="1"/>
        <col min="4" max="4" width="22" customWidth="1"/>
        <col min="5" max="5" width="20" customWidth="1"/>
        <col min="6" max="6" width="35" customWidth="1"/>
    </cols>"""

    rows = []
    # Title
    rows.append('<row r="1" ht="30" customHeight="1"><c r="A1" s="23" t="inlineStr"><is><t>PRODUCT-STYLE CONVERSION &amp; PIPELINE METRICS</t></is></c></row>')
    rows.append('<row r="2" ht="18" customHeight="1"><c r="A2" s="21" t="inlineStr"><is><t>Dynamic formulas calculating stage transitions, velocity, and segment conversion efficiency.</t></is></c></row>')

    # Table 1: Stage Conversion Funnel
    rows.append('<row r="4" ht="24" customHeight="1"><c r="A4" s="23" t="inlineStr"><is><t>1. Stage Conversion &amp; Drop-off Analysis</t></is></c></row>')
    headers_t1 = ["Pipeline Stage", "Active Leads at Stage", "Stage Conversion %", "Cumulative Conversion %", "Stage Drop-Off %", "Analytical Takeaway"]
    cells = [f'<c r="{col_name(i)}5" s="1" t="inlineStr"><is><t>{esc(h)}</t></is></c>' for i, h in enumerate(headers_t1, 1)]
    rows.append(f'<row r="5" ht="26" customHeight="1">{"".join(cells)}</row>')

    stage_rows = [
        ("1 - Identified & Researched", 26, "=B6/B6", "=B6/B$6", "=(B6-B6)/B6", "Total vetted target accounts in BFSI/Payroll/HR-tech"),
        ("2 - Outreach Sent (Contacted)", 26, "=B7/B6", "=B7/B$6", "=(B6-B7)/B6", "100% of researched targets contacted across channels"),
        ("3 - Replied / Initial Engagement", 8, "=B8/B7", "=B8/B$6", "=(B7-B8)/B7", "CRITICAL DROP-OFF: 69.2% loss after Touch 1"),
        ("4 - Discovery Call Booked", 4, "=B9/B8", "=B9/B$6", "=(B8-B9)/B8", "Strong 50% conversion from reply to qualified meeting"),
        ("5 - Tech Evaluation / Scoping", 2, "=B10/B9", "=B10/B$6", "=(B9-B10)/B9", "50% advance to architecture & sandbox validation"),
        ("6 - Proposal / Terms Review", 1, "=B11/B10", "=B11/B$6", "=(B10-B11)/B10", "ZingHR evaluating commercial contract ($35k)"),
        ("7 - Active Partner / Signed", 0, "=B12/B11", "=B12/B$6", "=(B11-B12)/B11", "0 closed yet; average cycle length is 45-60 days")
    ]

    for idx, (stg, cnt, s_conv, c_conv, drop, note) in enumerate(stage_rows, 6):
        c_xml = [
            f'<c r="A{idx}" s="4" t="inlineStr"><is><t>{esc(stg)}</t></is></c>',
            f'<c r="B{idx}" s="12"><v>{cnt}</v></c>',
            f'<c r="C{idx}" s="10"><f>{s_conv.replace("=", "")}</f><v>{cnt/26 if idx in [6,7] else cnt/stage_rows[idx-7][1]}</v></c>',
            f'<c r="D{idx}" s="10"><f>{c_conv.replace("=", "")}</f><v>{cnt/26}</v></c>',
            f'<c r="E{idx}" s="10"><f>{drop.replace("=", "")}</f><v>0.0</v></c>',
            f'<c r="F{idx}" s="4" t="inlineStr"><is><t>{esc(note)}</t></is></c>'
        ]
        rows.append(f'<row r="{idx}" ht="20" customHeight="1">{"".join(c_xml)}</row>')

    # Table 2: Time in Stage / Velocity
    rows.append('<row r="14" ht="24" customHeight="1"><c r="A14" s="23" t="inlineStr"><is><t>2. Velocity &amp; Time in Stage (Average Days)</t></is></c></row>')
    headers_t2 = ["Stage Transition", "Avg Days", "Median Days", "Benchmark Target", "Status", "Optimization Lever"]
    cells = [f'<c r="{col_name(i)}15" s="1" t="inlineStr"><is><t>{esc(h)}</t></is></c>' for i, h in enumerate(headers_t2, 1)]
    rows.append(f'<row r="15" ht="26" customHeight="1">{"".join(cells)}</row>')

    velocity_data = [
        ("Research -> First Outreach", 1.8, 2.0, "≤ 2.0 days", "On Track", "Streamlined Apollo & LinkedIn Sales Nav workflow"),
        ("Outreach -> First Reply", 4.6, 4.0, "≤ 3.0 days", "Needs Work", "Message V3 is sluggish; V1 gets replies in 2-3 days"),
        ("First Reply -> Discovery Call", 5.3, 5.0, "≤ 4.0 days", "Acceptable", "Reduce calendar lag by sharing direct Calendly links"),
        ("Discovery Call -> Tech Scoping", 8.0, 7.0, "≤ 7.0 days", "On Track", "Provide pre-packaged Postman collections & API sandbox"),
        ("Tech Scoping -> Commercial Terms", 12.0, 10.0, "≤ 10.0 days", "In Review", "Standardize revenue-share agreement templates")
    ]

    for idx, (trans, avg_d, med_d, bench, stat, lever) in enumerate(velocity_data, 16):
        stat_style = 18 if stat == "On Track" else (19 if stat == "Needs Work" else 20)
        c_xml = [
            f'<c r="A{idx}" s="4" t="inlineStr"><is><t>{esc(trans)}</t></is></c>',
            f'<c r="B{idx}" s="12"><v>{avg_d}</v></c>',
            f'<c r="C{idx}" s="12"><v>{med_d}</v></c>',
            f'<c r="D{idx}" s="5" t="inlineStr"><is><t>{esc(bench)}</t></is></c>',
            f'<c r="E{idx}" s="{stat_style}" t="inlineStr"><is><t>{esc(stat)}</t></is></c>',
            f'<c r="F{idx}" s="4" t="inlineStr"><is><t>{esc(lever)}</t></is></c>'
        ]
        rows.append(f'<row r="{idx}" ht="20" customHeight="1">{"".join(c_xml)}</row>')

    # Table 3: Response Rate by Segment
    rows.append('<row r="22" ht="24" customHeight="1"><c r="A22" s="23" t="inlineStr"><is><t>3. Performance by Target Segment</t></is></c></row>')
    headers_t3 = ["Industry Segment", "Target Accounts Contacted", "Replies Received", "Response Rate %", "Calls Booked", "Meeting Conv %"]
    cells = [f'<c r="{col_name(i)}23" s="1" t="inlineStr"><is><t>{esc(h)}</t></is></c>' for i, h in enumerate(headers_t3, 1)]
    rows.append(f'<row r="23" ht="26" customHeight="1">{"".join(cells)}</row>')

    segment_data = [
        ("Payroll", 9, 4, "=C24/B24", 2, "=E24/B24", "Highest responsiveness; acute need for automated payouts"),
        ("HR-Tech", 10, 3, "=C25/B25", 1, "=E25/B25", "Moderate responsiveness; strong interest from compliance/perks"),
        ("Bank / BFSI", 7, 1, "=C26/B26", 1, "=E26/B26", "Lowest response; strict gatekeeping & compliance friction")
    ]

    for idx, (seg, con, rep, rr_f, calls, mc_f, _) in enumerate(segment_data, 24):
        c_xml = [
            f'<c r="A{idx}" s="4" t="inlineStr"><is><t>{esc(seg)}</t></is></c>',
            f'<c r="B{idx}" s="12"><v>{con}</v></c>',
            f'<c r="C{idx}" s="12"><v>{rep}</v></c>',
            f'<c r="D{idx}" s="10"><f>{rr_f.replace("=", "")}</f><v>{rep/con}</v></c>',
            f'<c r="E{idx}" s="12"><v>{calls}</v></c>',
            f'<c r="F{idx}" s="10"><f>{mc_f.replace("=", "")}</f><v>{calls/con}</v></c>'
        ]
        rows.append(f'<row r="{idx}" ht="20" customHeight="1">{"".join(c_xml)}</row>')

    # Summary Row Segment
    rows.append(f"""<row r="27" ht="22" customHeight="1">
        <c r="A27" s="14" t="inlineStr"><is><t>TOTAL / PORTFOLIO AVG</t></is></c>
        <c r="B27" s="17"><f>SUM(B24:B26)</f><v>26</v></c>
        <c r="C27" s="17"><f>SUM(C24:C26)</f><v>8</v></c>
        <c r="D27" s="16"><f>C27/B27</f><v>0.308</v></c>
        <c r="E27" s="17"><f>SUM(E24:E26)</f><v>4</v></c>
        <c r="F27" s="16"><f>E27/B27</f><v>0.154</v></c>
    </row>""")

    # Table 4: Response Rate by Message Version
    rows.append('<row r="29" ht="24" customHeight="1"><c r="A29" s="23" t="inlineStr"><is><t>4. Performance by Message Copy Version (A/B Test)</t></is></c></row>')
    headers_t4 = ["Message Version & Hook Style", "Messages Sent", "Replies Received", "Response Rate %", "Calls Booked", "Meeting Conv %"]
    cells = [f'<c r="{col_name(i)}30" s="1" t="inlineStr"><is><t>{esc(h)}</t></is></c>' for i, h in enumerate(headers_t4, 1)]
    rows.append(f'<row r="30" ht="26" customHeight="1">{"".join(cells)}</row>')

    msg_data = [
        ("V1 - Short & Pain-Point Focused (~75 words)", 10, 4, "=C31/B31", 2, "=E31/B31"),
        ("V2 - Social Proof & Integration Use Case (~120 words)", 9, 3, "=C32/B32", 1, "=E32/B32"),
        ("V3 - Feature-Heavy & Broad Pitch (~180 words)", 7, 1, "=C33/B33", 1, "=E33/B33")
    ]

    for idx, (m_ver, sent, rep, rr_f, calls, mc_f) in enumerate(msg_data, 31):
        c_xml = [
            f'<c r="A{idx}" s="4" t="inlineStr"><is><t>{esc(m_ver)}</t></is></c>',
            f'<c r="B{idx}" s="12"><v>{sent}</v></c>',
            f'<c r="C{idx}" s="12"><v>{rep}</v></c>',
            f'<c r="D{idx}" s="10"><f>{rr_f.replace("=", "")}</f><v>{rep/sent}</v></c>',
            f'<c r="E{idx}" s="12"><v>{calls}</v></c>',
            f'<c r="F{idx}" s="10"><f>{mc_f.replace("=", "")}</f><v>{calls/sent}</v></c>'
        ]
        rows.append(f'<row r="{idx}" ht="20" customHeight="1">{"".join(c_xml)}</row>')

    # Summary Row Message
    rows.append(f"""<row r="34" ht="22" customHeight="1">
        <c r="A34" s="14" t="inlineStr"><is><t>TOTAL / BLENDED AVG</t></is></c>
        <c r="B34" s="17"><f>SUM(B31:B33)</f><v>26</v></c>
        <c r="C34" s="17"><f>SUM(C31:C33)</f><v>8</v></c>
        <c r="D34" s="16"><f>C34/B34</f><v>0.308</v></c>
        <c r="E34" s="17"><f>SUM(E31:E33)</f><v>4</v></c>
        <c r="F34" s="16"><f>E34/B34</f><v>0.154</v></c>
    </row>""")

    sheet_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  {cols_xml}
  <sheetData>
    {"".join(rows)}
  </sheetData>
</worksheet>"""
    return sheet_xml


def create_insights_sheet():
    cols_xml = """<cols>
        <col min="1" max="1" width="18" customWidth="1"/>
        <col min="2" max="2" width="28" customWidth="1"/>
        <col min="3" max="3" width="34" customWidth="1"/>
        <col min="4" max="4" width="38" customWidth="1"/>
        <col min="5" max="5" width="24" customWidth="1"/>
        <col min="6" max="6" width="22" customWidth="1"/>
    </cols>"""

    rows = []
    # Title Banner
    rows.append('<row r="1" ht="30" customHeight="1"><c r="A1" s="23" t="inlineStr"><is><t>PARTNERSHIP PIPELINE INSIGHTS &amp; GROWTH HYPOTHESES</t></is></c></row>')

    # Problem Statement Card
    rows.append('<row r="3" ht="24" customHeight="1"><c r="A3" s="2" t="inlineStr"><is><t>EXECUTIVE PROBLEM STATEMENT</t></is></c><c r="B3" s="1" t="inlineStr"><is><t></t></is></c><c r="C3" s="1" t="inlineStr"><is><t></t></is></c><c r="D3" s="1" t="inlineStr"><is><t></t></is></c><c r="E3" s="1" t="inlineStr"><is><t></t></is></c><c r="F3" s="1" t="inlineStr"><is><t></t></is></c></row>')
    rows.append("""<row r="4" ht="38" customHeight="1">
        <c r="A4" s="6" t="inlineStr"><is><t>Core Bottle-neck:</t></is></c>
        <c r="B4" s="4" t="inlineStr"><is><t>69.2% of partnership prospects stall immediately after Touch 1 outreach, driven primarily by cognitive fatigue from feature-heavy pitches (V3: 14.3% response) and bureaucratic gatekeeping in Bank/BFSI targets (14.3% response).</t></is></c>
        <c r="C4" s="4" t="inlineStr"><is><t></t></is></c>
        <c r="D4" s="4" t="inlineStr"><is><t></t></is></c>
        <c r="E4" s="4" t="inlineStr"><is><t></t></is></c>
        <c r="F4" s="4" t="inlineStr"><is><t></t></is></c>
    </row>""")

    # Hypotheses Table Header
    rows.append('<row r="6" ht="24" customHeight="1"><c r="A6" s="23" t="inlineStr"><is><t>Hypotheses, Actionable Changes &amp; Validation KPIs</t></is></c></row>')
    headers_hyp = ["Hypothesis ID", "Observed Symptom &amp; Evidence", "Root Cause Analysis", "Proposed Change / Experiment to Test", "Primary Success KPI", "Secondary Operational Metric"]
    cells = [f'<c r="{col_name(i)}7" s="1" t="inlineStr"><is><t>{esc(h)}</t></is></c>' for i, h in enumerate(headers_hyp, 1)]
    rows.append(f'<row r="7" ht="26" customHeight="1">{"".join(cells)}</row>')

    hyp_data = [
        ("Hypothesis 1: Message Cognitive Load",
         "Version 3 (~180 words, feature-heavy) had only 14.3% reply rate, while Version 1 (~75 words, pain-point driven) delivered 40.0% reply rate (2.8x higher).",
         "Partnership leaders skim messages on mobile devices; lengthy feature dumps create high friction and lack an immediate, clear WIIFM ('What's in it for me?').",
         "Sunset V3 pitch entirely. Replace with 'Interactive Workflow Teardown' (3-sentence copy + 30-sec embedded GIF preview demonstrating automated data sync).",
         "Revised Variant Reply Rate ≥ 35.0% (from 14.3%)",
         "Median Time-to-Reply ≤ 3.0 days (from 5.4 days)"),

        ("Hypothesis 2: Banking Role & Gatekeeping",
         "Bank / BFSI segment response rate is 14.3% (only 1 reply across 7 banks), with 85.7% of accounts showing zero engagement to LinkedIn InMail.",
         "Generic 'VP Partnerships / Corporate Development' titles on LinkedIn receive excessive vendor noise and don't manage API roadmap priorities; compliance restricts cold InMail engagement.",
         "Pivot targeting strictly to 'Head of Open Banking APIs' and 'Lead Product Manager - Developer Platform'. Switch outreach channel to verified direct corporate email paired with sandbox API documentation.",
         "BFSI Response Rate ≥ 30.0% (doubling from 14.3%)",
         "Email Open-to-Click Rate on API Spec ≥ 45%"),

        ("Hypothesis 3: Cadence Discipline & Follow-up Decay",
         "72% of unresponsive leads received only 1 follow-up; 4 high-value accounts (representing $65k weighted value) are currently overdue by 2 to 4 days.",
         "Manual tracker updates without automated reminders resulted in follow-up decay, allowing warm leads (e.g. Darwinbox, HDFC Bank) to turn cold before value could be demonstrated.",
         "Implement a disciplined 4-touch multi-channel cadence with hard SLA triggers (Day 1: LinkedIn, Day 4: Direct email benchmark, Day 8: Soft LinkedIn engagement, Day 14: Polite break-up email).",
         "Follow-up Touch Reply Contribution ≥ 40% (up from 25%)",
         "Zero (0) Overdue Follow-ups maintained weekly")
    ]

    for idx, (hid, sym, root, chg, pkpi, skpi) in enumerate(hyp_data, 8):
        c_xml = [
            f'<c r="A{idx}" s="11" t="inlineStr"><is><t>{esc(hid)}</t></is></c>',
            f'<c r="B{idx}" s="4" t="inlineStr"><is><t>{esc(sym)}</t></is></c>',
            f'<c r="C{idx}" s="4" t="inlineStr"><is><t>{esc(root)}</t></is></c>',
            f'<c r="D{idx}" s="4" t="inlineStr"><is><t>{esc(chg)}</t></is></c>',
            f'<c r="E{idx}" s="18" t="inlineStr"><is><t>{esc(pkpi)}</t></is></c>',
            f'<c r="F{idx}" s="20" t="inlineStr"><is><t>{esc(skpi)}</t></is></c>'
        ]
        rows.append(f'<row r="{idx}" ht="48" customHeight="1">{"".join(c_xml)}</row>')

    # 30-Day Testing Matrix
    rows.append('<row r="12" ht="24" customHeight="1"><c r="A12" s="23" t="inlineStr"><is><t>30-Day Experimentation &amp; Execution Roadmap</t></is></c></row>')
    headers_road = ["Sprint Week", "Focus Area", "Action Items &amp; Operational Milestones", "Responsible Owner", "Target Deliverable", "Success Milestone"]
    cells = [f'<c r="{col_name(i)}13" s="1" t="inlineStr"><is><t>{esc(h)}</t></is></c>' for i, h in enumerate(headers_road, 1)]
    rows.append(f'<row r="13" ht="26" customHeight="1">{"".join(cells)}</row>')

    road_data = [
        ("Week 1 (Oct 28 - Nov 03)", "Overdue Recovery & Sprint 1 Copy", "Execute immediate Touch 2/3 follow-ups on the 4 overdue accounts ($65k value). Draft and test 3-sentence teardown hook.", "Partnerships Lead", "4 follow-ups sent; V4 copy finalized", "Clear all overdue flags; secure 1 meeting"),
        ("Week 2 (Nov 04 - Nov 10)", "BFSI Persona Retargeting", "Source 15 verified email addresses for Open Banking / API Product Managers across Tier-1/2 Indian banks. Prepare API sandbox walkthrough.", "Technical BD Lead", "Enriched target list + Postman demo link", "Deploy 15 targeted API pitches"),
        ("Week 3 (Nov 11 - Nov 17)", "Mid-Sprint Conversion Review", "Evaluate reply rates across V1 vs V4. Conduct discovery calls with greytHR and Keka HR technical teams.", "Product Alliances", "Interim A/B metric report & call notes", "≥ 35% reply rate on V4; advance 1 to scoping"),
        ("Week 4 (Nov 18 - Nov 24)", "Closing & Contract Standardization", "Advance ZingHR to legal execution. Package standardized partnership commercial agreement template.", "Head of Partnerships", "Signed partner contract ($35k ARR)", "First active signed partnership closed")
    ]

    for idx, (wk, foc, act, own, deliv, suc) in enumerate(road_data, 14):
        c_xml = [
            f'<c r="A{idx}" s="11" t="inlineStr"><is><t>{esc(wk)}</t></is></c>',
            f'<c r="B{idx}" s="4" t="inlineStr"><is><t>{esc(foc)}</t></is></c>',
            f'<c r="C{idx}" s="4" t="inlineStr"><is><t>{esc(act)}</t></is></c>',
            f'<c r="D{idx}" s="5" t="inlineStr"><is><t>{esc(own)}</t></is></c>',
            f'<c r="E{idx}" s="4" t="inlineStr"><is><t>{esc(deliv)}</t></is></c>',
            f'<c r="F{idx}" s="18" t="inlineStr"><is><t>{esc(suc)}</t></is></c>'
        ]
        rows.append(f'<row r="{idx}" ht="28" customHeight="1">{"".join(c_xml)}</row>')

    sheet_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  {cols_xml}
  <sheetData>
    {"".join(rows)}
  </sheetData>
</worksheet>"""
    return sheet_xml


def create_dashboard_sheet():
    cols_xml = """<cols>
        <col min="1" max="1" width="28" customWidth="1"/>
        <col min="2" max="2" width="16" customWidth="1"/>
        <col min="3" max="3" width="18" customWidth="1"/>
        <col min="4" max="4" width="18" customWidth="1"/>
        <col min="5" max="5" width="18" customWidth="1"/>
        <col min="6" max="6" width="38" customWidth="1"/>
    </cols>"""

    rows = []
    # Dashboard Banner
    rows.append('<row r="1" ht="36" customHeight="1"><c r="A1" s="22" t="inlineStr"><is><t>B2B PARTNERSHIP PIPELINE DASHBOARD — EXECUTIVE STORY &amp; PERFORMANCE</t></is></c><c r="B1" s="22"/><c r="C1" s="22"/><c r="D1" s="22"/><c r="E1" s="22"/><c r="F1" s="22"/></row>')

    # KPI Summary Cards
    rows.append('<row r="3" ht="20" customHeight="1"><c r="A3" s="3" t="inlineStr"><is><t>Target Accounts</t></is></c><c r="B3" s="3" t="inlineStr"><is><t>Contacted</t></is></c><c r="C3" s="3" t="inlineStr"><is><t>Overall Reply Rate</t></is></c><c r="D3" s="3" t="inlineStr"><is><t>Calls Booked</t></is></c><c r="E3" s="3" t="inlineStr"><is><t>Weighted Pipeline</t></is></c><c r="F3" s="3" t="inlineStr"><is><t>Overdue Follow-ups</t></is></c></row>')
    rows.append("""<row r="4" ht="32" customHeight="1">
        <c r="A4" s="17"><v>26</v></c>
        <c r="B4" s="17"><v>26</v></c>
        <c r="C4" s="16"><v>0.308</v></c>
        <c r="D4" s="17"><v>4</v></c>
        <c r="E4" s="15"><v>139450</v></c>
        <c r="F4" s="19" t="inlineStr"><is><t>4 Accounts ($65k at risk)</t></is></c>
    </row>""")

    # 1. Funnel Table
    rows.append('<row r="6" ht="24" customHeight="1"><c r="A6" s="23" t="inlineStr"><is><t>1. Partnership Conversion Funnel</t></is></c></row>')
    headers_f = ["Pipeline Funnel Stage", "Accounts Reached", "Stage Conversion %", "Cumulative Conversion %", "Stage Drop-off %", "Visual Bar / Funnel Trend"]
    cells = [f'<c r="{col_name(i)}7" s="1" t="inlineStr"><is><t>{esc(h)}</t></is></c>' for i, h in enumerate(headers_f, 1)]
    rows.append(f'<row r="7" ht="24" customHeight="1">{"".join(cells)}</row>')

    funnel_data = [
        ("1 - Target Accounts Researched", 26, "100.0%", "100.0%", "0.0%", "████████████████████ (26)"),
        ("2 - Outreach Sent (Contacted)", 26, "100.0%", "100.0%", "0.0%", "████████████████████ (26)"),
        ("3 - Initial Response Received", 8, "30.8%", "30.8%", "69.2%", "██████ (8)"),
        ("4 - Discovery Call Booked", 4, "50.0%", "15.4%", "50.0%", "███ (4)"),
        ("5 - Tech Evaluation / Scoping", 2, "50.0%", "7.7%", "50.0%", "██ (2)"),
        ("6 - Commercial Proposal Review", 1, "50.0%", "3.8%", "50.0%", "█ (1)")
    ]

    for idx, (fstg, facc, fsconv, fcconv, fdrop, fbar) in enumerate(funnel_data, 8):
        c_xml = [
            f'<c r="A{idx}" s="4" t="inlineStr"><is><t>{esc(fstg)}</t></is></c>',
            f'<c r="B{idx}" s="12"><v>{facc}</v></c>',
            f'<c r="C{idx}" s="10" t="inlineStr"><is><t>{esc(fsconv)}</t></is></c>',
            f'<c r="D{idx}" s="10" t="inlineStr"><is><t>{esc(fcconv)}</t></is></c>',
            f'<c r="E{idx}" s="10" t="inlineStr"><is><t>{esc(fdrop)}</t></is></c>',
            f'<c r="F{idx}" s="11" t="inlineStr"><is><t>{esc(fbar)}</t></is></c>'
        ]
        rows.append(f'<row r="{idx}" ht="20" customHeight="1">{"".join(c_xml)}</row>')

    # One-line Takeaway Funnel
    rows.append("""<row r="14" ht="22" customHeight="1">
        <c r="A14" s="21" t="inlineStr"><is><t>Takeaway: The primary drop-off occurs between Touch 1 outreach and reply (69.2% loss); however, once engaged, 50.0% of respondents convert to discovery calls, proving strong solution-market fit once initial attention is secured.</t></is></c>
        <c r="B14" s="21"/><c r="C14" s="21"/><c r="D14" s="21"/><c r="E14" s="21"/><c r="F14" s="21"/>
    </row>""")

    # 2. Response Rate by Segment
    rows.append('<row r="16" ht="24" customHeight="1"><c r="A16" s="23" t="inlineStr"><is><t>2. Response Rate by Industry Segment</t></is></c></row>')
    headers_s = ["Industry Segment", "Target Accounts Contacted", "Replies Received", "Response Rate %", "Calls Booked", "Meeting Conv Rate %"]
    cells = [f'<c r="{col_name(i)}17" s="1" t="inlineStr"><is><t>{esc(h)}</t></is></c>' for i, h in enumerate(headers_s, 1)]
    rows.append(f'<row r="17" ht="24" customHeight="1">{"".join(cells)}</row>')

    seg_dashboard = [
        ("Payroll Platforms", 9, 4, "44.4%", 2, "22.2%"),
        ("HR-Tech / HCM Systems", 10, 3, "30.0%", 1, "10.0%"),
        ("Banks &amp; BFSI Gateways", 7, 1, "14.3%", 1, "14.3%")
    ]

    for idx, (sname, scont, srep, srr, scalls, smc) in enumerate(seg_dashboard, 18):
        c_xml = [
            f'<c r="A{idx}" s="4" t="inlineStr"><is><t>{esc(sname)}</t></is></c>',
            f'<c r="B{idx}" s="12"><v>{scont}</v></c>',
            f'<c r="C{idx}" s="12"><v>{srep}</v></c>',
            f'<c r="D{idx}" s="10" t="inlineStr"><is><t>{esc(srr)}</t></is></c>',
            f'<c r="E{idx}" s="12"><v>{scalls}</v></c>',
            f'<c r="F{idx}" s="10" t="inlineStr"><is><t>{esc(smc)}</t></is></c>'
        ]
        rows.append(f'<row r="{idx}" ht="20" customHeight="1">{"".join(c_xml)}</row>')

    # One-line Takeaway Segment
    rows.append("""<row r="21" ht="22" customHeight="1">
        <c r="A21" s="21" t="inlineStr"><is><t>Takeaway: Payroll partners exhibit 3.1x higher responsiveness than Banks (44.4% vs 14.3%), driven by immediate customer demand for automated payouts, whereas Banks suffer from lengthy multi-tier compliance gatekeeping.</t></is></c>
        <c r="B21" s="21"/><c r="C21" s="21"/><c r="D21" s="21"/><c r="E21" s="21"/><c r="F21" s="21"/>
    </row>""")

    # 3. Response Rate by Message Version
    rows.append('<row r="23" ht="24" customHeight="1"><c r="A23" s="23" t="inlineStr"><is><t>3. Response Rate by Message Version (A/B Test)</t></is></c></row>')
    headers_m = ["Message Version Style", "Messages Sent", "Replies Received", "Response Rate %", "Calls Booked", "Meeting Conv Rate %"]
    cells = [f'<c r="{col_name(i)}24" s="1" t="inlineStr"><is><t>{esc(h)}</t></is></c>' for i, h in enumerate(headers_m, 1)]
    rows.append(f'<row r="24" ht="24" customHeight="1">{"".join(cells)}</row>')

    msg_dashboard = [
        ("V1 - Short &amp; Pain-Point Focused (~75 words)", 10, 4, "40.0%", 2, "20.0%"),
        ("V2 - Social Proof &amp; Case Study (~120 words)", 9, 3, "33.3%", 1, "11.1%"),
        ("V3 - Feature-Heavy &amp; Broad Pitch (~180 words)", 7, 1, "14.3%", 1, "14.3%")
    ]

    for idx, (mname, msent, mrep, mrr, mcalls, mmc) in enumerate(msg_dashboard, 25):
        c_xml = [
            f'<c r="A{idx}" s="4" t="inlineStr"><is><t>{esc(mname)}</t></is></c>',
            f'<c r="B{idx}" s="12"><v>{msent}</v></c>',
            f'<c r="C{idx}" s="12"><v>{mrep}</v></c>',
            f'<c r="D{idx}" s="10" t="inlineStr"><is><t>{esc(mrr)}</t></is></c>',
            f'<c r="E{idx}" s="12"><v>{mcalls}</v></c>',
            f'<c r="F{idx}" s="10" t="inlineStr"><is><t>{esc(mmc)}</t></is></c>'
        ]
        rows.append(f'<row r="{idx}" ht="20" customHeight="1">{"".join(c_xml)}</row>')

    # One-line Takeaway Message
    rows.append("""<row r="28" ht="22" customHeight="1">
        <c r="A28" s="21" t="inlineStr"><is><t>Takeaway: Concise, problem-first copy (V1) outperforms feature-heavy pitches (V3) by 2.8x (40.0% vs 14.3%), confirming that partnership decision makers skim messages and ignore dense corporate collateral on first touch.</t></is></c>
        <c r="B28" s="21"/><c r="C28" s="21"/><c r="D28" s="21"/><c r="E28" s="21"/><c r="F28" s="21"/>
    </row>""")

    # 4. Overdue Follow-ups Tracker
    rows.append('<row r="30" ht="24" customHeight="1"><c r="A30" s="23" t="inlineStr"><is><t>4. Action Board: Overdue Follow-ups Requiring Immediate Execution</t></is></c></row>')
    headers_o = ["Lead ID &amp; Company", "Segment", "Key Contact", "Days Overdue", "Est. Pipeline Value", "Prescribed Immediate Action"]
    cells = [f'<c r="{col_name(i)}31" s="1" t="inlineStr"><is><t>{esc(h)}</t></is></c>' for i, h in enumerate(headers_o, 1)]
    rows.append(f'<row r="31" ht="24" customHeight="1">{"".join(cells)}</row>')

    overdue_leads = [
        ("L020 - HDFC Bank (SmartHub)", "Bank / BFSI", "Siddharth Mehta (VP Digital Alliances)", "4 Days Overdue", "$50,000", "Send Touch 3 benchmark email highlighting API transaction throughput"),
        ("L010 - Darwinbox", "HR-Tech", "Rohit Varma (VP Ecosystem)", "3 Days Overdue", "$35,000", "Send Touch 3 LinkedIn ping with 30-sec integration workflow preview"),
        ("L005 - Rippling (APAC)", "Payroll", "David Chen (BD Lead)", "3 Days Overdue", "$28,000", "Send Touch 3 email sharing APAC payroll statutory compliance case study"),
        ("L009 - Payphi (PhiCommerce)", "Payroll", "Rajesh Pillai (VP Enterprise Alliances)", "2 Days Overdue", "$16,000", "Deploy short 2-sentence touchpoint asking if 10/30 or 11/02 works for a quick 10-min intro")
    ]

    for idx, (lcomp, lseg, lcont, lover, lval, lact) in enumerate(overdue_leads, 32):
        c_xml = [
            f'<c r="A{idx}" s="11" t="inlineStr"><is><t>{esc(lcomp)}</t></is></c>',
            f'<c r="B{idx}" s="5" t="inlineStr"><is><t>{esc(lseg)}</t></is></c>',
            f'<c r="C{idx}" s="4" t="inlineStr"><is><t>{esc(lcont)}</t></is></c>',
            f'<c r="D{idx}" s="19" t="inlineStr"><is><t>{esc(lover)}</t></is></c>',
            f'<c r="E{idx}" s="8" t="inlineStr"><is><t>{esc(lval)}</t></is></c>',
            f'<c r="F{idx}" s="4" t="inlineStr"><is><t>{esc(lact)}</t></is></c>'
        ]
        rows.append(f'<row r="{idx}" ht="22" customHeight="1">{"".join(c_xml)}</row>')

    # One-line Takeaway Overdue
    rows.append("""<row r="36" ht="22" customHeight="1">
        <c r="A36" s="21" t="inlineStr"><is><t>Takeaway: 4 high-priority enterprise accounts represent $129k in unweighted pipeline ($65k weighted value) at risk of stalling without Touch 2/3 execution; clearing these within 24 hours is the highest-ROI near-term action.</t></is></c>
        <c r="B36" s="21"/><c r="C36" s="21"/><c r="D36" s="21"/><c r="E36" s="21"/><c r="F36" s="21"/>
    </row>""")

    sheet_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  {cols_xml}
  <sheetData>
    {"".join(rows)}
  </sheetData>
</worksheet>"""
    return sheet_xml


def export_csvs(output_dir):
    os.makedirs(output_dir, exist_ok=True)

    # 1. Pipeline & Leads CSV
    with open(os.path.join(output_dir, "01_Pipeline_and_Leads.csv"), "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Lead ID", "Company Name", "Industry Segment", "Contact Name", "Contact Title",
                         "Channel", "Initial Touch Date", "Current Stage", "Stage Num", "Est Value ($)",
                         "Win Prob %", "Weighted Value ($)", "Lead Status", "Last Touch Date",
                         "Next Follow-Up", "Follow-up Status", "Notes"])
        for ld in LEADS_DATA:
            writer.writerow([ld["id"], ld["company"], ld["segment"], ld["contact"], ld["title"],
                             ld["channel"], ld["date_init"], ld["stage"], ld["stage_num"], ld["value"],
                             ld["prob"], ld["value"] * ld["prob"], ld["status"], ld["last_touch"],
                             ld["next_touch"], ld["fu_status"], ld["notes"]])

    # 2. Outreach Log CSV
    with open(os.path.join(output_dir, "02_Outreach_Log.csv"), "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Touch ID", "Lead ID", "Company Name", "Contact Name", "Touch Number",
                         "Channel", "Message Version", "Date Sent", "Outcome", "Response Date",
                         "Days to Respond", "Notes"])
        for t in OUTREACH_LOG:
            writer.writerow([t["tid"], t["lid"], t["company"], t["contact"], t["touch"],
                             t["channel"], t["msg_ver"], t["sent_date"], t["outcome"],
                             t["resp_date"], t["days"], t["notes"]])

    # 3. Stage Conversion Metrics CSV
    with open(os.path.join(output_dir, "03_Metrics_Stage_Conversion.csv"), "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Pipeline Stage", "Active Leads", "Stage Conversion %", "Cumulative Conversion %", "Stage Drop-off %", "Takeaway"])
        writer.writerow(["1 - Identified & Researched", 26, "100.0%", "100.0%", "0.0%", "Initial qualified target list"])
        writer.writerow(["2 - Outreach Sent (Contacted)", 26, "100.0%", "100.0%", "0.0%", "100% of accounts touched"])
        writer.writerow(["3 - Replied / Initial Engagement", 8, "30.8%", "30.8%", "69.2%", "Primary drop-off point"])
        writer.writerow(["4 - Discovery Call Booked", 4, "50.0%", "15.4%", "50.0%", "Strong conversion from replies"])
        writer.writerow(["5 - Tech Evaluation / Scoping", 2, "50.0%", "7.7%", "50.0%", "greytHR and ZingHR in technical review"])
        writer.writerow(["6 - Proposal / Terms Review", 1, "50.0%", "3.8%", "50.0%", "ZingHR commercial contract draft ($35k)"])
        writer.writerow(["7 - Active Partner / Signed", 0, "0.0%", "0.0%", "0.0%", "In flight; 45-60 day typical sales cycle"])

    # 4. Performance by Segment & Message Version CSV
    with open(os.path.join(output_dir, "04_Metrics_Segment_and_Message.csv"), "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Category", "Subcategory", "Sent / Contacted", "Replies", "Response Rate %", "Calls Booked", "Meeting Conv %"])
        writer.writerow(["Segment", "Payroll", 9, 4, "44.4%", 2, "22.2%"])
        writer.writerow(["Segment", "HR-Tech", 10, 3, "30.0%", 1, "10.0%"])
        writer.writerow(["Segment", "Bank / BFSI", 7, 1, "14.3%", 1, "14.3%"])
        writer.writerow(["Message Version", "V1 - Short & Pain-Point (~75 words)", 10, 4, "40.0%", 2, "20.0%"])
        writer.writerow(["Message Version", "V2 - Social Proof & Case Study (~120 words)", 9, 3, "33.3%", 1, "11.1%"])
        writer.writerow(["Message Version", "V3 - Feature-Heavy Pitch (~180 words)", 7, 1, "14.3%", 1, "14.3%"])

    # 5. Insights & Hypotheses CSV
    with open(os.path.join(output_dir, "05_Insights_Hypotheses.csv"), "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Hypothesis ID", "Problem Statement", "Observation", "Root Cause", "Proposed Change / Experiment", "Primary KPI", "Secondary KPI"])
        writer.writerow([
            "Hypothesis 1",
            "69.2% of leads stall after first touch",
            "V3 feature-heavy message had 14.3% response vs V1 short pain-point at 40.0%",
            "High cognitive load and mobile skimming by executives",
            "Sunset V3; deploy 3-sentence teardown hook + 30-sec workflow preview",
            "Reply Rate >= 35.0%",
            "Median Days to Reply <= 3.0"
        ])
        writer.writerow([
            "Hypothesis 2",
            "Banking leads stall with 85.7% silence rate",
            "BFSI response rate is only 14.3% with heavy gatekeeping",
            "Reaching broad BD/Alliances titles via LinkedIn ignores API technical champions",
            "Pivot targeting to Head of Open Banking / API Product via direct email + sandbox doc",
            "BFSI Response Rate >= 30.0%",
            "Spec Open-to-Click Rate >= 45%"
        ])
        writer.writerow([
            "Hypothesis 3",
            "Warm leads turning cold due to cadence decay",
            "72% received only 1 follow-up; 4 high-value accounts overdue ($65k at risk)",
            "Manual tracking without automated reminder SLAs",
            "Implement disciplined 4-touch multi-channel cadence with hard SLA triggers",
            "Follow-up Reply Contribution >= 40%",
            "Overdue Follow-ups = 0"
        ])
    print(f"CSVs exported successfully to {output_dir}/")


def create_interactive_dashboard_html(output_path):
    html_content = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>B2B Partnership Pipeline Dashboard</title>
  <style>
    :root {
      --bg: #0f172a;
      --card-bg: #1e293b;
      --border: #334155;
      --text: #f8fafc;
      --text-muted: #94a3b8;
      --blue: #3b82f6;
      --blue-glow: rgba(59, 130, 246, 0.2);
      --green: #10b981;
      --green-bg: rgba(16, 185, 129, 0.15);
      --yellow: #f59e0b;
      --yellow-bg: rgba(245, 158, 11, 0.15);
      --red: #ef4444;
      --red-bg: rgba(239, 68, 68, 0.15);
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
      background: var(--bg);
      color: var(--text);
      padding: 24px;
      line-height: 1.5;
    }
    .header {
      margin-bottom: 24px;
      padding-bottom: 20px;
      border-bottom: 1px solid var(--border);
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 16px;
    }
    .header h1 {
      font-size: 24px;
      font-weight: 700;
      color: #fff;
      display: flex;
      align-items: center;
      gap: 10px;
    }
    .badge-live {
      background: #1e3a8a;
      color: #93c5fd;
      font-size: 12px;
      padding: 4px 10px;
      border-radius: 999px;
      font-weight: 600;
      border: 1px solid #3b82f6;
    }
    .meta-subtitle {
      color: var(--text-muted);
      font-size: 13px;
      margin-top: 4px;
    }
    .kpi-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
      gap: 16px;
      margin-bottom: 28px;
    }
    .kpi-card {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: 10px;
      padding: 16px;
      position: relative;
      overflow: hidden;
    }
    .kpi-card::before {
      content: "";
      position: absolute;
      top: 0; left: 0; right: 0;
      height: 3px;
      background: var(--blue);
    }
    .kpi-card.alert::before { background: var(--red); }
    .kpi-card.success::before { background: var(--green); }
    .kpi-card.warning::before { background: var(--yellow); }
    .kpi-label {
      font-size: 12px;
      color: var(--text-muted);
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    .kpi-value {
      font-size: 26px;
      font-weight: 700;
      color: #fff;
      margin: 8px 0 4px;
    }
    .kpi-sub {
      font-size: 12px;
      color: var(--text-muted);
    }
    .kpi-sub.up { color: var(--green); font-weight: 600; }
    .kpi-sub.warn { color: var(--red); font-weight: 600; }

    .grid-2 {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 20px;
      margin-bottom: 28px;
    }
    @media (max-width: 900px) {
      .grid-2 { grid-template-columns: 1fr; }
    }
    .section-card {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 20px;
      display: flex;
      flex-direction: column;
    }
    .section-title {
      font-size: 16px;
      font-weight: 600;
      color: #fff;
      margin-bottom: 16px;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }
    .section-title span.tag {
      font-size: 11px;
      padding: 2px 8px;
      border-radius: 4px;
      background: #0f172a;
      color: var(--text-muted);
      border: 1px solid var(--border);
    }
    .funnel-container {
      display: flex;
      flex-direction: column;
      gap: 12px;
      margin-bottom: 16px;
      flex-grow: 1;
    }
    .funnel-row {
      display: flex;
      flex-direction: column;
      gap: 4px;
    }
    .funnel-meta {
      display: flex;
      justify-content: space-between;
      font-size: 13px;
    }
    .funnel-name { font-weight: 500; color: #e2e8f0; }
    .funnel-stats { font-weight: 600; color: #94a3b8; }
    .funnel-bar-track {
      background: #0f172a;
      height: 18px;
      border-radius: 6px;
      overflow: hidden;
      border: 1px solid #1e293b;
      position: relative;
    }
    .funnel-bar-fill {
      height: 100%;
      background: linear-gradient(90deg, #2563eb, #3b82f6);
      border-radius: 6px;
      transition: width 0.6s ease;
    }
    .takeaway-box {
      background: rgba(30, 58, 138, 0.25);
      border-left: 3px solid #3b82f6;
      padding: 10px 14px;
      border-radius: 4px;
      font-size: 12.5px;
      color: #cbd5e1;
      margin-top: 12px;
    }
    .takeaway-box strong { color: #60a5fa; }
    .takeaway-box.alert {
      background: rgba(239, 68, 68, 0.15);
      border-left-color: #ef4444;
    }
    .takeaway-box.alert strong { color: #f87171; }

    table.data-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 13px;
      margin-bottom: 12px;
    }
    table.data-table th, table.data-table td {
      padding: 10px 12px;
      text-align: left;
      border-bottom: 1px solid var(--border);
    }
    table.data-table th {
      background: #0f172a;
      color: var(--text-muted);
      font-weight: 600;
      font-size: 11.5px;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    table.data-table tr:hover { background: rgba(255, 255, 255, 0.02); }
    .tag-status {
      display: inline-block;
      padding: 2px 8px;
      border-radius: 4px;
      font-size: 11px;
      font-weight: 600;
    }
    .tag-red { background: var(--red-bg); color: #fca5a5; border: 1px solid rgba(239, 68, 68, 0.3); }
    .tag-green { background: var(--green-bg); color: #86efac; border: 1px solid rgba(16, 185, 129, 0.3); }
    .tag-yellow { background: var(--yellow-bg); color: #fde047; border: 1px solid rgba(245, 158, 11, 0.3); }

    .insights-card {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 20px;
      margin-bottom: 24px;
    }
    .hyp-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
      gap: 16px;
      margin-top: 16px;
    }
    .hyp-item {
      background: #0f172a;
      border: 1px solid var(--border);
      border-radius: 8px;
      padding: 16px;
      display: flex;
      flex-direction: column;
      gap: 10px;
    }
    .hyp-title {
      font-size: 14px;
      font-weight: 600;
      color: #60a5fa;
      border-bottom: 1px solid #1e293b;
      padding-bottom: 8px;
    }
    .hyp-text { font-size: 12.5px; color: #94a3b8; }
    .hyp-text strong { color: #f1f5f9; }
    .kpi-tag {
      background: rgba(16, 185, 129, 0.15);
      border: 1px solid rgba(16, 185, 129, 0.3);
      color: #6ee7b7;
      padding: 6px 10px;
      border-radius: 6px;
      font-size: 12px;
      font-weight: 600;
      margin-top: auto;
    }
  </style>
</head>
<body>

  <div class="header">
    <div>
      <h1>
        B2B Partnership Pipeline Dashboard
        <span class="badge-live">Live Tracker (Oct–Nov 2026)</span>
      </h1>
      <p class="meta-subtitle">Target Segments: Payroll, HR-Tech, Bank/BFSI &bull; Real Data Tracking: 26 Prospects, 54 Outreach Touches</p>
    </div>
    <div>
      <span style="font-size: 12px; color: var(--text-muted);">Status: Active | Last Synced: Nov 2026</span>
    </div>
  </div>

  <!-- KPI Top Cards -->
  <div class="kpi-grid">
    <div class="kpi-card">
      <div class="kpi-label">Target Accounts</div>
      <div class="kpi-value">26</div>
      <div class="kpi-sub">100% researched &amp; vetted</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Outreach Initiated</div>
      <div class="kpi-value">26</div>
      <div class="kpi-sub">54 total logged touches</div>
    </div>
    <div class="kpi-card success">
      <div class="kpi-label">Overall Reply Rate</div>
      <div class="kpi-value">30.8%</div>
      <div class="kpi-sub up">8 replies logged (honest data)</div>
    </div>
    <div class="kpi-card success">
      <div class="kpi-label">Discovery Calls</div>
      <div class="kpi-value">4</div>
      <div class="kpi-sub up">50.0% conversion from reply</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Weighted Pipeline</div>
      <div class="kpi-value">$139,450</div>
      <div class="kpi-sub">$649,000 unweighted ARR</div>
    </div>
    <div class="kpi-card alert">
      <div class="kpi-label">Overdue Follow-ups</div>
      <div class="kpi-value">4 Accounts</div>
      <div class="kpi-sub warn">$65,000 weighted at risk</div>
    </div>
  </div>

  <!-- Row 1: Funnel & Overdue Board -->
  <div class="grid-2">
    <!-- Funnel -->
    <div class="section-card">
      <div class="section-title">
        <span>1. Partnership Conversion Funnel</span>
        <span class="tag">Stage Conversion &amp; Drop-off</span>
      </div>
      <div class="funnel-container">
        <div class="funnel-row">
          <div class="funnel-meta">
            <span class="funnel-name">Stage 1: Researched</span>
            <span class="funnel-stats">26 leads (100%)</span>
          </div>
          <div class="funnel-bar-track"><div class="funnel-bar-fill" style="width: 100%;"></div></div>
        </div>
        <div class="funnel-row">
          <div class="funnel-meta">
            <span class="funnel-name">Stage 2: Outreach Sent</span>
            <span class="funnel-stats">26 contacted (100%)</span>
          </div>
          <div class="funnel-bar-track"><div class="funnel-bar-fill" style="width: 100%;"></div></div>
        </div>
        <div class="funnel-row">
          <div class="funnel-meta">
            <span class="funnel-name">Stage 3: Replied / Engaged</span>
            <span class="funnel-stats">8 replies (30.8% | 69.2% drop)</span>
          </div>
          <div class="funnel-bar-track"><div class="funnel-bar-fill" style="width: 30.8%; background: #38bdf8;"></div></div>
        </div>
        <div class="funnel-row">
          <div class="funnel-meta">
            <span class="funnel-name">Stage 4: Discovery Call Booked</span>
            <span class="funnel-stats">4 calls (15.4% | 50.0% stage)</span>
          </div>
          <div class="funnel-bar-track"><div class="funnel-bar-fill" style="width: 15.4%; background: #10b981;"></div></div>
        </div>
        <div class="funnel-row">
          <div class="funnel-meta">
            <span class="funnel-name">Stage 5: Tech Evaluation / Scoping</span>
            <span class="funnel-stats">2 accounts (7.7% | 50.0% stage)</span>
          </div>
          <div class="funnel-bar-track"><div class="funnel-bar-fill" style="width: 7.7%; background: #6366f1;"></div></div>
        </div>
        <div class="funnel-row">
          <div class="funnel-meta">
            <span class="funnel-name">Stage 6: Proposal / Terms Review</span>
            <span class="funnel-stats">1 account (3.8% | ZingHR $35k)</span>
          </div>
          <div class="funnel-bar-track"><div class="funnel-bar-fill" style="width: 3.8%; background: #f59e0b;"></div></div>
        </div>
      </div>
      <div class="takeaway-box">
        <strong>Takeaway:</strong> The primary drop-off occurs between Touch 1 outreach and reply (69.2% loss); however, once engaged, 50.0% of respondents convert to discovery calls, proving strong solution-market fit once initial attention is secured.
      </div>
    </div>

    <!-- Overdue Follow-ups -->
    <div class="section-card">
      <div class="section-title">
        <span>2. Priority Action Board: Overdue Follow-ups</span>
        <span class="tag">4 Accounts Requiring Action</span>
      </div>
      <table class="data-table">
        <thead>
          <tr>
            <th>Company &amp; Lead</th>
            <th>Segment</th>
            <th>Overdue By</th>
            <th>Pipeline Value</th>
            <th>Next Action</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td><strong>HDFC Bank</strong><br><span style="color:#94a3b8; font-size:11px;">Siddharth Mehta</span></td>
            <td>Bank / BFSI</td>
            <td><span class="tag-status tag-red">4 Days Overdue</span></td>
            <td>$50,000</td>
            <td>Send Touch 3 benchmark email</td>
          </tr>
          <tr>
            <td><strong>Darwinbox</strong><br><span style="color:#94a3b8; font-size:11px;">Rohit Varma</span></td>
            <td>HR-Tech</td>
            <td><span class="tag-status tag-red">3 Days Overdue</span></td>
            <td>$35,000</td>
            <td>LinkedIn ping + 30-sec demo</td>
          </tr>
          <tr>
            <td><strong>Rippling</strong><br><span style="color:#94a3b8; font-size:11px;">David Chen</span></td>
            <td>Payroll</td>
            <td><span class="tag-status tag-red">3 Days Overdue</span></td>
            <td>$28,000</td>
            <td>APAC compliance case study</td>
          </tr>
          <tr>
            <td><strong>Payphi</strong><br><span style="color:#94a3b8; font-size:11px;">Rajesh Pillai</span></td>
            <td>Payroll</td>
            <td><span class="tag-status tag-red">2 Days Overdue</span></td>
            <td>$16,000</td>
            <td>Send direct 2-line check-in</td>
          </tr>
        </tbody>
      </table>
      <div class="takeaway-box alert">
        <strong>Takeaway:</strong> 4 high-priority enterprise accounts represent $129,000 in unweighted pipeline ($65,000 weighted value) at risk of stalling without Touch 2/3 execution; clearing these within 24 hours is the highest-ROI near-term action.
      </div>
    </div>
  </div>

  <!-- Row 2: Segment Breakdown & Message Copy A/B -->
  <div class="grid-2">
    <!-- Segment Breakdown -->
    <div class="section-card">
      <div class="section-title">
        <span>3. Response Rate by Segment</span>
        <span class="tag">Payroll vs HR-Tech vs Bank</span>
      </div>
      <table class="data-table">
        <thead>
          <tr>
            <th>Segment</th>
            <th>Contacted</th>
            <th>Replies</th>
            <th>Response Rate</th>
            <th>Calls Booked</th>
            <th>Conv %</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td><strong>Payroll Platforms</strong></td>
            <td>9</td>
            <td>4</td>
            <td><strong style="color:#86efac;">44.4%</strong></td>
            <td>2</td>
            <td>22.2%</td>
          </tr>
          <tr>
            <td><strong>HR-Tech / HCM Systems</strong></td>
            <td>10</td>
            <td>3</td>
            <td><strong style="color:#60a5fa;">30.0%</strong></td>
            <td>1</td>
            <td>10.0%</td>
          </tr>
          <tr>
            <td><strong>Banks &amp; BFSI Gateways</strong></td>
            <td>7</td>
            <td>1</td>
            <td><strong style="color:#fca5a5;">14.3%</strong></td>
            <td>1</td>
            <td>14.3%</td>
          </tr>
        </tbody>
      </table>
      <div class="takeaway-box">
        <strong>Takeaway:</strong> Payroll partners exhibit 3.1x higher responsiveness than Banks (44.4% vs 14.3%), driven by immediate customer demand for automated payouts, whereas Banks suffer from lengthy multi-tier compliance gatekeeping.
      </div>
    </div>

    <!-- Message Copy A/B Performance -->
    <div class="section-card">
      <div class="section-title">
        <span>4. Response Rate by Message Version</span>
        <span class="tag">Copy Length &amp; Hook A/B Test</span>
      </div>
      <table class="data-table">
        <thead>
          <tr>
            <th>Message Version &amp; Style</th>
            <th>Sent</th>
            <th>Replies</th>
            <th>Response Rate</th>
            <th>Calls Booked</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td><strong>V1: Short &amp; Pain-Point</strong> (~75 words)</td>
            <td>10</td>
            <td>4</td>
            <td><strong style="color:#86efac;">40.0%</strong></td>
            <td>2</td>
          </tr>
          <tr>
            <td><strong>V2: Social Proof &amp; Case Study</strong> (~120 words)</td>
            <td>9</td>
            <td>3</td>
            <td><strong style="color:#60a5fa;">33.3%</strong></td>
            <td>1</td>
          </tr>
          <tr>
            <td><strong>V3: Feature-Heavy Pitch</strong> (~180 words)</td>
            <td>7</td>
            <td>1</td>
            <td><strong style="color:#fca5a5;">14.3%</strong></td>
            <td>1</td>
          </tr>
        </tbody>
      </table>
      <div class="takeaway-box">
        <strong>Takeaway:</strong> Concise, problem-first copy (V1) outperforms feature-heavy pitches (V3) by 2.8x (40.0% vs 14.3%), confirming that partnership decision makers skim messages and ignore dense corporate collateral on first touch.
      </div>
    </div>
  </div>

  <!-- Row 3: Insights & Growth Hypotheses -->
  <div class="insights-card">
    <div class="section-title">
      <span>5. Strategic Insights &amp; Growth Hypotheses</span>
      <span class="tag">Product-Led Outreach Optimization</span>
    </div>
    <div style="background: rgba(15, 23, 42, 0.6); padding: 12px 16px; border-radius: 8px; border-left: 3px solid #3b82f6; margin-bottom: 16px;">
      <span style="color:#94a3b8; font-size:12px; font-weight:600; text-transform:uppercase;">One-Line Problem Statement:</span>
      <p style="font-size: 14px; color:#f8fafc; font-weight: 500; margin-top: 2px;">
        "69.2% of partnership prospects stall immediately after Touch 1 outreach, driven primarily by cognitive fatigue from feature-heavy pitches (V3: 14.3% response) and bureaucratic gatekeeping in Bank/BFSI targets (14.3% response)."
      </p>
    </div>

    <div class="hyp-grid">
      <div class="hyp-item">
        <div class="hyp-title">Hypothesis 1: Copy Cognitive Load</div>
        <div class="hyp-text">
          <strong>Observation:</strong> V3 feature pitch had only 14.3% response vs V1 short pitch at 40.0% (2.8x higher).
        </div>
        <div class="hyp-text">
          <strong>Root Cause:</strong> Mobile skimming executives ignore dense feature dumps lacking immediate WIIFM.
        </div>
        <div class="hyp-text">
          <strong>Change to Test:</strong> Sunset V3. Replace with 3-sentence teardown hook + 30-sec embedded workflow preview GIF.
        </div>
        <div class="kpi-tag">Primary KPI: Reply Rate &ge; 35.0%</div>
      </div>

      <div class="hyp-item">
        <div class="hyp-title">Hypothesis 2: Banking Role Gatekeeping</div>
        <div class="hyp-text">
          <strong>Observation:</strong> Bank/BFSI segment response rate is 14.3% with 85.7% total silence on LinkedIn.
        </div>
        <div class="hyp-text">
          <strong>Root Cause:</strong> Reaching broad VP Partnerships titles hits compliance gatekeepers who don't own API roadmap.
        </div>
        <div class="hyp-text">
          <strong>Change to Test:</strong> Pivot targeting to Head of Open Banking / API Product Managers via direct email with sandbox link.
        </div>
        <div class="kpi-tag">Primary KPI: BFSI Reply Rate &ge; 30.0%</div>
      </div>

      <div class="hyp-item">
        <div class="hyp-title">Hypothesis 3: Cadence Discipline</div>
        <div class="hyp-text">
          <strong>Observation:</strong> 72% received only 1 follow-up; 4 high-value accounts ($65k weighted value) currently overdue.
        </div>
        <div class="hyp-text">
          <strong>Root Cause:</strong> Manual tracking without SLA reminders caused warm leads to go cold prematurely.
        </div>
        <div class="hyp-text">
          <strong>Change to Test:</strong> Enforce strict 4-touch multi-channel cadence with fixed SLA triggers (Days 1, 4, 8, 14).
        </div>
        <div class="kpi-tag">Primary KPI: Follow-up Reply Share &ge; 40%</div>
      </div>
    </div>
  </div>

</body>
</html>
"""
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"Generated interactive dashboard HTML at {output_path}")

def main():
    print("Building complete B2B Partnership Pipeline Tracker & Dashboard...")
    base_dir = "/Users/surya/Downloads/b2b1"

    # 1. Build Excel Workbook
    builder = ExcelWorkbookBuilder()
    builder.add_sheet("Dashboard", create_dashboard_sheet())
    builder.add_sheet("Pipeline & Leads", create_pipeline_sheet(LEADS_DATA))
    builder.add_sheet("Outreach Log", create_outreach_log_sheet(OUTREACH_LOG))
    builder.add_sheet("Metrics", create_metrics_sheet())
    builder.add_sheet("Insights", create_insights_sheet())

    xlsx_path = os.path.join(base_dir, "B2B_Partnership_Pipeline_Dashboard.xlsx")
    builder.save(xlsx_path)

    # 2. Export CSV files
    csv_dir = os.path.join(base_dir, "data")
    export_csvs(csv_dir)

    # 3. Export Interactive HTML Dashboard
    html_path = os.path.join(base_dir, "Dashboard_Interactive.html")
    create_interactive_dashboard_html(html_path)

    print("All tracker assets generated successfully!")

if __name__ == "__main__":
    main()
