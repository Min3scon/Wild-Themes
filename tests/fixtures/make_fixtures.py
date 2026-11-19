"""Builds the sample files the regression tests theme: realistic Word, PowerPoint and Excel
files with the features a theme must not break (tables, lists, links, headers, charts,
animations, formulas, validation, a macro part...). Run: python3 make_fixtures.py <outdir>"""
import io, os, sys, zipfile
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from pptx import Presentation
from pptx.util import Inches as PI, Pt as PPt
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.dml.color import RGBColor as PRGB
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.styles import Font
from PIL import Image

OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
os.makedirs(OUT, exist_ok=True)


def png(w=120, h=80, color=(30, 120, 200)):
    b = io.BytesIO(); Image.new("RGB", (w, h), color).save(b, "PNG"); b.seek(0); return b


def hyperlink(paragraph, url, text):
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    h = OxmlElement("w:hyperlink"); h.set(qn("r:id"), r_id)
    r = OxmlElement("w:r"); rpr = OxmlElement("w:rPr"); st = OxmlElement("w:rStyle"); st.set(qn("w:val"), "Hyperlink")
    rpr.append(st); r.append(rpr); t = OxmlElement("w:t"); t.text = text; r.append(t); h.append(r); paragraph._p.append(h)


def docx_report(path):
    d = Document()
    d.core_properties.title = "Quarterly report"
    sec = d.sections[0]
    sec.header.paragraphs[0].text = "Company confidential"
    sec.footer.paragraphs[0].text = "Page footer text"
    d.add_heading("Quarterly report", 0)
    d.add_paragraph("A subtitle line", style="Subtitle")
    d.add_heading("Summary", 1)
    p = d.add_paragraph("Body text with ")
    r = p.add_run("hard-coded Office blue"); r.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)  # old theme accent1 as direct colour
    p.add_run(", a ")
    r = p.add_run("fixed Calibri run"); r.font.name = "Calibri"
    p.add_run(", ")
    r = p.add_run("red text"); r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)  # not a theme colour: must stay red
    p.add_run(" and a highlight ")
    r = p.add_run("here"); r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    p = d.add_paragraph("See "); hyperlink(p, "https://example.com/", "the website"); p.add_run(" for more.")
    d.add_heading("Details", 2)
    for t in ["First point", "Second point", "Third point"]:
        d.add_paragraph(t, style="List Bullet")
    for t in ["Step one", "Step two"]:
        d.add_paragraph(t, style="List Number")
    d.add_heading("Numbers", 3)
    tb = d.add_table(rows=4, cols=3); tb.style = "Table Grid"
    for i, row in enumerate([["Region", "Q1", "Q2"], ["North", "10", "12"], ["South", "7", "9"], ["Total", "17", "21"]]):
        for j, v in enumerate(row): tb.cell(i, j).text = v
    d.add_paragraph("A quote that matters.", style="Quote")
    d.add_paragraph("An intense quote.", style="Intense Quote")
    d.add_picture(png(), width=Inches(1.5))
    d.add_page_break()
    d.add_heading("Second page", 1)
    d.add_paragraph("More text on page two. " * 30)
    new = d.add_section()
    new.orientation = 1; new.page_width, new.page_height = new.page_height, new.page_width
    d.add_heading("Landscape section", 1)
    d.add_paragraph("Wide content. " * 20)
    d.save(path)


def docx_plain(path):
    d = Document()
    d.add_paragraph("Just one plain paragraph, no styles at all.")
    d.save(path)


def pptx_deck(path):
    p = Presentation()
    p.core_properties.title = "Team update"
    s = p.slides.add_slide(p.slide_layouts[0]); s.shapes.title.text = "Team update"; s.placeholders[1].text = "November"
    s = p.slides.add_slide(p.slide_layouts[1]); s.shapes.title.text = "Agenda"
    tf = s.placeholders[1].text_frame; tf.text = "Wins"; tf.add_paragraph().text = "Risks"; tf.add_paragraph().text = "Next steps"
    box = s.shapes.add_shape(1, PI(6), PI(5), PI(2), PI(1)); box.fill.solid(); box.fill.fore_color.rgb = PRGB(0x44, 0x72, 0xC4)
    box.text_frame.text = "Callout"
    s = p.slides.add_slide(p.slide_layouts[5]); s.shapes.title.text = "Numbers"
    cd = CategoryChartData(); cd.categories = ["Q1", "Q2", "Q3"]; cd.add_series("Sales", (3, 5, 7)); cd.add_series("Costs", (2, 3, 4))
    s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, PI(1), PI(1.6), PI(8), PI(4.5), cd)
    s = p.slides.add_slide(p.slide_layouts[5]); s.shapes.title.text = "Table"
    t = s.shapes.add_table(3, 3, PI(1), PI(1.8), PI(8), PI(2)).table
    for i, row in enumerate([["A", "B", "C"], ["1", "2", "3"], ["4", "5", "6"]]):
        for j, v in enumerate(row): t.cell(i, j).text = v
    s = p.slides.add_slide(p.slide_layouts[6])
    s.shapes.add_picture(png(300, 200), PI(2), PI(2))
    p.save(path)
    # add an entrance animation to slide 2's callout and a transition to slide 3 (python-pptx can't)
    z = zipfile.ZipFile(path); files = {n: z.read(n) for n in z.namelist()}; z.close()
    s2 = files["ppt/slides/slide2.xml"].decode()
    import re
    sp_id = re.findall(r'<p:cNvPr id="(\d+)" name="Rectangle', s2)[0]
    timing = ('<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
              '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par>'
              '<p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="4" fill="hold">'
              '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="5" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect">'
              '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="6" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
              f'<p:tgtEl><p:spTgt spid="{sp_id}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
              f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="7" dur="500"/><p:tgtEl><p:spTgt spid="{sp_id}"/></p:tgtEl></p:cBhvr></p:animEffect>'
              '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn>'
              '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
              '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>')
    s2 = s2.replace("</p:clrMapOvr>", "</p:clrMapOvr>" + timing) if "</p:clrMapOvr>" in s2 else s2.replace("</p:sld>", timing + "</p:sld>")
    files["ppt/slides/slide2.xml"] = s2.encode()
    s3 = files["ppt/slides/slide3.xml"].decode()
    s3 = s3.replace("</p:clrMapOvr>", '</p:clrMapOvr><p:transition spd="slow"><p:wipe dir="d"/></p:transition>') if "</p:clrMapOvr>" in s3 else s3.replace("</p:sld>", '<p:transition spd="slow"><p:wipe dir="d"/></p:transition></p:sld>')
    files["ppt/slides/slide3.xml"] = s3.encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in files.items(): z.writestr(n, b)


def xlsx_book(path, macro=False):
    wb = Workbook(); ws = wb.active; ws.title = "Sales"
    ws.append(["Region", "Q1", "Q2", "Total"])
    for r in [["North", 10, 12], ["South", 7, 9], ["East", 4, 11], ["West", 8, 6]]: ws.append(r)
    for i in range(2, 6): ws[f"D{i}"] = f"=SUM(B{i}:C{i})"
    ws["A1"].style = "Headline 1" if "Headline 1" in wb.named_styles else "Title"
    ws["F1"] = "Old blue"; ws["F1"].font = Font(color="FF4472C4")
    ws["F2"] = "Red"; ws["F2"].font = Font(color="FFC00000")
    t = Table(displayName="SalesTable", ref="A1:D5"); t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True); ws.add_table(t)
    ws.conditional_formatting.add("B2:C5", ColorScaleRule(start_type="min", start_color="FFF8696B", end_type="max", end_color="FF63BE7B"))
    dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True); ws.add_data_validation(dv); dv.add("E2:E5")
    wb.defined_names["Totals"] = DefinedName("Totals", attr_text="Sales!$D$2:$D$5")
    ch = BarChart(); ch.title = "By region"; ch.add_data(Reference(ws, min_col=2, min_row=1, max_col=3, max_row=5), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=5)); ws.add_chart(ch, "H2")
    ws.freeze_panes = "A2"
    ws2 = wb.create_sheet("Notes"); ws2["A1"] = "Notes sheet"; ws2["A2"] = "=Sales!D2*2"
    wb.save(path)
    if macro:  # add a (dummy) VBA project part the theme must carry through untouched
        z = zipfile.ZipFile(path); files = {n: z.read(n) for n in z.namelist()}; z.close()
        files["xl/vbaProject.bin"] = b"\xd0\xcf\x11\xe0WILD-THEMES-TEST-VBA" + bytes(range(256))
        ct = files["[Content_Types].xml"].decode()
        ct = ct.replace("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
        ct = ct.replace("</Types>", '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>')
        files["[Content_Types].xml"] = ct.encode()
        rels = files["xl/_rels/workbook.xml.rels"].decode()
        rels = rels.replace("</Relationships>", '<Relationship Id="rIdVBA" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>')
        files["xl/_rels/workbook.xml.rels"] = rels.encode()
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for n, b in files.items(): z.writestr(n, b)


docx_report(os.path.join(OUT, "report.docx"))
docx_plain(os.path.join(OUT, "plain.docx"))
pptx_deck(os.path.join(OUT, "deck.pptx"))
xlsx_book(os.path.join(OUT, "book.xlsx"))
xlsx_book(os.path.join(OUT, "macro.xlsm"), macro=True)
# assets the tests pass in place of canvas-made pictures
Image.new("RGB", (794, 1123), (40, 10, 60)).save(os.path.join(OUT, "backdrop.jpg"), quality=80)
Image.new("RGBA", (600, 800), (255, 255, 255, 230)).save(os.path.join(OUT, "panel.png"))
print("fixtures in", OUT)
