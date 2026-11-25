"""Generate small, realistic dummy Office/PDF files for testing sensitivity labels in a SharePoint test site.

Word and PowerPoint files are written as raw OOXML (no Office install used), so they carry
no sensitivity label from the machine that made them. Excel via openpyxl, PDF via reportlab.
Output: ./files/<Library>/<optional subfolder>/<file>. All fictional data.
"""
import os, random, zipfile, datetime as dt
from xml.sax.saxutils import escape
import openpyxl
from openpyxl.styles import Font, PatternFill
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

random.seed(42)
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "files")
ORG = "Harbourline Property Group"
FUNDS = ["Harbourline Diversified Fund", "Harbourline Industrial Trust", "Harbourline Retail Partners"]
ASSETS = ["12 Bayside Drive, Port Melbourne", "88 Logistics Way, Eastern Creek", "Westgate Plaza, Altona",
          "4 Quarry Road, Wacol", "Riverside Business Park, Rosehill"]
PEOPLE = ["Sam Ortega", "Priya Nair", "Tom Fitzgerald", "Mei Chen", "Jordan Blake", "Aisha Rahman",
          "Lachlan Burke", "Hannah Kowalski", "Daniel Osei", "Grace Liang"]
ROLES = ["Asset Manager", "Fund Accountant", "Development Manager", "Investor Relations Lead",
         "Leasing Executive", "Financial Controller", "HR Advisor", "Legal Counsel", "Analyst", "Project Coordinator"]

def path(lib, name, sub=None):
    d = os.path.join(OUT, lib, sub) if sub else os.path.join(OUT, lib)
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, name)

def core_xml(title):
    now = dt.datetime(2026, random.randint(1, 10), random.randint(1, 28), 9, 0).strftime("%Y-%m-%dT%H:%M:%SZ")
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f'<dc:title>{escape(title)}</dc:title><dc:creator>{random.choice(PEOPLE)}</dc:creator>'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified></cp:coreProperties>')

# ---------------- Word ----------------
def docx(p, title, paras):
    def para(text, style=None):
        ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ''
        return f'<w:p>{ppr}<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'
    body = para(title, "Title") + "".join(para(t[1], "Heading1") if isinstance(t, tuple) else para(t) for t in paras)
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
           f'{body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>')
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="22"/></w:rPr></w:style>'
              '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="40"/></w:rPr></w:style>'
              '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>'
              '</w:styles>')
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                   '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
                   '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>')
        z.writestr("_rels/.rels",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                   '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>')
        z.writestr("word/_rels/document.xml.rels",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
        z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", styles)
        z.writestr("docProps/core.xml", core_xml(title))

# ---------------- PowerPoint ----------------
NS_A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
NS_P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
NS_R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
THEME = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme {NS_A} name="Plain"><a:themeElements>'
         '<a:clrScheme name="Plain"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>'
         '<a:dk2><a:srgbClr val="1F2A44"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="1F6FB2"/></a:accent1>'
         '<a:accent2><a:srgbClr val="E07A1F"/></a:accent2><a:accent3><a:srgbClr val="5A9E3A"/></a:accent3><a:accent4><a:srgbClr val="8E44AD"/></a:accent4>'
         '<a:accent5><a:srgbClr val="2AA198"/></a:accent5><a:accent6><a:srgbClr val="B03A2E"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink>'
         '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>'
         '<a:fontScheme name="Plain"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
         '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
         '<a:fmtScheme name="Plain"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
         '<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
         '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
         '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>'
         '</a:fmtScheme></a:themeElements></a:theme>')
EMPTY_TREE = ('<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
              '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>')

def textbox(i, x, y, cx, cy, lines, size, bold=False):
    runs = "".join(f'<a:p><a:r><a:rPr lang="en-AU" sz="{size}" b="{1 if bold else 0}"/><a:t>{escape(l)}</a:t></a:r></a:p>' for l in lines)
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="TextBox {i}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>{runs}</p:txBody></p:sp>')

def pptx(p, title, slides):
    n = len(slides)
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        ov = "".join(f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' for i in range(1, n + 1))
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
                   '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
                   '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
                   '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
                   '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' + ov + '</Types>')
        z.writestr("_rels/.rels",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
                   '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>')
        sld_ids = "".join(f'<p:sldId id="{255 + i}" r:id="rId{i + 1}"/>' for i in range(1, n + 1))
        z.writestr("ppt/presentation.xml",
                   f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation {NS_A} {NS_R} {NS_P}>'
                   '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
                   f'<p:sldIdLst>{sld_ids}</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>')
        rels = '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
        rels += "".join(f'<Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>' for i in range(1, n + 1))
        rels += f'<Relationship Id="rId{n + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
        z.writestr("ppt/_rels/presentation.xml.rels", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>')
        z.writestr("ppt/slideMasters/slideMaster1.xml",
                   f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster {NS_A} {NS_R} {NS_P}><p:cSld><p:spTree>{EMPTY_TREE}</p:spTree></p:cSld>'
                   '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
                   '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>')
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
                   '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>')
        z.writestr("ppt/slideLayouts/slideLayout1.xml",
                   f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout {NS_A} {NS_R} {NS_P} type="blank"><p:cSld name="Blank"><p:spTree>{EMPTY_TREE}</p:spTree></p:cSld>'
                   '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>')
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>')
        z.writestr("ppt/theme/theme1.xml", THEME)
        for i, (head, bullets) in enumerate(slides, 1):
            shapes = textbox(2, 600000, 400000, 11000000, 1000000, [head], 3600, True) + \
                     textbox(3, 600000, 1600000, 11000000, 4800000, ["• " + b for b in bullets], 2000)
            z.writestr(f"ppt/slides/slide{i}.xml",
                       f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld {NS_A} {NS_R} {NS_P}><p:cSld><p:spTree>{EMPTY_TREE}{shapes}</p:spTree></p:cSld>'
                       '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>')
            z.writestr(f"ppt/slides/_rels/slide{i}.xml.rels",
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                       '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>')
        z.writestr("docProps/core.xml", core_xml(title))

# ---------------- Excel ----------------
def xlsx(p, sheets):
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    for name, header, rows in sheets:
        ws = wb.create_sheet(name[:31]); ws.append(header)
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="1F2A44")
        for r in rows: ws.append(r)
        for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = max(12, min(40, max(len(str(c.value or "")) for c in col) + 2))
    wb.save(p)

# ---------------- PDF ----------------
def pdf(p, title, lines):
    c = canvas.Canvas(p, pagesize=A4); w, h = A4
    c.setFont("Helvetica-Bold", 16); c.drawString(50, h - 60, title)
    c.setFont("Helvetica", 9); c.drawString(50, h - 78, ORG)
    c.setFont("Helvetica", 11); y = h - 110
    for l in lines:
        if y < 60: c.showPage(); c.setFont("Helvetica", 11); y = h - 60
        c.drawString(50, y, l); y -= 16
    c.save()

def money(a, b): return round(random.uniform(a, b), 2)

# ================= Content =================
months = ["Jul-25", "Aug-25", "Sep-25", "Oct-25", "Nov-25", "Dec-25", "Jan-26", "Feb-26", "Mar-26", "Apr-26", "May-26", "Jun-26"]

# Finance
xlsx(path("Finance", "FY26 Operating Budget.xlsx"), [
    ("Budget", ["Cost centre", "Account"] + months + ["FY26 Total"],
     [[cc, acc] + (v := [money(8000, 60000) for _ in months]) + [round(sum(v), 2)]
      for cc in ["Fund Admin", "Asset Mgmt", "Leasing", "Corporate"] for acc in ["Salaries", "Consultants", "Travel", "IT", "Marketing"]])])
xlsx(path("Finance", "Management Accounts Q1 FY26.xlsx", "Management Accounts"), [
    ("P&L", ["Line", "Jul-25", "Aug-25", "Sep-25", "Q1 Actual", "Q1 Budget", "Variance"],
     [[l, *(v := [money(100000, 900000) for _ in range(3)]), round(sum(v), 2), (b := money(500000, 2500000)), round(sum(v) - b, 2)]
      for l in ["Rental income", "Outgoings recovered", "Property expenses", "Management fees", "Interest expense", "Net profit"]])])
xlsx(path("Finance", "Management Accounts Q2 FY26.xlsx", "Management Accounts"), [
    ("P&L", ["Line", "Oct-25", "Nov-25", "Dec-25", "Q2 Actual"],
     [[l, *(v := [money(100000, 900000) for _ in range(3)]), round(sum(v), 2)] for l in ["Rental income", "Outgoings recovered", "Property expenses", "Net profit"]])])
xlsx(path("Finance", "Fund Valuation Summary Sep-26.xlsx"), [
    ("Valuations", ["Fund", "Asset", "Valuer", "Cap rate %", "Book value $", "Independent value $", "Movement $"],
     [[random.choice(FUNDS), a, random.choice(["CBRE", "JLL", "Colliers", "Knight Frank"]), round(random.uniform(4.5, 7.0), 2),
       (bv := money(20e6, 120e6)), (iv := money(20e6, 120e6)), round(iv - bv, 2)] for a in ASSETS])])
docx(path("Finance", "Board Paper - Capital Raise Proposal.docx"),
     "Board Paper: Proposed Capital Raise", [
         ("h", "Purpose"), f"To seek Board approval for a $45m equity raise for the {FUNDS[1]} to fund the acquisition of 88 Logistics Way, Eastern Creek.",
         ("h", "Background"), "The asset is 100% leased to a national 3PL operator with a WALE of 8.2 years. The raise will be offered to existing wholesale unitholders first.",
         ("h", "Recommendation"), "That the Board approves the raise at an issue price of $1.12 per unit, subject to final due diligence and debt approval.",
         ("h", "Risks"), "Interest rate movements, tenant concentration and settlement timing. Mitigations are set out in Appendix B."])
docx(path("Finance", "FY25 Audit Findings and Management Responses.docx"),
     "FY25 External Audit: Findings and Management Responses", [
         ("h", "Finding 1: Journal approval"), "Manual journals over $50,000 were not consistently approved before posting. Response: approval workflow added in the ERP from March 2026.",
         ("h", "Finding 2: Lease incentives"), "Two lease incentives were amortised over the wrong term. Response: lease schedule reconciled and corrected in Q2.",
         ("h", "Finding 3: Access reviews"), "Quarterly user access reviews were not evidenced. Response: reviews now logged in the compliance register."])
pdf(path("Finance", "Investor Distribution Statement - Sep 2026.pdf", "Distributions"), "Distribution Statement - Quarter ended 30 September 2026",
    [f"Fund: {FUNDS[0]}", "Distribution per unit: 2.15 cents", "Record date: 30 September 2026", "Payment date: 14 October 2026", "",
     "Unitholder        Units held      Gross distribution    Tax withheld    Net paid"] +
    [f"{random.choice(PEOPLE):<18}{(u := random.randint(50000, 900000)):>10,}      ${u * 0.0215:>14,.2f}      $0.00      ${u * 0.0215:>12,.2f}" for _ in range(12)])
pptx(path("Finance", "Finance Update Q3 FY26.pptx"), "Finance Update Q3 FY26", [
    ("Finance Update - Q3 FY26", [ORG, "Prepared for the Executive Committee"]),
    ("Headlines", ["Net property income 3.4% ahead of budget", "Occupancy 97.1% across the portfolio", "Gearing 31% (covenant 50%)"]),
    ("Capital management", ["$45m raise proposed for Industrial Trust", "Debt facility extended to 2029", "Hedging cover 62%"]),
    ("Next steps", ["Finalise FY27 budget by May", "Valuation cycle in June", "Audit planning meeting 12 May"])])

# HR
xlsx(path("HR", "Employee Register.xlsx"), [
    ("Staff", ["Employee ID", "Name", "Role", "Start date", "Location", "Employment type", "Manager"],
     [[f"E{1000 + i}", n, ROLES[i], dt.date(2018 + i % 8, (i % 12) + 1, 1 + i), random.choice(["Melbourne", "Sydney", "Brisbane"]),
       random.choice(["Full time", "Part time"]), PEOPLE[(i + 3) % 10]] for i, n in enumerate(PEOPLE)])])
xlsx(path("HR", "Remuneration Review 2026.xlsx", "Remuneration"), [
    ("Review", ["Employee ID", "Name", "Role", "Current base $", "Proposed increase %", "Proposed base $", "STI target %"],
     [[f"E{1000 + i}", n, ROLES[i], (b := random.randrange(85000, 260000, 500)), (inc := round(random.uniform(2, 6), 1)), round(b * (1 + inc / 100)), random.choice([10, 15, 20, 30])]
      for i, n in enumerate(PEOPLE)])])
docx(path("HR", "Leave Policy.docx", "Policies"), "Leave Policy", [
    ("h", "Scope"), f"This policy applies to all permanent employees of {ORG}.",
    ("h", "Annual leave"), "Full-time employees accrue 20 days of annual leave per year. Leave requests are submitted through the HR system at least two weeks in advance.",
    ("h", "Personal leave"), "Employees accrue 10 days of personal/carer's leave per year. A medical certificate is required for absences longer than two days.",
    ("h", "Review"), "This policy is reviewed every two years by the People & Culture team."])
docx(path("HR", "Position Description - Asset Manager.docx"), "Position Description: Asset Manager", [
    ("h", "Role purpose"), "Manage the performance of a portfolio of industrial and retail assets to deliver income and capital growth.",
    ("h", "Key responsibilities"), "Prepare asset plans and budgets. Lead leasing strategy with the Leasing team. Oversee capital works. Report to the Fund Manager monthly.",
    ("h", "Requirements"), "Degree in property, finance or commerce. 5+ years' commercial property experience."])
docx(path("HR", "Performance Review - Tom Fitzgerald.docx", "Performance Reviews"), "Performance Review FY26: Tom Fitzgerald", [
    ("h", "Overall rating"), "Meets expectations.",
    ("h", "Achievements"), "Delivered the Westgate Plaza refurbishment on time and 4% under budget.",
    ("h", "Development areas"), "Stakeholder updates to investors; financial modelling depth.",
    ("h", "Agreed actions"), "Attend the advanced modelling course in Q1 FY27."])
pdf(path("HR", "Onboarding Checklist.pdf", "Policies"), "New Starter Onboarding Checklist",
    ["[ ] Contract signed and returned", "[ ] IT account and laptop issued", "[ ] Building access card", "[ ] Code of Conduct acknowledged",
     "[ ] Privacy and cyber security training", "[ ] Meet with manager - week 1 goals", "[ ] Payroll and super details provided"])
pdf(path("HR", "Code of Conduct.pdf", "Policies"), "Code of Conduct",
    ["1. Act honestly and with integrity.", "2. Treat colleagues, tenants and investors with respect.", "3. Declare conflicts of interest.",
     "4. Protect confidential information.", "5. Report breaches to the Compliance Officer.", "", "Approved by the Board, March 2026."])

# Legal
xlsx(path("Legal", "Contract Register.xlsx"), [
    ("Contracts", ["Contract ID", "Counterparty", "Type", "Asset", "Start", "Expiry", "Value $", "Owner"],
     [[f"C-{2400 + i}", random.choice(["Apex Cleaning Pty Ltd", "Southern Security Services", "Brightline Electrical", "GreenScape Grounds", "Metro Lifts"]),
       random.choice(["Services", "Maintenance", "Consultancy"]), random.choice(ASSETS), dt.date(2025, random.randint(1, 12), 1),
       dt.date(2027, random.randint(1, 12), 1), money(20000, 400000), random.choice(PEOPLE)] for i in range(18)])])
docx(path("Legal", "Lease - Unit 4, 12 Bayside Drive.docx", "Leases"), "Lease Agreement: Unit 4, 12 Bayside Drive, Port Melbourne", [
    ("h", "Parties"), f"Landlord: {FUNDS[0]} (ACN 600 000 000). Tenant: Coastal Freight Solutions Pty Ltd.",
    ("h", "Term"), "Five (5) years commencing 1 February 2026, with one further option of five (5) years.",
    ("h", "Rent"), "$412,500 per annum plus GST, reviewed annually by 3.5% fixed increases.",
    ("h", "Outgoings"), "Tenant pays 100% of recoverable outgoings in proportion to lettable area."])
docx(path("Legal", "Lease - Shop 12, Westgate Plaza.docx", "Leases"), "Retail Lease: Shop 12, Westgate Plaza", [
    ("h", "Tenant"), "Bean There Cafe Pty Ltd", ("h", "Term"), "Three years from 1 July 2026.",
    ("h", "Rent"), "$96,000 per annum plus GST, CPI reviews.", ("h", "Permitted use"), "Cafe and takeaway food."])
docx(path("Legal", "Mutual NDA - Template.docx", "Templates"), "Mutual Non-Disclosure Agreement", [
    ("h", "1. Confidential Information"), "Each party may disclose information to the other for the Purpose. The recipient must keep it confidential.",
    ("h", "2. Exceptions"), "Information that is public, already known, or independently developed is not Confidential Information.",
    ("h", "3. Term"), "Obligations survive for three years after the last disclosure."])
pdf(path("Legal", "Legal Advice - Rezoning Eastern Creek.pdf"), "Privileged & Confidential - Legal Advice",
    ["Re: Proposed rezoning of 88 Logistics Way, Eastern Creek", "", "Summary of advice:",
     "- The planning proposal is unlikely to be determined before Q3 2027.", "- Existing use rights protect current warehouse operations.",
     "- We recommend a condition precedent in the contract tied to the rezoning outcome.", "", "Prepared by external counsel for the General Counsel."])

# Projects
xlsx(path("Projects", "Project Plan - Westgate Refurbishment.xlsx", "Westgate Plaza"), [
    ("Plan", ["ID", "Task", "Owner", "Start", "Finish", "Status", "% complete"],
     [[i + 1, t, random.choice(PEOPLE), dt.date(2026, 2 + i // 2, 1 + i), dt.date(2026, 3 + i // 2, 10 + i), random.choice(["Done", "In progress", "Not started"]), random.choice([0, 25, 50, 75, 100])]
      for i, t in enumerate(["Design brief", "Council approval", "Tender", "Contractor award", "Demolition", "Mall amenities upgrade", "Facade works", "Tenant fit-out coordination", "Practical completion", "Defects period"])])])
xlsx(path("Projects", "Risk Register.xlsx", "Westgate Plaza"), [
    ("Risks", ["Risk ID", "Description", "Likelihood", "Consequence", "Rating", "Mitigation", "Owner"],
     [[f"R{i + 1:02}", d, random.choice(["Low", "Medium", "High"]), random.choice(["Minor", "Moderate", "Major"]), random.choice(["Low", "Medium", "High"]), m, random.choice(PEOPLE)]
      for i, (d, m) in enumerate([("Cost overrun on facade", "Fixed-price contract"), ("Tenant disruption", "After-hours works"),
                                  ("Council delay", "Early engagement"), ("Asbestos found", "Pre-works survey"), ("Contractor insolvency", "Bank guarantee")])])])
xlsx(path("Projects", "Tender Evaluation - Facade Works.xlsx", "Westgate Plaza"), [
    ("Evaluation", ["Tenderer", "Price $", "Price score", "Experience", "Program", "Safety", "Weighted total"],
     [[t, money(1.8e6, 2.6e6), random.randint(6, 10), random.randint(5, 10), random.randint(5, 10), random.randint(6, 10), round(random.uniform(6, 9), 2)]
      for t in ["Buildcorp Victoria", "Ridgeline Constructions", "Atlas Facades"]])])
docx(path("Projects", "Status Report - Week 45.docx", "Westgate Plaza"), "Project Status Report: Week 45", [
    ("h", "Overall status"), "Amber: facade works two weeks behind due to wet weather.",
    ("h", "Progress this week"), "Mall amenities upgrade complete. Facade panels delivered for levels 1-2.",
    ("h", "Issues"), "Crane access on Tuesday conflicts with tenant deliveries; working with centre management.",
    ("h", "Next week"), "Install facade panels level 2. Defects walk for amenities."])
pptx(path("Projects", "Steering Committee - Westgate Plaza.pptx", "Westgate Plaza"), "Steering Committee - Westgate Plaza", [
    ("Westgate Plaza Refurbishment", ["Steering Committee - November 2026"]),
    ("Status", ["Overall: Amber", "Budget: $8.4m of $9.1m committed", "Practical completion forecast: March 2027"]),
    ("Decisions needed", ["Approve $180k variation for additional signage", "Endorse revised tenant communication plan"])])
pdf(path("Projects", "Site Inspection Report - Wacol.pdf", "Wacol"), "Site Inspection Report - 4 Quarry Road, Wacol",
    ["Inspection date: 18 November 2026", f"Inspector: {random.choice(PEOPLE)}", "", "Findings:",
     "- Roof sheeting corrosion on the north-east corner (medium priority).", "- Fire sprinkler service tags current.",
     "- Hardstand cracking near dock 3; recommend repair within 6 months.", "- Tenant housekeeping satisfactory."])
docx(path("Projects", "Feasibility Study - Wacol Expansion.docx", "Wacol"), "Feasibility Study: Wacol Warehouse Expansion", [
    ("h", "Proposal"), "Add 6,500 sqm of warehouse space on the vacant southern land.",
    ("h", "Cost estimate"), "$11.2m including professional fees and contingency.",
    ("h", "Returns"), "Yield on cost 6.8% assuming pre-commitment from the existing tenant.",
    ("h", "Recommendation"), "Proceed to design development, subject to tenant heads of agreement."])

total, size = 0, 0
for root, _, files in os.walk(OUT):
    for f in files:
        total += 1; size += os.path.getsize(os.path.join(root, f))
print(f"{total} files, {size / 1024:.0f} KB total -> {OUT}")
