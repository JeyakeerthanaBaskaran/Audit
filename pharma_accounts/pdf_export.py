"""Year-end PDF export (ReportLab)."""
from datetime import datetime
from io import BytesIO

from reportlab.lib import colors
from reportlab.lib.enums import TA_RIGHT
from reportlab.lib.pagesizes import A3, A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import LongTable, Paragraph, SimpleDocTemplate, Spacer, TableStyle
from xml.sax.saxutils import escape


def _txt(v):
    # Standard PDF fonts have no rupee glyph
    return escape(str(v).replace("₹", "Rs."))


class _NumberedCanvas(canvas.Canvas):
    """Adds 'Page X of Y' and a generated-on line to every page."""

    def __init__(self, *args, footer="", **kwargs):
        super().__init__(*args, **kwargs)
        self._saved = []
        self._footer = footer

    def showPage(self):
        self._saved.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total = len(self._saved)
        for state in self._saved:
            self.__dict__.update(state)
            w, _ = self._pagesize
            self.setFont("Helvetica", 7)
            self.setFillColor(colors.HexColor("#555555"))
            self.drawString(12 * mm, 7 * mm, self._footer)
            self.drawRightString(w - 12 * mm, 7 * mm, f"Page {self._pageNumber} of {total}")
            super().showPage()
        super().save()


def build_pdf(df, types, title, company="", notes=None, totals=None):
    """df: DataFrame of display text. types: {col: Text/Money/Number/Date}.
    totals: {col: formatted total text}. Returns PDF bytes."""
    notes = notes or []
    totals = totals or {}
    ncols = max(len(df.columns), 1)
    pagesize = landscape(A3) if ncols > 10 else landscape(A4)
    font_size = 8 if ncols <= 7 else 7 if ncols <= 10 else 6.5 if ncols <= 14 else 5.5

    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=pagesize,
        leftMargin=12 * mm,
        rightMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=14 * mm,
        title=title,
        author=company,
    )
    ss = getSampleStyleSheet()
    h_company = ParagraphStyle("c", parent=ss["Title"], fontSize=15, leading=18, spaceAfter=2)
    h_title = ParagraphStyle("t", parent=ss["Heading2"], fontSize=12, leading=15, spaceAfter=4)
    small = ParagraphStyle("s", parent=ss["Normal"], fontSize=8, leading=10, textColor=colors.HexColor("#444444"))
    cell = ParagraphStyle("cell", parent=ss["Normal"], fontSize=font_size, leading=font_size + 2)
    cell_r = ParagraphStyle("cellr", parent=cell, alignment=TA_RIGHT)
    head = ParagraphStyle("head", parent=cell, fontName="Helvetica-Bold", textColor=colors.white)
    head_r = ParagraphStyle("headr", parent=head, alignment=TA_RIGHT)
    tot = ParagraphStyle("tot", parent=cell, fontName="Helvetica-Bold")
    tot_r = ParagraphStyle("totr", parent=tot, alignment=TA_RIGHT)

    story = []
    if company:
        story.append(Paragraph(_txt(company), h_company))
    story.append(Paragraph(_txt(title), h_title))
    for n in notes:
        story.append(Paragraph(_txt(n), small))
    story.append(Spacer(1, 5 * mm))

    cols = list(df.columns)
    right = [types.get(c) in ("Money", "Number") for c in cols]
    data = [[Paragraph(_txt(c), head_r if r else head) for c, r in zip(cols, right)]]
    for row in df.itertuples(index=False):
        data.append([Paragraph(_txt(v), cell_r if r else cell) for v, r in zip(row, right)])
    has_total = bool(totals)
    if has_total:
        trow = []
        for i, c in enumerate(cols):
            if c in totals:
                trow.append(Paragraph(_txt(totals[c]), tot_r))
            elif i == 0:
                trow.append(Paragraph("TOTAL", tot))
            else:
                trow.append("")
        data.append(trow)

    # column widths proportional to content length
    avail = pagesize[0] - doc.leftMargin - doc.rightMargin
    lens = []
    for i, c in enumerate(cols):
        sample = [len(str(v)) for v in df.iloc[:300, i]] if len(df) else [0]
        lens.append(min(max([len(str(c))] + sample + [len(totals.get(c, ""))]), 40) + 2)
    total_len = sum(lens) or 1
    widths = [avail * l / total_len for l in lens]

    table = LongTable(data, colWidths=widths, repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E5F")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#9AA5AB")),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
    ]
    for r in range(1, len(data)):
        if r % 2 == 0:
            style.append(("BACKGROUND", (0, r), (-1, r), colors.HexColor("#EEF3F5")))
    if has_total:
        style.append(("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#D5E3E8")))
        style.append(("LINEABOVE", (0, -1), (-1, -1), 1, colors.black))
    table.setStyle(TableStyle(style))
    story.append(table)
    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph(f"Number of records: {len(df)}", small))

    footer = f"{company}  |  {title}  |  Generated {datetime.now().strftime('%d-%m-%Y %H:%M')}"
    doc.build(story, canvasmaker=lambda *a, **k: _NumberedCanvas(*a, footer=footer, **k))
    return buf.getvalue()
