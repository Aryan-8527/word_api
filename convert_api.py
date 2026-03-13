from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.units import inch
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from reportlab.pdfgen import canvas


details_pdf = "/tmp/details_page.pdf"

c = canvas.Canvas(details_pdf, pagesize=landscape(A4))

width, height = landscape(A4)

c.setFont("Helvetica-Bold", 26)
c.drawCentredString(width/2, height-80, "Document Details")

data = [
    ["Field", "Value"],
    ["Document Code", document_code],
    ["Client Name", client_name],
    ["Department", department],
    ["Document Type", document_type],
    ["Purpose", purpose],
    ["Created On", created_on],
    ["Created By", created_by],
]

table = Table(data, colWidths=[3*inch, 6*inch])

table.setStyle(TableStyle([
    ("BACKGROUND",(0,0),(1,0),colors.lightgrey),
    ("GRID",(0,0),(-1,-1),1,colors.black),
    ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
    ("ALIGN",(0,0),(-1,-1),"LEFT"),
]))

table.wrapOn(c, width, height)
table.drawOn(c, 150, height-450)

c.save()
