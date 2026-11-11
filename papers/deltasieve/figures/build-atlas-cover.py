"""Compose the cover from a saved Atlas view and native image excerpts."""
from pathlib import Path
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from PIL import Image
p = Path(__file__).resolve().parent
c = canvas.Canvas(str(p / 'atlas-cover.pdf'), pagesize=(900, 550))
c.setFillColorRGB(1, 1, 1)
c.rect(0, 0, 900, 550, fill=1, stroke=0)
c.setFillColorRGB(0.12, 0.2, 0.24)
c.setFont('Helvetica-Bold', 18)
c.drawString(16, 526, 'A  Worksheet Atlas')
c.drawString(556, 526, 'B  Inspect a visible change')
c.setFont('Helvetica', 11)
c.drawString(16, 506, 'Data.gov | nine workbooks | five transformation pipelines')
c.drawString(556, 506, 'Adage Fixed Factor | readme!A1:M22')
img = Image.open(p / 'atlas-stage09-light.png')
gallery = img.crop((240, 275, 770, 655))
c.drawImage(ImageReader(gallery), 16, 135, width=510, height=366, mask='auto')
colors = [('#fff0ce', 'Delta  Observed difference'), ('#e3f1eb', '=  Compared captures unchanged'), ('#e5edf2', '?  No retained pixel comparison')]
for i, (color, text) in enumerate(colors):
    y = 108 - i * 25
    c.setFillColor(color)
    c.rect(16, y, 13, 13, fill=1, stroke=0)
    c.setFillColorRGB(0.12, 0.2, 0.24)
    c.setFont('Helvetica', 12)
    c.drawString(37, y + 1, text)
c.setFont('Helvetica', 10)
c.drawString(16, 18, 'P  Capture plan     I  Dependency candidate     F  Conservative rule capture')
c.setStrokeColorRGB(0.78, 0.83, 0.86)
c.line(540, 15, 540, 510)
for y, label, file in [(365, 'Original', 'adage-reference-original.png'), (209, 'Inquire', 'adage-reference-changed.png')]:
    c.setFillColorRGB(0.12, 0.2, 0.24)
    c.setFont('Helvetica-Bold', 12)
    c.drawString(556, y + 117, label)
    c.drawImage(str(p / file), 556, y, width=328, height=112, mask='auto')
c.setFillColorRGB(0.65, 0.14, 0.14)
c.setFont('Helvetica-Bold', 13)
c.drawString(556, 181, '!  Title-band fill removed')
c.setFillColorRGB(0.12, 0.2, 0.24)
c.setFont('Helvetica', 11)
c.drawString(556, 161, '16,182 changed pixels in the matched reference')
for y, label, file in [(100, 'Preserved', 'adage-header-preserved.png'), (52, 'Shortened', 'adage-header-shortened.png')]:
    c.setFont('Helvetica', 10)
    c.drawString(556, y + 22, label)
    c.drawImage(str(p / file), 556, y, width=328, height=16, mask='auto')
c.setFont('Helvetica', 9)
c.drawString(556, 18, 'Native image excerpts. Overview differences remain unclassified.')
c.showPage()
c.save()
