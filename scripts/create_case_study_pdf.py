from pathlib import Path
from reportlab.platypus import SimpleDocTemplate,Paragraph,Spacer,Table,TableStyle,KeepTogether
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
import argparse
parser=argparse.ArgumentParser(description='Render the portfolio case study; requires ReportLab.')
parser.add_argument('--output',type=Path,default=Path(__file__).resolve().parents[1]/'.artifacts/case-study/ProductMatch_case_study.pdf')
output=parser.parse_args().output.resolve()
output.parent.mkdir(parents=True,exist_ok=True)
pdfmetrics.registerFontFamily('PM',normal='Helvetica',bold='Helvetica-Bold')
navy=colors.HexColor('#142B45');teal=colors.HexColor('#087F8C');gray=colors.HexColor('#EEF2F6');ink=colors.HexColor('#243A4D')
body=ParagraphStyle('body',fontName='Helvetica',fontSize=9.5,leading=13,textColor=ink,spaceAfter=7)
heading=ParagraphStyle('head',parent=body,fontName='Helvetica-Bold',fontSize=11,textColor=navy,spaceBefore=8,spaceAfter=5)
small=ParagraphStyle('small',parent=body,fontSize=8.5,leading=11)
story=[]
def p(t,style=body):return Paragraph(t,style)
story += [p('PORTFOLIO CASE STUDY',ParagraphStyle('eyebrow',parent=small,textColor=teal,fontName='Helvetica-Bold')),p('ProductMatch',ParagraphStyle('title',parent=body,fontName='Helvetica-Bold',fontSize=25,leading=30,textColor=navy)),p('Product Reconciliation &amp; Data Quality Workspace',ParagraphStyle('sub',parent=body,fontSize=12,leading=16)),p('<b>Ridla A. Gani</b> | Excel 365 | November 2026',small),Spacer(1,7)]
story += [p('The problem',heading),p('A fictional electronics and office-accessories retailer needs to reconcile 250 supplier records against 200 internal catalog rows. Inconsistent names, missing fields, retransmissions and similar product variants make a simple SKU lookup unreliable.')]
story += [p('What I built',heading),p('A macro-free Excel workspace with preserved raw feeds, transparent cleaning, brand aliases, duplicate review, a matching hierarchy, validated human decisions, field-level research, accepted-mapping exports and a filterable dashboard with product search.'),p('Matching prioritizes nonblank manufacturer identifiers, brand/model and complete variants before a refreshable fuzzy-name suggestion. Color, capacity, size/hardware version and pack conflicts block export. Approved IDs are stored as values so refreshed suggestions cannot silently change reviewer decisions.'),p('<b>Tools:</b> Excel tables, native formulas, validation, conditional formatting, hyperlinks and charts; Python for candidate refresh, CSV exports and independent QA; JavaScript for workbook authoring and rendering.',small)]
story += [p('Measured demonstration results',heading)]
rows=[['Input and outcome','Delivered baseline'],['Master / supplier rows','200 / 250'],['Accepted mappings exported','124'],['Awaiting review / unmatched / rejected','88 / 18 / 5'],['Supplier duplicate rows excluded','15'],['Remaining exception records','111'],['Required completeness, before / after','98.52% / 98.72%'],['Verified research fields / total fields','15 / 22']]
th=ParagraphStyle('th',parent=small,textColor=colors.white,fontName='Helvetica-Bold')
t=Table([[p(a,th if i==0 else small),p(b,th if i==0 else small)]for i,(a,b) in enumerate(rows)],colWidths=[324,180],hAlign='LEFT')
t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),navy),('TEXTCOLOR',(0,0),(-1,0),colors.white),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),9),('RIGHTPADDING',(0,0),(-1,-1),9),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),2),('ROWBACKGROUNDS',(0,1),(-1,-1),[gray,colors.white]),('LINEBELOW',(0,-1),(-1,-1),0.5,colors.HexColor('#D7E0E8'))]))
story += [t,Spacer(1,7),p('Supplier reconciliation: <b>124 + 88 + 18 + 5 + 15 = 250</b>. Completeness uses the same 2,500 required-cell denominator before and after; it measures presence, not validity or business impact.',small)]
story += [p('Evaluation and limits',heading),p('On deliberately constructed nonduplicate examples, 204 of 217 suggestions identify the expected master (94.0% precision), covering 204 of 215 known positive cases (94.9% recall). Thirteen suggestions are incorrect; eleven positive cases lack a correct suggestion. Human review remains necessary.',small),p('Trading data, prices and reviewer decisions are synthetic; 15 real models have official-source research. Full headless LibreOffice recalculation and programmatic tests are documented separately. Native Excel 365 and Google Sheets UI testing was unavailable. Fuzzy scores require Python refresh; normal review and dashboard formulas do not. Capacity: 300 master, 350 supplier and 100 research records. Specification is the only research field applied automatically. Manual inputs retain current state, not a complete automatic history.',small),p('<b>No client engagement, time saving, production accuracy improvement or financial impact is claimed.</b>',small)]
def footer(c,doc):
 c.setStrokeColor(teal);c.setLineWidth(1);c.line(45,32,550,32);c.setFont('Helvetica',8);c.setFillColor(ink);c.drawString(45,20,'Ridla A. Gani | ProductMatch | Demonstration dataset');c.drawRightString(550,20,str(doc.page))
doc=SimpleDocTemplate(str(output),pagesize=A4,rightMargin=45,leftMargin=45,topMargin=35,bottomMargin=45,title='ProductMatch - Portfolio Case Study',author='Ridla A. Gani')
doc.build(story,onFirstPage=footer,onLaterPages=footer)
print(output)
