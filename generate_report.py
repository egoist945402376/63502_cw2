from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BLACK = RGBColor(0, 0, 0)

def add_heading_black(doc, text, level):
    """Add a heading with forced black colour."""
    p = doc.add_heading(text, level=level)
    for run in p.runs:
        run.font.color.rgb = BLACK
    return p

def add_code_block(doc, code):
    """Add a grey-background monospace paragraph."""
    para = doc.add_paragraph()
    run = para.add_run(code)
    run.font.name = 'Courier New'
    run.font.size = Pt(9)
    run.font.color.rgb = BLACK
    pPr = para._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), 'F2F2F2')
    pPr.append(shd)
    return para

doc = Document()

# ---------- Title ----------
title = doc.add_heading('CW2 Report: Knowledge Graph and Semantic Table Annotation', level=1)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
for run in title.runs:
    run.font.color.rgb = BLACK

doc.add_paragraph()

# ---------- Part I ----------
add_heading_black(doc, 'Part I: SPARQL Query', level=2)
add_heading_black(doc, 'Task 1 (Country)', level=3)
add_heading_black(doc, 'Task 1.1', level=4)

intro = doc.add_paragraph()
r1 = intro.add_run('Introduction: ')
r1.bold = True
r1.font.color.rgb = BLACK
r2 = intro.add_run(
    'This query selects instances of dbo:Country that have a capital city, '
    'excludes historical countries by filtering out those with a dissolutionYear, '
    'and restricts labels to English.'
)
r2.font.color.rgb = BLACK

add_code_block(doc, """\
PREFIX dbo:  <http://dbpedia.org/ontology/>
PREFIX rdfs: <http://www.w3.org/2000/01/rdf-schema#>

SELECT DISTINCT ?country ?countryName
WHERE {
  ?country a dbo:Country ;
           rdfs:label ?countryName ;
           dbo:capital ?capital .
  FILTER NOT EXISTS { ?country dbo:dissolutionYear ?year }
  FILTER (LANG(?countryName) = "en")
}
ORDER BY ?countryName""")

note = doc.add_paragraph()
r = note.add_run('Countries #: ')
r.bold = True
r.font.color.rgb = BLACK
rn = note.add_run('~195 (varies by DBpedia snapshot)')
rn.font.color.rgb = BLACK

# ---------- save ----------
out_path = r'C:\Users\94540\Desktop\UoM\63502\cw2\CW2_Report.docx'
doc.save(out_path)
print(f'Saved to {out_path}')
