from pathlib import Path

from docx import Document
from lxml import html


ROOT = Path(__file__).resolve().parents[1]
DOCX = ROOT / "deliverables/BUS209_FactSet_Financial_Concepts_Syllabus_Fall_2026.docx"
WEB = ROOT / "deliverables/BUS209_FactSet_Financial_Concepts_Syllabus_Fall_2026.html"
CANVAS = ROOT / "deliverables/BUS209_FactSet_Financial_Concepts_Syllabus_Fall_2026_CANVAS.html"


def clean(text):
    return " ".join(text.split())


document = Document(DOCX)
word_table = next(
    table
    for table in document.tables
    if [clean(cell.text) for cell in table.rows[0].cells]
    == ["Week", "Week of", "Assignments Due / Weekly Class Topics"]
)
word_dates = [(clean(row.cells[0].text), clean(row.cells[1].text)) for row in word_table.rows[1:]]


def verify_web():
    tree = html.fromstring(WEB.read_text(encoding="utf-8"))
    schedule = tree.xpath(".//section[@id='schedule']")[0]
    rows = schedule.xpath("./table/tbody/tr")
    values = [[clean(cell.text_content()) for cell in row.xpath("./td")] for row in rows]
    return tree, values


def verify_canvas():
    tree = html.fromstring(CANVAS.read_text(encoding="utf-8"))
    heading = tree.xpath(".//h2[normalize-space()='Topic Outline & Timeline']")[0]
    rows = heading.getparent().xpath("./table/tbody/tr")
    values = [[clean(cell.text_content()) for cell in row.xpath("./td")] for row in rows]
    return tree, values


for filename, (tree, rows) in [(WEB.name, verify_web()), (CANVAS.name, verify_canvas())]:
    html_dates = [(week, date) for week, date, _ in rows]
    assert html_dates == word_dates, (filename, html_dates, word_dates)
    topics = {week: topic for week, _, topic in rows}
    assert topics["9"] == "To Be determined", (filename, topics["9"])
    assert topics["10"] == "To Be determined", (filename, topics["10"])
    page_text = clean(tree.text_content())
    assert "No Class (instructor-planned" not in page_text, filename
    print(f"PASS {filename}: {len(rows)} Word-matched dates; Nov. 9 and Nov. 16 are To Be determined")
