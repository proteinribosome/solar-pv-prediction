from copy import deepcopy
from pathlib import Path

from docx import Document


PAPER = Path(
    "/Users/chengyizhou/Documents/GitHub/solar-pv-prediction/"
    "Solar_PV_Research_Paper_Model_C_Revised.docx"
)


def find(document, prefix):
    matches = [
        paragraph
        for paragraph in document.paragraphs
        if paragraph.text.startswith(prefix)
    ]
    if len(matches) != 1:
        raise ValueError(f"Expected one paragraph for {prefix!r}")
    return matches[0]


document = Document(PAPER)
template = find(document, "6.2 Physics adds")
template_run_properties = deepcopy(template.runs[0]._r.rPr)

for prefix in (
    "6.3 No-history systems",
    "6.4 Bias remains unresolved",
):
    paragraph = find(document, prefix)
    run_element = paragraph.runs[0]._r
    if run_element.rPr is not None:
        run_element.remove(run_element.rPr)
    run_element.insert(0, deepcopy(template_run_properties))

document.save(PAPER)
print(f"Matched inserted heading formatting in {PAPER}")
