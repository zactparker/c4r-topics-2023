"""Build data/ and the page's data blob from the raw Topic Survey export (October 12, 2023).

Usage: python3 -I tools/build.py "path/to/Topic Survey.xlsx"

Writes:
  data/topic-survey-2023-deidentified.csv  every response; one identifying comment withheld
  data/topic-counts.csv                    times each topic was picked
  data/comment-codes.csv                   one row per (respondent, theme) for the coded comments
  index.html                               rewrites the DATA blob between /*DATA*/ markers
"""
import csv, json, re, sys
from collections import Counter
from pathlib import Path
import openpyxl

ROOT = Path(__file__).resolve().parent.parent
SRC = sys.argv[1]

ws = openpyxl.load_workbook(SRC, read_only=True).active
rows = [r[:4] for r in ws.iter_rows(min_row=2, values_only=True) if any(r[:4])]

def clean(v):
    return "" if v is None else str(v).strip()

recs = []
for i, r in enumerate(rows):
    topics = [t.strip() for t in re.split(r",(?=[A-Z])", str(r[0]))] if r[0] else []
    recs.append(dict(id=f"R{i+1:03d}", topics=topics, missed=clean(r[1]), comments=clean(r[2]), stage=clean(r[3])))

# Answers withheld from every output: R157's comment named the respondent's institution and two of
# its instructors. Rows are matched by position; the quote checks below fail if the export's row
# order ever differs from the one reviewed.
WITHHELD = {156: "comments"}  # row index (0-based) -> field
for i, field in WITHHELD.items():
    assert recs[i][field], f"R{i+1:03d} {field} is already empty; check the row order"
    recs[i][field] = ""



# ---- comment codes ---------------------------------------------------------
# Codes apply to a respondent's two open-ended answers together. R-ids are 1-based row order.
THEMES = {
 "INC": "Incentives & publication pressure",
 "PUB": "Publication bias & null results",
 "PRE": "Preregistration & registered reports",
 "REV": "Journals, peer review & accountability",
 "STA": "Statistical literacy",
 "CAU": "Causal inference & controls",
 "QUE": "Research questions, theory & interpretation",
 "MET": "Methods & materials reporting",
 "COD": "Code & software practices",
 "TRN": "Rigor training & education",
 "DES": "Experimental design (blinding, degrees of freedom)",
 "WRK": "Funding, resources & working conditions",
 "OTH": "Other",
}
CODES = {  # row index (0-based) -> themes; withheld answers and rows with only "no"/"na"/praise/jokes are uncoded
 1:"INC", 2:"INC PRE STA CAU DES", 5:"STA", 6:"CAU", 7:"REV", 11:"PRE", 14:"MET TRN", 16:"QUE",
 20:"REV TRN", 24:"PUB TRN", 25:"QUE", 40:"OTH", 41:"MET", 43:"INC PRE REV CAU QUE STA",
 45:"INC CAU", 46:"WRK", 55:"INC REV", 60:"INC OTH", 69:"DES COD QUE INC", 78:"INC DES TRN",
 79:"TRN", 80:"PUB", 85:"INC", 88:"MET REV", 91:"COD", 92:"INC", 93:"QUE", 100:"OTH", 102:"STA",
 106:"INC PUB", 107:"INC PUB", 109:"WRK", 124:"REV", 125:"INC", 128:"QUE", 129:"MET", 132:"MET",
 136:"INC", 140:"INC OTH", 143:"COD", 147:"INC PUB", 148:"INC STA PRE", 162:"MET WRK",
 164:"COD", 168:"WRK", 176:"INC WRK",
}
n_comment = sum(1 for x in recs if x["missed"] or x["comments"])
tc = Counter(c for v in CODES.values() for c in v.split())
themes = [dict(label=THEMES[k], n=tc[k]) for k in sorted(THEMES, key=lambda k: (k == "OTH", -tc[k], THEMES[k]))]

# ---- quotes (verbatim; "…" marks a cut) -------------------------------------
QUOTES = [
 (1, "missed", ["How the incentive structure of science is geared towards generating splashy findings, rather than dispassionate pursuit of truth."]),
 (11, "comments", ["Every biomed wet lab does p hacking. The only way to prevent it is to require that every experiment be registered before it starts in an electronic lab notebook."]),
 (2, "comments", ["I think one serious QRP not discussed enough is outlier removal.", "In fact, if you go to the reference for Tukey's Fences (Tukey 1976), he DOES NOT SAY that it can be used for data REMOVAL, only flagging data to examine for technical problems in measurements."]),
 (5, "missed", ["Statistical literacy! Not only the ability to plan a well controlled and powered experiment and perform a test that answers the question at hand, but *also* to be equipped to confidently evaluate the statistical claims of others in the literature and not let sloppy reasoning slide."]),
 (20, "comments", ["There should be a course on how to be a reviewer for academic papers.", "it's all for naught if the \"referees\" are not good at calling foul."]),
 (25, "missed", ["Formulating good research questions that can be answered rigorously with the tools at hand", "In other words, conceptual rigor."]),
 (43, "missed", ["Sources of spurious correlation are not widely recognized, such as conditioning variable selection on a collider relationship"]),
 (148, "missed", ["Many people believe that if they get a significant with N=10, they would also get it with a bigger sample because their significant effect size is so big but they don't realize that their effect size is largely overestimated"]),
 (45, "comments", ["Rigor must not be an “add-on” but integral part of all levels and facets of scientific endeavors."]),
 (46, "comments", ["I know some researchers who can't afford modern equipment and rely on equipment that has been out of production for decades."]),
 (132, "missed", ["reagents!!!! you have totally missed the authentication of key biological resources as well as software tools!"]),
 (143, "comments", ["Even openly available resources are often awkward to actually access and use", "If we want open resources to succeed and find widespread use, attention should be paid to the design of their usage."]),
 (164, "missed", ["Over-reliance on other peoples code (without understanding what it’s doing)"]),
 (128, "comments", ["Many studies seemingly retrofit narratives rather than arising from an important larger question"]),
 (168, "missed", ["I get rushed to move on to the next experiment before running the experiment several times"]),
]
quotes = []
for i, field, parts in QUOTES:
    src = recs[i][field]
    for p in parts: assert p in src, (i, p)
    quotes.append(dict(text=" … ".join(parts), stage=recs[i]["stage"] or "Career stage not given"))

# ---- write -------------------------------------------------------------------
d = ROOT / "data"; d.mkdir(exist_ok=True)
with open(d / "topic-survey-2023-deidentified.csv", "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["respondent", "career_stage", "topics_selected (" + str(len(cnt)) + " options, pick 1-5; separated by ' | ')",
                "Are there any you think we missed?", "Any additional comments you would like to offer on the subject?"])
    for x in recs: w.writerow([x["id"], x["stage"], " | ".join(x["topics"]), x["missed"], x["comments"]])
with open(d / "topic-counts.csv", "w", newline="") as f:
    w = csv.writer(f); w.writerow(["topic", "times_selected", "pct_of_respondents"])
    for t in topics: w.writerow([t["full"], t["n"], round(100 * t["n"] / n_topic, 1)])
with open(d / "comment-codes.csv", "w", newline="") as f:
    w = csv.writer(f); w.writerow(["respondent", "theme"])
    for i, v in sorted(CODES.items()):
        for c in v.split(): w.writerow([recs[i]["id"], THEMES[c]])

blob = dict(n=len(recs), nTopic=n_topic, nComment=n_comment, nCoded=len(CODES), topics=topics, options=options, themes=themes, quotes=quotes)
page = ROOT / "index.html"
s = page.read_text()
s = re.sub(r"/\*DATA\*/.*?/\*END\*/", lambda m: "/*DATA*/" + json.dumps(blob, ensure_ascii=False) + "/*END*/", s, flags=re.S)
page.write_text(s)
print(f"{len(recs)} responses, {n_topic} picked topics, {n_comment} wrote comments, {len(CODES)} coded")
for t in themes: print(t["n"], t["label"])
