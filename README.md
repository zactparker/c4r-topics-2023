# c4r-topics-2023

Results of the C4R topic survey (October 12, 2023 export): which topics respondents think hold back science the most, the coded written comments, and a selection of those comments.

- `index.html` is the whole site.
- `data/topic-survey-2023-deidentified.csv` has all 182 responses. One comment included respondent institution and two of its instructors and is withheld.
- `data/topic-counts.csv` gives the number of times each of the 18 topics was picked (181 respondents picked at least one).
- `data/comment-codes.csv` lists the theme codes for the 46 respondents whose comments had substance, one row per respondent and theme. The answers to the two open-ended questions were coded together. Answers like "no" or "na" and one joke were left uncoded.
- `tools/build.py` rebuilds every data file and the page's data from the raw export: `python3 -I tools/build.py "Topic Survey.xlsx"`. The codebook and the quote selection are in the script.
