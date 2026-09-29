# Theta Xi GPA and Recruitment

Theta Xi (UNL) grades and recruitment compared with the other IFC houses.

Open `analysis/Theta Xi GPA and Recruitment.xlsx`. That's the one to share. It has:

- By Semester: our GPA, size and new members next to the IFC every semester since Fall 2018
- Recruitment by Year and GPA by Year: each exec board (spring + fall) since 2011, best and worst years marked
- IFC Houses: every house ranked on GPA, size and new members
- Data and UNL Averages: the raw numbers

`analysis/Theta Xi vs UNL Fraternities.xlsx` is a longer version with more detail.

## Where the numbers come from

- `source-data/`: UNL's report cards (2018 to 2022), community scorecards (2023 to 2026) and the Fall 2023 grade report
- `data/`: the same numbers in two CSV files
- `legacy/`: the old GPA Data and AE List sheets (our numbers before 2018, rush chair names)
- `chapter-scholarship/`: scholarship program sheets from 2022 and 2023

From Fall 2018 on, UNL's numbers are used for every house. Our own records only fill spots UNL left blank: Spring 2020 GPA and new members, Spring 2022 new members, and Fall 2023 new members.

## Adding a semester

Paste the new IFC rows at the bottom of the Data tab. Or add them to `data/chapter_semester_data.csv` and run:

```
python3 scripts/build_simple_workbook.py
python3 scripts/build_workbook.py
```

## Gaps

- Spring 2020: no GPAs for other houses (pass/no-pass semester)
- Spring 2020, Spring 2022, Fall 2023: no new-member counts for other houses
- Nothing on other houses before Fall 2018
