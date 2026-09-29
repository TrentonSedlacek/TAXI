# Theta Xi (UNL) GPA and recruitment

Chapter scholarship and recruitment data for Theta Xi at the University of Nebraska-Lincoln, with a workbook that compares us to every other IFC house.

## Start here

`analysis/Theta Xi vs UNL Fraternities.xlsx` has these tabs:

- **Summary**: the latest semester at a glance, calendar-year recruitment, combined ranks, targets for the current exec team, and a written evaluation.
- **TX vs IFC**: one row per semester since Fall 2018, covering GPA, size, new members and retention against the IFC.
- **Combined Score**: every IFC house scored on GPA and recruitment over a window you pick.
- **Recruitment Index**: the reworked recruitment index (v2), by calendar year to match the November-to-November exec term.
- **Old Index Review**: what was off in the old recruitment score and GPA math.

The Read Me tab inside the workbook explains every tab and how to add a new semester.

## Folders

| Folder | What's in it |
|---|---|
| `analysis/` | The combined GPA + recruitment workbook |
| `data/` | Tidy CSVs the workbook is built from: `chapter_semester_data.csv` (one row per chapter per semester, all councils) and `campus_stats.csv` (all-UNL GPAs by group) |
| `source-data/ifc-report-cards/` | UNL IFC report cards, Fall 2018 to Fall 2022 (PDF) |
| `source-data/community-scorecards/` | UNL community scorecards, Spring 2023 to Spring 2026 (PDF and CSV) |
| `chapter-scholarship/` | Chapter scholarship-program workbooks for 2022 and 2023 (member grade checks) |
| `legacy/` | The original `GPA Data.xlsx` (multi-university) and `AE List.xlsx` (old recruitment index) |
| `scripts/` | `build_workbook.py` rebuilds the workbook from `data/`; `workbook_text.py` holds the tab text |

## Adding a semester

Either paste the new rows into the workbook's Chapter Data tab (the formulas pick them up), or add them to `data/chapter_semester_data.csv` and run:

```
python3 scripts/build_workbook.py
```

The script writes formulas only. Excel calculates them when the file opens.

## Known gaps

- There is no Fall 2023 scorecard. Theta Xi's Fall 2023 numbers come from chapter records.
- UNL's Spring 2020 report has no GPAs (pass/no-pass semester).
- The Spring 2020 and Spring 2022 reports leave out new-member counts.
- Campus-wide GPAs after Spring 2024 are not in the repo. UNL posts them in the All-Community Academic Report on fsl.unl.edu.
