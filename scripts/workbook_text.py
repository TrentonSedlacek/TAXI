"""Long text for the workbook's Summary, Old Index Review and Read Me tabs.

The evaluation reflects the data through Spring 2026. Rewrite it when new
semesters are added; the tables on the Summary tab update on their own.
"""

EVALUATION = [
    "Bottom line: Theta Xi is a below-average IFC house on grades, and it has taken in fewer men than it lost "
    "in every calendar year from 2020 through 2025. Over Spring 2024 to Spring 2026 we rank 18th of 23 active IFC houses "
    "on GPA, 20th on recruitment and 20th combined.",

    "GPA: we were below the IFC average (all IFC men) in every reported semester since Fall 2018 except Spring "
    "2022 (+0.010). The gap shrank from -0.42 in Fall 2018 to within 0.06 in Spring 2021, Fall 2021 and Spring "
    "2024 (11th of 25 that term). It opened up again in 2025: -0.32 in Spring 2025, -0.30 in Fall 2025 and -0.23 "
    "in Spring 2026 (3.163 vs 3.392, 17th of 23).",

    "Against all UNL men we were ahead in 4 of the 7 semesters from Spring 2021 to Spring 2024. Campus-wide "
    "numbers after Spring 2024 are not in the repo yet (see Campus Stats), so that comparison stops there.",

    "Size: from 2018 to 2023 we were a median IFC house (62 to 86 men, about 13th of 25). Since Spring 2024 we "
    "have had 43 to 48 men, about 55% of the median house (78 in Spring 2026). Over the same years total IFC "
    "membership grew from 1,641 men (Fall 2018) to 1,980 (Fall 2025), so the drop is ours, not the market's.",

    "Recruiting volume: new members per calendar year were 34 (2019), 34 (2020), 24 (2021), 21 (2022), "
    "19 (2023), 17 (2024) and 15 (2025). The average IFC house took 24 to 27 a year, so our recruiting index "
    "fell from 143 in 2019 to 99 in 2021, 68 in 2024 and 56 in 2025. Our share of all IFC new members fell from "
    "5.4% to 2.3%, where an equal share is about 4%.",

    "Replacement: the replacement index has been under 100 every year from 2020 through 2025: 81 (2020), 86 (2021), "
    "88 (2022), 63 (2023), 59 (2024), 79 (2025). Each year we lost more men than we brought in, which is how "
    "86 men in Fall 2019 became 43 in Spring 2026.",

    "For our size, recruiting is closer to normal: the 2025 size-adjusted index was 89. So the problem is less "
    "that each member recruits badly and more that a smaller house brings in a smaller class, which shrinks the "
    "next class too. Getting back to a stable size means recruiting above the IFC rate for a few years, not just "
    "matching it.",

    "Retention is not the main leak now. UNL's new-member retention for us was 100% (Fall 2024), 75% (Spring "
    "2025, 3 men), 92% (Fall 2025) and 100% (Spring 2026), close to the IFC median each term, and active retention "
    "was 90% to 98%. Our own records show heavier pledge losses in the 2020 to 2023 classes (30 of 86, 35%), so "
    "this has improved.",

    "The current exec team is off to a good start: 6 new members in Spring 2026, 5th most in the IFC (the "
    "average house took 4.1). To match the 2025 average IFC house for the year, Fall 2026 needs about 21 new "
    "members; to replace a 2025-sized loss it needs about 13 (Fall 2025 brought in 12, against an IFC median "
    "of 25).",

    "Houses that score well on both GPA and recruitment over Spring 2024 to Spring 2026 are Pi Kappa Alpha, "
    "Phi Kappa Theta and Sigma Chi: all three recruit big classes and keep GPAs above 3.38. Farmhouse, Sigma Phi "
    "Epsilon and Beta Theta Pi lead on GPA but sit in the bottom half on recruitment, mostly because they bring "
    "in fewer new men per member.",

    "Caveats: there is no Fall 2023 report in the repo, and the 2025 and 2026 CSV exports do not include the "
    "campus-wide GPAs. UNL's Spring 2020 report has no GPAs (pass/no-pass semester). Some old chapter spreadsheets "
    "disagree with UNL's reports (for example Fall 2019: 77 vs 86 members). The UNL report is used everywhere, "
    "and each difference is written in the Chapter Data notes.",
]

OLD_INDEX_NOTES = [
    ("AE List.xlsx, Cleaned tab: Expected Rush (cols A-B)",
     "It divides our new members by our own historical fall or spring average. So 100 means 'normal for Theta "
     "Xi', and most of that history is 2010-2017, when we had 23 to 48 men. A class can look fine against our "
     "past and still be far behind the other houses. v2 divides by the average IFC house in the same year "
     "instead."),
    ("Replacement Rush Index (col C)",
     "The formula is (new / actives) / (average new / average actives). That is a ratio of ratios. It tells you "
     "whether our pledge-to-active ratio beat our historical ratio, but it never uses how many men actually "
     "left, so it can read above 100 in a year we shrank. v2 estimates departures directly (last Fall's roster + "
     "the year's new members - this Fall's roster) and divides new members by that."),
    ("Spring weighting (D39 = 0.25, cols F, G and J)",
     "The 0.25 / 0.75 split was typed in by hand. Because col A already normalizes fall and spring separately, the "
     "weighting works as a weighted average, but the weight itself is arbitrary. The yearly 'Grand Rush' also "
     "appears twice with different math: col J is weighted, while col N is an unweighted average of cols I and L. "
     "v2 needs no weight because every house is compared within the same year, and spring classes are small "
     "for everyone. Your calendar-year grouping (Spring + Fall of the same year) was right, and v2 keeps it."),
    ("Averages (rows 32-34)",
     "B32 averages rows 3:30, but C32, D32 and E32 stop at row 29, so Spring 2024 is left out. The fall average "
     "(D33) covers 14 falls and the spring average (D34) 14 springs, but D32 covers 27 semesters (14 falls, "
     "13 springs). That is why the note says 9.1 x 2 does not equal 14.1 + 3.8."),
    ("GPA block (cols M-S)",
     "'IFC avg GPA' (col N) is really the Greek-men GPA from the UNL tab of GPA Data.xlsx, not an IFC figure. "
     "'TX / IFC avg GPA' (col R) is typed in rather than calculated on every row except the last. Col O "
     "multiplies by 100 and col Q does not, so the two '% change' columns use different units. v2 builds the IFC "
     "average from every chapter's reported GPA, weighted by members."),
    ("Member counts",
     "Several semesters in this tab disagree with UNL's report for the same term: Fall 2019 (77 vs 86 members), "
     "Fall 2020 (81 vs 78 members, 20 vs 28 new members), Spring 2023 (68 vs 71) and Spring 2024 (58 vs 48). v2 "
     "uses the UNL report everywhere and keeps the old value in the Chapter Data notes."),
    ("AE List.xlsx, Disgusting tab: Fall 2024 projection (BD3-BD9)",
     "The projection subtracts drops and graduates from the fall class but leaves out the spring class (about "
     "+5). The drop figure (BD5) is a per-class average but is used as a chapter-wide loss. It also mixes full "
     "class sizes with Fall 2023 'still in' counts. In the history block, the 2022-2023 rush-index cells "
     "(AZ92-AZ94) show reference errors."),
    ("GPA Data.xlsx, UNL tab: statistics",
     "The STDEV row (31) reads rows 1:27, so it includes the header and leaves out Fall 2023 and Spring 2024. The "
     "'paired' t-test lines semesters up by position (Spring 17 with Fall 21, and so on), which is not a real "
     "pairing; the unequal-variance test in row 51 is the one to use. In a few semesters, All Non-Greek is a plain "
     "average of men and women, which ignores that there are more women. Fall 10 'All Greek' equals 'Greek Male' "
     "(3.117), which is probably a typo."),
    ("Scholarship workbooks (2022-2023)",
     "The biweekly grade-check summaries have range errors. The averages skip row 2. In Autumn 2022, the Fall 2021 "
     "class range (rows 36:52) overlaps the Spring 2021 range (rows 35:40). In Fall 2023, the summary block is "
     "shifted one row, so 'Active' shows the chapter average and 'Official' shows a reference error. These files track individual "
     "members, so they stay as they are and are not merged into this workbook."),
]

README = [
    ("What this is", [
        ("Purpose",
         "Compares Theta Xi (UNL) with every other IFC house on grades and recruitment. Every house is covered "
         "semester by semester from Fall 2018, and our own history goes back to 2010. It replaces the "
         "multi-university GPA Data.xlsx and the recruitment index in AE List.xlsx, both kept in legacy/."),
        ("Our chapter",
         "Set on the Lists tab (yellow cell). Type any IFC house name there and every tab recalculates for that "
         "house."),
    ]),
    ("Tabs", [
        ("Summary", "Latest-semester snapshot, calendar-year recruitment, combined ranks, targets and the written "
                    "evaluation."),
        ("TX vs IFC", "One row per semester: GPA, size, new members and retention against the IFC, with charts."),
        ("Combined Score", "Every IFC house over a window you choose. GPA and recruitment are scored as percentiles "
                           "and blended with weights you set."),
        ("Recruitment Index", "Recruitment Index v2 by calendar year (Spring + Fall, one exec term): recruiting, "
                              "size-adjusted, replacement and growth vs the IFC."),
        ("GPA / Members / New Members Matrix", "Every IFC house by semester, with IFC averages, medians and our "
                                               "rank at the bottom."),
        ("TX History", "Our GPA and size since 2010 against the campus benchmarks, plus pledge-class retention."),
        ("Chapter Data", "The raw numbers from UNL's reports: one row per chapter per semester, all four councils."),
        ("Campus Stats", "All-UNL GPAs by group (men, women, Greek, non-Greek)."),
        ("IFC Size (legacy)", "Average IFC chapter size before 2018, from AE List.xlsx."),
        ("Lists", "Semester order, IFC chapter list and settings."),
        ("Old Index Review", "What was off in the old recruitment score and GPA math, and what v2 does instead."),
    ]),
    ("Adding a new semester", [
        ("1. Chapter Data",
         "Paste the new scorecard's rows at the bottom (Semester, Council, Chapter, Total Members, New Members, "
         "GPA, ranks, retention). Type the semester exactly like 'Fall 2026', and spell chapters the way the Lists "
         "tab does (for example 'Farmhouse'). Column B fills itself in. Formulas read rows 2 to 2500."),
        ("2. New chapters",
         "If a new IFC house appears, add it to column D on the Lists tab. The matrices and Combined Score have "
         "five spare rows for new houses."),
        ("3. Campus Stats",
         "Type in the all-campus GPAs from UNL's All-Community Academic Report (fsl.unl.edu). Orange cells are "
         "the ones still missing."),
        ("4. Settings",
         "Update 'Latest semester' on the Lists tab, the recruitment year on Summary and the window on Combined "
         "Score. Then reread the written evaluation on Summary, since it does not update itself."),
        ("Room to grow",
         "Semesters are set up through Spring 2028. To go further, add semesters to Lists column A, add rows to "
         "Campus Stats, copy the last column of each matrix one to the right, and extend the TX vs IFC and "
         "Recruitment Index rows."),
    ]),
    ("Definitions", [
        ("GPA", "The term GPA printed on UNL's report. 'IFC average' is weighted by members, so it equals the GPA of "
                "all IFC men. 'IFC median' is the middle house."),
        ("New members", "As printed by UNL. Scorecards from 2024 on say 'new members initiated', so pledges who "
                        "dropped before initiation are not counted. Spring classes are small for every house."),
        ("Calendar year", "Spring + Fall of the same year. This matches the exec term: elected in November, "
                          "trained in December, running recruitment from January."),
        ("Indexes", "100 = matches the benchmark, 150 = 50% better, 50 = half as good."),
        ("Percentiles", "Among IFC houses in the Combined Score window. 100% = best, 0% = worst."),
    ]),
    ("Sources and gaps", [
        ("Reports", "UNL IFC report cards for Fall 2018 to Fall 2022 and Community Scorecards for Spring 2023 to "
                    "Spring 2026, all in source-data/. The data is also kept as a CSV in data/."),
        ("Campus GPAs", "The UNL tab of GPA Data.xlsx, through Spring 2024."),
        ("Missing", "Fall 2023 scorecard. GPAs for Spring 2020 (pass/no-pass semester). New-member counts for other "
                    "houses in Spring 2020 and Spring 2022. Campus GPAs after Spring 2024. The Sigma Tau Gamma row in "
                    "Fall 2022, which is blacked out on the PDF."),
        ("Chapter records", "Theta Xi numbers before Fall 2018 and for Fall 2023 come from AE List.xlsx and GPA "
                            "Data.xlsx. Where UNL left our Spring 2020 GPA or our Spring 2020 and Spring 2022 new "
                            "members blank, the chapter records fill the gap, as noted in Chapter Data."),
    ]),
]
