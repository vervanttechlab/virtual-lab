# DAL3 Virtual Lab

A mobile-first virtual lab for the TESDA **Data Analytics Level III** course (running case: *Tindahan ni Aling Nena Mini-Mart*).
Trainees can learn and practice on a phone or a PC: lessons, an Excel-like spreadsheet with real formulas, the Analysis ToolPak
(Descriptive Statistics, t-Test, Regression), charts, auto-checked activities and self-check quizzes. The trainer watches live
progress, opens or closes activities, and grades written answers.

- **Students:** open the link, sign in with the class code, student ID and PIN, then **Add to Home Screen**. The app works offline;
  your answers are sent when you are back online.
- **Trainer board:** add `#/trainer` to the link.

Content available: Day 2 (Descriptive Statistics) and Day 3 (Probability, Correlation, Regression, t-Test). More days are added as
the course goes on.

Built as a static site (no build step): `index.html`, `js/` (ES modules), `data/` (generated day content, answers stored only as
salted hashes). The back end is a Google Apps Script web app set in `js/config.js`.

Third-party: [HyperFormula](https://hyperformula.handsontable.com/) (GPL-v3, see `vendor/HYPERFORMULA-LICENSE.txt`) and
[Chart.js](https://www.chartjs.org/) (MIT, `vendor/CHARTJS-LICENSE.md`). This project's code is released under the GPL-3.0 to match.
