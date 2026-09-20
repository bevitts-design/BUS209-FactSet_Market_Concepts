# Build a 10-Stock Portfolio — infographic source notes

Created September 18, 2026 for BUS209. Standalone classroom graphic; not published or added to the course website.

## Authentic screenshot insets

At Bethany's request, the final graphic includes an enlarged reference strip containing actual pixels from the signed-in Online Assistant screenshot supplied by the originating task. The toolbar icon and dropdown are cropped from the source capture, not generated imitations. The source screenshot is low resolution; nearest-neighbor enlargement preserves its pixels and readable editorial labels identify the controls.

- Raw source: `factset-creating-portfolio-help.png` (1363 × 768), copied without alteration from the originating task's original screen-capture image bytes. It shows the Creating a Portfolio section of https://my.apps.factset.com/oa/pages/20747#portfolio.
- Original six-step artwork: `factset-10-stock-portfolio-base.png`.
- Derived crops: `factset-create-edit-watchlist-icon.png` and `factset-ticker-shares-dropdown.png`.
- Maintained composition source: `../scripts/build-portfolio-infographic.cjs`. Crop coordinates and enlargement factors are recorded there. This inserts a 320-pixel reference strip below the title and leaves all six-step artwork intact.
- Final output: `factset-10-stock-portfolio.png` (1024 × 1856). Visually reviewed for legibility, complete six-step sequence and recognizable authentic controls.
- Rebuild with the bundled Node runtime and its `node_modules` directory in `NODE_PATH`, running `scripts/build-portfolio-infographic.cjs`.

## Scope and verification

Revised September 18, 2026 to use FactSet Web's Watchlist Editor workflow: Create/Edit Watchlist → Ticker + Shares → copy 10 Excel data rows → paste and check → Save → reopen and verify.

The originating live-help task (01a0b513-5202-7943-a2f0-92bb327fe2d1) verified the signed-in FactSet Online Assistant page by screen capture on September 18, 2026 and supplied the documented labels and behavior. This revision uses that verified documentation; completed end-to-end portfolio creation has not been verified. The final image was visually inspected for legibility, sequence, exact labels, and removal of obsolete instructions.

## Sources

- FactSet Online Assistant, “Managing Watchlists and Basic Portfolios,” page 20747: https://my.apps.factset.com/oa/pages/20747 (signed-in documentation).
- Verified behavior: Watchlist Editor creates watchlists and basic non-time-series portfolios. FactSet Web launches it through Create/Edit Watchlist. Ticker + Shares accepts stock identifiers in the first column and share counts in the second, copied from Excel. Save allows a directory and filename without spaces or prohibited punctuation. FactSet saves an OFDB file and automatically creates a corresponding ACCT account for Portfolio Analysis / Portfolio View.
- Basic portfolios do not retain historical holdings over time. The graphic therefore contains no holdings-date entry or historical portfolio instructions.
- This source replaces the previous legacy Quick Portfolio citations. Quick Portfolio/@QP and Modify Settings/Publish for PA are removed from the graphic.

## Teaching assumptions

- The student selects 10 distinct stocks under the assignment's own rules; no securities are recommended.
- The $100,000 equal-weight example is optional and hypothetical, not an assignment requirement. $100,000 / 10 = $10,000; $10,000 / $50 = 200 shares.
- Equal numbers of shares generally do not produce equal portfolio weights. Rounded whole-share quantities may leave cash and weights may subsequently drift.
- The chart describes entering simulated holdings, not placing trades.
- The miniature Excel table uses “Stock 1” and “Stock 2” as layout placeholders; students must copy actual resolved stock identifiers from their Excel file.
- Personal is the default individual classroom save location unless the instructor specifies a class folder.
