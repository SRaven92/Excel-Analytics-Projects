# TCR Car Collection & Race Tracker

An Excel-based inventory and performance tracking system created to organise a vintage TCR car collection and its historical race-testing data.

What started as a spreadsheet project for my dad became a structured system for managing **250+ cars**, preserving historical test results, and supporting future race testing.

## 🚗 Features

- Central **Car Garage** containing each car and its specifications
- Permanent unique `TCR-####` identifiers for individual cars
- Dedicated **Add a Car** workflow for extending the collection
- Separate **Inside Lane** and **Outside Lane** performance testing
- Five-lap test entry with automatic fastest, slowest, and average calculations
- Automatic combined-average ranking across both lanes
- Data validation and duplicate-ID checks
- Automatic flags for missing or inconsistent results
- Preserved source data for traceability and auditing
- Dropdown-controlled fields for consistent data entry

## 📊 Excel Techniques Used

The workbook uses structured Excel tables, formulas, validation, and conditional formatting, including:

- `XLOOKUP`
- `COUNTIF` / `COUNTIFS`
- `SUMIF`
- `SUMPRODUCT`
- `AVERAGE`
- `MIN`
- `MAX`
- Structured table references
- Data validation
- Conditional formatting

## 🔍 Data Quality & Traceability

The original files contained historical summary results rather than all of the individual lap times used to produce them.

Rather than inventing missing data or silently correcting discrepancies, the workbook:

- Preserves the original recorded values
- Identifies the original source data
- Flags inconsistent historical results for review
- Leaves genuinely missing data blank
- Recalculates rankings only where the required data is available

This was an intentional **audit-first approach to working with imperfect real-world data**.

## 🛠️ Skills Demonstrated

`Excel` `Data Cleaning` `Data Modelling` `Data Validation` `Data Quality` `XLOOKUP` `Structured Tables` `Conditional Formatting` `Automated Calculations` `Traceability`

## 💡 Background

This project was originally built for my dad's vintage TCR car collection, so some of the workbook interface contains deliberately informal instructions and humour.

The underlying system, however, was designed using the same principles I would apply to a professional data solution:

- Clear and consistent identifiers
- Controlled data entry
- Structured source data
- Automated calculations
- Exception reporting
- Data-quality checking
- Preservation of original records
- Traceability between source and processed data

## 📁 Project File

`TCR_Car_Collection_Tracker.xlsx`

---

*Built as a real-world Excel project to turn a collection of inconsistent historical spreadsheets into a structured, maintainable, and user-friendly data system.*