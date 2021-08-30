# recruitment-scripts
Scripts that automate recruitment processes and tasks in excel.

## [ATS Data Processing](ats-data-process/ats_data_process.ts)
Processes data exported by an ATS (Applicant Tracking System). 
The data is thus made ready for use in Microsoft Power BI by parsing comments and statuses that are formatted in strings by the ATS.

The result is outputted in columns that are appended to the original sheet. 
The output data consists of simple dates that can be easily processed in Power BI.

## [English Evaluation Tracker](English-evaluation-tracker/tracker_eng.ts)
### [tracker_eng.ts](English-evaluation-tracker/tracker_eng.ts)
Counts how many candidates are scheduled for the oral English evaluation.
The script tracks data for the next 2 days and skips weekends.

Candidates are scheduled by Recruitment in availabilities set by the Language department.
Availabilities are read from "Tracker Programari" sheet and updated by users of the database.

Tracking data is wrote in a tabel in "Tracker Programari" sheet.
Scheduling data is referenced in the "Tracker Programari" sheet from "ORAL SCHEDULE" sheet.
Columns J, K, L contain back-end data of the script and are hidden from the user.

### [tracker_sheet_data_fill.ts](English-evaluation-tracker/tracker_sheet_data_fill.ts)
This script populates a number of cells in a single column with data.

The data is a combination of excel functions that reference data from "ORAL SCHEDULE" sheet.
The referenced data is formatted as "date#language#interval". 
The end data will be processed by the "tracker_eng.ts" script.

## [Vlookup Filler](results-vlookup-filler/results_vlookup_fill.ts)
Fills a sheet with vlookup functions. 
Helps to easily access candidate results in an excel database.
