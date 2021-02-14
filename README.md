# recruitment-scripts
Scripts that automate excel processes of a recruitment department.

I chose to add these scripts to github as they are more of a personal project and not part of my day-to-day work responsibilities as a recruiter.

## [English Evaluation Tracker](English-evaluation-tracker/tracker_eng.ts)
### [tracker_eng.ts](English-evaluation-tracker/tracker_eng.ts)
Counts how many candidates are scheduled for the oral English evaluation.
The script tracks data for the next 2 days and skips weekends.

Candidates are scheduled by Recruitment in availabilities set by the Language department.
Availabilities are read from "Tracker Programari" sheet and updated by users of the database.

Tracking data is wrote in a tabel in "Tracker Programari" sheet.
Scheduling data is referenced in the "Tracker Programari" sheet from "ORAL SCHEDULE" sheet.
Columns J, K, L contain 'back-end data' of the script and are hidden from the user.

### [tracker_sheet_data_fill.ts](English-evaluation-tracker/tracker_sheet_data_fill.ts)
This script populates a number of cells in a single column with data.

The data is a combination of excel functions that reference data from "ORAL SCHEDULE" sheet.
The referenced data is formatted as "date#language#interval". 
The end data will be processed by the "tracker_eng.ts" script.

## [Results Vlookup Filler](results-vlookup-filler/results_vlookup_fill.ts)
Fills an excel sheet with vlookup functions that search for a candidate's language results.
Sometimes, we might need to check if the candidates have language results from the past and we might do so for a large number of them.
This script greatly simplifies the process.
