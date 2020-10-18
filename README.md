# recruitment-scripts
Scripts that automate excel processes of a recruitment department.

I chose to add these scripts to github as they are more of a personal project and not part of my day-to-day work as a recruiter.

## [English Evaluation Tracker](English-evaluation-tracker/programati_eng.ts)
### [programati_eng.ts](English-evaluation-tracker/programati_eng.ts)
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
The end data will be processed by the "programati_eng.ts" script.
