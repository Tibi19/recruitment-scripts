
/**
 * This script populates a number of cells in a single column with data.
 * In format "WWWW + m XXXX + m YYYY + m ZZZZ",
 * Where "m" is a variable modifier.
 * Data is first written with positive and 0 modifiers and then negative modifiers.
 * 
 * The data is a combination of excel functions that reference data from "ORAL SCHEDULE" sheet.
 * The referenced data is formatted as "date#language#interval". 
 * The end data will be processed by the "tracker_eng.ts" script.
 */
function main(workbook: ExcelScript.Workbook) {
    // How many negatively modified datasets to write.
    let negModifiedDatasetsNr = 400;
    // Column where to write data.
    let column = workbook
     .getWorksheet("Tracker Programari")
     .getUsedRange()
     .getColumn(9);
    // Number of row where to start writing.
    let rowNr = 1;
    
    // How many datasets to write with a positive modifier.
    let posModifiedDatasetsNr = 2;
    // Write positively and 0 modified datasets.
    for(let i = posModifiedDatasetsNr; i >= 0; i--) {
        let modifier = "+" + i;
        column
         .getCell(rowNr++, 0)
         .setValue(getData(modifier));
    }

    // Fill cells with data.
    // From now on write only negatively modified datasets.
    for(let i = -1; i > (-negModifiedDatasetsNr); i--) {
        column
         .getCell(rowNr++, 0)
         .setValue(getData(i.toString()));
        console.log(i);
    }
}

/**
 * Gets the data to be written in excel based on constant chunks of data and a modifier added between them.
 * 
 * @param modifier - The modifier to be added between data chunks.
 * @returns - The data to be written in excel.
 */
function getData(modifier: string): string {
    // Data chunks to be written in every cell with a modifier added between them.
    // Start and Date data.
    const startDateData = "=CONCAT(INDIRECT(ADDRESS(K2";
    // Date and Language data.
    const dateLanguageData = ", 1,,,\"ORAL SCHEDULE\")), \"#\", INDIRECT(ADDRESS(K2";
    // Language and Interval data.
    const languageIntervalData = ", 4,,,\"ORAL SCHEDULE\")), \"#\", TEXT(INDIRECT(ADDRESS(K2";
    // Interval and End data.
    const intervalEndData = ", 6,,,\"ORAL SCHEDULE\")), \"hh:mm\"))";

    return startDateData + modifier + dateLanguageData + modifier + languageIntervalData + modifier + intervalEndData;
}