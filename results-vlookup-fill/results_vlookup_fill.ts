
/**
 * Fills an excel sheet with vlookup functions.
 * The vlookup functions look for the language evaluation results on oral, reading, grammar and writing.
 * The vlookup functions will look based on the name found in @LOOK_UP_COLUMN_VALUE of @VlookUpParams enum.
 */
function main(workbook: ExcelScript.Workbook) {
    // Column where to write data.
    let columnOral = workbook
      .getWorksheet("Sheet1")
      .getUsedRange()
      .getColumn(ExcelWritingValues.START_COL + 0);
    let columnRead = workbook
      .getWorksheet("Sheet1")
      .getUsedRange()
      .getColumn(ExcelWritingValues.START_COL + 1);
    let columnGram = workbook
      .getWorksheet("Sheet1")
      .getUsedRange()
      .getColumn(ExcelWritingValues.START_COL + 2);
    let columnWrit = workbook
      .getWorksheet("Sheet1")
      .getUsedRange()
      .getColumn(ExcelWritingValues.START_COL + 3);
    // Number of row where to write.
    let rowNr = ExcelWritingValues.START_ROW;
  
    // Fill cells with data.
    for (let i = 1; i <= ExcelWritingValues.DATASETS_TO_WRITE; i++) {
      const modifier = (ExcelWritingValues.START_ROW + i).toString();
      columnOral
        .getCell(rowNr, 0)
        .setValue(getData(modifier, DataType.ORAL));
      columnRead
        .getCell(rowNr, 0)
        .setValue(getData(modifier, DataType.READ));
      columnGram
        .getCell(rowNr, 0)
        .setValue(getData(modifier, DataType.GRAM));
      columnWrit
        .getCell(rowNr, 0)
        .setValue(getData(modifier, DataType.WRIT));
  
      rowNr++;
    }
  }
  
  /**
   * Gets the data to be written in excel based on constant chunks of data and a modifier added between them.
   * 
   * @param modifier - The modifier to be added between data chunks.
   * @returns - The data to be written in excel.
   */
  function getData(modifier: string, typeOfData: DataType): string {
    // Data chunks to be written in every cell with a modifier added between them.
    // Start and Date data.
    const lookupData = "=VLOOKUP(" + VlookUpParams.LOOK_UP_COLUMN_VALUE;
    // Date and Language data.
    const tableArrayData = ", " + VlookUpParams.TABLE_ARRAY +", ";
    // Language and Interval data.
    const rangeLookupData = ", " + VlookUpParams.RANGE_LOOKUP +")";
  
    return lookupData + modifier + tableArrayData + typeOfData + rangeLookupData;
  }
  
  // In what column to look for data.
  enum DataType {
    ORAL = 5, // Oral results are on column 5 
    READ, // Reading results are on column 6
    GRAM, // Grammar results are on column 7
    WRIT // Writing results are on column 8
  }
  
  // Modify next enums to change vlookup reading and this script's writing.
  
  enum VlookUpParams {
    LOOK_UP_COLUMN_VALUE = "G",
    TABLE_ARRAY = "DATABASE!B1:K323",
    RANGE_LOOKUP = "FALSE"
  }
  
  enum ExcelWritingValues {
    START_ROW = 0,
    START_COL = 7,
    DATASETS_TO_WRITE = 20
  }