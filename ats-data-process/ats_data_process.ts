function main(workbook: ExcelScript.Workbook) {
    const worksheet: ExcelScript.Worksheet = workbook.getWorksheet(WritingSettings.WRITING_SHEET);
    const data: string[][] = worksheet.getUsedRange().getTexts(); // Get data before appending the header and creating empty range.
  
    writeHeader(worksheet); // Pre-write output header so that columns for writing are available in the used range.
    writeData(worksheet.getUsedRange(), processData(data));
  }
  
  function writeHeader(worksheet: ExcelScript.Worksheet) {
    worksheet.getCell(0, WritingSettings.COLUMN_HIRED).setValue(WritingSettings.TITLE_HIRED);
    worksheet.getCell(0, WritingSettings.COLUMN_HR_INTERVIEW).setValue(WritingSettings.TITLE_HR_INTERVIEW);
    worksheet.getCell(0, WritingSettings.COLUMN_LAST_CONTACT).setValue(WritingSettings.TITLE_LAST_CONTACT);
    worksheet.getCell(0, WritingSettings.COLUMN_TECH_INTERVIEW).setValue(WritingSettings.TITLE_TECH_INTERVIEW);
    worksheet.getCell(0, WritingSettings.COLUMN_FINAL_INTERVIEW).setValue(WritingSettings.TITLE_FINAL_INTERVIEW);
  }
  
  function writeData(worksheetRange: ExcelScript.Range, data: string[][]) {
    const columnHired: ExcelScript.Range = worksheetRange.getColumn(WritingSettings.COLUMN_HIRED);
    const columnLastContact: ExcelScript.Range = worksheetRange.getColumn(WritingSettings.COLUMN_LAST_CONTACT);
    const columnHr: ExcelScript.Range = worksheetRange.getColumn(WritingSettings.COLUMN_HR_INTERVIEW);
    const columnTech: ExcelScript.Range = worksheetRange.getColumn(WritingSettings.COLUMN_TECH_INTERVIEW);
    const columnFinal: ExcelScript.Range = worksheetRange.getColumn(WritingSettings.COLUMN_FINAL_INTERVIEW);
  
    const dataHired: string[][] = getColumnData(data, ProcessParams.INDEX_HIRED, WritingSettings.TITLE_HIRED);
    const dataLastContact: string[][] = getColumnData(data, ProcessParams.INDEX_LAST_CONTACT, WritingSettings.TITLE_LAST_CONTACT);
    const dataHr: string[][] = getColumnData(data, ProcessParams.INDEX_HR_INTERVIEW, WritingSettings.TITLE_HR_INTERVIEW);
    const dataTech: string[][] = getColumnData(data, ProcessParams.INDEX_TECH_INTERVIEW, WritingSettings.TITLE_TECH_INTERVIEW);
    const dataFinal: string[][] = getColumnData(data, ProcessParams.INDEX_FINAL_INTERVIEW, WritingSettings.TITLE_FINAL_INTERVIEW);
  
    columnHired.setValues(dataHired);
    columnLastContact.setValues(dataLastContact);
    columnHr.setValues(dataHr);
    columnTech.setValues(dataTech);
    columnFinal.setValues(dataFinal);
  }
  
  function getColumnData(data: string[][], dataColumnIndex: number, columnTitle: string): string[][] {
    let columnData: string[][] = new Array<string[]>(data.length + 1); // +1 to store header as well.
    columnData[0] = [columnTitle]; // Store column header because the entire column data has to be written.
  
    for (let i = 1; i < columnData.length; i++) {
      columnData[i] = [data[i - 1][dataColumnIndex]];
    }
  
    return columnData;
  }
  
  function processData(data: string[][]): string[][] {
    let processedData: string[][] = new Array<string[]>();
  
    // Process data starting with 2nd row (index 1). Because the first row of data coming from excel is the header.
    for (let i = 1; i < data.length; i++) {
      const processedRow = processRow(data[i]);
      processedData.push(processedRow);
    }
  
    return processedData;
  }
  
  function processRow(dataRow: string[]): string[] {
    let processedRow: string[] = new Array<string>(ProcessParams.COLUMNS);
    const statusData: string = dataRow[ReadingSettings.COLUMN_STATUS];
    const commentsData: string = dataRow[ReadingSettings.COLUMN_COMMENTS];
    const applicationDate: string = dataRow[ReadingSettings.COLUMN_APPLICATION];
  
    if (statusData === '')
      return processedRow;
  
    // All data can be found in the status, except for the last date of contact.
    // Last date of contact can be obtained from both status and comments.
    processedRow[ProcessParams.INDEX_HIRED] = getHiredData(statusData);
    processedRow[ProcessParams.INDEX_LAST_CONTACT] = getLastContactData(statusData, commentsData, applicationDate);
    processedRow[ProcessParams.INDEX_HR_INTERVIEW] = getHrInterviewData(statusData);
    processedRow[ProcessParams.INDEX_TECH_INTERVIEW] = getTechInterviewData(statusData);
    processedRow[ProcessParams.INDEX_FINAL_INTERVIEW] = getFinalInterviewData(statusData);
  
    return processedRow;
  }
  
  function getDateFromStatusWithKey(status: string, stepKey: string): string {
    const statusChunks: string[] = status.split(ReadingSettings.DELIMITER_CHUNKS);

    // Iterate inversely through status to get the most recent date.
    for (let i = statusChunks.length - 1; i >= 0; i--) {
      const statusChunkData: string[] = statusChunks[i].split(ReadingSettings.DELIMITER_DATA);
      if (statusChunkData[ReadingSettings.INDEX_STATUS_STEP] === stepKey)
        return statusChunkData[ReadingSettings.INDEX_STATUS_DATE];
    }
  
    return '';
  }
  
  const getHrInterviewData = (status: string) => getDateFromStatusWithKey(status, ReadingSettings.KEY_STATUS_HR);
  const getTechInterviewData = (status: string) => getDateFromStatusWithKey(status, ReadingSettings.KEY_STATUS_TECH);
  const getFinalInterviewData = (status: string) => getDateFromStatusWithKey(status, ReadingSettings.KEY_STATUS_FINAL);
  
  function getHiredData(status: string): string {
    const statusChunks: string[] = status.split(ReadingSettings.DELIMITER_CHUNKS);
    // Correct hired details can only be in the last chunk.
    const lastStatusData: string[] = statusChunks[statusChunks.length - 1].split(ReadingSettings.DELIMITER_DATA);
  
    if (lastStatusData[ReadingSettings.INDEX_STATUS_STEP] === ReadingSettings.KEY_STATUS_HIRED)
      return lastStatusData[ReadingSettings.INDEX_STATUS_DATE];
  
    return '';
  }
  
  function getLastContactData(status: string, comments: string, applicationDate: string): string {
    const recontactedDate: string = getRecontactedDateFromComments(comments);
    if(recontactedDate)
      return recontactedDate;
      
    const contactedDate: string = getDateFromStatusWithKey(status, ReadingSettings.KEY_STATUS_CONTACT);
    if(contactedDate)
      return contactedDate;

    return applicationDate;
  }

  function getRecontactedDateFromComments(comments: string): string {
    if(comments === '')
      return '';

    const commentsChunks: string[] = comments.split(ReadingSettings.DELIMITER_CHUNKS);
  
    // Inverse iteration because we can stop when we find the last recontacted date.
    for (let i = commentsChunks.length - 1; i >= 0; i--) {
      const comment: string[] = commentsChunks[i].split(ReadingSettings.DELIMITER_DATA);
      const words: string[] = comment[ReadingSettings.INDEX_COMMENT_INFO].split(' ');
      const hasRecontacted: boolean = isWordSimilarToCorrectWord(words[0], ReadingSettings.KEY_COMMENT_RECONTACTED);
      const hasRecontactat: boolean = isWordSimilarToCorrectWord(words[0], ReadingSettings.KEY_COMMENT_RECONTACTAT);
      if (hasRecontacted || hasRecontactat)
        return comment[ReadingSettings.INDEX_COMMENT_DATE];
    }

    return '';
  }
  
  // Checks how similar a human-written word is to the word we are looking for ('recontacted').
  // The human-written word can be mistaken by a letter and the check would still be true.
  // E.g. 'rcontacted' would be considered the same as 'recontacted'. 
  function isWordSimilarToCorrectWord(word: string, correctWord: string): boolean {
    word = word.toLowerCase();
    correctWord = correctWord.toLowerCase();
    const sensitivity: number = 1;

    if(word.length < correctWord.length - sensitivity || word.length > correctWord.length + sensitivity)
        return false;
  
    let mistakenLetters: number = 0;
    for (let iWord = 0, iRecontacted = 0; iWord < word.length && iRecontacted < correctWord.length; iWord++ , iRecontacted++) {
      const char: string = word.charAt(iWord);
      const correctChar: string = correctWord.charAt(iRecontacted);
  
      if (char === correctChar) continue;
  
      if (++mistakenLetters > sensitivity) return false;
  
      const wordCheckMistake: WordCheckMistake = getWordCheckMistake(word.length, correctWord.length);
      if (wordCheckMistake == WordCheckMistake.MISSING_LETTER)
        iWord--; // Only recontacted index should increment, so decrement word index before next iteration.
      else if (wordCheckMistake == WordCheckMistake.EXTRA_LETTER)
        iRecontacted--; // Only word index should increment, so decrement recontacted index before next iteration.
    }
  
    return true;
  }
  
  function getWordCheckMistake(checkedWordLength: number, correctWordLength: number): WordCheckMistake {
    if (checkedWordLength == correctWordLength) return WordCheckMistake.WRONG_LETTER;
    if (checkedWordLength < correctWordLength) return WordCheckMistake.MISSING_LETTER;
    if (checkedWordLength > correctWordLength) return WordCheckMistake.EXTRA_LETTER;
  }
  
  enum WordCheckMistake {
    MISSING_LETTER,
    EXTRA_LETTER,
    WRONG_LETTER
  }
  
  enum ProcessParams {
    COLUMNS = 5,
    INDEX_HIRED = 0,
    INDEX_LAST_CONTACT = 1,
    INDEX_HR_INTERVIEW = 2,
    INDEX_TECH_INTERVIEW = 3,
    INDEX_FINAL_INTERVIEW = 4
  }
  
  enum ReadingSettings {
    COLUMN_APPLICATION = 3,
    COLUMN_STATUS = 7,
    COLUMN_COMMENTS = 11,
    DELIMITER_CHUNKS = '|',
    DELIMITER_DATA = ';',
    KEY_STATUS_CONTACT = 'Contacted',
    KEY_STATUS_HR = 'HR Interview',
    KEY_STATUS_TECH = 'Technical Interview',
    KEY_STATUS_FINAL = 'Final Interview',
    KEY_STATUS_HIRED = 'Hired',
    KEY_COMMENT_RECONTACTED = 'recontacted',
    KEY_COMMENT_RECONTACTAT = 'recontactat',
    INDEX_STATUS_DATE = 0,
    INDEX_STATUS_STEP = 1,
    INDEX_COMMENT_DATE = 0,
    INDEX_COMMENT_INFO = 1
  }
  
  enum WritingSettings {
    COLUMN_HIRED = 15,
    COLUMN_LAST_CONTACT = 16,
    COLUMN_HR_INTERVIEW = 17,
    COLUMN_TECH_INTERVIEW = 18,
    COLUMN_FINAL_INTERVIEW = 19,
    TITLE_HIRED = 'Data angajării',
    TITLE_LAST_CONTACT = 'Data Ultimei Contactări',
    TITLE_HR_INTERVIEW = 'Data HR Interview',
    TITLE_TECH_INTERVIEW = 'Data Technical Interview',
    TITLE_FINAL_INTERVIEW = 'Data Final Interview',
    WRITING_SHEET = 'Sheet'
  }