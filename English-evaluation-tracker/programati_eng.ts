/**
 * Counts how many candidates are scheduled for the oral English evaluation.
 * The script tracks data for the next 2 days and skips weekends.
 * 
 * Candidates are scheduled by Recruitment in availabilities set by the Language department.
 * Availabilities are read from "Tracker Programari" sheet and updated by users of the database.
 *  
 * Tracking data is wrote in a tabel in "Tracker Programari" sheet.
 * Scheduling data is referenced in the "Tracker Programari" sheet from "ORAL SCHEDULE" sheet.
 * Columns J, K, L contain back-end data of the script and are hidden from the user.
 */
function main(workbook: ExcelScript.Workbook) {
  // Save start time of the script for debugging.
  // const startMillis = Date.now();

  // The sheet of the tracker.
  const trackerSheet = workbook.getWorksheet("Tracker Programari");
  // Read data that will be processed from column J.
  const excelData = trackerSheet
    .getUsedRange()
    .getColumn(9)
    .getTexts();

  // Create and initialize tracker.
  const tracker = new Tracker(trackerSheet);
  tracker.init();

  // Update tracker.
  updateTracker(excelData, tracker);
  // Write tracker to excel.
  tracker.writeAppointmentsNumbers();
  // Inform user of the last time the script was run.
  writeLastUpdate(trackerSheet);

  // Write script time.
  // writeScriptTime(startMillis, trackerSheet);
}

/**
 * Debugging tool to write script time in column L.
 */
function writeScriptTime(startMillis: number, sheet: ExcelScript.Worksheet) {
  // Get size of column L, the value in K4.
  const size = sheet
    .getCell(3, 10)
    .getValue();
  // Write in column L's last row the difference in time between
  // The start of the script and now in seconds.
  sheet
    .getCell(size, 11)
    .setValue((Date.now() - startMillis) / 1000);
}

/**
 * Updates the tracker to count appointments.
 * 
 * @param excelData - The data to be processed by the tracker.
 * @param tracker - The tracker that counts language evaluation appointments.
 */
function updateTracker(excelData: string[][], tracker: Tracker) {
  // Until what date we should interate.
  const breakDay = Utils.getDay(-7);

  for (let i = 1; i < excelData.length; i++) {
    // Each row represents the data of an appointment for a language evaluation.
    const appointment = new Appointment(excelData[i][0]);
    appointment.init();

    // If we have reached an error in excel, stop iteration.
    if (appointment.hasError()) { break; }
    // If appointment is not valid, continue to next iteration.
    if (!appointment.isValid()) { continue; }

    const appDate = appointment.getDate();
    // If iteration reached a candidate scheduled on the break day,
    // Or before it, stop iteration.
    if (Utils.isBeforeDate(appDate, breakDay)) { break; }

    // Finds the index of the appointment date in the next days array.
    // -1 if the appointment date is not a day we are tracking appointments for. 
    const appDayIndex = Utils.findDate(appDate, tracker.getDays());
    // If appointment day index is less than 0, we shouldn't track this appointment;
    // Or if appointment language is not a language we are tracking;
    // Continue to next iteration.
    if ( (appDayIndex < 0) || (!Utils.canFindString(appointment.getLanguage(), tracker.getLanguage())) ) { continue; }

    // The time interval in which the candidate has been scheduled.  
    const appInterval = appointment.getInterval();
    processAppInterval(tracker, appInterval, appDayIndex);
  }
}

/**
 * Informs the excel user of the last time this script was run.
 * 
 * @param sheet - The sheet where to write last update.
 */
function writeLastUpdate(sheet: ExcelScript.Worksheet) {
  const currentTime = new Date(Date.now());

  // Write today's date in cell B10.
  sheet.getCell(9, 1)
    .setValue(currentTime.toLocaleDateString());
  // Write current hour in cell D10.  
  sheet.getCell(9, 3)
    .setValue(currentTime.toLocaleTimeString());
}

/**
 * Processes an interval where a candidate has been scheduled.
 * Checks if the candidate has been scheduled in a correct availability and
 * Updates the tracker to count this appointment.
 * 
 * @param tracker - The tracker object that handles appointment tracking.
 * @param interval - The interval of an appointment that is being processed.
 * @param appDayIndex - The day index, starting from 0, where the appointment is being processed.
 */
function processAppInterval(tracker: Tracker, interval: string, appDayIndex: number) {
  const intervalIndex = findInterval(interval, tracker.getSchedule()[appDayIndex]);
  if (intervalIndex < 0) { return; }
  tracker.incrementAppsNr(appDayIndex, intervalIndex);
}

/**
 * Finds which interval we have, if it's part of our availabilities.
 * 
 * @param interval - The interval to be checked if it exists in availabilities.
 * @param availabilities - Array of availabilities where the interval should be searched for.
 * @returns The index of the availability where the interval was found or -1 if it isn't found.
 */
function findInterval(interval: string, availabilities: Availability[]): number {
  const intBounds = getIntervalBounds(interval);
  // Check each appointment time if it can be part of each availability,
  // Return availability's index once one is found.
  for (let time of intBounds) {
    for (let i = 0; i < availabilities.length; i++) {
      if (availabilities[i].canHaveTime(time)) { return i; }
    }
  }
  return -1;
}

/**
 * Separates an interval in its start time and end time.
 * 
 * @param interval - The interval of an availability or appointment;
 * Expected interval format: "xx:xx-yy:yy" or "xx:xx".
 * @returns The start time and end time of the interval as a string array;
 * If interval is in format "xx:xx", the array will only contain 1 element, the parameter.
 */
function getIntervalBounds(interval: string): string[] {
  // Remove any white space from interval.
  interval.replace(' ', "");
  // Interval times are separated by a dash.
  return interval.split("-");
}

/**
 * Describes when a person has been scheduled
 * by date and interval.
 */
class Tracker {

  // The sheet where the tracker is located.
  private sheet: ExcelScript.Worksheet;
  // Array containing the tracked language and variants of its naming.
  private language: Array<string> = ["Engleza", "English", "Englsih"];
  // Array containing the tracked dates.
  private days: Array<Date> = new Array<Date>(2);
  // Bidimensional array containing the availabilities when candidates can be scheduled.
  private schedule: Availability[][] = [
    [null, null],
    [null, null]
  ];

  constructor(sheet: ExcelScript.Worksheet) {
    this.sheet = sheet;
  }

  public init() {
    this.initSchedule();
    this.setupNextDays();
  }

  /**
   * Initializes the schedule,
   * the avilabilities when candidates can be scheduled.
   */
  private initSchedule() {
    // iterate through columns in excel, the days of the availabilities.
    for (let col = 0; col < this.schedule.length; col++) {
      // iterate through rows in excel, the intervals of the availabilities.
      for (let row = 0; row < this.schedule[0].length; row++) {
        // The interval in excel.
        // Cells B8, B9, D8, D9.
        const cellRow = 7 + row;
        const cellCol = 1 + col * 2;
        const interval = this.sheet
          .getCell(cellRow, cellCol)
          .getValue();

        // Availabilities are grouped per column in excel for each day,
        // We are going to group them per row for each day, so reverse col and row;
        // Then create availabilities.
        this.schedule[col][row] = new Availability(interval, 0);
        this.schedule[col][row].init();
      }
    }
  }

  /**
   * Write how many appointments have been tracked in excel.
   */
  public writeAppointmentsNumbers() {
    // iterate through columns in excel, the days of the availabilities.
    for (let col = 0; col < this.schedule.length; col++) {
      // iterate through rows in excel, the intervals of the availabilities.
      for (let row = 0; row < this.schedule[0].length; row++) {
        // Write amount of tracked appointments,
        // In cells C8, C9, E8, E9.
        const cellRow = 7 + row;
        const cellCol = 2 + col * 2;
        this.sheet
          .getCell(cellRow, cellCol)
          .setValue(this.schedule[col][row].numberOfApps());
      }
    }
  }

  /**
   * Increment number of appointments found in schedule for day at appDayIndex in interval at intervalIndex.
   * @param appDayIndex - Index of the tracked appointment day.
   * @param intervalIndex  - Index of the tracked interval.
   */
  public incrementAppsNr(appDayIndex: number, intervalIndex: number) {
    this.schedule[appDayIndex][intervalIndex]
      .increment();
  }


  /**
   * Initialize and write the days we are tracking appointments for in excel.
   */
  private setupNextDays() {
    // The 6th row, same for each cell.
    const row = 5;

    for (let i = 0; i < this.days.length; i++) {
      // The date after (i + 1) days.
      const futureDay = Utils.getDay(i + 1);
      // Write next days in excel.
      this.sheet
        .getCell(row, i * 2 + 1)
        .setValue(futureDay.toLocaleDateString());
      // Update dates array.
      this.days[i] = futureDay;
    }
  }

  /**
   * @returns The availabilities we are tracking appointments for.
   */
  public getSchedule() {
    return this.schedule;
  }

  /**
   * @returns The days we are tracking appointments for.
   */
  public getDays(): Array<Date> {
    return this.days;
  }

  /**
   * @returns The language we are tracking appointments for.
   */
  public getLanguage(): Array<string> {
    return this.language;
  }

}

/**
 * Describes a time interval when candidates can be scheduled for an evaluation,
 * and how many candidates have been scheduled then.
 */
class Availability {

  // How many appointments have been scheduled in this availability.
  private appsNr: number;
  // Interval of the availability, in format "xx:xx-yy:yy".
  private interval: string;
  // The time when the interval starts, in format "xx.xx".
  private lowerBound: number;
  // The time when the interval ends in, format "yy.yy".
  private upperBound: number;

  constructor(interval: string, scheduledNr: number) {
    this.interval = interval;
    this.appsNr = scheduledNr;
  }

  public init() {
    const times = getIntervalBounds(this.interval);

    // If there aren't 2 times after splitting the interval, assign bounds to null and stop.
    if (times.length != 2) {
      this.lowerBound = -1;
      this.upperBound = -1;
      return;
    }

    // Assign availability bounds and subtract/add 0.00001 to account for floating point comparisons.
    this.lowerBound = Utils.getTimeValue(times[0]) - 0.00001;
    this.upperBound = Utils.getTimeValue(times[1]) + 0.00001;
  }

  // Increment how many candidates have been scheduled in this availability.
  public increment() {
    this.appsNr++;
  }

  /**
   * Whether a time is during this availability.
   * 
   * @param time - A time string in format "xx:xx".
   * @returns True if time is inside this object's interval.
   */
  public canHaveTime(time: string): boolean {
    const timeValue = Utils.getTimeValue(time);
    return (timeValue <= this.upperBound) && (timeValue >= this.lowerBound);
  }

  /**
   * @returns The time interval when candidates are scheduled.
   */
  public getInterval(): string {
    return this.interval;
  }

  /**
   * @returns The number of appointments made in this availability
   */
  public numberOfApps(): number {
    return this.appsNr;
  }
}

/**
 * The appointment of the candidate for a language test,
 * based on an excel row in the appointments sheet. 
 * Also abbreviated as "app".
 */
class Appointment {

  // the data from Excel that represents an appointment 
  // Of a candidate for an oral language test.
  private data: string;
  // Date of the appointment.
  private date: Date;
  // Language of the appointment.
  private language: string;
  // Interval of the appointment.
  private interval: string;

  // Whether this is a valid appointment.
  private valid: boolean;
  // Whether there was an error finding an appointment.
  private error: boolean;

  /**
   * Default contructor storing the appointment data.
   * @param data - The excel data that represents an appointment. 
   */
  constructor(data: string) {
    this.data = data;
  }

  /**
   * Initializes variables of this instance based on its data.
   */
  public init() {
    // Datasets to be processed are separated by "#".
    const dataArray = this.data.split("#");

    // If length of the array is 2, the data is an excel error: "#ERROR!".
    if (dataArray.length == 2) {
      this.error = true;
      return;
    }

    // Transform the Date data into a number;
    const dateValue = parseInt(dataArray[0]);
    // If Date data is not a number or is missing, 
    // This is not a valid appointment.
    if (isNaN(dateValue)) {
      this.valid = false;
      return;
    }

    // At this point we can safely assign the state of the appointment.
    this.date = Utils.convertExcelDate(dateValue);
    this.language = dataArray[1];
    this.interval = dataArray[2];
    this.valid = true;
    this.error = false;
  }

  /**
   * @returns The Date of the appointment.
   */
  public getDate(): Date {
    return this.date;
  }

  /**
   * @returns The Language test for which the candidate has been scheduled
   */
  public getLanguage(): string {
    return this.language;
  }

  /**
   * @returns A time interval as a string representing
   * when the candidate has been scheduled.
   */
  public getInterval(): string {
    return this.interval;
  }

  /**
   * @returns Whether this is a valid appointment.
   */
  public isValid(): boolean {
    return this.valid;
  }

  /**
   * @returns Whether we have an error in excel.
   */
  public hasError(): boolean {
    return this.error;
  }

}

class Utils {

  /**
   * Amount of milliseconds in a day.
   */
  public static readonly MILLIS_IN_A_DAY = 24 * 60 * 60 * 1000;
  /**
   * Amount of milliseconds in a minute.
   */
  public static readonly MILLIS_IN_A_MINUTE = 60 * 1000;

  /**
   * Converts a date in excel format to an instance of Date.
   * Excel date - counts days since 01.01.1900.
   * Date class - counts millis since beginning of Unix epoch, 01.01.1970.
   * 
   * @param excelDate - How many days since 01.01.1900.
   * @returns Instance of Date based on excel date.
   */
  public static convertExcelDate(excelDate: number): Date {
    // Days between 01.01.1900
    // And beginning of Unix epoch.
    const offsetDays = 25569;
    const daysSinceUnix = excelDate - offsetDays;
    return new Date(daysSinceUnix * Utils.MILLIS_IN_A_DAY);
  }

  /**
   * If we can find the searchee in a words array.
   * 
   * @param searchee - The string we are searching for.
   * @param words - The string array where we are searching for searchee.
   * @returns True if searchee is found in words.
   */
  public static canFindString(searchee: string, words: Array<string>): boolean {
    let isFound = false;
    for (let word of words) {
      if (searchee.localeCompare(word) == 0) {
        isFound = true;
        break;
      }
    }
    return isFound;
  }
  
  /**
   * Finds a date in a date array.
   * 
   * @param searchee - Date to look for in dates.
   * @param dates - Array of dates in which to search for searchee.
   * @returns Searched date's index in dates array or -1 if it isn't found.
   */
  public static findDate(searchee: Date, dates: Array<Date>): number {
    for (let i = 0; i < dates.length; i++) {
      if (Utils.isSameDate(searchee, dates[i])) {
        return i;
      }
    }
    return -1;
  }

  /**
   * Checks whether a date is before a compared date,
   * The check is inclusive when both parameters are the same dates. 
   * 
   * @param date - The date we are checking if it's before a date.
   * @param comparedDate - The date we are checking against.
   * @returns True if the checked date is before the compared date or they are the same day.
   */
  public static isBeforeDate(date: Date, comparedDate: Date): boolean {
    return Utils.getDateValue(date) <= Utils.getDateValue(comparedDate);
  }
  
  /**
   * Checks whether 2 dates are the same date.
   * 
   * @param date - First date we are checking. 
   * @param comparedDate - Second date we are checking.
   * @returns True if the 2 dates are the same day.
   */
  public static isSameDate(date: Date, comparedDate: Date): boolean {
    return Utils.getDateValue(date) == Utils.getDateValue(comparedDate);
  }

  /** 
   * Standardizes dates as numeric values.
   * To be used for comparisons.
   * 
   * @param date - The date we are standardizing.
   * @returns The number of days since beginning of the Unix epoc until date parameter.
   */
  public static getDateValue(date: Date): number {
    return Math.trunc(date.getTime() / Utils.MILLIS_IN_A_DAY);;
  }

  /**
   * Gets a day in the future or in the past,
   * By a number of days from today.
   * 
   * @param daysNrDiff - How many days to count from today.
   * @returns The date that is daysNrDiff days away from today.
   */
  public static getDay(daysNrDiff: number): Date {
    // Get current time.
    const date = new Date(Date.now());

    // If daysNrDiff is positive, we are going to add,
    // Otherwise, we'll subtract.
    const increment = Math.sign(daysNrDiff);
    while (daysNrDiff != 0) {
      // Increment our current date.
      date.setDate(date.getDate() + increment);
      // Update condition only if we have a week day.
      if (date.getDay() % 6) {
        daysNrDiff -= increment;
      }
    }
    return date;
  }

  /**
   * Standardizes time strings as float values to be used for comparisons.
   * 
   * @param time - A time in format "xx:xx".
   * @returns The time in format "xx.xx"
   */
  public static getTimeValue(time: string): number {
    const times = time.split(":");
    return Number.parseInt(times[0]) + Number.parseInt(times[1]) / 100;
  }

}