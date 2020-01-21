/*

GLOBAL SETTINGS

*/

const ARC_APP_DEBUG_MODE: boolean = true;

/*

SERVER TYPES (COPIED FROM CLIENT)

*/

type ServerResponse<T> = {
  error: boolean;
  val: T;
  message: string;
};

enum AskStatus {
  LOADING = 'LOADING',
  LOADED = 'LOADED',
  ERROR = 'ERROR'
}

type Ask<T> = AskLoading | AskFinished<T>;

type AskLoading = { status: AskStatus.LOADING };
type AskFinished<T> = AskLoaded<T> | AskError;

type AskError = {
  status: AskStatus.ERROR;
  message: string;
};

type AskLoaded<T> = {
  status: AskStatus.LOADED;
  val: T;
};

/*

UTILITIES

*/

function roundDownToDay(utcTime: number) {
  const tzOffset = new Date(utcTime).getTimezoneOffset() * 60 * 1000;
  return utcTime - (utcTime % (24 * 60 * 60 * 1000)) + tzOffset;
}

function formatAttendanceModDataString(mod: number, minutes: number) {
  return `${Number(mod)} ${Number(minutes)}`;
}

function onlyKeepUnique<T>(arr: T[]): T[] {
  const x = {};
  for (let i = 0; i < arr.length; ++i) {
    x[JSON.stringify(arr[i])] = arr[i];
  }
  const result = [];
  for (const key of Object.getOwnPropertyNames(x)) {
    result.push(x[key]);
  }
  return result;
}

// polyfill of the typical Object.values()
function Object_values<T>(o: ObjectMap<T>): T[] {
  const result: T[] = [];
  for (const i of Object.getOwnPropertyNames(o)) {
    result.push(o[i]);
  }
  return result;
}

function recordCollectionToArray(r: RecCollection): Rec[] {
  const x = [];
  for (const i of Object.getOwnPropertyNames(r)) {
    x.push(r[i]);
  }
  return x;
}

// This function converts mod numbers (ie. 11) into A-B-day strings (ie. 1B).
function stringifyMod(mod: number) {
  if (1 <= mod && mod <= 10) {
    return String(mod) + 'A';
  } else if (11 <= mod && mod <= 20) {
    return String(mod - 10) + 'B';
  }
  throw new Error(`mod ${mod} isn't serializable`);
}

function stringifyError(error: any): string {
  if (error instanceof Error) {
    return JSON.stringify(error, Object.getOwnPropertyNames(error));
  }
  try {
    return JSON.stringify(error);
  } catch (unusedError) {
    return String(error);
  }
}

type ObjectMap<T> = {
  [key: string]: T;
};

abstract class Field {
  name: string;
  abstract parse(x: any): any;
  abstract serialize(x: any): any;

  constructor(name: string = null) {
    this.name = name;
  }
}

class BooleanField extends Field {
  constructor(name: string = null) {
    super(name);
  }
  parse(x: any) {
    return x === 'true' || x === true ? true : false;
  }
  serialize(x: any) {
    return x ? true : false;
  }
}

class NumberField extends Field {
  constructor(name: string = null) {
    super(name);
  }
  parse(x: any) {
    return Number(x);
  }
  serialize(x: any) {
    return Number(x);
  }
}

class StringField extends Field {
  constructor(name: string = null) {
    super(name);
  }
  parse(x: any): string {
    return String(x);
  }
  serialize(x: any): string {
    return String(x);
  }
}

// Dates are treated as numbers.
class DateField extends Field {
  constructor(name: string = null) {
    super(name);
  }
  parse(x: any) {
    if (x === '' || x === -1) {
      return -1;
    } else {
      return Number(x);
    }
  }
  serialize(x: any) {
    if (x === -1 || x === '') {
      return '';
    } else {
      return new Date(x);
    }
  }
}

class JsonField extends Field {
  constructor(name: string = null) {
    super(name);
  }
  parse(x: any) {
    return JSON.parse(x);
  }
  serialize(x: any) {
    return JSON.stringify(x);
  }
}

interface TableInfo {
  sheetName: string;
  fields: Field[];
  isForm?: boolean;
}

// Short for Record. Avoids naming collision.
type Rec = {
  [field: string]: any;
  id: number;
  date: number;
};
type RecCollection = {
  [id: string]: Rec;
};

/*

TABLE CLASS

*/

class Table {
  idCounter: number = new Date().getTime();

  static makeBasicStudentConfig() {
    return [
      new StringField('friendlyFullName'),
      new StringField('friendlyName'),
      new StringField('firstName'),
      new StringField('lastName'),
      new NumberField('grade'),
      new NumberField('studentId'),
      new StringField('email'),
      new StringField('phone'),
      new StringField('contactPref'),
      new StringField('homeroom'),
      new StringField('homeroomTeacher'),
      new StringField('attendanceAnnotation')
    ];
  }

  static tableInfoAll: ObjectMap<TableInfo> = {
    tutors: {
      sheetName: '$tutors',
      fields: [
        new NumberField('id'),
        new DateField('date'),
        ...Table.makeBasicStudentConfig(),
        new JsonField('mods'),
        new JsonField('modsPref'),
        new StringField('subjectList'),
        new JsonField('attendance'),
        new JsonField('dropInMods'),
        new StringField('afterSchoolAvailability'),
        new NumberField('additionalHours')
      ]
    },
    learners: {
      sheetName: '$learners',
      fields: [
        new NumberField('id'),
        new DateField('date'),
        ...Table.makeBasicStudentConfig(),
        new JsonField('attendance')
      ]
    },
    requests: {
      sheetName: '$requests',
      fields: [
        new NumberField('id'),
        new DateField('date'),
        new NumberField('learner'),
        new JsonField('mods'),
        new StringField('subject'),
        new BooleanField('isSpecial'),
        new StringField('annotation'),
        new NumberField('step'),
        new JsonField('chosenBookings')
      ]
    },
    requestSubmissions: {
      sheetName: '$request-submissions',
      fields: [
        new NumberField('id'),
        new DateField('date'),
        ...Table.makeBasicStudentConfig(),
        new JsonField('mods'),
        new StringField('subject'),
        new BooleanField('isSpecial'),
        new StringField('annotation'),
        new StringField('status')
      ]
    },
    bookings: {
      sheetName: '$bookings',
      fields: [
        new NumberField('id'),
        new DateField('date'),
        new NumberField('request'),
        new NumberField('tutor'),
        new NumberField('mod'),
        new StringField('status')
      ]
    },
    matchings: {
      sheetName: '$matchings',
      fields: [
        new NumberField('id'),
        new DateField('date'),
        new NumberField('learner'),
        new NumberField('tutor'),
        new StringField('subject'),
        new NumberField('mod'),
        new StringField('annotation')
      ]
    },
    requestForm: {
      sheetName: '$request-form',

      // this means that the ID field is automatically generated from the date field
      isForm: true,

      fields: [
        // THE ORDER OF THE FIELDS MATTERS! They must match the order of the form's questions.
        new DateField('date'),
        new StringField('firstName'),
        new StringField('lastName'),
        new StringField('friendlyFullName'),
        new NumberField('studentId'),
        new StringField('grade'),
        new StringField('subject'),
        new StringField('modDataA1To5'),
        new StringField('modDataB1To5'),
        new StringField('modDataA6To10'),
        new StringField('modDataB6To10'),
        new StringField('homeroom'),
        new StringField('homeroomTeacher'),
        new StringField('email'),
        new StringField('phone'),
        new StringField('contactPref'),
        new StringField('iceCreamQuestion')
      ]
    },
    specialRequestForm: {
      sheetName: '$special-request-form',
      isForm: true,
      fields: [
        new DateField('date'),
        new StringField('teacherName'),
        new StringField('teacherEmail'),
        new StringField('numLearners'),
        new StringField('subject'),
        new StringField('tutoringDateInformation'),
        new NumberField('room'),
        new StringField('abDay'),
        new NumberField('mod1To10'),
        new StringField('studentNames'),
        new StringField('additionalInformation')
      ]
    },
    attendanceForm: {
      sheetName: '$attendance-form',
      isForm: true,
      fields: [
        new DateField('date'),
        new DateField('dateOfAttendance'), // optional in the form
        new NumberField('mod1To10'),
        new NumberField('studentId'),
        new StringField('presence')
      ]
    },
    tutorRegistrationForm: {
      sheetName: '$tutor-registration-form',
      isForm: true,
      fields: [
        new DateField('date'),
        new StringField('firstName'),
        new StringField('lastName'),
        new StringField('friendlyName'),
        new StringField('friendlyFullName'),
        new NumberField('studentId'),
        new StringField('grade'),
        new StringField('email'),
        new StringField('phone'),
        new StringField('contactPref'),
        new StringField('homeroom'),
        new StringField('homeroomTeacher'),
        new StringField('afterSchoolAvailability'),
        new StringField('modDataA1To5'),
        new StringField('modDataB1To5'),
        new StringField('modDataA6To10'),
        new StringField('modDataB6To10'),
        new StringField('modDataPrefA1To5'),
        new StringField('modDataPrefB1To5'),
        new StringField('modDataPrefA6To10'),
        new StringField('modDataPrefB6To10'),
        new StringField('subjects0'),
        new StringField('subjects1'),
        new StringField('subjects2'),
        new StringField('subjects3'),
        new StringField('subjects4'),
        new StringField('subjects5'),
        new StringField('subjects6'),
        new StringField('subjects7'),
        new StringField('iceCreamQuestion'),
        new NumberField('numberGuessQuestion')
      ]
    },
    attendanceLog: {
      // this table is merged into the JSON of tutor.fields.attendance
      // the table will get quite large, so we will hand-archive it from time to time
      // ASSUMPTION: (thus...) the table DOESN'T contain all of the attendance data. Some of it
      // will be archived somewhere else. The JSON will be merged with the attendance log.
      sheetName: '$attendance-log',
      fields: [
        new NumberField('id'),
        new DateField('date'),
        new DateField('dateOfAttendance'), // rounded to nearest day
        new StringField('validity'), // filled with an error message if the form entry was typed wrong
        new NumberField('mod'),

        // one of these ID fields will be left as -1 (blank).
        new NumberField('tutor'),
        new NumberField('learner'),
        new NumberField('minutesForTutor'),
        new NumberField('minutesForLearner'),
        new StringField('presenceForTutor'),
        new StringField('presenceForLearner'),

        // used for reset purposes
        new StringField('markForReset')
      ]
    },
    attendanceDays: {
      sheetName: '$attendance-days',
      fields: [
        new NumberField('id'),
        new DateField('date'),
        new DateField('dateOfAttendance'),
        new StringField('abDay'),

        // we add a functionality to reset a day's attendance absences
        new StringField('status') // upcoming, finished, finalized, unreset, reset
      ]
    },
    operationLog: {
      sheetName: '$operation-log',
      fields: [
        new NumberField('id'),
        new DateField('date'),
        new JsonField('args')
      ]
    }
  };

  name: string;
  tableInfo: TableInfo;
  sheet: GoogleAppsScript.Spreadsheet.Sheet;
  isForm: boolean;
  sheetLastColumn: number;

  constructor(name: string) {
    this.name = name;
    this.tableInfo = Table.tableInfoAll[name];

    // Forms have important limitations in the app. We want to protect form data at all costs.
    this.isForm = !!this.tableInfo.isForm;

    if (this.tableInfo === undefined) {
      throw new Error(`table ${name} not found in tableInfoAll`);
    }
    this.rebuildSheetIfNeeded();
  }

  rebuildSheetIfNeeded() {
    this.sheet = SpreadsheetApp.getActive().getSheetByName(
      this.tableInfo.sheetName
    );
    if (this.sheet === null) {
      if (this.isForm) {
        throw new Error(
          `table ${this.name} not found, and it's supposed to be a form`
        );
      } else {
        this.sheet = SpreadsheetApp.getActive().insertSheet(
          this.tableInfo.sheetName
        );
        this.rebuildSheetHeadersIfNeeded();
      }
    } else {
      this.sheetLastColumn = this.sheet.getLastColumn();
    }
  }

  rebuildSheetHeadersIfNeeded() {
    const col = this.sheet.getLastColumn();
    if (!this.isForm) {
      this.sheet.getRange(1, 1, 1, col === 0 ? 1 : col).clearContent();
      this.sheet
        .getRange(1, 1, 1, this.tableInfo.fields.length)
        .setValues([this.tableInfo.fields.map(field => field.name)]);
    }
    this.sheetLastColumn = col;
  }

  resetEntireSheet() {
    if (!this.isForm) {
      this.sheet.getDataRange().clearContent();
    }
    this.rebuildSheetHeadersIfNeeded();
  }

  // This is useful for debug. It rewrites each cell with the content that the app *thinks* is inside the cell.
  rewriteEntireSheet() {
    if (!this.isForm) {
      this.updateAllRecords(Object_values(this.retrieveAllRecords()));
    }
    this.rebuildSheetHeadersIfNeeded();
  }

  retrieveAllRecords(): RecCollection {
    if (ARC_APP_DEBUG_MODE) {
      this.rebuildSheetHeadersIfNeeded();
    }
    if (this.sheetLastColumn !== this.tableInfo.fields.length) {
      throw new Error(
        `something's wrong with the columns of table ${this.name} (${this.tableInfo.fields.length})`
      );
    }
    const raw = this.sheet.getDataRange().getValues();
    const res = {};
    for (let i = 1; i < raw.length; ++i) {
      const rec = this.parseRecord(raw[i]);
      res[String(rec.id)] = rec;
    }
    return res;
  }

  parseRecord(raw: any[]): Rec {
    const rec: { [key: string]: any } = {};
    for (let i = 0; i < this.tableInfo.fields.length; ++i) {
      const field = this.tableInfo.fields[i];
      // this accounts for blanks in the last field
      rec[field.name] = field.parse(raw[i] === undefined ? '' : raw[i]);
    }
    if (this.isForm) {
      // forms don't have an id field, so copy the date into the id
      rec.id = rec.date;
    }
    return rec as Rec;
  }

  // Turn a JS object into an array of raw cell data, using each field object's built-in serializer.
  serializeRecord(record: Rec): any[] {
    return this.tableInfo.fields.map(field =>
      field.serialize(record[field.name])
    );
  }

  createRecord(record: Rec): Rec {
    if (record.date === -1) {
      record.date = Date.now();
    }
    // REMEMBER: forms don't have id fields!
    if (!this.isForm && record.id === -1) {
      record.id = this.createNewId();
    }
    this.sheet.appendRow(this.serializeRecord(record));
    return record;
  }

  getRowById(id: number): number {
    if (this.isForm) {
      throw new Error('getRowById not supported for forms');
    }

    // because the first row is headers, we ignore it and start from the second row
    const mat: any[][] = this.sheet
      .getRange(2, 1, this.sheet.getLastRow() - 1)
      .getValues();
    let rowNum = -1;
    for (let i = 0; i < mat.length; ++i) {
      const cell: number = mat[i][0];
      if (typeof cell !== 'number') {
        throw new Error(
          `id at location ${String(i)} is not a number in table ${String(
            this.name
          )}`
        );
      }
      if (cell === id) {
        if (rowNum !== -1) {
          throw new Error(
            `duplicate ID ${String(id)} in table ${String(this.name)}`
          );
        }
        rowNum = i + 2; // i = 0 <=> second row (rows are 1-indexed)
      }
    }
    if (rowNum == -1) {
      throw new Error(
        `ID ${String(id)} not found in table ${String(this.name)}`
      );
    }
    return rowNum;
  }

  updateRecord(editedRecord: Rec, rowNum?: number): void {
    if (rowNum === undefined) {
      rowNum = this.getRowById(editedRecord.id);
    }
    this.sheet
      .getRange(rowNum, 1, 1, this.sheetLastColumn)
      .setValues([this.serializeRecord(editedRecord)]);
  }

  updateAllRecords(editedRecords: Rec[]): void {
    if (this.sheet.getLastRow() === 1) {
      return; // the sheet is empty, and trying to select it will result in an error
    }
    // because the first row is headers, we ignore it and start from the second row
    const mat: any[][] = this.sheet
      .getRange(2, 1, this.sheet.getLastRow() - 1)
      .getValues();
    let idRowMap: ObjectMap<number> = {};
    for (let i = 0; i < mat.length; ++i) {
      idRowMap[String(mat[i][0])] = i + 2; // i = 0 <=> second row (rows are 1-indexed)
    }
    for (const r of editedRecords) {
      this.updateRecord(r, idRowMap[String(r.id)]);
    }
  }

  deleteRecord(id: number): void {
    if (this.isForm) {
      throw new Error('cannot delete records from a form');
    }
    this.sheet.deleteRow(this.getRowById(id));
  }

  createNewId(): number {
    if (this.isForm) {
      // There's no point! Forms don't have IDs, because the Date field is treated like an ID!
      throw new Error('ID generation not supported for forms');
    }
    const time = new Date().getTime();
    if (time <= this.idCounter) {
      ++this.idCounter;
    } else {
      this.idCounter = time;
    }
    return this.idCounter;
  }

  // This is auto-called by the website (client) which is also known as the frontend.
  processClientAsk(args: any[]): any {
    if (this.isForm) {
      throw new Error('cannot access form tables from the client API');
    }
    if (args[0] === 'retrieveAll') {
      return this.retrieveAllRecords();
    }
    if (args[0] === 'update') {
      this.updateRecord(args[1]);
      onClientNotification(['update', this.name, args[1]]);
      return null;
    }
    if (args[0] === 'create') {
      const newRecord = this.createRecord(args[1]);
      onClientNotification(['create', this.name, newRecord]);
      return newRecord;
    }
    if (args[0] === 'delete') {
      this.deleteRecord(args[1]);
      onClientNotification(['delete', this.name, args[1]]);
      return null;
    }
    if (args[0] === 'debugheaders') {
      this.rebuildSheetHeadersIfNeeded();
      return null;
    }
    if (args[0] === 'debugreseteverything') {
      this.resetEntireSheet();
      return null;
    }
    throw new Error('args not matched');
  }
}

// This is a way to prevent constructing six table objects every time the script is called.
type TableMap = {
  [name: string]: () => Table;
};

// The point of this whole thing is so we don't read 10+ tables each time the script is run.
// Tables are only loaded once you write the magic words: tableMap.NAMEOFTABLE().
const tableMap: TableMap = {
  ...tableMapBuild('tutors'),
  ...tableMapBuild('learners'),
  ...tableMapBuild('requests'),
  ...tableMapBuild('requestSubmissions'),
  ...tableMapBuild('bookings'),
  ...tableMapBuild('matchings'),
  ...tableMapBuild('requestForm'),
  ...tableMapBuild('specialRequestForm'),
  ...tableMapBuild('attendanceForm'),
  ...tableMapBuild('attendanceLog'),
  ...tableMapBuild('operationLog'),
  ...tableMapBuild('attendanceDays'),
  ...tableMapBuild('tutorRegistrationForm')
};

// This fancy code makes it so the table isn't loaded twice if you call tableMap.NAMEOFTABLE() twice.
function tableMapBuild(name: string) {
  return {
    [name]: (() => {
      let table: Table = null;
      return () => {
        if (table === null) {
          return new Table(name);
        } else {
          return table;
        }
      };
    })()
  };
}

/*

IMPORTANT EVENT HANDLERS
(CODE THAT DOES ALL THE USER ACTIONS NECESSARY IN THE BACKEND)
(ALSO CODE THAT HANDLES SERVER-CLIENT INTERACTIONS)

*/

function doGet() {
  // The operation log is reset every time the website is opened.
  tableMap.operationLog().resetEntireSheet();

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ARC App')
    .setFaviconUrl(
      'https://arc-app-frontend-server.netlify.com/dist/favicon.ico'
    );
}

function processClientAsk(args: any[]): any {
  const resourceName: string = args[0];
  if (resourceName === undefined) {
    throw new Error('no args, or must specify resource name');
  }
  if (resourceName === 'command') {
    if (args[1] === 'syncDataFromForms') {
      return onSyncForms();
    }
    if (args[1] === 'recalculateAttendance') {
      return onRecalculateAttendance();
    }
    if (args[1] === 'generateSchedule') {
      return onGenerateSchedule();
    }
    if (args[1] === 'retrieveMultiple') {
      return onRetrieveMultiple(args[2] as string[]);
    }
    throw new Error('unknown command');
  }
  if (tableMap[resourceName] === undefined) {
    throw new Error(`resource ${String(resourceName)} not found`);
  }
  const resource = tableMap[resourceName]();
  return resource.processClientAsk(args.slice(1));
}

// this is the MAIN ENTRYPOINT that the client uses to ask the server for data.
function onClientAsk(args: any[]): string {
  let returnValue = {
    error: true,
    val: null,
    message: 'Mysterious error'
  };
  try {
    returnValue = {
      error: false,
      val: processClientAsk(args),
      message: null
    };
  } catch (err) {
    returnValue = {
      error: true,
      val: null,
      message: stringifyError(err)
    };
  }
  // If you send a too-big object, Google Apps Script doesn't let you do it, and null is returned. But if you stringify it, you're fine.
  return JSON.stringify(returnValue);
}

function onClientNotification(args: any[]): void {
  // this is a way to record the logs
  // originally, this was supposed to have the client read them every 20 seconds
  // so they know the things that other clients have done
  // in the case that multiple clients are open at once
  // but that plan was deemed unnecessary

  tableMap.operationLog().createRecord({
    id: -1,
    date: -1,
    args
  });
}

function debugClientApiTest() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt('Enter args as JSON array');
    ui.alert(
      JSON.stringify(onClientAsk(JSON.parse(response.getResponseText())))
    );
  } catch (err) {
    Logger.log(stringifyError(err));
    throw err;
  }
}

// This is a useful debug. It rewrites all the sheet headers to what the app thinks the sheet headers "should" be.
function debugHeaders() {
  try {
    for (const name of Object.getOwnPropertyNames(tableMap)) {
      tableMap[name]().rebuildSheetHeadersIfNeeded();
    }
  } catch (err) {
    Logger.log(stringifyError(err));
    throw err;
  }
}

// Resets every table with length < 5. ONLY FOR DEMO PURPOSES. DO NOT RUN IN PRODUCTION.
function debugResetAllSmallTables() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      'Leave the box below blank to cancel debug operation.'
    );
    if (response.getResponseText() === 'DEBUG_SMALL_RESET') {
      for (const name of Object.getOwnPropertyNames(tableMap)) {
        const table = tableMap[name]();
        if (Object.getOwnPropertyNames(table.retrieveAllRecords()).length < 5) {
          table.resetEntireSheet();
        }
      }
    }
  } catch (err) {
    Logger.log(stringifyError(err));
    throw err;
  }
}

// Completely wipes every single table in the database (except forms).
function debugResetEverything() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      'Leave the box below blank to cancel debug operation.'
    );
    if (response.getResponseText() === 'DEBUG_RESET') {
      for (const name of Object.getOwnPropertyNames(tableMap)) {
        tableMap[name]().resetEntireSheet();
      }
    }
  } catch (err) {
    Logger.log(stringifyError(err));
    throw err;
  }
}

function debugRewriteEverything() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      'Leave the box below blank to cancel debug operation.'
    );
    if (response.getResponseText() === 'DEBUG_REWRITE') {
      for (const name of Object.getOwnPropertyNames(tableMap)) {
        tableMap[name]().rewriteEntireSheet();
      }
    }
  } catch (err) {
    Logger.log(stringifyError(err));
    throw err;
  }
}

// This is a utility designed for onSyncForms().
// Syncs between the formTable and the actualTable that we want to associate with it.
// Basically, we use formRecordToActualRecord() to convert form records to actual records.
// Then the actual records go in the actualTable.
// But we only do this for form records that have dates that don't exist as IDs in actualTable.
// (Remember that a form date === a record ID.)
// There is NO DELETING RECORDS! No matter what!
function doFormSync(
  formTable: Table,
  actualTable: Table,
  formRecordToActualRecord: (formRecord: Rec) => Rec
): number {
  const actualRecords = actualTable.retrieveAllRecords();
  const formRecords = formTable.retrieveAllRecords();

  let numOfThingsSynced = 0;

  // create an index of actualdata >> date.
  // Then iterate over all formdata and find the ones that are missing from the index.
  const index: { [date: string]: Rec } = {};
  for (const idKey of Object.getOwnPropertyNames(actualRecords)) {
    const record = actualRecords[idKey];
    const dateIndexKey = String(record.date);
    index[dateIndexKey] = record;
  }
  for (const idKey of Object.getOwnPropertyNames(formRecords)) {
    const record = formRecords[idKey];
    const dateIndexKey = String(record.date);
    if (index[dateIndexKey] === undefined) {
      actualTable.createRecord(formRecordToActualRecord(record));
      ++numOfThingsSynced;
    }
  }

  return numOfThingsSynced;
}

const MINUTES_PER_MOD = 38;

function onRetrieveMultiple(resourceNames: string[]) {
  const result = {};
  for (const resourceName of resourceNames) {
    result[resourceName] = tableMap[resourceName]().retrieveAllRecords();
  }
  return result;
}
function uiSyncForms() {
  try {
    const result = onSyncForms();
    SpreadsheetApp.getUi().alert(
      `Finished sync! ${result as number} new form submits found.`
    );
  } catch (err) {
    Logger.log(stringifyError(err));
    throw err;
  }
}

// The main "sync forms" function that's crammed with form data formatting.
function onSyncForms(): number {
  // tables
  const tutors = tableMap.tutors().retrieveAllRecords();
  const matchings = tableMap.matchings().retrieveAllRecords();
  const attendanceDays = tableMap.attendanceDays().retrieveAllRecords();
  const attendanceDaysIndex = {};
  for (const day of Object_values(attendanceDays)) {
    attendanceDaysIndex[day.dateOfAttendance] = day.abDay;
  }

  // parsing contact preferences
  function parseContactPref(s: string) {
    if (s === 'Phone') return 'phone';
    if (s === 'Email') return 'email';
    return 'either';
  }

  // parsing grade
  function parseGrade(g: string) {
    if (g === 'Freshman') return 9;
    if (g === 'Sophomore') return 10;
    if (g === 'Junior') return 11;
    if (g === 'Senior') return 12;
    return 0;
  }

  function parseModInfo(abDay: string, mod1To10: number): number {
    if (abDay.toLowerCase().charAt(0) === 'a') {
      return mod1To10;
    }
    if (abDay.toLowerCase().charAt(0) === 'b') {
      return mod1To10 + 10;
    }
    throw new Error(`${String(abDay)} does not start with A or B`);
  }

  function parseModData(modData: string[]): number[] {
    function doParse(d: string) {
      return d
        .split(',')
        .map(x => x.trim())
        .filter(x => x !== '' && x !== 'None')
        .map(x => parseInt(x));
    }
    const mA15: number[] = doParse(modData[0]);
    const mB15: number[] = doParse(modData[1]).map(x => x + 10);
    const mA60: number[] = doParse(modData[2]);
    const mB60: number[] = doParse(modData[3]).map(x => x + 10);
    return mA15
      .concat(mA60)
      .concat(mB15)
      .concat(mB60);
  }

  function parseStudentConfig(r: Rec): ObjectMap<any> {
    return {
      firstName: r.firstName,
      lastName: r.lastName,
      friendlyName: r.friendlyName ? r.friendlyName : r.firstName,
      friendlyFullName: r.friendlyFullName
        ? r.friendlyFullName
        : r.firstName + ' ' + r.lastName,
      grade: parseGrade(r.grade),
      studentId: r.studentId,
      email: r.email,
      phone: r.phone,
      contactPref: parseContactPref(r.contactPref),
      homeroom: r.homeroom,
      homeroomTeacher: r.homeroomTeacher,
      attendanceAnnotation: ''
    };
  }

  function processRequestFormRecord(r: Rec): Rec {
    return {
      id: -1,
      date: r.date, // the date MUST be the date from the form; this is used for syncing
      subject: r.subject,
      mods: parseModData([
        r.modDataA1To5,
        r.modDataB1To5,
        r.modDataA6To10,
        r.modDataB6To10
      ]),
      specialRoom: '',
      ...parseStudentConfig(r),
      status: 'unchecked',
      homeroom: r.homeroom,
      homeroomTeacher: r.homeroomTeacher,
      chosenBookings: [],
      isSpecial: false,
      annotation: ''
    };
  }
  function processTutorRegistrationFormRecord(r: Rec): Rec {
    function parseSubjectList(d: string[]) {
      return d
        .join(',')
        .split(',') // remember that within each string there are commas
        .map(x => x.trim())
        .filter(x => x !== '' && x !== 'None')
        .map(x => String(x))
        .join(', ');
    }
    return {
      id: -1,
      date: r.date,
      ...parseStudentConfig(r),
      mods: parseModData([
        r.modDataA1To5,
        r.modDataB1To5,
        r.modDataA6To10,
        r.modDataB6To10
      ]),
      modsPref: parseModData([
        r.modDataPrefA1To5,
        r.modDataPrefB1To5,
        r.modDataPrefA6To10,
        r.modDataPrefB6To10
      ]),
      subjectList: parseSubjectList([
        r.subjects0,
        r.subjects1,
        r.subjects2,
        r.subjects3,
        r.subjects4,
        r.subjects5,
        r.subjects6,
        r.subjects7
      ]),
      attendance: {},
      dropInMods: [],
      afterSchoolAvailability: r.afterSchoolAvailability,
      attendanceAnnotation: '',
      additionalHours: 0
    };
  }
  function processSpecialRequestFormRecord(r: Rec): Rec {
    let annotation = '';
    annotation += `TEACHER EMAIL: ${r.teacherEmail}; `;
    annotation += `# LEARNERS: ${r.numLearners}; `;
    annotation += `TUTORING DATE INFO: ${r.tutoringDateInformation}; `;
    if (r.studentNames.trim() !== '') {
      annotation += `STUDENT NAMES: ${r.studentNames}; `;
    }
    if (r.additionalInformation.trim() !== '') {
      annotation += `ADDITIONAL INFO: ${r.additionalInformation}; `;
    }
    return {
      id: -1,
      date: r.date, // the date MUST be the date from the form
      friendlyFullName: '[special request]',
      friendlyName: '[special request]',
      firstName: '[special request]',
      lastName: '[special request]',
      grade: -1,
      studentId: -1,
      email: '[special request]',
      phone: '[special request]',
      contactPref: 'either',
      homeroom: '[special request]',
      homeroomTeacher: '[special request]',
      attendanceAnnotation: '[special request]',
      mods: [parseModInfo(r.abDay, r.mod1To10)],
      subject: r.subject,
      isSpecial: true,
      annotation,
      status: 'unchecked'
    };
  }
  function processAttendanceFormRecord(r: Rec): Rec {
    let tutor = -1;
    let learner = -1;
    let validity = '';
    let minutesForTutor = -1;
    let minutesForLearner = -1;
    let presenceForTutor = '';
    let presenceForLearner = '';
    let mod = -1;
    const processedDateOfAttendance = roundDownToDay(
      r.dateOfAttendance === -1 ? r.date : r.dateOfAttendance
    );

    if (attendanceDaysIndex[processedDateOfAttendance] === undefined) {
      validity = 'date does not exist in attendance days index';
    } else {
      mod = parseModInfo(
        attendanceDaysIndex[processedDateOfAttendance],
        r.mod1To10
      );
    }

    // give the tutor their time
    minutesForTutor = MINUTES_PER_MOD;
    presenceForTutor = 'P';

    // figure out who the tutor is, by student ID
    const xTutors = recordCollectionToArray(tutors).filter(
      x => x.studentId === r.studentId
    );
    if (xTutors.length === 0) {
      validity = 'tutor student ID does not exist';
    } else if (xTutors.length === 1) {
      tutor = xTutors[0].id;
    } else {
      throw new Error(`duplicate tutor student id ${String(r.studentId)}`);
    }

    // does the tutor have a learner?
    const xMatchings = recordCollectionToArray(matchings).filter(
      x => x.tutor === tutor && x.mod === mod
    );
    if (xMatchings.length === 0) {
      learner = -1;
    } else if (xMatchings.length === 1) {
      learner = xMatchings[0].learner;
    } else {
      throw new Error(
        `the tutor ${xMatchings[0].friendlyFullName} is matched twice on the same mod`
      );
    }

    // ATTENDANCE LOGIC
    if (validity === '') {
      if (r.presence === 'Yes') {
        minutesForLearner = MINUTES_PER_MOD;
        presenceForLearner = 'P';
      } else if (r.presence === 'No') {
        minutesForLearner = 0;
        presenceForLearner = 'A';
      } else if (r.presence === `I don't have a learner assigned`) {
        if (learner === -1) {
          minutesForLearner = -1;
        } else {
          // so there really is a learner...
          validity = `tutor said they don't have a learner assigned, but they do!`;
        }
      } else if (
        r.presence === `No, but the learner doesn't need any tutoring today`
      ) {
        if (learner === -1) {
          validity = `tutor said they have a learner assigned, but they don't!`;
        } else {
          minutesForLearner = 1; // TODO: this is a hacky solution; fix it
          presenceForLearner = 'E';
        }
      } else {
        throw new Error(`invalid presence (${String(r.presence)})`);
      }
    }

    return {
      id: r.id,
      date: r.date,
      dateOfAttendance: processedDateOfAttendance,
      validity,
      mod,
      tutor,
      learner,
      minutesForTutor,
      minutesForLearner,
      presenceForTutor,
      presenceForLearner,
      markForReset: ''
    };
  }

  let numOfThingsSynced = 0;
  numOfThingsSynced += doFormSync(
    tableMap.requestForm(),
    tableMap.requestSubmissions(),
    processRequestFormRecord
  );
  numOfThingsSynced += doFormSync(
    tableMap.specialRequestForm(),
    tableMap.requestSubmissions(),
    processSpecialRequestFormRecord
  );
  numOfThingsSynced += doFormSync(
    tableMap.attendanceForm(),
    tableMap.attendanceLog(),
    processAttendanceFormRecord
  );
  numOfThingsSynced += doFormSync(
    tableMap.tutorRegistrationForm(),
    tableMap.tutors(),
    processTutorRegistrationFormRecord
  );
  return numOfThingsSynced;
}

function uiRecalculateAttendance() {
  try {
    const numAttendancesChanged = onRecalculateAttendance();
    SpreadsheetApp.getUi().alert(
      `Finished attendance update. ${numAttendancesChanged} attendances were changed.`
    );
  } catch (err) {
    Logger.log(stringifyError(err));
    throw err;
  }
}

// This recalculates the attendance.
function onRecalculateAttendance() {
  let numAttendancesChanged = 0;
  function calculateIsBDay(x: string) {
    if (x.toLowerCase().charAt(0) === 'a') {
      return false;
    } else if (x.toLowerCase().charAt(0) === 'b') {
      return true;
    } else {
      throw new Error('unrecognized attendance day letter');
    }
  }

  function whenTutorFormNotFilledOutLogic(
    tutorId: number,
    learnerId: number,
    mod: number,
    day: Rec
  ) {
    const tutor = tutors[tutorId];
    const learner = learnerId === -1 ? null : learners[learnerId];
    const date = day.dateOfAttendance;
    // mark tutor as absent
    let alreadyExists = false; // if a presence or absence exists, don't add an absence
    if (tutor.attendance[date] === undefined) {
      tutor.attendance[date] = [];
    }
    if (learner !== null && learner.attendance[date] === undefined) {
      learner.attendance[date] = [];
    } else {
      for (const attendanceModDataString of tutor.attendance[date]) {
        const tokens = attendanceModDataString.split(' ');
        if (Number(tokens[0]) === mod) {
          alreadyExists = true;
        }
      }
    }
    if (!alreadyExists) {
      // add an absence for the tutor
      tutor.attendance[date].push(formatAttendanceModDataString(mod, 0));
      // add an excused absence for the learner, if exists
      if (learnerId !== -1) {
        learners[learnerId].attendance[date].push(
          formatAttendanceModDataString(mod, 1)
        );
      }
      numAttendancesChanged += 2;
    }
  }

  function applyAttendanceForStudent(
    attendance: any,
    entry: Rec,
    minutes: number
  ) {
    let isNew = true;
    if (attendance[entry.dateOfAttendance] === undefined) {
      attendance[entry.dateOfAttendance] = [];
    } else {
      for (const attendanceModDataString of attendance[
        entry.dateOfAttendance
      ]) {
        const tokens = attendanceModDataString.split(' ');
        if (entry.mod === Number(tokens[0])) {
          isNew = false;
        }
      }
    }
    // RESET BEHAVIOR: remove attendance logs marked with a reset flag
    if (entry.markForReset === 'doreset') {
      if (!isNew) {
        attendance[entry.dateOfAttendance] = attendance[
          entry.dateOfAttendance
        ].filter((attendanceModDataString: string) => {
          const tokens = attendanceModDataString.split(' ');
          if (Number(tokens[0]) === entry.mod) {
            ++numAttendancesChanged;
            return false;
          } else {
            return true;
          }
        });
      }
    } else {
      if (isNew) {
        // SPECIAL DATA FORMAT TO SAVE SPACE
        attendance[entry.dateOfAttendance].push(
          formatAttendanceModDataString(entry.mod, minutes)
        );
        ++numAttendancesChanged;
      }
    }
  }

  // read tables
  const tutors = tableMap.tutors().retrieveAllRecords();
  const learners = tableMap.learners().retrieveAllRecords();
  const attendanceLog = tableMap.attendanceLog().retrieveAllRecords();
  const attendanceDays = tableMap.attendanceDays().retrieveAllRecords();
  const matchings = tableMap.matchings().retrieveAllRecords();
  const tutorsArray = Object_values(tutors);
  const learnersArray = Object_values(learners);
  const matchingsArray = Object_values(matchings);
  const attendanceLogArray = Object_values(attendanceLog);

  // PROCESS EACH ATTENDANCE LOG ENTRY
  for (const entry of attendanceLogArray) {
    if (entry.validity !== '') {
      // exclude the entry
      continue;
    }
    if (entry.tutor !== -1) {
      applyAttendanceForStudent(
        tutors[entry.tutor].attendance,
        entry,
        entry.minutesForTutor
      );
    }
    if (entry.learner !== -1) {
      applyAttendanceForStudent(
        learners[entry.learner].attendance,
        entry,
        entry.minutesForLearner
      );
    }
  }

  // INDEX TUTORS: FIGURE OUT WHICH TUTORS/MODS HAVE UNSUBMITTED FORMS
  // For each tutor, keep track of which mods they need to fill out forms.
  // Figure out which tutors HAVEN'T filled out their forms.
  // Also keep track of learner ID associated with each mod (-1 for null).
  type TutorAttendanceFormIndexEntry = {
    id: number;
    mod: { [mod: number]: { wasFormSubmitted: boolean; learnerId: number } };
  };
  type TutorAttendanceFormIndex = {
    [id: number]: TutorAttendanceFormIndexEntry;
  };
  const tutorAttendanceFormIndex: TutorAttendanceFormIndex = {};
  for (const tutor of tutorsArray) {
    tutorAttendanceFormIndex[tutor.id] = {
      id: tutor.id,
      mod: {}
    };
    // handle drop-ins
    for (const mod of tutor.dropInMods) {
      tutorAttendanceFormIndex[tutor.id].mod[mod] = {
        wasFormSubmitted: false,
        learnerId: -1
      };
    }
  }
  for (const matching of matchingsArray) {
    tutorAttendanceFormIndex[matching.tutor].mod[matching.mod] = {
      wasFormSubmitted: false,
      learnerId: matching.learner
    };
  }

  // DEAL WITH THE UNSUBMITTED FORMS
  for (const day of Object_values(attendanceDays)) {
    if (
      day.status === 'ignore' ||
      day.status === 'isdone' ||
      day.status === 'isreset'
    ) {
      continue;
    }
    if (day.status !== 'doit' && day.status !== 'doreset') {
      throw new Error('unknown day status');
    }

    // modify dateOfAttendance; round down to the nearest 24-hour day
    day.dateOfAttendance = roundDownToDay(day.dateOfAttendance);

    const isBDay: boolean = calculateIsBDay(day.abDay);

    // iterate through all tutors & hunt down the ones that didn't fill out the form
    if (day.status === 'doit') {
      for (const tutor of tutorsArray) {
        for (let i = isBDay ? 10 : 0; i < (isBDay ? 20 : 10); ++i) {
          const x = tutorAttendanceFormIndex[tutor.id].mod[i];
          if (x !== undefined && !x.wasFormSubmitted) {
            whenTutorFormNotFilledOutLogic(tutor.id, x.learnerId, i, day);
          }
        }
        if (
          tutor.attendance[day.dateOfAttendance] !== undefined &&
          tutor.attendance[day.dateOfAttendance].length === 0
        ) {
          delete tutor.attendance[day.dateOfAttendance];
        }
      }
    } else if (day.status === 'doreset') {
      for (const tutor of tutorsArray) {
        if (tutor.attendance[day.dateOfAttendance] !== undefined) {
          // delete EVERY ABSENCE for that day (but keep excused)
          // this gets rid of anything automatically generated for that day
          tutor.attendance[day.dateOfAttendance] = tutor.attendance[
            day.dateOfAttendance
          ].filter((attendanceModDataString: string) => {
            const tokens = attendanceModDataString.split(' ');
            if (Number(tokens[1]) === 0) {
              ++numAttendancesChanged;
              return false;
            } else {
              return true;
            }
          });
        }
      }
    }

    // change day status
    if (day.status === 'doreset') {
      day.status = 'isreset';
    }
    if (day.status === 'doit') {
      day.status = 'isdone';
    }

    // update record
    tableMap.attendanceDays().updateRecord(day);
  }

  // THAT'S ALL! UPDATE TABLES
  tableMap.tutors().updateAllRecords(tutorsArray);
  tableMap.learners().updateAllRecords(learnersArray);

  return numAttendancesChanged;
}

function uiGenerateSchedule() {
  try {
    onGenerateSchedule();
    SpreadsheetApp.getUi().alert('Success');
  } catch (err) {
    Logger.log(stringifyError(err));
    throw err;
  }
}

// This is the KEY FUNCTION for generating the "schedule" sheet.
function onGenerateSchedule() {
  type ScheduleEntry = {
    mod: number;
    tutorName: string;
    info: string;
    isDropIn: boolean;
  };

  const sortComparator = (a: ScheduleEntry, b: ScheduleEntry) => {
    if (a.isDropIn < b.isDropIn) return -1;
    if (a.isDropIn > b.isDropIn) return 1;
    const x = a.tutorName.toLowerCase();
    const y = b.tutorName.toLowerCase();
    if (x < y) return -1;
    if (x > y) return 1;
    return 0;
  };

  // Delete & insert sheet
  let ss = SpreadsheetApp.openById(
    '1VbrZTMXGju_pwSrY7-M0l8citrYTm5rIv7RPuKN9deY'
  );
  let sheet = ss.getSheetByName('schedule');
  if (sheet !== null) {
    ss.deleteSheet(sheet);
  }
  sheet = ss.insertSheet('schedule', 0);

  // Get all matchings, tutors, and learners
  const matchings = tableMap.matchings().retrieveAllRecords();
  const tutors = tableMap.tutors().retrieveAllRecords();
  const learners = tableMap.learners().retrieveAllRecords();

  // Header
  sheet.appendRow(['ARC SCHEDULE']);
  sheet.appendRow(['']);

  // Create a list of [ mod, tutorName, info about matching, as string ]
  const scheduleInfo: ScheduleEntry[] = []; // mod = -1 means that the tutor has no one

  // Figure out all matchings that are finalized, indexed by tutor
  const index: {
    [tutorId: string]: { matchedMods: number[]; dropInMods: number[] };
  } = {};

  // create a bunch of blanks in the index
  for (const x of Object_values(tutors)) {
    index[String(x.id)] = {
      matchedMods: [],
      dropInMods: []
    };
  }

  // fill in index with matchings
  for (const x of Object_values(matchings)) {
    index[x.tutor].matchedMods.push(x.mod);
    scheduleInfo.push({
      isDropIn: false,
      mod: x.mod,
      tutorName: tutors[x.tutor].friendlyFullName,
      info:
        (x.learner === -1
          ? ' (SPECIAL)'
          : `(w/${learners[x.learner].friendlyFullName})`) +
        (x.annotation === '' ? '' : ` (INFO: ${x.annotation})`)
    });
  }

  const unscheduledTutorNames: string[] = [];

  // fill in index with drop-ins
  for (const x of Object_values(tutors)) {
    for (const mod of x.dropInMods) {
      // if the tutor is matched, they don't count for drop-in
      let matchedAtMod = false;
      for (const mod2 of index[x.id].matchedMods) {
        if (mod2 === mod) {
          matchedAtMod = true;
          break;
        }
      }
      if (matchedAtMod) continue;

      index[x.id].dropInMods.push(mod);

      scheduleInfo.push({
        isDropIn: true,
        mod,
        tutorName: x.friendlyFullName,
        info: ''
      });
    }
    if (
      index[String(x.id)].matchedMods.length === 0 &&
      index[String(x.id)].dropInMods.length === 0
    ) {
      unscheduledTutorNames.push(x.friendlyFullName);
    }
  }

  // Print!
  // COLOR SCHEME
  const COLOR_SCHEME_STRING = 'ff99c8-fcf6bd-d0f4de-a9def9-e4c1f9';
  const COLOR_SCHEME = COLOR_SCHEME_STRING.split('-');
  const DAY_OF_WEEK = new Date().getDay();
  // calculate the color based on the day of the week
  const PRIMARY_COLOR =
    '#' +
    COLOR_SCHEME[
      {
        0: 0,
        1: 1,
        2: 2,
        3: 3,
        4: 4,
        5: 2,
        6: 3
      }[DAY_OF_WEEK]
    ];
  const SECONDARY_COLOR = '#383f51';

  // FORMAT SHEET
  sheet.setHiddenGridlines(true);

  // CHANGE COLUMNS
  sheet.deleteColumns(5, sheet.getMaxColumns() - 5);
  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 30);
  sheet.setColumnWidth(4, 300);
  sheet.setColumnWidth(5, 30);

  // HEADER
  sheet.getRange(1, 1, 3, 5).mergeAcross();
  sheet
    .getRange(1, 1)
    .setValue('ARC Schedule')
    .setFontSize(36)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(2, 15);
  sheet.setRowHeight(3, 15);
  sheet
    .getRange(4, 2)
    .setValue('A Days')
    .setFontSize(18)
    .setHorizontalAlignment('center');
  sheet
    .getRange(4, 4)
    .setValue('B Days')
    .setFontSize(18)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(5, 30);
  sheet.getRange(1, 1, 4, 5).setBackground(PRIMARY_COLOR);

  const layoutMatrix: [ScheduleEntry[], ScheduleEntry[]][] = []; // [mod0to9][abday]
  for (let i = 0; i < 10; ++i) {
    layoutMatrix.push([
      scheduleInfo.filter(x => x.mod === i + 1).sort(sortComparator), // A days
      scheduleInfo.filter(x => x.mod === i + 11).sort(sortComparator) // B days
    ]);
  }

  // LAYOUT
  let nextRow = 6;
  for (let i = 0; i < 10; ++i) {
    const scheduleRowSize = Math.max(
      layoutMatrix[i][0].length,
      layoutMatrix[i][1].length,
      1
    );
    // LABEL
    sheet.getRange(nextRow, 1, scheduleRowSize).merge();
    sheet
      .getRange(nextRow, 1)
      .setValue(`${i + 1}`)
      .setFontSize(18)
      .setVerticalAlignment('top');

    // CONTENT
    if (layoutMatrix[i][0].length > 0) {
      sheet
        .getRange(nextRow, 2, layoutMatrix[i][0].length)
        .setValues(layoutMatrix[i][0].map(x => [`${x.tutorName} ${x.info}`]))
        .setWrap(true)
        .setFontSize(12)
        .setFontColors(
          layoutMatrix[i][0].map(x => [x.isDropIn ? 'black' : 'red'])
        );
    }
    if (layoutMatrix[i][1].length > 0) {
      sheet
        .getRange(nextRow, 4, layoutMatrix[i][1].length)
        .setValues(layoutMatrix[i][1].map(x => [`${x.tutorName} ${x.info}`]))
        .setWrap(true)
        .setFontSize(12)
        .setFontColors(
          layoutMatrix[i][1].map(x => [x.isDropIn ? 'black' : 'red'])
        );
    }

    // SET THE NEXT ROW
    nextRow += scheduleRowSize;

    // GUTTER
    sheet.getRange(nextRow, 1, 1, 5).merge();
    sheet.setRowHeight(nextRow, 60);
    ++nextRow;
  }

  // UNSCHEDULED TUTORS
  sheet
    .getRange(nextRow, 2, 1, 3)
    .merge()
    .setValue(`Unscheduled tutors`)
    .setFontSize(18)
    .setFontStyle('italic')
    .setHorizontalAlignment('center')
    .setWrap(true);
  sheet
    .getRange(nextRow, 2, unscheduledTutorNames.length + 1, 3)
    .setBorder(
      true,
      true,
      true,
      true,
      null,
      null,
      PRIMARY_COLOR,
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );
  ++nextRow;
  sheet
    .getRange(nextRow, 2, unscheduledTutorNames.length, 3)
    .mergeAcross()
    .setHorizontalAlignment('center');
  sheet
    .getRange(nextRow, 2, unscheduledTutorNames.length)
    .setFontSize(12)
    .setValues(unscheduledTutorNames.map(x => [x]));
  nextRow += unscheduledTutorNames.length;

  // FOOTER
  sheet.getRange(nextRow, 1, 1, 5).merge();
  sheet.setRowHeight(nextRow, 20);
  ++nextRow;

  sheet
    .getRange(nextRow, 1, 1, 5)
    .merge()
    .setValue(`Schedule auto-generated on ${new Date()}`)
    .setFontSize(10)
    .setFontColor('white')
    .setBackground(SECONDARY_COLOR)
    .setHorizontalAlignment('center');
  ++nextRow;

  sheet
    .getRange(nextRow, 1, 1, 5)
    .merge()
    .setValue(`ARC App designed by Suhao Jeffrey Huang`)
    .setFontSize(10)
    .setFontColor('white')
    .setBackground(SECONDARY_COLOR)
    .setHorizontalAlignment('center');
  ++nextRow;

  // FIT ROWS/COLUMNS
  sheet.deleteRows(
    sheet.getLastRow() + 1,
    sheet.getMaxRows() - sheet.getLastRow()
  );

  // FONT
  sheet.getDataRange().setFontFamily('Helvetica');

  return null;
}

function uiDeDuplicateTutorsAndLearners() {
  const PROMPT_TEXT =
    'This command will de-duplicate tutors and learners. Old form submissions are replaced with newer form submissions. Type the word "proceed" to proceed. Leave the box blank to cancel.';

  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(PROMPT_TEXT);
    if (response.getResponseText() === 'proceed') {
      const numberOfOperations = deDuplicateTutorsAndLearners();
      SpreadsheetApp.getUi().alert(
        `${numberOfOperations} de-duplications performed`
      );
    }
  } catch (err) {
    Logger.log(stringifyError(err));
    throw err;
  }
}

function deDuplicateTutorsAndLearners(): number {
  let numberOfOperations = 0;
  numberOfOperations += deDuplicateTableByStudentId(tableMap.tutors());
  numberOfOperations += deDuplicateTableByStudentId(tableMap.learners());
  return numberOfOperations;
}

function deDuplicateTableByStudentId(table: Table): number {
  let numberOfOperations = 0;
  type ReplacementMap = { [studentId: number]: { recordIds: number[] } };

  const replacementMap: ReplacementMap = {};

  const recordObject = table.retrieveAllRecords();
  for (const record of Object_values(recordObject)) {
    if (replacementMap[record.studentId] === undefined) {
      replacementMap[record.studentId] = { recordIds: [record.id] };
    } else {
      replacementMap[record.studentId].recordIds.push(record.id);
    }
  }
  // If a student ID has multiple records, delete all but the newest (highest ID)
  // and the remaining record should have ID edited to the oldest record
  for (const mapItem of Object_values(replacementMap)) {
    if (mapItem.recordIds.length >= 2) {
      let lowestId = mapItem.recordIds[0];
      let highestId = mapItem.recordIds[0];
      for (let i = 1; i < mapItem.recordIds.length; ++i) {
        if (mapItem.recordIds[i] < lowestId) {
          lowestId = mapItem.recordIds[i];
        }
        if (mapItem.recordIds[i] > highestId) {
          highestId = mapItem.recordIds[i];
        }
      }

      // delete all records not equal to the highest ID
      for (let i = 0; i < mapItem.recordIds.length; ++i) {
        if (highestId !== mapItem.recordIds[i]) {
          table.deleteRecord(mapItem.recordIds[i]);
          ++numberOfOperations;
        }
      }

      // This edits the ID of the record, which is very uncommon,
      // so the second argument of updateRecord is invoked.
      recordObject[highestId].id = lowestId;
      table.updateRecord(recordObject[highestId], table.getRowById(highestId));
      ++numberOfOperations;
    }
  }
  return numberOfOperations;
}

function onOpen(_ev: any) {
  const menu = SpreadsheetApp.getUi().createMenu('ARC APP');
  menu.addItem('Sync data from forms', 'uiSyncForms');
  menu.addItem('Generate schedule', 'uiGenerateSchedule');
  menu.addItem('Recalculate attendance', 'uiRecalculateAttendance');
  menu.addItem(
    'De-duplicate tutors and learners',
    'uiDeDuplicateTutorsAndLearners'
  );
  if (ARC_APP_DEBUG_MODE) {
    menu.addItem('Debug: autoformat', 'debugAutoformat');
    menu.addItem('Debug: test client API', 'debugClientApiTest');
    menu.addItem('Debug: reset all tables', 'debugResetEverything');
    menu.addItem('Debug: reset all small tables', 'debugResetAllSmallTables');
    menu.addItem('Debug: rebuild all headers', 'debugHeaders');
    menu.addItem('Debug: rewrite all tables', 'debugRewriteEverything');
    menu.addItem('Debug: run temporary script', 'debugRunTemporaryScript');
  }
  menu.addToUi();
}

function debugRunTemporaryScript() {
  /*
  Clean any attendance times that are not an "integer" number of days.

  const tutors = Object_values(tableMap.tutors().retrieveAllRecords());
  const learners = Object_values(tableMap.learners().retrieveAllRecords());
  let count = 0;
  for (const tutor of tutors) {
    for (const x of Object.getOwnPropertyNames(tutor.attendance)) {
      if (roundDownToDay(Number(x)) !== Number(x)) {
        ++count;
        delete tutor.attendance[x];
      }
    }
  }
  for (const learner of learners) {
    for (const x of Object.getOwnPropertyNames(learner.attendance)) {
      if (roundDownToDay(Number(x)) !== Number(x)) {
        ++count;
        delete learner.attendance[x];
      }
    }
  }
  tableMap.tutors().updateAllRecords(tutors);
  tableMap.learners().updateAllRecords(learners);
  SpreadsheetApp.getUi().alert(String(count));
  */
}
