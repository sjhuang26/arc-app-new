/**
 * @OnlyCurrentDoc
 */

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
    const timeInThisTimezone = utcTime - 4 * 3600 * 1000;
    return Math.floor(timeInThisTimezone / 1000 / 86400) * 86400 * 1000 + 4 * 3600 * 1000;
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
}

abstract class Field {
    name: string;
    abstract parse(x: any): any
    abstract serialize(x: any): any

    constructor(name: string = null) {
        this.name = name;
    }
}

class BooleanField extends Field {
    constructor(name: string = null) {
        super(name);
    }
    parse(x: any) {
        return x === 'true' ? 'true' : 'false';
    }
    serialize(x: any) {
        return x ? 'true' : 'false';
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
}
type RecCollection = {
    [id: string]: Rec;
}

/*

TABLE CLASS

*/

class Table {
    idCounter: number = (new Date()).getTime();

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
            new StringField('contactPref')
        ]
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
                new JsonField('dropInMods')
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
                new StringField('specialRoom')
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
                new NumberField('studentId'),
                new StringField('specialRoom'),
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
                new StringField('status'),
                new StringField('specialRoom')
            ]
        },
        requestForm: {
            sheetName: '$request-form',

            // this means that the ID field is automatically generated from the date field
            isForm: true,

            fields: [
                /* Layout of form
                    Timestamp
                    Legal first name
                    Legal last name
                    Short friendly name (ie. "Jeffrey")
                    Full friendly name (ie. "Jeffrey Huang")
                    Student ID (ie. "20186")
                    Grade
                    What subject(s) do you want to be tutored in?
                    Mods available (ie. study halls, frees) (part 1) [A]
                    Mods available (ie. study halls, frees) (part 1) [B]
                    Mods available (part 2) [A]
                    Mods available (part 2) [B]
                    Email
                    Phone (XXX-XXX-XXXX)
                    What kind of contact do you prefer?
                    What is your favorite flavor of ice cream?
                */
                // THE ORDER OF THE FIELDS MATTERS! They must match the order of the form's questions.
                new DateField('date'),
                new StringField('firstName'),
                new StringField('lastName'),
                new StringField('friendlyName'),
                new StringField('friendlyFullName'),
                new NumberField('studentId'),
                new NumberField('grade'),
                new StringField('subject'),
                new StringField('modDataA1To5'),
                new StringField('modDataB1To5'),
                new StringField('modDataA6To10'),
                new StringField('modDataB6To10'),
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
                /*
                Layout of form:
                    Timestamp
                    Legal first name
                    Legal last name
                    Short friendly name (ie. "Jeffrey")
                    Full friendly name (ie. "Jeffrey Huang")
                    Student ID (ie. "20186")
                    Grade
                    What subject(s) do you want to be tutored in? You can write "all subjects" if needed.
                    A days or B days?
                    Which mod?
                    What room number?
                    What is your favorite flavor of ice cream?
                */
                new DateField('date'),
                new StringField('firstName'),
                new StringField('lastName'),
                new StringField('friendlyName'),
                new StringField('friendlyFullName'),
                new NumberField('studentId'),
                new NumberField('grade'),
                new StringField('subject'),
                new StringField('abDay'),
                new NumberField('mod1To10'),
                new StringField('specialRoom'),
                new StringField('iceCreamQuestion')
                // contact info is intentionally omitted
            ]
        },
        attendanceForm: {
            sheetName: '$attendance-form',
            isForm: true,
            fields: [
                /*
                Layout of form:
                    Timestamp
                    Date (LEAVE BLANK if taking today's attendance)
                    A or B day?
                    Mod?
                    Student ID?
                    Is your learner/tutor present with you?
                */
                new DateField('date'),
                new DateField('dateOfAttendance'), // optional in the form
                new StringField('abDay'),
                new NumberField('mod1To10'),
                new NumberField('studentId'),
                new StringField('learnerOrTutor'),
                new StringField('presence')
            ]
        },
        tutorRegistrationForm: {
            /*
            Layout of form:
                Timestamp
                Legal first name
                Legal last name
                Short friendly name (ie. "Jeffrey")
                Full friendly name (ie. "Jeffrey Huang")
                Student ID (ie. "20186")
                Grade
                Email
                Phone (XXX-XXX-XXXX)
                What kind of contact do you prefer?
                All mods available (part 1) [A]
                All mods available (part 1) [B]
                All mods available (part 2) [A]
                All mods available (part 2) [B]
                Mods preferred (part 1) [A]
                Mods preferred (part 1) [B]
                Mods preferred (part 2) [A]
                Mods preferred (part 2) [B]
                Subjects: English
                Subjects: Social Studies
                Subjects: Math
                Subjects: Foreign languages
                Subjects: Science
                Subjects: Computer Science
                Subjects: Other
                Subjects: Anything else?
                Favorite flavor of ice cream?
                You win if you pick the number that the fewest other people choose .... Think wisely, and be unique!
            */
            sheetName: '$tutor-registration-form',
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
                new NumberField('minutesForLearner')
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
        this.sheet = SpreadsheetApp.getActive().getSheetByName(this.tableInfo.sheetName);
        if (this.sheet === null) {
            if (this.isForm) {
                throw new Error(`table ${this.name} not found, and it's supposed to be a form`);
            } else {
                this.sheet = SpreadsheetApp.getActive().insertSheet(this.tableInfo.sheetName);
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
            this.sheet.getRange(1, 1, 1, this.tableInfo.fields.length).setValues([this.tableInfo.fields.map(field => field.name)]);
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
            throw new Error(`something's wrong with the columns of table ${this.name} (${this.tableInfo.fields.length})`)
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
        return this.tableInfo.fields.map(field => field.serialize(record[field.name]));
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
        const mat: any[][] = this.sheet.getRange(2, 1, this.sheet.getLastRow() - 1).getValues();
        let rowNum = -1;
        for (let i = 0; i < mat.length; ++i) {
            const cell: number = mat[i][0];
            if (typeof cell !== 'number') {
                throw new Error(`id at location ${String(i)} is not a number in table ${String(this.name)}`);
            }
            if (cell === id) {
                if (rowNum !== -1) {
                    throw new Error(`duplicate ID ${String(id)} in table ${String(this.name)}`);
                }
                rowNum = i + 2; // i = 0 <=> second row (rows are 1-indexed)
            }
        }
        if (rowNum == -1) {
            throw new Error(`ID ${String(id)} not found in table ${String(this.name)}`);
        }
        return rowNum;
    }

    updateRecord(editedRecord: Rec, rowNum?: number): void {
        if (rowNum === undefined) {
            rowNum = this.getRowById(editedRecord.id);
        }
        this.sheet.getRange(rowNum, 1, 1, this.sheetLastColumn).setValues([this.serializeRecord(editedRecord)]);
    }

    updateAllRecords(editedRecords: Rec[]): void {
        if (this.sheet.getLastRow() === 1) {
            return; // the sheet is empty, and trying to select it will result in an error
        }
        // because the first row is headers, we ignore it and start from the second row
        const mat: any[][] = this.sheet.getRange(2, 1, this.sheet.getLastRow() - 1).getValues();
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
        const time = (new Date()).getTime();
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
}

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
    // TODO: fix this
    tableMap.operationLog().resetEntireSheet();

    return HtmlService.createHtmlOutputFromFile('index');
}



function processClientAsk(args: any[]): ServerResponse<any> {
    const resourceName: string = args[0];
    if (resourceName === undefined) {
        throw new Error('no args, or must specify resource name');
    }
    if (tableMap[resourceName] === undefined) {
        throw new Error(`resource ${String(resourceName)} not found`);
    }
    const resource = tableMap[resourceName]();
    return {
        error: false,
        val: resource.processClientAsk(args.slice(1)),
        message: null
    };
}

// this is the MAIN ENTRYPOINT that the client uses to ask the server for data.
function onClientAsk(args: any[]): string {
    let returnValue = {
        error: true,
        val: null,
        message: 'Mysterious error'
    };
    try {
        returnValue = processClientAsk(args);   
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

// This, and anything related to it, is 100% TODO.
function onClientNotification(args: any[]): void {
    // we record the logs
    // TODO: have the client read them every 20 seconds so they know the things that other clients have done
    // in the case that multiple clients are open at once
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
        ui.alert(JSON.stringify(onClientAsk(JSON.parse(response.getResponseText()))));
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
        const response = ui.prompt('Leave the box below blank to cancel debug operation.');
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
        const response = ui.prompt('Leave the box below blank to cancel debug operation.');
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
        const response = ui.prompt('Leave the box below blank to cancel debug operation.');
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
function doFormSync(formTable: Table, actualTable: Table, formRecordToActualRecord: (formRecord: Rec) => Rec): number {
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

// The main "sync forms" function that's crammed with form data formatting.
function onSyncForms() {
    // tables
    const tutors = tableMap.tutors().retrieveAllRecords();
    const learners = tableMap.learners().retrieveAllRecords();
    const matchings = tableMap.matchings().retrieveAllRecords();

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
        return mA15.concat(mA60).concat(mB15).concat(mB60);
    }

    function parseStudentConfig(r: Rec): ObjectMap<any> {
        return {
            firstName: r.firstName,
            lastName: r.lastName,
            friendlyName: r.friendlyName,
            friendlyFullName: r.friendlyFullName,
            grade: parseGrade(r.grade),
            studentId: r.studentId,
            email: r.email,
            phone: r.phone,
            contactPref: parseContactPref(r.contactPref)
        };
    }

    function processRequestFormRecord(r: Rec): Rec {
        return {
            id: -1,
            date: r.date, // the date MUST be the date from the form; this is used for syncing
            subject: r.subject,
            mods: parseModData([r.modDataA1To5, r.modDataB1To5, r.modDataA6To10, r.modDataB6To10]),
            specialRoom: '',
            ...parseStudentConfig(r),
            status: 'unchecked'
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
            mods: parseModData([r.modDataA1To5, r.modDataB1To5, r.modDataA6To10, r.modDataB6To10]),
            modsPref: parseModData([r.modDataPrefA1To5, r.modDataPrefB1To5, r.modDataPrefA6To10, r.modDataPrefB6To10]),
            subjectList: parseSubjectList([r.subjects0, r.subjects1, r.subjects2, r.subjects3, r.subjects4, r.subjects5, r.subjects6, r.subjects7]),
            attendance: {},
            dropInMods: []
        };
    }
    function processSpecialRequestFormRecord(r: Rec): Rec {
        return {
            id: -1,
            date: r.date, // the date MUST be the date from the form
            subject: r.subject,
            specialRoom: r.specialRoom,
            mods: [parseModInfo(r.abDay, r.mod1To10)],
            ...parseStudentConfig(r),

            status: 'unchecked'
        };
    }
    function processAttendanceFormRecord(r: Rec): Rec {
        const mod = parseModInfo(r.abDay, r.mod1To10);
        let tutor = -1;
        let learner = -1;
        let validity = '';
        let minutesForTutor = -1;
        let minutesForLearner = -1;
        if (r.learnerOrTutor === 'Learner') {
            // give the learner their time
            minutesForLearner = MINUTES_PER_MOD;

            // figure out who the learner is, by student ID
            const xLearners = recordCollectionToArray(learners).filter(x => x.studentId === r.studentId);
            if (xLearners.length === 0) {
                validity = 'learner student ID does not exist';
            } else if (xLearners.length === 1) {
                learner = xLearners[0].id;
            } else {
                throw new Error(`duplicate learner ID#${String(r.studentId)}`);
            }

            // learners always have tutors --- figure out who the tutor is, by looking through matchings
            const xMatchings = recordCollectionToArray(matchings).filter(x => x.status === 'finalized' && x.learner === learner && x.mod === mod);
            if (xMatchings.length === 0) {
                validity = 'matching does not exist';
            } else if (xMatchings.length === 1) {
                tutor = xMatchings[0].tutor;
            } else {
                throw new Error(`the learner ${xMatchings[0].friendlyFullName} is matched twice on the same mod`);
            }

            // see if the tutor showed up
            if (validity === '') {
                if (r.presence === 'Yes') {
                    minutesForTutor = MINUTES_PER_MOD;
                } else if (r.presence === 'No') {
                    minutesForTutor = 0;
                } else if (r.presence === `I'm a tutor and don't have a learner assigned`) {
                    // but they said that they were a learner! nonsense!
                    validity = 'incompatible set of form answers';
                } else {
                    throw new Error(`invalid presence (${String(r.presence)})`)
                }
            }
        } else if (r.learnerOrTutor === 'Tutor') {
            // give the tutor their time
            minutesForTutor = MINUTES_PER_MOD;

            // figure out who the tutor is, by student ID
            const xTutors = recordCollectionToArray(tutors).filter(x => x.studentId === r.studentId);
            if (xTutors.length === 0) {
                validity = 'tutor student ID does not exist';
            } else if (xTutors.length === 1) {
                tutor = xTutors[0].id;
            } else {
                throw new Error(`duplicate tutor student id ${String(r.studentId)}`);
            }

            // does the tutor have a learner?
            const xMatchings = recordCollectionToArray(matchings).filter(x => x.status === 'finalized' && x.tutor === tutor && x.mod === mod);
            if (xMatchings.length === 0) {
                learner = -1;
            } else if (xMatchings.length === 1) {
                learner = xMatchings[0].learner;
            } else {
                throw new Error(`the tutor ${xMatchings[0].friendlyFullName} is matched twice on the same mod`);
            }

            // see if the learner showed up
            if (validity === '') {
                if (r.presence === 'Yes') {
                    minutesForLearner = MINUTES_PER_MOD;
                } else if (r.presence === 'No') {
                    minutesForLearner = 0;
                } else if (r.presence === `I'm a tutor and don't have a learner assigned`) {
                    if (learner === -1) {
                        minutesForLearner = -1;
                    } else {
                        // so there really is a learner...
                        validity = `tutor said they don't have a learner assigned, but they do!`
                    }
                } else {
                    throw new Error(`invalid presence (${String(r.presence)})`)
                }
            }
        } else {
            throw new Error('process attendance error: learner/tutor naming issue')
        }

        return {
            id: r.id,
            date: r.date,
            dateOfAttendance: roundDownToDay(r.dateOfAttendance === -1 ? r.date : r.dateOfAttendance),
            validity,
            mod,
            tutor,
            learner,
            minutesForTutor,
            minutesForLearner
        };
    }

    try {
        let numOfThingsSynced = 0;
        numOfThingsSynced += doFormSync(tableMap.requestForm(), tableMap.requestSubmissions(), processRequestFormRecord);
        numOfThingsSynced += doFormSync(tableMap.specialRequestForm(), tableMap.requestSubmissions(), processSpecialRequestFormRecord);
        numOfThingsSynced += doFormSync(tableMap.attendanceForm(), tableMap.attendanceLog(), processAttendanceFormRecord);
        numOfThingsSynced += doFormSync(tableMap.tutorRegistrationForm(), tableMap.tutors(), processTutorRegistrationFormRecord);
        SpreadsheetApp.getUi().alert(`Finished sync! ${numOfThingsSynced} new form submits found.`);
    } catch (err) {
        Logger.log(stringifyError(err));
        throw err;
    }
}

// This recalculates the attendance.
function onRecalculateAttendance() {
    try {
        let numAttendancesChanged = 0;
        
        // read tables
        const tutors = tableMap.tutors().retrieveAllRecords();
        const tutorsArray = Object_values(tutors);
        const learners = tableMap.learners().retrieveAllRecords();
        const attendanceLog = tableMap.attendanceLog().retrieveAllRecords();
        const attendanceDays = tableMap.attendanceDays().retrieveAllRecords();
        const matchings = tableMap.matchings().retrieveAllRecords();
        // combine presences
        for (const x of Object_values(attendanceLog)) {
            if (x.tutor !== -1) {
                const y = tutors[String(x.tutor)];
                if (y.attendance[String(x.dateOfAttendance)] === undefined) {
                    y.attendance[String(x.dateOfAttendance)] = [];
                }
                let isNew = true;
                for (const z of y.attendance[String(x.dateOfAttendance)]) {
                    if (x.mod === z.mod) {
                        isNew = false;
                    }
                }
                if (isNew) {
                    y.attendance[String(x.dateOfAttendance)].push({ date: x.dateOfAttendance, mod: x.mod, minutes: x.minutesForTutor });
                    ++numAttendancesChanged;
                }
            }
            if (x.learner !== -1) {
                const y = learners[String(x.learner)];
                if (y.attendance[String(x.dateOfAttendance)] === undefined) {
                    y.attendance[String(x.dateOfAttendance)] = [];
                }
                let isNew = true;
                for (const z of y.attendance[String(x.dateOfAttendance)]) {
                    if (x.mod === z.mod) {
                        isNew = false;
                    }
                }
                if (isNew) {
                    y.attendance[String(x.dateOfAttendance)].push({ date: x.dateOfAttendance, mod: x.mod, minutes: x.minutesForLearner });
                    ++numAttendancesChanged;
                }
            }
        }
        // index whether a tutor should be showing up
        const tutorShouldBeShowingUp: { [id: string]: { [mod: string]: boolean } } = {};
        for (const tutor of tutorsArray) {
            tutorShouldBeShowingUp[String(tutor.id)] = {};
            // handle drop-ins
            for (const mod of tutor.dropInMods) {
                tutorShouldBeShowingUp[String(tutor.id)][mod] = true;
            }
        }
        for (const matching of Object_values(matchings)) {
            if (matching.status === 'finalized') {
                tutorShouldBeShowingUp[String(matching.tutor)][matching.mod] = true;
            }
        }

        Logger.log(JSON.stringify(tutorShouldBeShowingUp));

        // mark absences
        for (const day of Object_values(attendanceDays)) {
            // round down to the nearest 24-hour day
            day.dateOfAttendance = roundDownToDay(day.dateOfAttendance);

            if (day.status === 'upcoming' || day.status === 'finalized' || day.status === 'reset') {
                // ignore
            } else if (day.status === 'unfinalized' || day.status === 'unreset') {
                let isBDay: boolean = null;
                if (day.abDay.toLowerCase().charAt(0) === 'a') {
                    isBDay = false;
                } else if (day.abDay.toLowerCase().charAt(0) === 'b') {
                    isBDay = true;
                } else {
                    throw new Error('unrecognized attendance day letter');
                }
                for (const tutor of tutorsArray) {
                    Logger.log(JSON.stringify(tutor));
                    for (const mod of tutor.mods) {
                        if (isBDay ? (10 < mod) : (mod <= 10)) {
                            if (tutorShouldBeShowingUp[String(tutor.id)][String(mod)] === true) {
                                // mark tutor as (un-?)absent at a specific date and mod
                                if (day.status === 'unreset') {
                                    if (tutor.attendance[day.dateOfAttendance] !== undefined) {
                                        tutor.attendance[day.dateOfAttendance] = tutor.attendance[day.dateOfAttendance].filter((x: any) => {
                                            if (x.mod === mod && x.minutes === 0) {
                                                ++numAttendancesChanged;
                                                return false;
                                            } else {
                                                return true;
                                            }
                                        });
                                    }
                                }
                                if (day.status === 'unfinalized') {
                                    let alreadyExists = false; // if a presence or absence exists, don't add an absence
                                    if (tutor.attendance[day.dateOfAttendance] === undefined) {
                                        tutor.attendance[day.dateOfAttendance] = [];
                                    } else {
                                        for (const x of tutor.attendance[day.dateOfAttendance]) {
                                            if (x.mod === mod) {
                                                alreadyExists = true;
                                            }
                                        }
                                    }
                                    if (!alreadyExists) {
                                        // add an absence!
                                        tutor.attendance[day.dateOfAttendance].push({ date: day.dateOfAttendance, mod, minutes: 0 });
                                        ++numAttendancesChanged;
                                    }
                                }
                            }
                        }
                    }
                    if (tutor.attendance[day.dateOfAttendance] !== undefined && tutor.attendance[day.dateOfAttendance].length === 0) {
                        delete tutor.attendance[day.dateOfAttendance];
                    }
                }

                // change day status
                if (day.status === 'unreset') {
                    day.status = 'reset';
                }
                if (day.status === 'unfinalized') {
                    day.status = 'finalized';
                }
                tableMap.attendanceDays().updateRecord(day);
            } else {
                throw new Error('unknown day status');
            }
        }

        // update table
        tableMap.tutors().updateAllRecords(tutorsArray);

        SpreadsheetApp.getUi().alert(`Finished attendance update. ${numAttendancesChanged} attendances were changed.`);
    } catch (err) {
        Logger.log(stringifyError(err));
        throw err;
    }
}

// This is the KEY FUNCTION for generating the "schedule" sheet.
function onGenerateSchedule() {
    type ScheduleEntry = { mod: number, tutorName: string, info: string; isDropIn: boolean };

    const sortComparator = (a: ScheduleEntry, b: ScheduleEntry) => {
        const x = a.tutorName.toLowerCase();
        const y = b.tutorName.toLowerCase();
        if (x < y) return -1;
        if (x > y) return 1;
        return 0;
    };

    try {
        // Delete & insert sheet
        let ss = SpreadsheetApp.getActive();
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
        sheet.appendRow([`Automatically generated on ${new Date()}`]);

        // Create a list of [ mod, tutorName, info about matching, as string ]
        const scheduleInfo: ScheduleEntry[] = []; // mod = -1 means that the tutor has no one

        // Figure out all matchings that are finalized, indexed by tutor
        const index: { [tutorId: string]: { isMatched: boolean; hasBeenScheduled: boolean; } } = {};

        // create a bunch of blanks in the index
        for (const x of Object_values(tutors)) {
            index[String(x.id)] = {
                isMatched: false,
                hasBeenScheduled: false
            }
        }

        // fill in index with matchings
        for (const x of Object_values(matchings)) {
            if (x.status === 'finalized') {
                const name = learners[x.learner].friendlyFullName;
                index[x.tutor].isMatched = true;
                index[x.tutor].hasBeenScheduled = true;
                scheduleInfo.push({
                    isDropIn: false,
                    mod: x.mod,
                    tutorName: tutors[x.tutor].friendlyFullName,
                    info: (x.specialRoom === '' || x.specialRoom === undefined) ? `(w/${name})` : `(w/${name} SPECIAL @room ${x.specialRoom})`
                });
            }
        }

        const unscheduledTutorNames: string[] = [];

        // fill in index with drop-ins
        for (const x of Object_values(tutors)) {
            if (!index[String(x.id)].isMatched) {
                for (const mod of x.dropInMods) {
                    index[String(x.id)].hasBeenScheduled = true;
                    scheduleInfo.push({
                        isDropIn: true,
                        mod,
                        tutorName: x.friendlyFullName,
                        info: '(drop in)'
                    });
                }
                // unscheduled?
                if (!index[String(x.id)].hasBeenScheduled) {
                    unscheduledTutorNames.push(x.friendlyFullName);
                }
            }
        }

        // Print!
        // CHANGE COLUMNS
        sheet.deleteColumns(5, sheet.getMaxColumns() - 5);
        sheet.setColumnWidth(1, 30);
        sheet.setColumnWidth(2, 300);
        sheet.setColumnWidth(3, 30);
        sheet.setColumnWidth(4, 300);
        sheet.setColumnWidth(5, 30);

        // HEADER
        sheet.getRange(1, 1, 3, 5).mergeAcross();
        sheet.getRange(1, 1).setValue('ARC Schedule').setFontSize(36).setHorizontalAlignment('center');
        sheet.getRange(2, 1).setValue('Automatically generated on ' + new Date().toISOString()).setFontSize(14).setHorizontalAlignment('center');
        sheet.setRowHeight(3, 30);
        sheet.getRange(4, 2).setValue('A Days').setFontSize(18).setHorizontalAlignment('center');
        sheet.getRange(4, 4).setValue('B Days').setFontSize(18).setHorizontalAlignment('center');
        sheet.setRowHeight(5, 30);

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
            const scheduleRowSize = Math.max(layoutMatrix[i][0].length, layoutMatrix[i][1].length);
            // LABEL
            sheet.getRange(nextRow, 1, scheduleRowSize).merge();
            sheet.getRange(nextRow, 1).setValue(`${i + 1}`).setFontSize(18).setVerticalAlignment('top');

            // CONTENT
            sheet.getRange(nextRow, 2, layoutMatrix[i][0].length).setValues(layoutMatrix[i][0].map(x => [`${x.tutorName} ${x.info}`])).setWrap(true).setFontColors(layoutMatrix[i][0].map(x => [x.isDropIn ? 'black' : 'red']));
            sheet.getRange(nextRow, 4, layoutMatrix[i][1].length).setValues(layoutMatrix[i][1].map(x => [`${x.tutorName} ${x.info}`])).setWrap(true).setFontColors(layoutMatrix[i][1].map(x => [x.isDropIn ? 'black' : 'red']));
            
            // SET THE NEXT ROW
            nextRow += scheduleRowSize;

            // GUTTER
            sheet.getRange(nextRow, 1, 1, 5).merge();
            sheet.setRowHeight(nextRow, 60);
            ++nextRow;
        }
        
        // UNSCHEDULED TUTORS
        sheet.getRange(nextRow, 2, 1, 3).merge().setValue(`Unscheduled tutors`).setFontSize(18).setFontStyle('italic').setHorizontalAlignment('center').setWrap(true);
        ++nextRow;
        sheet.getRange(nextRow, 2, unscheduledTutorNames.length, 3).mergeAcross().setHorizontalAlignment('center');
        sheet.getRange(nextRow, 2, unscheduledTutorNames.length).setValues(unscheduledTutorNames.map(x => [x + ' (unscheduled)']));
        nextRow += unscheduledTutorNames.length;

        // FOOTER
        sheet.getRange(nextRow, 2, 1, 4).merge().setValue(`That's all!`).setFontSize(18).setFontStyle('italic').setHorizontalAlignment('center');
        ++nextRow;

        // FIT ROWS/COLUMNS
        sheet.deleteRows(sheet.getLastRow() + 1, sheet.getMaxRows() - sheet.getLastRow());
    } catch (err) {
        Logger.log(stringifyError(err));
        throw err;
    }
}

function onOpen(_ev: any) {
    const menu = SpreadsheetApp.getUi().createMenu('ARC APP');
    menu.addItem('Sync data from forms', 'onSyncForms');
    menu.addItem('Generate schedule', 'onGenerateSchedule');
    menu.addItem('Recalculate attendance', 'onRecalculateAttendance');
    if (ARC_APP_DEBUG_MODE) {
        menu.addItem('Debug: test client API', 'debugClientApiTest');
        menu.addItem('Debug: reset all tables', 'debugResetEverything');
        menu.addItem('Debug: reset all small tables', 'debugResetAllSmallTables');
        menu.addItem('Debug: rebuild all headers', 'debugHeaders');
        menu.addItem('Debug: rewrite all tables', 'debugRewriteEverything');
    }
    menu.addToUi();
}