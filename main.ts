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

// This function converts mod numbers (ie. 11) into A-B-day strings (ie. 1B).
export function stringifyMod(mod: number) {
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
    if (typeof error === 'object') {
        return JSON.stringify(error);
    }
    return String(error);
}

interface To<T> {
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

// The handling of dates is special ...
// Parsing dates returns a number, which is used internally
// Serializing dates takes an internally-used number, which is written to the spreadsheet as a date
class DateField extends Field {
    constructor(name: string = null) {
        super(name);
    }
    parse(x: any) {
        return new Date(x).getTime();
    }
    serialize(x: any) {
        return new Date(x);
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

    static tableInfoAll: To<TableInfo> = {
        tutors: {
            sheetName: '$tutors',
            fields: [
                new NumberField('id'),
                new DateField('date'),
                ...Table.makeBasicStudentConfig(),
                new JsonField('mods'),
                new JsonField('modsPref'),
                new StringField('subjectList')
            ]
        },
        learners: {
            sheetName: '$learners',
            fields: [
                new NumberField('id'),
                new DateField('date'),
                ...Table.makeBasicStudentConfig()
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

    constructor(name: string) {
        this.name = name;
        this.tableInfo = Table.tableInfoAll[name];
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
        }
    }

    rebuildSheetHeadersIfNeeded() {
        if (!this.isForm) {
            const col = this.sheet.getLastColumn();
            this.sheet.getRange(1, 1, 1, col === 0 ? 1 : col).clearContent();
            this.sheet.getRange(1, 1, 1, this.tableInfo.fields.length).setValues([this.tableInfo.fields.map(field => field.name)]);
        }
    }

    resetEntireSheet() {
        if (!this.isForm) {
            this.sheet.getDataRange().clearContent();
        }
        this.rebuildSheetHeadersIfNeeded();
    }

    retrieveAllRecords(): RecCollection {
        if (ARC_APP_DEBUG_MODE) {
            this.rebuildSheetHeadersIfNeeded();
        }
        if (this.sheet.getLastColumn() !== this.tableInfo.fields.length) {
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

    updateRecord(editedRecord: Rec): void {
        this.sheet.getRange(this.getRowById(editedRecord.id), 1, 1, this.sheet.getLastColumn()).setValues([this.serializeRecord(editedRecord)]);
    }

    deleteRecord(id: number): void {
        if (this.isForm) {
            throw new Error('cannot delete records from a form');
        }
        this.sheet.deleteRow(this.getRowById(id));
    }

    createNewId(): number {
        if (this.isForm) {
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


/*

MAIN EVENTS

*/

const tableMap: TableMap = {
    ...tableMapBuild('tutors'),
    ...tableMapBuild('learners'),
    ...tableMapBuild('requests'),
    ...tableMapBuild('requestSubmissions'),
    ...tableMapBuild('bookings'),
    ...tableMapBuild('matchings'),
    ...tableMapBuild('requestForm'),
    ...tableMapBuild('specialRequestForm'),
    ...tableMapBuild('operationLog')
};

function doGet() {
    // TODO: fix this
    tableMap.operationLog().resetEntireSheet();

    return HtmlService.createHtmlOutputFromFile('index');
}

// This is a way to prevent constructing six table objects every time the script is called.
type TableMap = {
    [name: string]: () => Table;
}
function tableMapBuild(name: string) {
    return {
        [name]: () => new Table(name)
    };
}

const ARC_APP_DEBUG_MODE: boolean = true;

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

function onClientNotification(args: any[]): void {
    // we record the logs, and the client reads them every 20 seconds
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

function doFormSync(formTable: Table, actualTable: Table, formRecordToActualRecord: (formRecord: Rec) => Rec): number {
    const actualRecords = actualTable.retrieveAllRecords();
    const formRecords = formTable.retrieveAllRecords();

    let numOfThingsSynced = 0;

    // create an index of actualdata >> date.
    // Then iterate over all formdata and find the ones that are missing from the index.
    const index: { [date: string]: Rec } = {};
    for (const idKey of Object.getOwnPropertyNames(actualRecords)) {
        const record = actualRecords[idKey];
        if (record.status === 'unchecked') { // record status MUST be unchecked for it to count!
            const dateIndexKey = String(record.date);
            index[dateIndexKey] = record;
        }
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

function onSyncForms() {
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

    function processRequestFormRecord(r: Rec): Rec {
        // parsing mod data
        function parseCommas(d: string) {
            return d
                .split(',')
                .map(x => x.trim())
                .filter(x => x !== '' && x !== 'None')
                .map(x => parseInt(x));
        }
        const mA15: number[] = parseCommas(r.modDataA1To5);
        const mB15: number[] = parseCommas(r.modDataB1To5).map(x => x + 10);
        const mA60: number[] = parseCommas(r.modDataA6To10);
        const mB60: number[] = parseCommas(r.modDataB6To10).map(x => x + 10);

        return {
            id: -1,
            date: r.date, // the date MUST be the date from the form
            firstName: r.firstName,
            lastName: r.lastName,
            friendlyName: r.friendlyName,
            friendlyFullName: r.friendlyFullName,
            studentId: r.studentId,
            grade: parseGrade(r.grade),
            subject: r.subject,
            mods: mA15.concat(mA60).concat(mB15).concat(mB60),
            email: r.email,
            phone: r.phone,
            contactPref: parseContactPref(r.contactPref),
            specialRoom: '',

            status: 'unchecked'
        };
    }
    function processSpecialRequestFormRecord(r: Rec): Rec {
        function parseModInfo(abDay: string, mod1To10: number): number {
            // polyfill
            // https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/String/startsWith#Polyfill
            function startsWith(search, pos) {
                pos = !pos || pos < 0 ? 0 : +pos;
                return this.substring(pos, pos + search.length) === search;
            }

            if (abDay.toLowerCase().charAt(0) === 'a') {
                return mod1To10;
            }
            if (abDay.toLowerCase().charAt(0) === 'b') {
                return mod1To10 + 10;
            }
            throw new Error(`${String(abDay)} does not start with A or B`);
        }
        return {
            id: -1,
            date: r.date, // the date MUST be the date from the form
            firstName: r.firstName,
            lastName: r.lastName,
            friendlyName: r.friendlyName,
            friendlyFullName: r.friendlyFullName,
            studentId: r.studentId,
            grade: parseGrade(r.grade),
            subject: r.subject,
            specialRoom: r.specialRoom,

            email: '',
            phone: '',
            contactPref: 'either',
            mods: [parseModInfo(r.abDay, r.mod1To10)],

            status: 'unchecked'
        };
    }

    try {
        let numOfThingsSynced = 0;
        numOfThingsSynced += doFormSync(tableMap.requestForm(), tableMap.requestSubmissions(), processRequestFormRecord);
        numOfThingsSynced += doFormSync(tableMap.specialRequestForm(), tableMap.requestSubmissions(), processSpecialRequestFormRecord);
        SpreadsheetApp.getUi().alert(`Finished sync! ${numOfThingsSynced} new tutor requests found.`);
    } catch (err) {
        Logger.log(stringifyError(err));
        throw err;
    }
}

// This generates a schedule that is written to a new sheet.
function onGenerateSchedule() {
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

        // Figure out all matchings that are finalized, indexed by tutor
        const index: { [tutorId: string]: Rec[] } = {};
        for (const tutorKey of Object.getOwnPropertyNames(tutors)) {
            const r = tutors[tutorKey];
            const idx = r.id;
            index[idx] = [];
        }
        for (const idKey of Object.getOwnPropertyNames(matchings)) {
            const r = matchings[idKey];
            const idx = r.tutor;
            if (r.status === 'finalized') {
                index[idx].push(r);
            }
        }

        // Sort the index, first reformatting it as a 2D matrix
        const tutorAndRecordsArray: [ string, Rec[] ][] = [];
        for (const i of Object.getOwnPropertyNames(index)) {
            tutorAndRecordsArray.push([tutors[i].friendlyFullName, index[i]]);
        }
        tutorAndRecordsArray.sort((a, b) => {
            const x = a[0].toLowerCase();
            const y = b[0].toLowerCase();
            if (x < y) return -1;
            if (x > y) return 1;
            return 0;
        });
        
        // Print!
        for (const [ tutorName, matchings ] of tutorAndRecordsArray) {
            for (const matching of matchings) {
                sheet.appendRow([tutorName, stringifyMod(matching.mod), learners[matching.learner].friendlyFullName]);
            }
            if (matchings.length === 0) {
                sheet.appendRow([tutorName, 'no one']);
            }
        }
        sheet.appendRow([`That's all!`]);
    } catch (err) {
        Logger.log(stringifyError(err));
        throw err;
    }
}


function onOpen(_ev: any) {
    const menu = SpreadsheetApp.getUi().createMenu('ARC APP');
    menu.addItem('Sync data from request forms', 'onSyncForms');
    menu.addItem('Generate schedule', 'onGenerateSchedule');
    if (ARC_APP_DEBUG_MODE) {
        menu.addItem('Debug: test client API', 'debugClientApiTest');
        menu.addItem('Debug: reset all tables', 'debugResetEverything');
        menu.addItem('Debug: reset all small tables', 'debugResetAllSmallTables');
        menu.addItem('Debug: rebuild all headers', 'debugHeaders');
    }
    menu.addToUi();
}