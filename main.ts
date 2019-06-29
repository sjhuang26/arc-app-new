/*

SERVER TYPES (COPIED FROM CLIENT)

*/

export type ServerResponse<T> = {
    error: boolean;
    val: T;
    message: string;
};

export enum AskStatus {
    LOADING = 'LOADING',
    LOADED = 'LOADED',
    ERROR = 'ERROR'
}

export type Ask<T> = AskLoading | AskFinished<T>;

export type AskLoading = { status: AskStatus.LOADING };
export type AskFinished<T> = AskLoaded<T> | AskError;

export type AskError = {
    status: AskStatus.ERROR;
    message: string;
};

export type AskLoaded<T> = {
    status: AskStatus.LOADED;
    val: T;
};


/*

UTILITIES

*/

export function stringifyError(error: any): string {
    if (error instanceof Error) {
        return JSON.stringify(error, Object.getOwnPropertyNames(error));
    }
    if (typeof error === 'object') {
        return JSON.stringify(error);
    }
    return String(error);
}

export interface To<T> {
    [key: string]: T;
}

export abstract class Field {
    name: string;
    abstract parse(x: any): any
    abstract serialize(x: any): any

    constructor(name: string) {
        this.name = name;
    }
}

export class BooleanField extends Field {
    constructor(name: string) {
        super(name);
    }
    parse(x: any) {
        return x === 'true' ? 'true' : 'false';
    }
    serialize(x: any) {
        return x ? 'true' : 'false';
    }
}

export class NumberField extends Field {
    constructor(name: string) {
        super(name);
    }
    parse(x: any) {
        return x;
    }
    serialize(x: any) {
        return x;
    }
}

export class StringField extends Field {
    constructor(name: string) {
        super(name);
    }
    parse(x: any) {
        return x;
    }
    serialize(x: any) {
        return x;
    }
}

export class DateField extends Field {
    constructor(name: string) {
        super(name);
    }
    parse(x: any) {
        return x;
    }
    serialize(x: any) {
        return x;
    }
}

export class IdField extends Field {
    constructor() {
        // ID fields must be named "id"
        super('id');
    }
    parse(x: any) {
        return x;
    }
    serialize(x: any) {
        return x;
    }
}

export class JsonField extends Field {
    constructor(name: string) {
        super(name);
    }
    parse(x: any) {
        return JSON.parse(x);
    }
    serialize(x: any) {
        return JSON.stringify(x);
    }
}

export interface TableInfo {
    sheetName: string;
    fields: Field[];
}

export interface Rec {
    [field: string]: any;
}

/*

TABLE CLASS

*/

export class Table {
    idCounter: number = (new Date()).getTime();

    static tableInfoAll: To<TableInfo> = {
        tutors: {
            sheetName: '$tutors',
            fields: [
                new IdField(),
                new DateField('date'),
                new StringField('friendlyFullName'),
                new StringField('friendlyName'),
                new StringField('firstName'),
                new StringField('lastName'),
                new NumberField('grade'),
                new JsonField('mods'),
                new JsonField('modsPref'),
                new StringField('subjectList')
            ]
        },
        learners: {
            sheetName: '$learners',
            fields: [
                new IdField(),
                new DateField('date'),
                new StringField('friendlyFullName'),
                new StringField('friendlyName'),
                new StringField('firstName'),
                new StringField('lastName'),
                new NumberField('grade'),
            ]
        },
        requests: {
            sheetName: '$requests',
            fields: [
                new IdField(),
                new DateField('date'),
                new NumberField('learner'),
                new JsonField('mods'),
                new StringField('subject')
            ]
        },
        requestSubmissions: {
            sheetName: '$request-submissions',
            fields: [
                new IdField(),
                new DateField('date'),
                new StringField('friendlyFullName'),
                new StringField('friendlyName'),
                new StringField('firstName'),
                new StringField('lastName'),
                new NumberField('grade'),
                new JsonField('mods'),
                new StringField('subject')
            ]
        },
        bookings: {
            sheetName: '$bookings',
            fields: [
                new IdField(),
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
                new IdField(),
                new DateField('date'),
                new NumberField('learner'),
                new NumberField('tutor'),
                new StringField('subject'),
                new NumberField('mod'),
                new StringField('status')
            ]
        }
    };

    name: string;
    tableInfo: TableInfo;
    sheet: GoogleAppsScript.Spreadsheet.Sheet;

    constructor(name: string) {
        this.name = name;
        this.tableInfo = Table.tableInfoAll[name];
        if (this.tableInfo === undefined) {
            throw new Error(`table ${name} not found in tableInfoAll`);
        }
        this.sheet = SpreadsheetApp.getActive().getSheetByName(this.tableInfo.sheetName);
    }

    retrieveAllRecords() {
        const raw = this.sheet.getDataRange().getValues();
        const res = {};
        for (let i = 1; i < raw.length; ++i) {
            const rec = this.parseRecord(raw[i]);
            res[rec.id] = rec;
        }
        return res;
    }

    parseRecord(raw: any[]): Rec {
        const rec: Rec = {};
        for (let i = 0; i < raw.length; ++i) {
            const field = this.tableInfo.fields[i];
            rec[field.name] = field.parse(raw[i]);
        }
        return rec;
    }

    serializeRecord(record: Rec): any[] {
        return this.tableInfo.fields.map(field => field.serialize(record[field.name]));
    }

    addRecord(record: Rec) {
        record.id = this.createNewId();
        this.sheet.appendRow(this.serializeRecord(record));
    }

    updateRecord(editedRecord: Rec) {
        const idColumn: any[][] = this.sheet.getRange(1, 1, this.sheet.getLastRow()).getValues();
        let matchingRow = -1;
        for (let i = 0; i < idColumn.length; ++i) {
            const v: number = i[0]; // only one column in the matrix
            if (v === editedRecord.id) {
                if (matchingRow !== -1) {
                    throw new Error(`duplicate primary key ${v} in table ${this.name}`);
                }
                matchingRow = v;
            }
        }
        if (matchingRow == -1) {
            throw new Error(`primary key ${editedRecord.id} not found in table ${this.name}`);
        }
        this.sheet.getRange(matchingRow, 1, 1, this.sheet.getLastColumn()).setValues(this.serializeRecord(editedRecord));
    }

    createNewId() {
        const time = (new Date()).getTime();
        if (time <= this.idCounter) {
            ++this.idCounter;
        } else {
            this.idCounter = time;
        }
        return this.idCounter;
    }

    processClientAsk(args: any[]): ServerResponse<any> {
        if (args[0] === 'retrieveAll') {
            return null;
        }
        if (args[0] === 'update') {
            this.updateRecord(args[1]);
            onClientNotification(['update', this.name, args[1]]);
            return null;
        }
        if (args[0] === 'create') {
            if (args[1].date === -1) {
                args[1].date = Date.now();
            }
            if (args[1].id === -1) {
                args[1].id = this.createNewId();
            }
            this.addRecord(args[1]);
            onClientNotification(['create', this.name, args[1]]);
            return args[1];
        }
        if (args[0] === 'delete') {
            // TODO
            // this.deleteRecord(args[1]);
            onClientNotification(['delete', this.name, args[1]]);
            return null;
        }
        throw new Error('args not matched');
    }
}


/*

MAIN EVENTS

*/

function doGet() {
    return HtmlService.createHtmlOutputFromFile('index');
}

function onClientAsk(args: any[]): ServerResponse<any> {
    function processClientAsk(path: any[]) {
        if (path[0] == 'tutors') {
            path.shift();
            return new Table('tutors').processClientAsk(path);
        }
        if (path[0] == 'learners') {
            path.shift();
            return new Table('learners').processClientAsk(path);
        }
        if (path[0] == 'requests') {
            path.shift();
            return new Table('requests').processClientAsk(path);
        }
        if (path[0] == 'requestSubmissions') {
            path.shift();
            return new Table('requestSubmissions').processClientAsk(path);
        }
        if (path[0] == 'bookings') {
            path.shift();
            return new Table('bookings').processClientAsk(path);
        }
        if (path[0] == 'matchings') {
            path.shift();
            return new Table('matchings').processClientAsk(path);
        }
    }
    try {
        return {
            error: false,
            val: processClientAsk(args),
            message: null
        };
    } catch (err) {
        return {
            error: true,
            val: null,
            message: stringifyError(err)
        };
    }
}

function onClientNotification(args: any[]): void {
    // TODO: notify client
}

function onTest() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt('Enter API endpoint path');
    ui.alert(JSON.stringify(onClientAsk(JSON.parse(response.getResponseText()))));
}

function onOpen(_e: any) {
    SpreadsheetApp.getUi().createMenu('APP TEST')
      .addItem('API', 'onTest')
      .addToUi();
}