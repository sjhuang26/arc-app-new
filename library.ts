namespace U {
    export function isDate(x: any): boolean {
        // see https://stackoverflow.com/questions/643782 how-to-check-whether-an-object-is-a-date
        return Object.prototype.toString.call(x) === '[object Date]';
    }

    export function isId(x: any): boolean {
        return (typeof x == 'number') && (Math.floor(x) == x) && x >= 0;
    }
}

interface To<T> {
    [key: string]: T;
}

abstract class Field {
    name: string;
    abstract parse(x: any): any
    abstract serialize(x: any): any

    constructor(name: string) {
        this.name = name;
    }

    failTypeValidation(value) {
        throw new Error(`value "${value}" failed the type validation for field name "${this.name}"`);
    }
}

class BooleanField extends Field {
    constructor(name: string) {
        super(name);
    }
    parse(x: any) {
        if (typeof x !== 'boolean') this.failTypeValidation(x);
        return x;
    }
    serialize(x: any) {
        if (typeof x !== 'boolean') this.failTypeValidation(x);
        return x;
    }
}

class NumberField extends Field {
    constructor(name: string) {
        super(name);
    }
    parse(x: any) {
        if (typeof x !== 'number') this.failTypeValidation(x);
        return x;
    }
    serialize(x: any) {
        if (typeof x !== 'number') this.failTypeValidation(x);
        return x;
    }
}

class StringField extends Field {
    constructor(name: string) {
        super(name);
    }
    parse(x: any) {
        if (typeof x !== 'string') this.failTypeValidation(x);
        return x;
    }
    serialize(x: any) {
        if (typeof x !== 'string') this.failTypeValidation(x);
        return x;
    }
}

class DateField extends Field {
    constructor(name: string) {
        super(name);
    }
    parse(x: any) {
        if (!U.isDate(x)) this.failTypeValidation(x);
        return x;
    }
    serialize(x: any) {
        if (!U.isDate(x)) this.failTypeValidation(x);
        return x;
    }
}

class IdField extends Field {
    constructor() {
        // ID fields must be named "id"
        super('id');
    }
    parse(x: any) {
        if (!U.isId(x)) this.failTypeValidation(x);
        return x;
    }
    serialize(x: any) {
        if (!U.isId(x)) this.failTypeValidation(x);
        return x;
    }
}

class RefField extends Field {
    constructor(name: string) {
        super(name);
    }
    parse(x: any) {
        if (!U.isId(x)) this.failTypeValidation(x);
        return x;
    }
    serialize(x: any) {
        if (!U.isId(x)) this.failTypeValidation(x);
        return x;
    }
}

class EnumField extends Field {
    possibilities: string[];
    constructor(name: string, possibilities: string[]) {
        super(name);
        this.possibilities = possibilities;
    }
    validateEnum(x: string) {
        for (const y of this.possibilities) {
            if (x == y) return true;
        }
        return false;
    }

    parse(x: any) {
        if (!this.validateEnum(x)) this.failTypeValidation(x);
        return x;
    }
    serialize(x: any) {
        if (!this.validateEnum(x)) this.failTypeValidation(x);
        return x;
    }
}

class JsonField extends Field {
    constructor(name: string) {
        super(name);
    }
    parse(x: any) {
        if (typeof x !== 'string') this.failTypeValidation(x);
        return JSON.parse(x);
    }
    serialize(x: any) {
        return JSON.stringify(x);
    }
}

interface TableInfo {
    sheetName: string;
    fields: Field[];
}

interface Rec {
    [field: string]: any;
}

class Table {
    static idCounter = (new Date()).getTime();

    static tableInfoAll: To<TableInfo> = {
        tutors: {
            sheetName: '$tutors',
            fields: [
                new IdField(),
                new StringField('name'),
                new NumberField('grade'),
                new JsonField('mods'),
                new EnumField('status', ['unchecked', 'checked'])
            ]
        },
        learners: {
            sheetName: '$learners',
            fields: [
                new IdField(),
                new StringField('name'),
                new NumberField('grade'),
                new JsonField('mods'),
                new EnumField('status', ['unchecked', 'checked'])
            ]
        },
        requests: {
            sheetName: '$requests',
            fields: [
                new IdField(),
                new RefField('learner'),
                new EnumField('status', ['unchecked', 'inProgress', 'booked'])
            ]
        },
        bookings: {
            sheetName: '$bookings',
            fields: [
                new IdField(),
                new RefField('request'),
                new RefField('tutor'),
                new RefField('mod')
            ]
        },
        matchings: {
            sheetName: '$matchings',
            fields: [
                new IdField(),
                new RefField('learner'),
                new RefField('tutor'),
                new RefField('mod')
            ]
        }
    };

    name: string;

    tableInfo: TableInfo;
    sheet: GoogleAppsScript.Spreadsheet.Sheet;

    records: Rec[];

    idRecordIndex: To<Rec>;

    constructor(name: string) {
        this.name = name;
        this.tableInfo = Table.tableInfoAll[name];
        if (this.tableInfo === undefined) {
            throw new Error(`table ${name} not found in tableInfoAll`);
        }
        this.sheet = SpreadsheetApp.getActive().getSheetByName(this.tableInfo.sheetName);
        this.loadRecordsFromSheet();
    }

    loadRecordsFromSheet() {
        const raw = this.sheet.getDataRange().getValues();
        const res = [];
        for (let i = 1; i < raw.length; ++i) {
            res.push(this.parseRecord(raw[i]));
        }
        this.records = res;

        this.rebuildIdRecordIndex();
    }

    rebuildIdRecordIndex() {
        const r = {};
        for (const x of this.records) {
            if (r[x.id] !== undefined) {
                throw new Error(`primary key repeated in table ${this.name}`);
            }
            r[x.id] = x;
        }
        this.idRecordIndex = r;
    }

    parseRecord(raw: any[]): Rec {
        const rec: Rec = {};
        for (let i = 0; i < raw.length; ++i) {
            const field = this.tableInfo.fields[i];
            rec[field.name] = field.parse(raw[i]);
        }
        return rec;
    }

    find(id: number): Rec {
        const r = this.search(id);
        if (r === null) {
            throw new Error(`cannot find id ${id} in table ${this.name}`);
        }
        return r;
    }

    search(id: number): Rec {
        const r = this.idRecordIndex[id];
        Logger.log(JSON.stringify(this.idRecordIndex));
        return r === undefined ? null : r;
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

    updateRecordField(recordId: number, fieldName: string, newValue: any) {
        this.idRecordIndex[recordId][fieldName] = newValue;
        this.updateRecord(this.idRecordIndex[recordId]);
    }

    queryRecords(predicate: (record: Rec) => boolean) {
        return this.records.filter(record => predicate(record));
    }

    createNewId() {
        const time = (new Date()).getTime();
        if (time <= Table.idCounter) {
            ++Table.idCounter;
        } else {
            Table.idCounter = time;
            return time;
        }
    }

    apiEndpoint(path: any[]) {
        if (path.length == 0) {
            return this.records;
        }
        if (path.length == 2 && typeof path[0] == 'number' && path[1] == 'delete') {
            // TODO this.deleteRecord(path[0]);
        }
        if (path.length == 2 && path[0] == 'add' && typeof path[1] == 'object') {
            // the API call is assumed to have the correct types (no conversions) so it is added directly
            // TODO: what about Date objects?
            this.addRecord(path[1]);
            return true;
        }
        if (path.length == 1 && typeof path[0] == 'number') {
            return this.search(path[0]);
        }
        if (path.length == 2 && path[0] == 'status' && typeof path[1] == 'string') {
            return this.queryRecords(record => record.status == path[1]);
        }
        throw new Error(`api endpoint ${this.name} >>> ${JSON.stringify(path)} not matched`);
    }
}
