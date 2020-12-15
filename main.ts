/*

GLOBAL SETTINGS

*/

const ARC_APP_DEBUG_MODE: boolean = true

/*

Shared with the other file

*/

export enum ModStatus {
  UNFREE,
  FREE,
  DROP_IN,
  BOOKED,
  MATCHED,
  FREE_PREF,
  DROP_IN_PREF,
}

export enum SchedulingReference {
  BOOKING,
  MATCHING,
}

export function schedulingTutorIndex(
  tutorRecords: TableInfo.RecCollection<'tutors'>,
  bookingRecords: TableInfo.RecCollection<'bookings'>,
  matchingRecords: TableInfo.RecCollection<'matchings'>
) {
  const tutorIndex: {
    [id: number]: {
      id: number
      modStatus: ModStatus[]
      refs: [SchedulingReference, number][]
      isInMod: boolean
    }
  } = {}
  for (const tutor of Object.values(tutorRecords)) {
    tutorIndex[tutor.id] = {
      id: tutor.id,
      modStatus: Array(20).fill(ModStatus.UNFREE),
      refs: [],
      isInMod: false,
    }
    const index = tutorIndex[tutor.id]
    const st = index.modStatus
    for (const mod of tutor.mods) {
      st[mod - 1] = ModStatus.FREE
      index.isInMod = true
    }
    for (const mod of tutor.dropInMods) {
      st[mod - 1] = ModStatus.DROP_IN
      index.isInMod = true
    }
    for (const mod of tutor.modsPref) {
      switch (st[mod - 1]) {
        case ModStatus.FREE:
        case ModStatus.UNFREE:
          st[mod - 1] = ModStatus.FREE_PREF
          break
        case ModStatus.DROP_IN:
          st[mod - 1] = ModStatus.DROP_IN_PREF
          break
        default:
          throw new Error()
      }
    }
  }
  for (const booking of Object.values(bookingRecords)) {
    if (booking.status !== "ignore" && booking.status !== "rejected") {
      if (booking.mod !== undefined) {
        tutorIndex[booking.tutor].isInMod = true
        tutorIndex[booking.tutor].modStatus[booking.mod - 1] = ModStatus.BOOKED
      }
      tutorIndex[booking.tutor].refs.push([
        SchedulingReference.BOOKING,
        booking.id,
      ])
    }
  }
  for (const matching of Object.values(matchingRecords)) {
    if (matching.mod !== undefined) {
      tutorIndex[matching.tutor].isInMod = true
      tutorIndex[matching.tutor].modStatus[matching.mod - 1] = ModStatus.MATCHED
      tutorIndex[matching.tutor].refs.push([
        SchedulingReference.MATCHING,
        matching.id,
      ])
    }
  }
  return tutorIndex
}

/*

/END

*/

/*

UTILITIES

*/

function roundDownToDay(utcTime: number) {
  const tzOffset = new Date(utcTime).getTimezoneOffset() * 60 * 1000
  return utcTime - (utcTime % (24 * 60 * 60 * 1000)) + tzOffset
}

function formatAttendanceModDataString(mod: number, minutes: number) {
  return `${Number(mod)} ${Number(minutes)}`
}

function onlyKeepUnique<T>(arr: T[]): T[] {
  const x = {}
  for (let i = 0; i < arr.length; ++i) {
    x[JSON.stringify(arr[i])] = arr[i]
  }
  const result = []
  for (const key of Object.getOwnPropertyNames(x)) {
    result.push(x[key])
  }
  return result
}

// polyfill of the typical Object.values()
function Object_values<T>(o: ObjectMap<T>): T[] {
  const result: T[] = []
  for (const i of Object.getOwnPropertyNames(o)) {
    result.push(o[i])
  }
  return result
}

function recordCollectionToArray<T extends TableInfo.KeyType>(r: TableInfo.RecCollection<T>): TableInfo.Rec<T>[] {
  const x = []
  for (const i of Object.getOwnPropertyNames(r)) {
    x.push(r[i])
  }
  return x
}

// This function converts mod numbers (ie. 11) into A-B-day strings (ie. 1B).
function stringifyMod(mod: number) {
  if (1 <= mod && mod <= 10) {
    return String(mod) + "A"
  } else if (11 <= mod && mod <= 20) {
    return String(mod - 10) + "B"
  }
  throw new Error(`mod ${mod} isn't serializable`)
}

function stringifyError(error: any): string {
  if (error instanceof Error) {
    return JSON.stringify(error, Object.getOwnPropertyNames(error))
  }
  try {
    return JSON.stringify(error)
  } catch (unusedError) {
    return String(error)
  }
}

type ObjectMap<T> = {
  [key: string]: T
}

enum FieldType {
  BOOLEAN,
  NUMBER,
  STRING,
  DATE,
  JSON,
}

/*

Spec: tables

MUST have IO carefully linked to
  Google Sheets API
  processClientAsk
MUST have ability to cache tables
new table
what you need to perform the operations
  IO (Google Sheets API)
  key
  sheetName
MUST have serializeRecord
  JS object into an array of raw cell data.
a way to configure tables from start to finish
  TABLE_INFO_MASTER_ARRAY
  TODO: special format for forms

TODO: a way to lock tables from write
  isWriteAllowed
*/

namespace TableInfo {
  export type KeyType = keyof TableInfo.InfoType;

  type LiteralString<T> = T extends string ? (string extends T ? never : T) : never;
  function StringField<T>(key: LiteralString<T>) {
    return [FieldType.STRING, key] as const;
  }
  function NumberField<T>(key: LiteralString<T>) {
    return [FieldType.NUMBER, key] as const;
  }
  function DateField<T>(key: LiteralString<T>) {
    return [FieldType.DATE, key] as const
  }
  function BooleanField<T>(key: LiteralString<T>) {
    return [FieldType.BOOLEAN, key] as const
  }
  function JsonField<T>(key: LiteralString<T>) {
    return [FieldType.JSON, key] as const
  }

  const arrayOfIdAndDate = [NumberField('id'), DateField('date')] as const;
  const arrayOfDate = [DateField('date')] as const;

  export type ParseRecordType<T extends KeyType> = {
    [U in (InfoType[T]['fields'][number]) as U[1]]: ParseFieldType<U[0]>
  }

  /*

  // PRACTICE CODE
  // https://github.com/Microsoft/TypeScript/issues/27272

  type Generic<T> = { something: T }
  type Union = {a: 1} | {b: 1} | {c: 1}
  type Fancy<T> = T extends T ? Generic<T> : never;
  type UnionOfGeneric = Fancy<Union>
  
  */



  export type Rec<T extends KeyType> = ParseRecordType<T>
  export type RecCollection<T extends KeyType> = {
    [id: number]: Rec<T>
  }

  export type ParseFieldType<T extends FieldType> =
    FieldType.STRING extends T ? string
    : FieldType.NUMBER extends T ? number
    : FieldType.DATE extends T ? number
    : FieldType.BOOLEAN extends T ? boolean
    : FieldType.JSON extends T ? any
    : never

  export type FieldsType = InfoType[KeyType]['fields'];

  const basicStudentConfig = [
    StringField('friendlyFullName'),
    StringField('firstName'),
    StringField('lastName'),
    NumberField('grade'),
    NumberField('studentId'),
    StringField('email'),
    StringField('phone'),
    StringField('contactPref'),
    StringField('homeroom'),
    StringField('homeroomTeacher'),
    StringField('attendanceAnnotation'),
    JsonField("attendance"),
  ] as const;

  // This type has been problematic in the clasp compiler so we have extracted it from the method.
  type ReformatType = { readonly [U in ((typeof tableInfoData)[number]) as U[0]]: {
    readonly key: U[0],
    readonly sheetName: U[1],
    readonly fields: (true extends U[3] ? [...(typeof arrayOfDate), ...U[2]] : [...(typeof arrayOfIdAndDate), ...U[2]]),
    readonly isForm: true extends U[3] ? true : false
  } }

  function reformat(x: typeof tableInfoData): ReformatType {
    return Object.fromEntries(x.map((y) => [y[0], {
      key: y[0],
      sheetName: y[1],
      fields: y[3] == true ? [...arrayOfDate, ...y[2]] : [...arrayOfIdAndDate, ...y[2]],
      isForm: !!y[3]
    }])) as any;
  }

  const tableInfoData = [
    [
      "studentInfo",
      "$student-info",
      basicStudentConfig
    ],
    [
      "tutors",
      "$tutors",
      [
        NumberField('studentInfo'),
        JsonField("mods"),
        JsonField("modsPref"),
        StringField("subjectList"),
        JsonField("dropInMods"),
        StringField("afterSchoolAvailability"),
        NumberField("additionalHours"),
      ],
    ],
    ["learners", "$learners", [NumberField("studentInfo")]],
    [
      "requests",
      "$requests",
      [
        NumberField("learner"),
        JsonField("mod"),
        StringField("subject"),
        StringField("annotation"),
        NumberField("step"),
        JsonField("chosenBookings"),
      ],
    ],
    [
      "requestSubmissions",
      "$request-submissions",
      [
        ...basicStudentConfig,
        JsonField("mods"),
        StringField("subject"),
        BooleanField("isSpecial"),
        StringField("annotation"),
        StringField("status"),
      ],
    ],
    [
      "bookings",
      "$bookings",
      [
        NumberField("request"),
        NumberField("tutor"),
        NumberField("mod"),
        StringField("status"),
      ],
    ],
    [
      "matchings",
      "$matchings",
      [
        NumberField("learner"),
        NumberField("tutor"),
        StringField("subject"),
        NumberField("mod"),
        StringField("annotation"),
      ],
    ],
    [
      "attendanceLog",
      "$attendance-log",
      [
        NumberField("id"),
        DateField("date"),
        DateField("dateOfAttendance"), // rounded to nearest day
        StringField("validity"), // filled with an error message if the form entry was typed wrong
        NumberField("mod"),

        // one of these ID fields will be left as -1 (blank).
        NumberField("tutor"),
        NumberField("learner"),
        NumberField("minutesForTutor"),
        NumberField("minutesForLearner"),
        StringField("presenceForTutor"),
        StringField("presenceForLearner"),

        // used for reset purposes
        StringField("markForReset"),
      ],
    ],
    [
      "attendanceDays",
      "$attendance-days",
      [
        NumberField("id"),
        DateField("date"),
        DateField("dateOfAttendance"),
        StringField("abDay"),
        // we add a functionality to reset a day's attendance absences
        StringField("status"), // upcoming, finished, finalized, unreset, reset
      ],
    ],
    [
      'requestForm',
      '$request-form',
      [
        StringField('firstName'),
        StringField('lastName'),
        StringField('friendlyFullName'),
        NumberField('studentId'),
        StringField('grade'),
        StringField('subject'),
        StringField('modDataA1To5'),
        StringField('modDataB1To5'),
        StringField('modDataA6To10'),
        StringField('modDataB6To10'),
        StringField('homeroom'),
        StringField('homeroomTeacher'),
        StringField('email'),
        StringField('phone'),
        StringField('contactPref'),
        StringField('iceCreamQuestion')
      ],
      true
    ],
    [
      'specialRequestForm',
      '$special-request-form',
      [
        StringField('teacherName'),
        StringField('teacherEmail'),
        StringField('numLearners'),
        StringField('subject'),
        StringField('tutoringDateInformation'),
        NumberField('room'),
        StringField('abDay'),
        NumberField('mod1To10'),
        StringField('studentNames'),
        StringField('additionalInformation')
      ],
      true
    ],
    [
      'attendanceForm',
      '$attendance-form',
      [
        DateField('dateOfAttendance'), // optional in the form
        NumberField('mod1To10'),
        NumberField('studentId'),
        StringField('presence')
      ],
      true,
    ],
    [

      'tutorRegistrationForm',
      '$tutor-registration-form',
      [
        StringField('firstName'),
        StringField('lastName'),
        StringField('friendlyName'),
        StringField('friendlyFullName'),
        NumberField('studentId'),
        StringField('grade'),
        StringField('email'),
        StringField('phone'),
        StringField('contactPref'),
        StringField('homeroom'),
        StringField('homeroomTeacher'),
        StringField('afterSchoolAvailability'),
        StringField('modDataA1To5'),
        StringField('modDataB1To5'),
        StringField('modDataA6To10'),
        StringField('modDataB6To10'),
        StringField('modDataPrefA1To5'),
        StringField('modDataPrefB1To5'),
        StringField('modDataPrefA6To10'),
        StringField('modDataPrefB6To10'),
        StringField('subjects0'),
        StringField('subjects1'),
        StringField('subjects2'),
        StringField('subjects3'),
        StringField('subjects4'),
        StringField('subjects5'),
        StringField('subjects6'),
        StringField('subjects7'),
        StringField('iceCreamQuestion'),
        NumberField('numberGuessQuestion')
      ],
      true
    ],
    [
      // this table is merged into the JSON of tutor.fields.attendance
      // the table will get quite large, so we will hand-archive it from time to time
      // ASSUMPTION: (thus...) the table DOESN'T contain all of the attendance data. Some of it
      // will be archived somewhere else. The JSON will be merged with the attendance log.
      'attendanceLog',
      '$attendance-log',
      [
        DateField('dateOfAttendance'), // rounded to nearest day
        StringField('validity'), // filled with an error message if the form entry was typed wrong
        NumberField('mod'),

        // one of these ID fields will be left as -1 (blank).
        NumberField('tutor'),
        NumberField('learner'),
        NumberField('minutesForTutor'),
        NumberField('minutesForLearner'),
        StringField('presenceForTutor'),
        StringField('presenceForLearner'),

        // used for reset purposes
        StringField('markForReset')
      ]
    ],
    [
      'attendanceDays',
      '$attendance-days',
      [
        DateField('dateOfAttendance'),
        StringField('abDay'),

        // we add a functionality to reset a day's attendance absences
        StringField('status') // upcoming, finished, finalized, unreset, reset
      ],
    ],
    [
      'operationLog',
      '$operation-log',
      [
        JsonField('args')
      ]
    ]
  ] as const;

  export const info = reformat(tableInfoData);
  export type InfoType = typeof info;
  // You have to include the "T extends T" thing because this forces Typescript to consider EVERY ITEM IN THE UNION
  type FilterBetweenNonFormAndForm<T extends InfoType[keyof InfoType], U> = T extends T ? (U extends T['isForm'] ? T : never) : never;

  export type NonFormInfoType = FilterBetweenNonFormAndForm<InfoType[keyof InfoType], false>;
  export type FormInfoType = FilterBetweenNonFormAndForm<InfoType[keyof InfoType], true>;
}

/*

Spec: GlobalId

Date.now() is the number of milliseconds since the UNIX epoch and that is the global ID.

*/

namespace GlobalId {
  let globalId = -1

  export function getNextGlobalId() {
    const now = Date.now()
    if (now <= globalId) {
      ++globalId
    } else {
      globalId = now
    }
    return globalId
  }
}

/*

Spec: TableIOInfo

primarily for caching

*/

namespace TableIO {
  export function retrieveAllRecords<T extends TableInfo.KeyType>(key: T, sheet: GoogleAppsScript.Spreadsheet.Sheet): { [id: string]: TableInfo.ParseRecordType<T> } {
    type MakeUnion<U extends TableInfo.InfoType[TableInfo.KeyType]> = U extends U ? U : never;
    type MakeUnion2<U extends TableInfo.KeyType> = U extends U ? TableInfo.ParseRecordType<U> : never;

    // https://stackoverflow.com/questions/50870423/discriminated-union-of-generic-type
    // We use types to make info into a union so that the Discriminated Union works right

    const info: MakeUnion<TableInfo.InfoType[TableInfo.KeyType]> = TableInfo.info[key]
    if (info.isForm === true) {
      const raw = sheet.getDataRange().getValues()
      const res: { [id: string]: TableInfo.ParseRecordType<T> } = {}
      for (let i = 1; i < raw.length; ++i) {
        const rec: MakeUnion2<typeof info['key']> = parseRecord<typeof info['key']>(raw[i], info.key)
        res[String(rec.date)] = rec as TableInfo.ParseRecordType<T>
      }
      return res
    } else if (info.isForm === false) {
      const raw = sheet.getDataRange().getValues()
      const res: { [id: string]: TableInfo.ParseRecordType<T> } = {}
      for (let i = 1; i < raw.length; ++i) {
        const rec: MakeUnion2<typeof info['key']> = parseRecord<typeof info['key']>(raw[i], info.key)
        res[String(rec.id)] = rec as TableInfo.ParseRecordType<T>
      }
      return res
    } else {
      throw 'Error';
    }
  }

  export function parseRecord<T extends TableInfo.KeyType>(raw: any[], key: T): TableInfo.ParseRecordType<T> {
    const rec = {}

    const fields = TableInfo.info[key].fields;
    for (let i = 0; i < fields.length; ++i) {
      const field = fields[i]
      // this accounts for blanks in the last field
      rec[field[1]] = Parsing.parseField(
        raw[i] === undefined ? "" : raw[i],
        field[0]
      )
    }
    return rec as any;
  }
}

class Table<T extends TableInfo.KeyType> {
  static cache: { [U in TableInfo.KeyType]?: GoogleAppsScript.Spreadsheet.Sheet } = {};

  key: T
  sheetName: TableInfo.InfoType[T]['sheetName']
  sheet: GoogleAppsScript.Spreadsheet.Sheet
  isWriteAllowed: boolean
  fields: TableInfo.InfoType[T]['fields']

  constructor(key: T) {
    this.key = key
    this.sheetName = TableInfo.info[key].sheetName;
    // In the best case, you would use the ??= operator, but there appears to be a clasp compiler bug here. So here is a workaround.
    // use of ??= operator -- only if the left-hand side is undefined/null, assign it to the right-hand side
    if (Table.cache[key] === undefined) {
      Table.cache[key] = SpreadsheetApp.getActive().getSheetByName(
        this.sheetName);
    }
    this.sheet = Table.cache[key];
  }

  static readAllRecords<T extends TableInfo.KeyType>(key: T): TableInfo.RecCollection<T> {
    return new Table(key).readAllRecords()
  }

  readAllRecords() {
    return TableIO.retrieveAllRecords(
      this.key,
      this.sheet
    )
  }

  serializeRecord(record: TableInfo.Rec<T>): any[] {
    return (this.fields as any[]).map((field) =>
      Parsing.serializeField(record[field[1]], field[0])
    )
  }

  createRecord(record: TableInfo.Rec<T>): TableInfo.Rec<T> {
    this.checkWritePermission()
    if (record['date'] === -1) {
      record['date'] = Date.now()
    }
    if (record['id'] === -1) {
      record['id'] = GlobalId.getNextGlobalId()
    }
    this.sheet.appendRow(this.serializeRecord(record))
    return record
  }

  getRowById(id: number): number {
    // because the first row is headers, we ignore it and start from the second row
    const mat: any[][] = this.sheet
      .getRange(2, 1, this.sheet.getLastRow() - 1)
      .getValues()
    let rowNum = -1
    for (let i = 0; i < mat.length; ++i) {
      const cell: number = mat[i][0]
      if (typeof cell !== "number") {
        throw new Error(
          `id at location ${String(i)} is not a number in table ${String(
            this.key
          )}`
        )
      }
      if (cell === id) {
        if (rowNum !== -1) {
          throw new Error(
            `duplicate ID ${String(id)} in table ${String(this.key)}`
          )
        }
        rowNum = i + 2 // i = 0 <=> second row (rows are 1-indexed)
      }
    }
    if (rowNum == -1) {
      throw new Error(
        `ID ${String(id)} not found in table ${String(this.key)}`
      )
    }
    return rowNum
  }

  updateRecord(editedRecord: TableInfo.Rec<T>, rowNum?: number): void {
    this.checkWritePermission()
    if (rowNum === undefined) {
      rowNum = this.getRowById(editedRecord['id'])
    }
    this.sheet
      .getRange(rowNum, 1, 1, this.sheet.getLastColumn())
      .setValues([this.serializeRecord(editedRecord)])
  }

  updateAllRecords(editedRecords: TableInfo.Rec<T>[]): void {
    this.checkWritePermission()
    if (this.sheet.getLastRow() === 1) {
      return // the sheet is empty, and trying to select it will result in an error
    }
    // because the first row is headers, we ignore it and start from the second row
    const mat: any[][] = this.sheet
      .getRange(2, 1, this.sheet.getLastRow() - 1)
      .getValues()
    let idRowMap: ObjectMap<number> = {}
    for (let i = 0; i < mat.length; ++i) {
      idRowMap[String(mat[i][0])] = i + 2 // i = 0 <=> second row (rows are 1-indexed)
    }
    for (const r of editedRecords) {
      this.updateRecord(r, idRowMap[String(r['id'])])
    }
  }

  deleteRecord(id: number): void {
    this.checkWritePermission()
    this.sheet.deleteRow(this.getRowById(id))
  }

  rebuildSheetHeadersIfNeeded() {
    this.checkWritePermission()
    const col = this.sheet.getLastColumn()
    this.sheet.getRange(1, 1, 1, col === 0 ? 1 : col).clearContent()
    this.sheet
      .getRange(1, 1, 1, this.fields.length)
      .setValues([(this.fields as any[]).map((x) => x[1])])
  }

  checkWritePermission(): void {
    if (!this.isWriteAllowed) {
      throw new Error()
    }
  }
}

/*

IMPORTANT EVENT HANDLERS
(CODE THAT DOES ALL THE USER ACTIONS NECESSARY IN THE BACKEND)
(ALSO CODE THAT HANDLES SERVER-CLIENT INTERACTIONS)

*/

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("ARC App")
    .setFaviconUrl(
      "https://arc-app-frontend-server.netlify.com/dist/favicon.ico"
    )
}

/*

processClientAsk

Spec:

MUST call 'tryThis many times
  if none of them work, throw an error
  args[0]: top level
  args[1]: second level
  args[2]: third level
there are exceptions to 'tryThis
*/

function processClientAsk(args: any[]): any {
  function tryThis(arg: number, key: string, then: () => any) {
    if (args[arg] === key) return then()
  }
  tryThis(0, "command", () => {
    tryThis(1, "syncDataFromForms", onSyncForms)
    tryThis(1, "recalculateAttendance", onRecalculateAttendance)
    tryThis(1, "generateSchedule", onGenerateSchedule)
    tryThis(1, "syncDataFromForms", onSyncForms)
    tryThis(1, "retrieveMultiple", () => onRetrieveMultiple(args[2]))
  })
  // might throw
  const table = new Table(args[0])
  tryThis(1, "retrieveAll", () => Table.readAllRecords(args[2]))
  tryThis(1, "update", () => table.updateRecord(args[2]))
  tryThis(1, "create", () => table.createRecord(args[2]))
  tryThis(1, "delete", () => table.deleteRecord(args[2]))
  throw new Error("error-?")
}

// this is the MAIN ENTRYPOINT that the client uses to ask the server for data.
function onClientAsk(args: any[]): string {
  let returnValue = {
    error: true,
    val: null,
    message: "Mysterious error",
  }
  try {
    returnValue = {
      error: false,
      val: processClientAsk(args),
      message: null,
    }
  } catch (err) {
    returnValue = {
      error: true,
      val: null,
      message: stringifyError(err),
    }
  }
  // If you send a too-big object, Google Apps Script doesn't let you do it, and null is returned. But if you stringify it, you're fine.
  return JSON.stringify(returnValue)
}

function debugClientApiTest() {
  try {
    const ui = SpreadsheetApp.getUi()
    const response = ui.prompt("Enter args as JSON array")
    ui.alert(
      JSON.stringify(onClientAsk(JSON.parse(response.getResponseText())))
    )
  } catch (err) {
    Logger.log(stringifyError(err))
    throw err
  }
}

// This is a useful debug. It rewrites all the sheet headers to what the app thinks the sheet headers "should" be.
function debugHeaders() {
  try {
    // TODO!
    throw new Error()
  } catch (err) {
    Logger.log(stringifyError(err))
    throw err
  }
}

// This is a utility designed for onSyncForms().
// Syncs between the formTable and the actualTable that we want to associate with it.
// Basically, we use formRecordToActualRecord() to convert form records to actual records.
// Then the actual records go in the actualTable.
// But we only do this for form records that have dates that don't exist as IDs in actualTable.
// (Remember that a form date === a record ID.)
// There is NO DELETING RECORDS! No matter what!
function doFormSync<T extends TableInfo.KeyType, U extends TableInfo.KeyType>(
  formTable: Table<T>,
  actualTable: Table<U>,
  processFormRecord: (formRecord: TableInfo.Rec<T>) => void
): number {
  const actualRecords = actualTable.readAllRecords()
  const formRecords = formTable.readAllRecords()

  let numOfThingsSynced = 0

  // create an index of actualdata >> date.
  // Then iterate over all formdata and find the ones that are missing from the index.
  const index: { [date: string]: TableInfo.Rec<U> } = {}
  for (const idKey of Object.getOwnPropertyNames(actualRecords)) {
    const record: TableInfo.Rec<U> = actualRecords[idKey] as any
    const dateIndexKey = String(record['date'])
    index[dateIndexKey] = record
  }
  for (const idKey of Object.getOwnPropertyNames(formRecords)) {
    const record: TableInfo.Rec<T> = formRecords[idKey] as any
    const dateIndexKey = String(record['date'])
    if (index[dateIndexKey] === undefined) {
      processFormRecord(record)
      ++numOfThingsSynced
    }
  }

  return numOfThingsSynced
}

const MINUTES_PER_MOD = 38

function onRetrieveMultiple(resourceNames: TableInfo.KeyType[]) {
  const result = {}
  for (const resourceName of resourceNames) {
    result[resourceName] = Table.readAllRecords(resourceName)
  }
  return result
}
function uiSyncForms() {
  try {
    const result = onSyncForms()
    SpreadsheetApp.getUi().alert(
      `Finished sync! ${result as number} new form submits found.`
    )
  } catch (err) {
    Logger.log(stringifyError(err))
    throw err
  }
}

namespace Parsing {
  /*

  Spec: field parsing

  Dates are treated as numbers.

  */

  export function parseField<T extends FieldType>(x: any, fieldType: T): TableInfo.ParseFieldType<T> {
    switch (fieldType) {
      case FieldType.BOOLEAN:
        return (x === "true" || x === true ? true : false) as TableInfo.ParseFieldType<T>
      case FieldType.NUMBER:
        return Number(x) as TableInfo.ParseFieldType<T>
      case FieldType.STRING:
        return String(x) as TableInfo.ParseFieldType<T>
      case FieldType.DATE:
        if (x === "" || x === -1) {
          return -1 as TableInfo.ParseFieldType<T>
        } else {
          return Number(x) as TableInfo.ParseFieldType<T>
        }
      case FieldType.JSON:
        return JSON.parse(x) as TableInfo.ParseFieldType<T>
    }
  }
  export function serializeField(x: any, fieldType: FieldType): any {
    switch (fieldType) {
      case FieldType.BOOLEAN:
        return x ? true : false
      case FieldType.NUMBER:
        return Number(x)
      case FieldType.STRING:
        return String(x)
      case FieldType.DATE:
        if (x === -1 || x === "") {
          return ""
        } else {
          return new Date(x)
        }
      case FieldType.JSON:
        return JSON.stringify(x)
    }
  }
  export function contactPref(s: string) {
    if (s === "Phone") return "phone"
    if (s === "Email") return "email"
    return "either"
  }
  export function grade(g: string) {
    if (g === "Freshman") return 9
    if (g === "Sophomore") return 10
    if (g === "Junior") return 11
    if (g === "Senior") return 12
    return 0
  }
  export function modInfo(abDay: string, mod1To10: number): number {
    if (abDay.toLowerCase().charAt(0) === "a") {
      return mod1To10
    }
    if (abDay.toLowerCase().charAt(0) === "b") {
      return mod1To10 + 10
    }
    throw new Error(`${String(abDay)} does not start with A or B`)
  }
  export function modData(modData: string[]): number[] {
    function doParse(d: string) {
      return d
        .split(",")
        .map((x) => x.trim())
        .filter((x) => x !== "" && x !== "None")
        .map((x) => parseInt(x))
    }
    const mA15: number[] = doParse(modData[0])
    const mB15: number[] = doParse(modData[1]).map((x) => x + 10)
    const mA60: number[] = doParse(modData[2])
    const mB60: number[] = doParse(modData[3]).map((x) => x + 10)
    return mA15.concat(mA60).concat(mB15).concat(mB60)
  }
}
// The main "sync forms" function that's crammed with form data formatting.
function onSyncForms(): number {
  // tables
  const tutors = Table.readAllRecords("tutors")
  const matchings = Table.readAllRecords("matchings")
  const attendanceDays = Table.readAllRecords("attendanceDays")
  const studentInfo = Table.readAllRecords('studentInfo');
  const attendanceDaysIndex = {}
  for (const day of Object_values(attendanceDays)) {
    attendanceDaysIndex[day.dateOfAttendance] = day.abDay
  }

  function processRequestFormRecord(r: TableInfo.Rec<"requestForm">) {
    new Table('requestSubmissions').createRecord({
      // student info
      friendlyFullName: r.friendlyFullName,
      firstName: r.firstName,
      lastName: r.lastName,
      grade: Parsing.grade(r.grade),
      studentId: r.studentId,
      email: r.email,
      phone: r.phone,
      contactPref: r.contactPref,
      homeroom: r.homeroom,
      homeroomTeacher: r.homeroomTeacher,
      // These two fields are never actually used!
      attendanceAnnotation: undefined,
      attendance: undefined,

      id: -1,
      date: r.date, // the date MUST be the date from the form; this is used for syncing
      mods: Parsing.modData([
        r.modDataA1To5,
        r.modDataB1To5,
        r.modDataA6To10,
        r.modDataB6To10,
      ]),
      subject: r.subject,
      status: "unchecked",
      isSpecial: false,
      annotation: "",
    });
  }
  function processTutorRegistrationFormRecord(r: TableInfo.Rec<"tutorRegistrationForm">) {
    function parseSubjectList(d: string[]) {
      return d
        .join(",")
        .split(",") // remember that within each string there are commas
        .map((x) => x.trim())
        .filter((x) => x !== "" && x !== "None")
        .map((x) => String(x))
        .join(", ")
    }
    const studentInfo = new Table('studentInfo').createRecord({
      id: -1,
      date: r.date,
      friendlyFullName: r.friendlyFullName,
      firstName: r.firstName,
      lastName: r.lastName,
      grade: Parsing.grade(r.grade),
      studentId: r.studentId,
      email: r.email,
      phone: r.phone,
      contactPref: r.contactPref,
      homeroom: r.homeroom,
      homeroomTeacher: r.homeroomTeacher,
      attendanceAnnotation: '',
      attendance: {},
    });
    new Table('tutors').createRecord({
      studentInfo: studentInfo.id,
      id: -1,
      date: r.date,
      mods: Parsing.modData([
        r.modDataA1To5,
        r.modDataB1To5,
        r.modDataA6To10,
        r.modDataB6To10,
      ]),
      modsPref: Parsing.modData([
        r.modDataPrefA1To5,
        r.modDataPrefB1To5,
        r.modDataPrefA6To10,
        r.modDataPrefB6To10,
      ]),
      subjectList: parseSubjectList([
        r.subjects0,
        r.subjects1,
        r.subjects2,
        r.subjects3,
        r.subjects4,
        r.subjects5,
        r.subjects6,
        r.subjects7,
      ]),
      dropInMods: [],
      afterSchoolAvailability: r.afterSchoolAvailability,
      additionalHours: 0,
    })
  }
  function processSpecialRequestFormRecord(r: TableInfo.Rec<"specialRequestForm">) {
    let annotation = ""
    annotation += `TEACHER EMAIL: ${r.teacherEmail}; `
    annotation += `# LEARNERS: ${r.numLearners}; `
    annotation += `TUTORING DATE INFO: ${r.tutoringDateInformation}; `
    if (r.studentNames.trim() !== "") {
      annotation += `STUDENT NAMES: ${r.studentNames}; `
    }
    if (r.additionalInformation.trim() !== "") {
      annotation += `ADDITIONAL INFO: ${r.additionalInformation}; `
    }
    new Table('requestSubmissions').createRecord({
      id: -1,
      date: r.date, // the date MUST be the date from the form
      friendlyFullName: "[special request]",
      firstName: "[special request]",
      lastName: "[special request]",
      grade: -1,
      studentId: -1,
      email: "[special request]",
      phone: "[special request]",
      contactPref: "either",
      homeroom: "[special request]",
      homeroomTeacher: "[special request]",
      attendanceAnnotation: undefined, // never used
      attendance: undefined, // never used
      mods: [Parsing.modInfo(r.abDay, r.mod1To10)],
      subject: r.subject,
      isSpecial: true,
      annotation,
      status: "unchecked",
    })
  }
  function processAttendanceFormRecord(r: TableInfo.Rec<"attendanceForm">) {
    let tutor = -1
    let learner = -1
    let validity = ""
    let minutesForTutor = -1
    let minutesForLearner = -1
    let presenceForTutor = ""
    let presenceForLearner = ""
    let mod = -1
    const processedDateOfAttendance = roundDownToDay(
      r.dateOfAttendance === -1 ? r.date : r.dateOfAttendance
    )

    if (attendanceDaysIndex[processedDateOfAttendance] === undefined) {
      validity = "date does not exist in attendance days index"
    } else {
      mod = Parsing.modInfo(
        attendanceDaysIndex[processedDateOfAttendance],
        r.mod1To10
      )
    }

    // give the tutor their time
    minutesForTutor = MINUTES_PER_MOD
    presenceForTutor = "P"

    // figure out who the tutor is, by student ID
    const xTutors = recordCollectionToArray(tutors).filter(
      (x) => studentInfo[x.studentInfo].studentId === r.studentId
    )
    if (xTutors.length === 0) {
      validity = "tutor student ID does not exist"
    } else if (xTutors.length === 1) {
      tutor = xTutors[0].id
    } else {
      throw new Error(`duplicate tutor student id ${String(r.studentId)}`)
    }

    // does the tutor have a learner?
    const xMatchings = recordCollectionToArray(matchings).filter(
      (x) => x.tutor === tutor && x.mod === mod
    )
    if (xMatchings.length === 0) {
      learner = -1
    } else if (xMatchings.length === 1) {
      learner = xMatchings[0].learner
    } else {
      throw new Error(
        `the tutor ${studentInfo[tutors[xMatchings[0].tutor].studentInfo].friendlyFullName} is matched twice on the same mod`
      )
    }

    // ATTENDANCE LOGIC
    if (validity === "") {
      if (r.presence === "Yes") {
        minutesForLearner = MINUTES_PER_MOD
        presenceForLearner = "P"
      } else if (r.presence === "No") {
        minutesForLearner = 0
        presenceForLearner = "A"
      } else if (r.presence === `I don't have a learner assigned`) {
        if (learner === -1) {
          minutesForLearner = -1
        } else {
          // so there really is a learner...
          validity = `tutor said they don't have a learner assigned, but they do!`
        }
      } else if (
        r.presence === `No, but the learner doesn't need any tutoring today`
      ) {
        if (learner === -1) {
          validity = `tutor said they have a learner assigned, but they don't!`
        } else {
          minutesForLearner = 1 // TODO: this is a hacky solution; fix it
          presenceForLearner = "E"
        }
      } else {
        throw new Error(`invalid presence (${String(r.presence)})`)
      }
    }

    new Table('attendanceLog').createRecord({
      id: -1,
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
      markForReset: "",
    })
  }

  let numOfThingsSynced = 0
  numOfThingsSynced += doFormSync(
    new Table("requestForm"),
    new Table("requestSubmissions"),
    processRequestFormRecord
  )
  numOfThingsSynced += doFormSync(
    new Table("specialRequestForm"),
    new Table("requestSubmissions"),
    processSpecialRequestFormRecord
  )
  numOfThingsSynced += doFormSync(
    new Table("attendanceForm"),
    new Table("attendanceLog"),
    processAttendanceFormRecord
  )
  numOfThingsSynced += doFormSync(
    new Table("tutorRegistrationForm"),
    new Table("tutors"),
    processTutorRegistrationFormRecord
  )
  return numOfThingsSynced
}

function uiRecalculateAttendance() {
  try {
    const numAttendancesChanged = onRecalculateAttendance()
    SpreadsheetApp.getUi().alert(
      `Finished attendance update. ${numAttendancesChanged} attendances were changed.`
    )
  } catch (err) {
    Logger.log(stringifyError(err))
    throw err
  }
}

// This recalculates the attendance.
function onRecalculateAttendance() {
  // read tables
  const tutors = Table.readAllRecords("tutors")
  const learners = Table.readAllRecords("learners")
  const studentInfo = Table.readAllRecords('studentInfo')
  const attendanceLog = Table.readAllRecords("attendanceLog")
  const attendanceDays = Table.readAllRecords("attendanceDays")
  const matchings = Table.readAllRecords("matchings")
  const tutorsArray = Object_values(tutors)
  const learnersArray = Object_values(learners)
  const matchingsArray = Object_values(matchings)
  const attendanceLogArray = Object_values(attendanceLog)
  const studentInfoArray = Object_values(studentInfo)

  let numAttendancesChanged = 0
  function calculateIsBDay(x: string) {
    if (x.toLowerCase().charAt(0) === "a") {
      return false
    } else if (x.toLowerCase().charAt(0) === "b") {
      return true
    } else {
      throw new Error("unrecognized attendance day letter")
    }
  }

  function whenTutorFormNotFilledOutLogic(
    tutorId: number,
    learnerId: number,
    mod: number,
    day: TableInfo.Rec<'attendanceDays'>,
    studentInfoRecs: TableInfo.RecCollection<'studentInfo'>
  ) {
    const tutor = tutors[tutorId]
    const learner = learnerId === -1 ? null : learners[learnerId]
    const date = day.dateOfAttendance
    if (tutor === undefined) {
      throw new Error(`tutor does not exist (${tutorId})`)
    }
    if (learner === undefined) {
      throw new Error(`learner does not exist (${learnerId})`)
    }
    // mark tutor as absent
    let alreadyExists = false // if a presence or absence exists, don't add an absence
    if (studentInfoRecs[tutor.studentInfo].attendance[date] === undefined) {
      studentInfoRecs[tutor.studentInfo].attendance[date] = []
    }
    if (learner !== null && studentInfoRecs[learner.studentInfo].attendance[date] === undefined) {
      studentInfoRecs[learner.studentInfo].attendance[date] = []
    } else {
      for (const attendanceModDataString of studentInfoRecs[tutor.studentInfo].attendance[date]) {
        const tokens = attendanceModDataString.split(" ")
        if (Number(tokens[0]) === mod) {
          alreadyExists = true
        }
      }
    }
    if (!alreadyExists) {
      // add an absence for the tutor
      studentInfoRecs[tutor.studentInfo].attendance[date].push(formatAttendanceModDataString(mod, 0))
      // add an excused absence for the learner, if exists
      if (learnerId !== -1) {
        studentInfoRecs[learners[learnerId].studentInfo].attendance[date].push(
          formatAttendanceModDataString(mod, 1)
        )
      }
      numAttendancesChanged += 2
    }
  }

  function applyAttendanceForStudent(
    collection: TableInfo.RecCollection<'tutors'> | TableInfo.RecCollection<'learners'>,
    id: number,
    entry: TableInfo.Rec<'attendanceLog'>,
    minutes: number
  ) {
    const record = collection[id]
    if (record === undefined) {
      // This is not a fatal error, especially since we want to retroactively
      // run attendance calculations, but give a log message
      Logger.log(`attendance: record not found ${id}`)
      return
    }
    const attendance = studentInfo[record.studentInfo].attendance
    let isNew = true
    if (attendance[entry.dateOfAttendance] === undefined) {
      attendance[entry.dateOfAttendance] = []
    } else {
      for (const attendanceModDataString of attendance[
        entry.dateOfAttendance
      ]) {
        const tokens = attendanceModDataString.split(" ")
        if (entry.mod === Number(tokens[0])) {
          isNew = false
        }
      }
    }
    // RESET BEHAVIOR: remove attendance logs marked with a reset flag
    if (entry.markForReset === "doreset") {
      if (!isNew) {
        attendance[entry.dateOfAttendance] = attendance[
          entry.dateOfAttendance
        ].filter((attendanceModDataString: string) => {
          const tokens = attendanceModDataString.split(" ")
          if (Number(tokens[0]) === entry.mod) {
            ++numAttendancesChanged
            return false
          } else {
            return true
          }
        })
      }
    } else {
      if (isNew) {
        // SPECIAL DATA FORMAT TO SAVE SPACE
        attendance[entry.dateOfAttendance].push(
          formatAttendanceModDataString(entry.mod, minutes)
        )
        ++numAttendancesChanged
      }
    }
  }

  // PROCESS EACH ATTENDANCE LOG ENTRY
  for (const entry of attendanceLogArray) {
    if (entry.validity !== "") {
      // exclude the entry
      continue
    }
    if (entry.tutor !== -1) {
      applyAttendanceForStudent(
        tutors,
        entry.tutor,
        entry,
        entry.minutesForTutor
      )
    }
    if (entry.learner !== -1) {
      applyAttendanceForStudent(
        learners,
        entry.learner,
        entry,
        entry.minutesForLearner
      )
    }
  }

  // INDEX TUTORS: FIGURE OUT WHICH TUTORS/MODS HAVE UNSUBMITTED FORMS
  // For each tutor, keep track of which mods they need to fill out forms.
  // Figure out which tutors HAVEN'T filled out their forms.
  // Also keep track of learner ID associated with each mod (-1 for null).
  type TutorAttendanceFormIndexEntry = {
    id: number
    mod: { [mod: number]: { wasFormSubmitted: boolean; learnerId: number } }
  }
  type TutorAttendanceFormIndex = {
    [id: number]: TutorAttendanceFormIndexEntry
  }
  const tutorAttendanceFormIndex: TutorAttendanceFormIndex = {}
  for (const tutor of tutorsArray) {
    tutorAttendanceFormIndex[tutor.id] = {
      id: tutor.id,
      mod: {},
    }
    // handle drop-ins
    for (const mod of tutor.dropInMods) {
      tutorAttendanceFormIndex[tutor.id].mod[mod] = {
        wasFormSubmitted: false,
        learnerId: -1,
      }
    }
  }
  for (const matching of matchingsArray) {
    tutorAttendanceFormIndex[matching.tutor].mod[matching.mod] = {
      wasFormSubmitted: false,
      learnerId: matching.learner,
    }
  }

  // DEAL WITH THE UNSUBMITTED FORMS
  for (const day of Object_values(attendanceDays)) {
    if (
      day.status === "ignore" ||
      day.status === "isdone" ||
      day.status === "isreset"
    ) {
      continue
    }
    if (day.status !== "doit" && day.status !== "doreset") {
      throw new Error("unknown day status")
    }

    // modify dateOfAttendance; round down to the nearest 24-hour day
    day.dateOfAttendance = roundDownToDay(day.dateOfAttendance)

    const isBDay: boolean = calculateIsBDay(day.abDay)

    // iterate through all tutors & hunt down the ones that didn't fill out the form
    if (day.status === "doit") {
      for (const tutor of tutorsArray) {
        for (let i = isBDay ? 10 : 0; i < (isBDay ? 20 : 10); ++i) {
          const x = tutorAttendanceFormIndex[tutor.id].mod[i]
          if (x !== undefined && !x.wasFormSubmitted) {
            whenTutorFormNotFilledOutLogic(tutor.id, x.learnerId, i, day, studentInfo)
          }
        }
        if (
          studentInfo[tutor.studentInfo].attendance[day.dateOfAttendance] !== undefined &&
          studentInfo[tutor.studentInfo].attendance[day.dateOfAttendance].length === 0
        ) {
          delete studentInfo[tutor.studentInfo].attendance[day.dateOfAttendance]
        }
      }
    } else if (day.status === "doreset") {
      for (const tutor of tutorsArray) {
        if (studentInfo[tutor.studentInfo].attendance[day.dateOfAttendance] !== undefined) {
          // delete EVERY ABSENCE for that day (but keep excused)
          // this gets rid of anything automatically generated for that day
          studentInfo[tutor.studentInfo].attendance[day.dateOfAttendance] = studentInfo[tutor.studentInfo].attendance[
            day.dateOfAttendance
          ].filter((attendanceModDataString: string) => {
            const tokens = attendanceModDataString.split(" ")
            if (Number(tokens[1]) === 0) {
              ++numAttendancesChanged
              return false
            } else {
              return true
            }
          })
        }
      }
    }

    // change day status
    if (day.status === "doreset") {
      day.status = "isreset"
    }
    if (day.status === "doit") {
      day.status = "isdone"
    }

    // update record
    new Table("attendanceDays").updateRecord(day)
  }

  // THAT'S ALL! UPDATE TABLES
  new Table("tutors").updateAllRecords(tutorsArray)
  new Table("learners").updateAllRecords(learnersArray)
  new Table("studentInfo").updateAllRecords(studentInfoArray)

  return numAttendancesChanged
}

function uiGenerateSchedule() {
  try {
    onGenerateSchedule()
    SpreadsheetApp.getUi().alert("Success")
  } catch (err) {
    Logger.log(stringifyError(err))
    throw err
  }
}

// This is the KEY FUNCTION for generating the "schedule" sheet.
function onGenerateSchedule() {
  type ScheduleEntry = {
    tutorName: string
    info: string
    isDropIn: boolean
  }

  const sortComparator = (a: ScheduleEntry, b: ScheduleEntry) => {
    if (a.isDropIn < b.isDropIn) return -1
    if (a.isDropIn > b.isDropIn) return 1
    const x = a.tutorName.toLowerCase()
    const y = b.tutorName.toLowerCase()
    if (x < y) return -1
    if (x > y) return 1
    return 0
  }

  // Delete & insert sheet
  let ss = SpreadsheetApp.openById(
    "1VbrZTMXGju_pwSrY7-M0l8citrYTm5rIv7RPuKN9deY"
  )
  let sheet = ss.getSheetByName("schedule")
  if (sheet !== null) {
    ss.deleteSheet(sheet)
  }
  sheet = ss.insertSheet("schedule", 0)

  // Get all matchings, tutors, and learners
  const matchings = Table.readAllRecords("matchings")
  const tutors = Table.readAllRecords("tutors")
  const learners = Table.readAllRecords("learners")
  const studentInfo = Table.readAllRecords('studentInfo')

  // Header
  sheet.appendRow(["ARC SCHEDULE"])
  sheet.appendRow([""])

  // Note: Bookings are deliberately ignored.
  const tutorIndex = schedulingTutorIndex(tutors, {}, matchings)
  const otherTutorStrings: string[] = []
  const layoutMatrix: [ScheduleEntry[], ScheduleEntry[]][] = [] // [mod0to9][abday]

  for (let i = 0; i < 10; ++i) {
    layoutMatrix[i] = [[], []]
  }

  function matchingToText(matching: TableInfo.Rec<'matchings'>) {
    let result = ""
    if (matching.learner !== -1) {
      result += `(w/${studentInfo[learners[matching.learner].studentInfo].friendlyFullName})`
    }
    if (matching.annotation !== "") {
      result += `(${matching.annotation})`
    }
    return result
  }

  function matchingsInfo(mod: number, refs: [SchedulingReference, number][]) {
    for (const [sr, srid] of refs) {
      if (sr === SchedulingReference.MATCHING) {
        const matching = matchings[srid]
        if (matching.mod !== mod) continue
        return matchingToText(matching)
      }
    }
  }

  for (const x of Object_values(tutorIndex)) {
    if (x.isInMod) {
      for (let i = 0; i < 20; ++i) {
        switch (x.modStatus[i]) {
          case ModStatus.DROP_IN:
          case ModStatus.DROP_IN_PREF:
            layoutMatrix[i % 10][i < 10 ? 0 : 1].push({
              tutorName: studentInfo[tutors[x.id].studentInfo].friendlyFullName,
              info: "(drop-in)",
              isDropIn: true,
            })
            break
          case ModStatus.MATCHED:
            layoutMatrix[i % 10][i < 10 ? 0 : 1].push({
              tutorName: studentInfo[tutors[x.id].studentInfo].friendlyFullName,
              info: "(tutoring) " + matchingsInfo(i + 1, x.refs),
              isDropIn: true,
            })
            break
        }
      }
    } else {
      otherTutorStrings.push(
        String(studentInfo[tutors[x.id].studentInfo].friendlyFullName) +
        x.refs
          .filter(
            ([sr, srid]) =>
              sr === SchedulingReference.MATCHING &&
              matchings[srid].mod === -1
          )
          .map(([sr, srid]) => matchingToText(matchings[srid]))
          .join(" ")
      )
    }
  }

  for (let i = 0; i < 10; ++i) {
    for (let j = 0; j < 2; ++j) {
      layoutMatrix[i][j].sort(sortComparator)
    }
  }

  // Print!
  // COLOR SCHEME
  const COLOR_SCHEME_STRING = "ff99c8-fcf6bd-d0f4de-a9def9-e4c1f9"
  const COLOR_SCHEME = COLOR_SCHEME_STRING.split("-")
  const DAY_OF_WEEK = new Date().getDay()
  // calculate the color based on the day of the week
  const PRIMARY_COLOR =
    "#" +
    COLOR_SCHEME[
    {
      0: 0,
      1: 1,
      2: 2,
      3: 3,
      4: 4,
      5: 2,
      6: 3,
    }[DAY_OF_WEEK]
    ]
  const SECONDARY_COLOR = "#383f51"

  // FORMAT SHEET
  sheet.setHiddenGridlines(true)

  // CHANGE COLUMNS
  sheet.deleteColumns(5, sheet.getMaxColumns() - 5)
  sheet.setColumnWidth(1, 50)
  sheet.setColumnWidth(2, 300)
  sheet.setColumnWidth(3, 30)
  sheet.setColumnWidth(4, 300)
  sheet.setColumnWidth(5, 30)

  // HEADER
  sheet.getRange(1, 1, 3, 5).mergeAcross()
  sheet
    .getRange(1, 1)
    .setValue("ARC Schedule")
    .setFontSize(36)
    .setHorizontalAlignment("center")
  sheet.setRowHeight(2, 15)
  sheet.setRowHeight(3, 15)
  sheet
    .getRange(4, 2)
    .setValue("A Days")
    .setFontSize(18)
    .setHorizontalAlignment("center")
  sheet
    .getRange(4, 4)
    .setValue("B Days")
    .setFontSize(18)
    .setHorizontalAlignment("center")
  sheet.setRowHeight(5, 30)
  sheet.getRange(1, 1, 4, 5).setBackground(PRIMARY_COLOR)

  // LAYOUT
  let nextRow = 6
  for (let i = 0; i < 10; ++i) {
    const scheduleRowSize = Math.max(
      layoutMatrix[i][0].length,
      layoutMatrix[i][1].length,
      1
    )
    // LABEL
    sheet.getRange(nextRow, 1, scheduleRowSize).merge()
    sheet
      .getRange(nextRow, 1)
      .setValue(`${i + 1}`)
      .setFontSize(18)
      .setVerticalAlignment("top")

    // CONTENT
    if (layoutMatrix[i][0].length > 0) {
      sheet
        .getRange(nextRow, 2, layoutMatrix[i][0].length)
        .setValues(layoutMatrix[i][0].map((x) => [`${x.tutorName} ${x.info}`]))
        .setWrap(true)
        .setFontSize(12)
        .setFontColors(
          layoutMatrix[i][0].map((x) => [x.isDropIn ? "black" : "red"])
        )
    }
    if (layoutMatrix[i][1].length > 0) {
      sheet
        .getRange(nextRow, 4, layoutMatrix[i][1].length)
        .setValues(layoutMatrix[i][1].map((x) => [`${x.tutorName} ${x.info}`]))
        .setWrap(true)
        .setFontSize(12)
        .setFontColors(
          layoutMatrix[i][1].map((x) => [x.isDropIn ? "black" : "red"])
        )
    }

    // SET THE NEXT ROW
    nextRow += scheduleRowSize

    // GUTTER
    sheet.getRange(nextRow, 1, 1, 5).merge()
    sheet.setRowHeight(nextRow, 60)
    ++nextRow
  }

  // UNSCHEDULED TUTORS
  sheet
    .getRange(nextRow, 2, 1, 3)
    .merge()
    .setValue(`Other tutors`)
    .setFontSize(18)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setWrap(true)
  sheet
    .getRange(nextRow, 2, otherTutorStrings.length + 1, 3)
    .setBorder(
      true,
      true,
      true,
      true,
      null,
      null,
      PRIMARY_COLOR,
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    )
  ++nextRow
  sheet
    .getRange(nextRow, 2, otherTutorStrings.length, 3)
    .mergeAcross()
    .setHorizontalAlignment("center")
  sheet
    .getRange(nextRow, 2, otherTutorStrings.length)
    .setFontSize(12)
    .setValues(otherTutorStrings.map((x) => [x]))
  nextRow += otherTutorStrings.length

  // FOOTER
  sheet.getRange(nextRow, 1, 1, 5).merge()
  sheet.setRowHeight(nextRow, 20)
  ++nextRow

  sheet
    .getRange(nextRow, 1, 1, 5)
    .merge()
    .setValue(`Schedule auto-generated on ${new Date()}`)
    .setFontSize(10)
    .setFontColor("white")
    .setBackground(SECONDARY_COLOR)
    .setHorizontalAlignment("center")
  ++nextRow

  sheet
    .getRange(nextRow, 1, 1, 5)
    .merge()
    .setValue(`ARC App designed by Suhao Jeffrey Huang`)
    .setFontSize(10)
    .setFontColor("white")
    .setBackground(SECONDARY_COLOR)
    .setHorizontalAlignment("center")
  ++nextRow

  // FIT ROWS/COLUMNS
  sheet.deleteRows(
    sheet.getLastRow() + 1,
    sheet.getMaxRows() - sheet.getLastRow()
  )

  // FONT
  sheet.getDataRange().setFontFamily("Helvetica")

  return null
}

function uiDeDuplicateTutorsAndLearners() {
  const PROMPT_TEXT =
    'This command will de-duplicate tutors and learners. Old form submissions are replaced with newer form submissions. Type the word "proceed" to proceed. Leave the box blank to cancel.'

  try {
    const ui = SpreadsheetApp.getUi()
    const response = ui.prompt(PROMPT_TEXT)
    if (response.getResponseText() === "proceed") {
      let numberOfOperations = 0
      numberOfOperations += deDuplicateTableByStudentId("tutors")
      numberOfOperations += deDuplicateTableByStudentId("learners")
      SpreadsheetApp.getUi().alert(
        `${numberOfOperations} de-duplication operations performed`
      )
    }
  } catch (err) {
    Logger.log(stringifyError(err))
    throw err
  }
}

function deDuplicateTableByStudentId(tableKey: TableInfo.KeyType): number {
  let numberOfOperations = 0
  type ReplacementMap = { [studentId: number]: { recordIds: number[] } }

  const replacementMap: ReplacementMap = {}

  const table = new Table(tableKey)
  const recordObject = table.readAllRecords()
  for (const record of Object_values(recordObject)) {
    if (replacementMap[record.studentId] === undefined) {
      replacementMap[record.studentId] = { recordIds: [record.id] }
    } else {
      replacementMap[record.studentId].recordIds.push(record.id)
    }
  }
  // If a student ID has multiple records, delete all but the newest (highest ID)
  // and the remaining record should have ID edited to the oldest record
  for (const mapItem of Object_values(replacementMap)) {
    if (mapItem.recordIds.length >= 2) {
      let lowestId = mapItem.recordIds[0]
      let highestId = mapItem.recordIds[0]
      for (let i = 1; i < mapItem.recordIds.length; ++i) {
        if (mapItem.recordIds[i] < lowestId) {
          lowestId = mapItem.recordIds[i]
        }
        if (mapItem.recordIds[i] > highestId) {
          highestId = mapItem.recordIds[i]
        }
      }

      // delete all records not equal to the highest ID
      for (let i = 0; i < mapItem.recordIds.length; ++i) {
        if (highestId !== mapItem.recordIds[i]) {
          table.deleteRecord(mapItem.recordIds[i])
          ++numberOfOperations
        }
      }

      // This edits the ID of the record, which is very uncommon,
      // so the second argument of updateRecord is invoked.
      recordObject[highestId].id = lowestId

      // IMPORTANT OPERATION: Copy the old attendance into the new attendance,
      // the old attendanceAnnotation into the new attendanceAnnotation, etc.
      // We want to preserve all of these fields!!!
      for (const field of [
        "attendance",
        "attendanceAnnotation",
        "dropInMods",
        "additionalHours",
      ]) {
        if (recordObject[lowestId].hasOwnProperty(field)) {
          recordObject[highestId][field] = recordObject[lowestId][field]
        }
      }
      table.updateRecord(recordObject[highestId], table.getRowById(highestId))
      ++numberOfOperations
    }
  }
  return numberOfOperations
}

function onOpen(_ev: any) {
  const menu = SpreadsheetApp.getUi().createMenu("ARC APP")
  menu.addItem("Sync data from forms", "uiSyncForms")
  menu.addItem("Generate schedule", "uiGenerateSchedule")
  menu.addItem("Recalculate attendance", "uiRecalculateAttendance")
  menu.addItem(
    "De-duplicate tutors and learners",
    "uiDeDuplicateTutorsAndLearners"
  )
  if (ARC_APP_DEBUG_MODE) {
    menu.addItem("Debug: test client API", "debugClientApiTest")
    menu.addItem("Debug: rebuild all headers", "debugHeaders")
    menu.addItem("Debug: run temporary script", "debugRunTemporaryScript")
  }
  menu.addToUi()
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
