/*

GLOBAL SETTINGS

*/

const ARC_APP_DEBUG_MODE: boolean = true

/*

Shared with the other file

*/

export type Rec = {
  id: number
  date: number
  [others: string]: any
}
export type RecCollection = {
  [id: string]: Rec
}

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
  tutorRecords: RecCollection,
  bookingRecords: RecCollection,
  matchingRecords: RecCollection
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

function recordCollectionToArray(r: RecCollection): Rec[] {
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
  function StringField(key: string): [FieldType, string] {
    return [FieldType.STRING, key]
  }
  function NumberField(key: string): [FieldType, string] {
    return [FieldType.NUMBER, key]
  }
  function DateField(key: string): [FieldType, string] {
    return [FieldType.DATE, key]
  }
  function BooleanField(key: string): [FieldType, string] {
    return [FieldType.BOOLEAN, key]
  }
  function JsonField(key: string): [FieldType, string] {
    return [FieldType.JSON, key]
  }
  export const ARRAY_1: [string, string, [FieldType, string][]][] = [
    [
      "studentInfo",
      "$student-info",
      [
        StringField("name"),
        NumberField("grade"),
        NumberField("studentId"),
        StringField("email"),
        StringField("phone"),
        StringField("contactPref"),
        StringField("homeroom"),
        StringField("homeroomTeacher"),
        StringField("attendanceAnnotation"),
        StringField("interestVirtualTutoring"),
        StringField("interestVirtualTutoringMiddleSchool"),
        JsonField("attendance"),
      ],
    ],
    [
      "formInfo",
      "$form-info",
      [
        StringField("origin"),
        StringField("hash"),
        JsonField("json"),
        StringField("archive"),
      ],
    ],
    [
      "tutors",
      "$tutors",
      [
        NumberField("studentInfo"),
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
  ]
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
  const cache: ObjectMap<Info> = {}

  export type Info = {
    key: string
    sheetName: string
    sheet: GoogleAppsScript.Spreadsheet.Sheet
    isWriteAllowed: boolean
    fields: [FieldType, string][]
  }

  export function use(key: string) {
    if (cache.hasOwnProperty(key)) {
      return cache[key]
    } else {
      for (const ti of TableInfo.ARRAY_1) {
        if (ti[0] === key) {
          const sheetName = ti[1]
          const sheet = SpreadsheetApp.getActive().getSheetByName(ti[1])
          const fields = ti[2]
          const isWriteAllowed = true // TODO Spec: forms
          cache[key] = {
            key,
            sheetName,
            sheet,
            isWriteAllowed,
            fields,
          }
          return cache[key]
        }
      }
    }
  }
}

/*

Spec: forms

MUST have every form linked to the table $form-info
MUST have a single-stop conversion utility
  Update forms
MUST have a custom record parser
TODO: start with attendance core logic
*/

const FORM_INFO_MASTER_ARRAY = [
  ["requestForm", "$request-form"],
  ["specialRequestForm", "$special-request-form"],
  ["attendanceForm", "$attendance-form"],
  ["tutorRegistrationForm", "$tutor-registration-form"],
]

class Table {
  io: TableIO.Info

  static readAllRecords(key: string) {
    return new Table(key).retrieveAllRecords()
  }

  constructor(key: string) {
    this.io = TableIO.use(key)
  }

  retrieveAllRecords(): RecCollection {
    if (this.io.sheet.getLastColumn() !== this.io.fields.length) {
      throw new Error()
    }
    const raw = this.io.sheet.getDataRange().getValues()
    const res = {}
    for (let i = 1; i < raw.length; ++i) {
      const rec = this.parseRecord(raw[i])
      res[String(rec.id)] = rec
    }
    return res
  }
  parseRecord(raw: any[]): Rec {
    const rec: { [key: string]: any } = {}
    for (let i = 0; i < this.io.fields.length; ++i) {
      const field = this.io.fields[i]
      // this accounts for blanks in the last field
      rec[field[1]] = Parsing.parseField(
        raw[i] === undefined ? "" : raw[i],
        field[0]
      )
    }
    if (rec.id === undefined) {
      // TODO Spec: forms
      rec.id = rec.date
    }
    return rec as Rec
  }

  serializeRecord(record: Rec): any[] {
    return this.io.fields.map((field) =>
      Parsing.serializeField(record[field[1]], field[0])
    )
  }

  createRecord(record: Rec): Rec {
    this.checkWritePermission()
    if (record.date === -1) {
      record.date = Date.now()
    }
    if (record.id === -1) {
      record.id = GlobalId.getNextGlobalId()
    }
    this.io.sheet.appendRow(this.serializeRecord(record))
    return record
  }

  getRowById(id: number): number {
    // because the first row is headers, we ignore it and start from the second row
    const mat: any[][] = this.io.sheet
      .getRange(2, 1, this.io.sheet.getLastRow() - 1)
      .getValues()
    let rowNum = -1
    for (let i = 0; i < mat.length; ++i) {
      const cell: number = mat[i][0]
      if (typeof cell !== "number") {
        throw new Error(
          `id at location ${String(i)} is not a number in table ${String(
            this.io.key
          )}`
        )
      }
      if (cell === id) {
        if (rowNum !== -1) {
          throw new Error(
            `duplicate ID ${String(id)} in table ${String(this.io.key)}`
          )
        }
        rowNum = i + 2 // i = 0 <=> second row (rows are 1-indexed)
      }
    }
    if (rowNum == -1) {
      throw new Error(
        `ID ${String(id)} not found in table ${String(this.io.key)}`
      )
    }
    return rowNum
  }

  updateRecord(editedRecord: Rec, rowNum?: number): void {
    this.checkWritePermission()
    if (rowNum === undefined) {
      rowNum = this.getRowById(editedRecord.id)
    }
    this.io.sheet
      .getRange(rowNum, 1, 1, this.io.sheet.getLastColumn())
      .setValues([this.serializeRecord(editedRecord)])
  }

  updateAllRecords(editedRecords: Rec[]): void {
    this.checkWritePermission()
    if (this.io.sheet.getLastRow() === 1) {
      return // the sheet is empty, and trying to select it will result in an error
    }
    // because the first row is headers, we ignore it and start from the second row
    const mat: any[][] = this.io.sheet
      .getRange(2, 1, this.io.sheet.getLastRow() - 1)
      .getValues()
    let idRowMap: ObjectMap<number> = {}
    for (let i = 0; i < mat.length; ++i) {
      idRowMap[String(mat[i][0])] = i + 2 // i = 0 <=> second row (rows are 1-indexed)
    }
    for (const r of editedRecords) {
      this.updateRecord(r, idRowMap[String(r.id)])
    }
  }

  deleteRecord(id: number): void {
    this.checkWritePermission()
    this.io.sheet.deleteRow(this.getRowById(id))
  }

  rebuildSheetHeadersIfNeeded() {
    this.checkWritePermission()
    const col = this.io.sheet.getLastColumn()
    this.io.sheet.getRange(1, 1, 1, col === 0 ? 1 : col).clearContent()
    this.io.sheet
      .getRange(1, 1, 1, this.io.fields.length)
      .setValues([this.io.fields.map((x) => x[1])])
  }

  checkWritePermission(): void {
    if (!this.io.isWriteAllowed) {
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
  tryThis(1, "retrieveAll", table.retrieveAllRecords)
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
    for (const x of TableInfo.ARRAY_1) {
      new Table(x[0]).rebuildSheetHeadersIfNeeded()
    }
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
function doFormSync(
  formTable: Table,
  actualTable: Table,
  formRecordToActualRecord: (formRecord: Rec) => Rec
): number {
  const actualRecords = actualTable.retrieveAllRecords()
  const formRecords = formTable.retrieveAllRecords()

  let numOfThingsSynced = 0

  // create an index of actualdata >> date.
  // Then iterate over all formdata and find the ones that are missing from the index.
  const index: { [date: string]: Rec } = {}
  for (const idKey of Object.getOwnPropertyNames(actualRecords)) {
    const record = actualRecords[idKey]
    const dateIndexKey = String(record.date)
    index[dateIndexKey] = record
  }
  for (const idKey of Object.getOwnPropertyNames(formRecords)) {
    const record = formRecords[idKey]
    const dateIndexKey = String(record.date)
    if (index[dateIndexKey] === undefined) {
      actualTable.createRecord(formRecordToActualRecord(record))
      ++numOfThingsSynced
    }
  }

  return numOfThingsSynced
}

const MINUTES_PER_MOD = 38

function onRetrieveMultiple(resourceNames: string[]) {
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

  export function parseField(x: any, fieldType: FieldType): any {
    switch (fieldType) {
      case FieldType.BOOLEAN:
        return x === "true" || x === true ? true : false
      case FieldType.NUMBER:
        return Number(x)
      case FieldType.STRING:
        return String(x)
      case FieldType.DATE:
        if (x === "" || x === -1) {
          return -1
        } else {
          return Number(x)
        }
      case FieldType.JSON:
        return JSON.parse(x)
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
  const attendanceDaysIndex = {}
  for (const day of Object_values(attendanceDays)) {
    attendanceDaysIndex[day.dateOfAttendance] = day.abDay
  }

  function processRequestFormRecord(r: Rec): Rec {
    return {
      id: -1,
      date: r.date, // the date MUST be the date from the form; this is used for syncing
      subject: r.subject,
      mods: Parsing.modData([
        r.modDataA1To5,
        r.modDataB1To5,
        r.modDataA6To10,
        r.modDataB6To10,
      ]),
      specialRoom: "",
      ...parseStudentConfig(r),
      status: "unchecked",
      homeroom: r.homeroom,
      homeroomTeacher: r.homeroomTeacher,
      chosenBookings: [],
      isSpecial: false,
      annotation: "",
    }
  }
  function processTutorRegistrationFormRecord(r: Rec): Rec {
    function parseSubjectList(d: string[]) {
      return d
        .join(",")
        .split(",") // remember that within each string there are commas
        .map((x) => x.trim())
        .filter((x) => x !== "" && x !== "None")
        .map((x) => String(x))
        .join(", ")
    }
    return {
      id: -1,
      date: r.date,
      ...parseStudentConfig(r),
      mods: parseModData([
        r.modDataA1To5,
        r.modDataB1To5,
        r.modDataA6To10,
        r.modDataB6To10,
      ]),
      modsPref: parseModData([
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
      attendance: {},
      dropInMods: [],
      afterSchoolAvailability: r.afterSchoolAvailability,
      attendanceAnnotation: "",
      additionalHours: 0,
    }
  }
  function processSpecialRequestFormRecord(r: Rec): Rec {
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
    return {
      id: -1,
      date: r.date, // the date MUST be the date from the form
      friendlyFullName: "[special request]",
      friendlyName: "[special request]",
      firstName: "[special request]",
      lastName: "[special request]",
      grade: -1,
      studentId: -1,
      email: "[special request]",
      phone: "[special request]",
      contactPref: "either",
      homeroom: "[special request]",
      homeroomTeacher: "[special request]",
      attendanceAnnotation: "[special request]",
      mods: [parseModInfo(r.abDay, r.mod1To10)],
      subject: r.subject,
      isSpecial: true,
      annotation,
      status: "unchecked",
    }
  }
  function processAttendanceFormRecord(r: Rec): Rec {
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
      mod = parseModInfo(
        attendanceDaysIndex[processedDateOfAttendance],
        r.mod1To10
      )
    }

    // give the tutor their time
    minutesForTutor = MINUTES_PER_MOD
    presenceForTutor = "P"

    // figure out who the tutor is, by student ID
    const xTutors = recordCollectionToArray(tutors).filter(
      (x) => x.studentId === r.studentId
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
        `the tutor ${xMatchings[0].friendlyFullName} is matched twice on the same mod`
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
      markForReset: "",
    }
  }

  let numOfThingsSynced = 0
  numOfThingsSynced += doFormSync(
    tableMap.requestForm(),
    tableMap.requestSubmissions(),
    processRequestFormRecord
  )
  numOfThingsSynced += doFormSync(
    tableMap.specialRequestForm(),
    tableMap.requestSubmissions(),
    processSpecialRequestFormRecord
  )
  numOfThingsSynced += doFormSync(
    tableMap.attendanceForm(),
    tableMap.attendanceLog(),
    processAttendanceFormRecord
  )
  numOfThingsSynced += doFormSync(
    tableMap.tutorRegistrationForm(),
    tableMap.tutors(),
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
    day: Rec
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
    if (tutor.attendance[date] === undefined) {
      tutor.attendance[date] = []
    }
    if (learner !== null && learner.attendance[date] === undefined) {
      learner.attendance[date] = []
    } else {
      for (const attendanceModDataString of tutor.attendance[date]) {
        const tokens = attendanceModDataString.split(" ")
        if (Number(tokens[0]) === mod) {
          alreadyExists = true
        }
      }
    }
    if (!alreadyExists) {
      // add an absence for the tutor
      tutor.attendance[date].push(formatAttendanceModDataString(mod, 0))
      // add an excused absence for the learner, if exists
      if (learnerId !== -1) {
        learners[learnerId].attendance[date].push(
          formatAttendanceModDataString(mod, 1)
        )
      }
      numAttendancesChanged += 2
    }
  }

  function applyAttendanceForStudent(
    collection: RecCollection,
    id: number,
    entry: Rec,
    minutes: number
  ) {
    const record = collection[id]
    if (record === undefined) {
      // This is not a fatal error, especially since we want to retroactively
      // run attendance calculations, but give a log message
      Logger.log(`attendance: record not found ${id}`)
      return
    }
    const attendance = record.attendance
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

  // read tables
  const tutors = Table.readAllRecords("tutors")
  const learners = Table.readAllRecords("learners")
  const attendanceLog = Table.readAllRecords("attendanceLog")
  const attendanceDays = Table.readAllRecords("attendanceDays")
  const matchings = Table.readAllRecords("matchings")
  const tutorsArray = Object_values(tutors)
  const learnersArray = Object_values(learners)
  const matchingsArray = Object_values(matchings)
  const attendanceLogArray = Object_values(attendanceLog)

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
            whenTutorFormNotFilledOutLogic(tutor.id, x.learnerId, i, day)
          }
        }
        if (
          tutor.attendance[day.dateOfAttendance] !== undefined &&
          tutor.attendance[day.dateOfAttendance].length === 0
        ) {
          delete tutor.attendance[day.dateOfAttendance]
        }
      }
    } else if (day.status === "doreset") {
      for (const tutor of tutorsArray) {
        if (tutor.attendance[day.dateOfAttendance] !== undefined) {
          // delete EVERY ABSENCE for that day (but keep excused)
          // this gets rid of anything automatically generated for that day
          tutor.attendance[day.dateOfAttendance] = tutor.attendance[
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

  // Header
  sheet.appendRow(["ARC SCHEDULE"])
  sheet.appendRow([""])

  // Note: Bookings are deliberately ignored.
  const tutorIndex = schedulingTutorIndex(tutors, {}, matchings)
  const unscheduledTutorNames: string[] = []
  const layoutMatrix: [ScheduleEntry[], ScheduleEntry[]][] = [] // [mod0to9][abday]

  for (let i = 0; i < 10; ++i) {
    layoutMatrix[i] = [[], []]
  }

  function matchingsInfo(mod: number, refs: [SchedulingReference, number][]) {
    for (const [sr, srid] of refs) {
      if (sr === SchedulingReference.MATCHING) {
        const matching = matchings[srid]
        if (matching.mod !== mod) continue
        let result = ""
        if (matching.learner !== -1) {
          result += `(w/${learners[matching.learner].name})`
        }
        if (matching.annotation !== "") {
          result += `(${matching.annotation})`
        }
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
              tutorName: tutors[x.id].name,
              info: "(drop-in)",
              isDropIn: true,
            })
            break
          case ModStatus.MATCHED:
            layoutMatrix[i % 10][i < 10 ? 0 : 1].push({
              tutorName: tutors[x.id].name,
              info: "(tutoring) " + matchingsInfo(i + 1, x.refs),
              isDropIn: true,
            })
            break
        }
      }
    } else {
      unscheduledTutorNames.push(tutors[x.id].name)
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
    .setValue(`Unscheduled tutors`)
    .setFontSize(18)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setWrap(true)
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
    )
  ++nextRow
  sheet
    .getRange(nextRow, 2, unscheduledTutorNames.length, 3)
    .mergeAcross()
    .setHorizontalAlignment("center")
  sheet
    .getRange(nextRow, 2, unscheduledTutorNames.length)
    .setFontSize(12)
    .setValues(unscheduledTutorNames.map((x) => [x]))
  nextRow += unscheduledTutorNames.length

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
      numberOfOperations += deDuplicateTableByStudentId(new Table("tutors"))
      numberOfOperations += deDuplicateTableByStudentId(new Table("learners"))
      SpreadsheetApp.getUi().alert(
        `${numberOfOperations} de-duplication operations performed`
      )
    }
  } catch (err) {
    Logger.log(stringifyError(err))
    throw err
  }
}

function deDuplicateTableByStudentId(table: Table): number {
  let numberOfOperations = 0
  type ReplacementMap = { [studentId: number]: { recordIds: number[] } }

  const replacementMap: ReplacementMap = {}

  const recordObject = table.retrieveAllRecords()
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
