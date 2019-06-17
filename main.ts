const tutorTable = new Table('tutors');
const learnerTable  = new Table('learners');
const requestTable = new Table('requests');
const bookingTable = new Table('bookings');
const matchingTable = new Table('matchings');
const tutorToMatchingIndex = buildTutorToMatchingIndex();
const tutorToBookingsIndex = buildTutorToBookingsIndex();

function buildTutorToMatchingIndex(): To<number> {
    const r = {};
    for (const tutor of tutorTable.records) {
        r[tutor.id] = null;
    }
    for (const matching of matchingTable.records) {
        if (r[matching.tutor] !== null) {
            throw new Error(`multiple matchings for tutor ${matching.tutor}`);
        }
        r[matching.tutor] = matching.id;
    }
    return r;
}

function buildTutorToBookingsIndex(): To<number[]> {
    const r = {};
    for (const tutor of tutorTable.records) {
        r[tutor.id] = [];
    }
    for (const booking of bookingTable.records) {
        r[booking.tutor].push(booking.id);
    }
    return r;
}

function suggestTutorsForRequest(request: Rec) {
    function tutorWorksForReq(tutorId: number, modId: number) {
        const tutor = tutorTable.find(tutorId);
        if (!tutor.mods.contains(modId)) return false;
        if (tutorToMatchingIndex[tutor.id] !== null) return false;
        if (tutorToBookingsIndex[tutor.id].length > 0) return false;
    }
    const r = {};
    for (const modId of request.mods) {
        r[modId] = [];
        for (const tutor of tutorTable.records) {
            if (tutorWorksForReq(tutor.id, modId)) {
                r[modId].push(tutor.id);
            }
        }
    }
    return r;
}

function apiEndpoint(path: any[]) {
    if (path[0] == 'tutors') {
        path.shift();
        return tutorTable.apiEndpoint(path);
    }
    if (path[0] == 'learners') {
        path.shift();
        return learnerTable.apiEndpoint(path);
    }
    if (path[0] == 'requests') {
        path.shift();
        if (path[0] == 'unchecked') {
            return requestTable.queryRecords(request => request.status == 'unchecked');
        }
        if (path[0] == 'unbooked') {
            return requestTable.queryRecords(request => request.status == 'unbooked');
        }
        return requestTable.apiEndpoint(path);
    }
    if (path[0] == 'bookings') {
        path.shift();
        if (path[0] == 'unsent') {
            return bookingTable.queryRecords(x => x.status == 'unsent');
        }
        if (path[0] == 'unreplied') {
            return bookingTable.queryRecords(x => x.status == 'noReply' || x.status == 'tutorReply' || x.status == 'learnerReply');
        }
        if (path[0] == 'unfinalized') {
            return bookingTable.queryRecords(x => x.status == 'unfinalized');
        }
        return bookingTable.apiEndpoint(path);
    }
    if (path[0] == 'matchings') {
        path.shift();
        if (path[0] == 'unfinalized') {
            matchingTable.queryRecords(x => x.status == 'unfinalized');
        }
        return matchingTable.apiEndpoint(path);
    }
}

function test() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt('Enter API endpoint path');
    ui.alert(JSON.stringify(apiEndpoint(JSON.parse(response.getResponseText()))));
}

function onOpen(_e: any) {
    SpreadsheetApp.getUi().createMenu('APP TEST')
      .addItem('API', 'test')
      .addToUi();
}