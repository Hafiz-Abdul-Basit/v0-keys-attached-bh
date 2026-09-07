/* Esign test documents where NO key is written as <<KEY>>.
 * The app must recognise key names written as normal text
 * (Student Name, studentname, PARENTNAME, @campusname, student_id …)
 * and replace them, while the checkboxes / textboxes get mapped.
 *
 *   node scripts/make-esign-test-docx.js            → mock-docx/
 *   node scripts/make-esign-test-docx.js some/dir   → some/dir
 */
const path = require("path");
const H = require("./docx-mock-helpers");
const {
  png, boxPng, linePng, run, B, I, para, heading, spacer,
  ccCheckbox, ccText, ccDate, legacyCheckbox, legacyText,
  wingSym, wingChar, picture, textBoxShape, tbl, IN, buildDocx,
} = H;
H.setOut(process.argv[2] || path.join(__dirname, "..", "mock-docx"));

const U = '<w:rPr><w:u w:val="single"/></w:rPr>';
const RED = '<w:rPr><w:b/><w:color w:val="C00000"/></w:rPr>';

/* ═══════════ 1. SIMPLE — one page, every key as plain words ═══════════ */
buildDocx("Esign_Simple_NoBrackets.docx", [
  heading("FIELD TRIP PERMISSION SLIP"),
  para(run("Campus Name") + run("  ·  School Year  ·  Todays Date")),
  spacer(),
  para(run("Dear Parent Name,")),
  para(run("Your child ") + run("Student Name", B) + run(" in grade Student Grade has been invited to the science museum trip.")),
  spacer(),
  para(run("Please tick:", B)),
  para(ccCheckbox(false) + run("  My child may attend the trip.")),
  para(ccCheckbox(false) + run("  My child may travel by school bus.")),
  para(ccCheckbox(true) + run("  I have read the safety rules (already ticked).")),
  spacer(),
  para(run("Parent phone: ") + ccText("Click here to enter text.")),
  para(run("Emergency contact: ________________________")),
  para(run("Signature: ____________________     Date: ") + ccDate()),
  spacer(),
  para(run("Thank you,")),
  para(run("Principal Name", B)),
  para(run("Attendance Officer Name · Campus Name")),
].join(""));

/* ═══════════ 2. MEDIUM — mixed spellings, split runs, symbols, table ═══════════ */
buildDocx("Esign_Medium_NoBrackets.docx", [
  heading("ATTENDANCE NOTICE — activeisd"),
  para(run("Notice Date: noticedate      Semester: Active Semester")),
  spacer(),
  // key split across formatted runs the way Word does after editing
  para(run("To the parent/guardian of ") + run("Student", B) + run("First", B) + run("Name ", B) + run("STUDENTLASTNAME", RED) + run(" (student_id), grade @studentgrade at Campus-Name.")),
  para(run("Our records show ") + run("ALLABSENCESCOUNT", B) + run(" unexcused absences this semester. Dates: ") + run("all absences dates", I) + run(".")),
  spacer(),
  para(run("Reason (select one):", B)),
  para(wingSym("F06F") + run("  Illness")),
  para(wingSym("F06F") + run("  Family emergency")),
  para(wingChar("F0FE") + run("  Other — discussed with attendance officer name (pre-checked)")),
  spacer(),
  para(run("☐ I will attend the parent conference      ☑ I received the handbook")),
  para(run("[ ] Call me      [x] Email me      [ ] Send a letter")),
  spacer(),
  tbl([
    [para(run("Parent name", B)), para(run("parentname"))],
    [para(run("Best phone", B)), para(run("__________________"))],
    [para(run("Preferred language", B)), para(ccText("Choose"))],
    [para(run("Teacher", B)), para(run("Teacher Name"))],
  ]),
  spacer(),
  para(run("Sincerely,")),
  para(run("ATTENDANCE OFFICER NAME", B)),
  para(run("for principalname, Campus Name")),
].join(""), {
  header: para(run("campusname · Attendance Office · todaysdate"), '<w:pPr><w:jc w:val="right"/></w:pPr>'),
  footer: para(run("Student ID: student id · School Year: school-year"), '<w:pPr><w:jc w:val="center"/></w:pPr>'),
});

/* ═══════════ 3. COMPLEX — 2 pages, pictures, shapes, legacy fields, labels vs values ═══════════ */
const media = {
  rIdBox: { name: "box.png", data: boxPng(32, false) },
  rIdBoxChecked: { name: "box_checked.png", data: boxPng(32, true) },
  rIdLine: { name: "line.png", data: linePng(320, 24) },
  rIdLogo: { name: "logo.png", data: png(180, 70, (x, y) => ((x >> 3) + (y >> 3)) % 2 ? [31, 56, 100, 255] : [200, 210, 230, 255]) },
  rIdMid: { name: "mid.png", data: boxPng(64, false) },
};
buildDocx("Esign_Complex_NoBrackets.docx", [
  para(picture("rIdLogo", Math.round(1.8 * IN), Math.round(0.7 * IN), "District logo") + run("   ") + run("STUDENT SERVICES ENROLLMENT PACKET", '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr>')),
  para(run("activeisd · Campus Name · School Year")),
  spacer(),
  para(run("SECTION A — STUDENT", B)),
  tbl([
    [para(run("Name of student")), para(run("Student Name"))],
    [para(run("Student ID")), para(run("student_id"))],
    [para(run("Grade")), para(run("studentgrade"))],
    [para(run("Age")), para(run("Student Age"))],
    [para(run("Teacher")), para(run("teacher name"))],
  ]),
  spacer(),
  para(run("SECTION B — PARENT / GUARDIAN", B)),
  para(run("Father: ") + run("Father Name", U) + run("      Mother: ") + run("Mother Name", U) + run("      Guardian: ") + run("guardianname", U)),
  para(run("Parent full address: ") + run("Parent Full Address")),
  para(run("Phone: ") + legacyText("Phone") + run("   Email: ") + legacyText("Email")),
  spacer(),
  para(run("SECTION C — SERVICES REQUESTED", B)),
  tbl([
    [para(run("Service", B)), para(run("Yes", B)), para(run("No", B))],
    [para(run("Counseling")), para(ccCheckbox(false)), para(ccCheckbox(false))],
    [para(run("Tutoring")), para(ccCheckbox(true)), para(ccCheckbox(false))],
    [para(run("Speech therapy")), para(wingSym("F06F")), para(wingSym("F06F"))],
    [para(run("Transportation")), para(run("[ ]")), para(run("[ ]"))],
    [para(run("Free / reduced lunch")), para(picture("rIdBox", Math.round(0.2 * IN), Math.round(0.2 * IN), "")), para(picture("rIdBox", Math.round(0.2 * IN), Math.round(0.2 * IN), ""))],
    [para(run("Photo release")), para(picture("rIdBoxChecked", Math.round(0.2 * IN), Math.round(0.2 * IN), "checked")), para(picture("rIdBox", Math.round(0.2 * IN), Math.round(0.2 * IN), ""))],
  ]),
  spacer(),
  para(run("Details of request:", B)),
  para(textBoxShape(Math.round(6 * IN), Math.round(0.6 * IN), "")),
  para(run("Preferred contact time: ") + picture("rIdLine", Math.round(3 * IN), Math.round(0.22 * IN), "") + run("   Language: ") + ccText("Choose")),
  para(run("Attach a photo of the ID card here: ") + picture("rIdMid", Math.round(0.9 * IN), Math.round(0.9 * IN), "") + run("  (medium picture — the app should ask what it is)")),
  '<w:p><w:r><w:br w:type="page"/></w:r></w:p>',
  heading("SECTION D — ACKNOWLEDGEMENT"),
  para(legacyCheckbox("Ack1", false) + run("  I understand services depend on availability at Campus Name.")),
  para(legacyCheckbox("Ack2", false) + run("  I consent to the release of records for Student Name.")),
  para(legacyCheckbox("Ack3", true) + run("  I received the parent handbook for school year (pre-checked).")),
  spacer(),
  para(run("Parent signature: ________________________     Date: ") + ccDate()),
  para(run("Parent First Name (print): ") + ccText("Click here")),
  spacer(),
  para(run("Markers the team already typed — numbering must continue after them:", I)),
  para(run("<c1> Agree to policy     <c2c> Newsletter (checked)     Preferred name: <t1>")),
  spacer(),
  heading("OFFICE USE ONLY"),
  tbl([
    [para(run("Received by")), para(run("Attendance Officer Name"))],
    [para(run("Reviewed by")), para(run("Assistant Principal Name"))],
    [para(run("Approved")), para(legacyCheckbox("App", false))],
    [para(run("Notes")), para(run("________________________________________"))],
  ]),
  para(run("Principal: PRINCIPALNAME      Date: todays date      Month: monthname")),
].join(""), {
  media,
  header: para(run("activeisd · campus address") , '<w:pPr><w:jc w:val="right"/></w:pPr>'),
  footer: para(run("Student Name · student id · Page 1"), '<w:pPr><w:jc w:val="center"/></w:pPr>'),
});

console.log("done");
