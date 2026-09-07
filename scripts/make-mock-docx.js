/* Generate realistic mock .docx files for testing the Esign Templates module. */
const path = require("path");
const H = require("./docx-mock-helpers");
const { png, boxPng, linePng, run, B, I, para, heading, spacer, ccCheckbox, ccText, ccDate, legacyCheckbox, legacyText, wingSym, wingChar, picture, textBoxShape, tbl, IN, buildDocx } = H;
H.setOut(process.argv[2] || path.join(__dirname, "..", "mock-docx"));

/* ═══════════════ Mock 1: Word content controls (Developer tab) ═══════════════ */
buildDocx("Mock1_ContentControls_ConsentForm.docx", [
  heading("PARENT CONSENT FORM"),
  para(run("Campus: ") + run("<<CAMPUSNAME>>", B) + run("        Date: ") + run("<<CURRENTDATE>>")),
  para(run("Dear ") + run("<<PARENTNAME>>", B) + run(",")),
  para(run("Your child ") + run("<<STUDENTNAME>>", I) + run(" (Grade ") + run("<<STUDENTGRADE>>") + run(", ID ") + run("<<STUDENTID>>") + run(") has been selected for the after-school program. Please review and complete the sections below.")),
  spacer(),
  para(run("Please check all that apply:", B)),
  para(ccCheckbox(false) + run("  I give permission for my child to participate.")),
  para(ccCheckbox(false) + run("  I give permission for transportation by school bus.")),
  para(ccCheckbox(true) + run("  I have read the program handbook (pre-checked).")),
  para(ccCheckbox(false) + run("  My child has allergies or medical conditions (explain below).")),
  spacer(),
  tbl([
    [para(run("Parent / Guardian full name:", B)), para(ccText("Click or tap here to enter text."))],
    [para(run("Phone number:", B)), para(ccText("Click or tap here to enter text."))],
    [para(run("Emergency contact:", B)), para(ccText("Click or tap here to enter text."))],
    [para(run("Date signed:", B)), para(ccDate())],
  ]),
  spacer(),
  para(run("Medical notes: ") + ccText("Enter any allergies or conditions.")),
  spacer(),
  para(run("Sincerely,")),
  para(run("<<PRINCIPALNAME>>", B)),
  para(run("Principal, <<CAMPUSNAME>>")),
].join(""));

/* ═══════════════ Mock 2: symbols, typed boxes, underscores, split keys ═══════════════ */
buildDocx("Mock2_Symbols_TypedBoxes_Attendance.docx", [
  heading("ATTENDANCE WARNING NOTICE"),
  para(run("<<ACTIVEISD>>") + run("  ·  ") + run("<<CAMPUSNAME>>")),
  // key split across runs the way Word does after editing
  para(run("To the parent of ") + run("<<STUDENT", B) + run("FIRSTNAME>> ", B) + run("<<STUDENTLAST") + run("NAME>>") + run(", grade @STUDENTGRADE:")),
  para(run("Our records show ") + run("<<ALLABSENCESCOUNT>>", B) + run(" absences this semester (") + run("<<Active Semester>>") + run("). Dates: ") + run("<<ALLABSENCESDATES>>", I)),
  spacer(),
  para(run("Reason for absences (check one):", B)),
  para(wingSym("F06F") + run("  Illness")),
  para(wingSym("F06F") + run("  Family emergency")),
  para(wingSym("F0FE") + run("  Other — already discussed with ATTENDANCEOFFICERNAME (pre-checked)")),
  spacer(),
  para(run("Typed with the keyboard:", B)),
  para(run("☐ I will attend the parent conference      ☑ I have received the handbook")),
  para(run("[ ] Please call me      [x] Please email me      [ ] Send a letter")),
  spacer(),
  para(run("Parent signature: ______________________      Date: __________")),
  para(run("Best phone number: ") + run("__________________") + run("   Alternate: ") + run("_______________")),
  spacer(),
  para(wingChar("F06F") + run("  Wingdings box typed as a character      ") + wingChar("F0FE") + run("  checked character")),
  spacer(),
  para(run("Attendance Officer: ") + run("<<ATTENDANCEOFFICERNAME>>", B)),
].join(""), {
  header: para(run("<<CAMPUSNAME>> · Attendance Office · <<CURRENTDATE>>"), '<w:pPr><w:jc w:val="right"/></w:pPr>'),
  footer: para(run("Page ") + run("1") + run(" · Student ID <<STUDENTID>>"), '<w:pPr><w:jc w:val="center"/></w:pPr>'),
});

/* ═══════════════ Mock 3: pictures of boxes, shapes, legacy fields, pre-mapped markers ═══════════════ */
const media3 = {
  rIdBox: { name: "box.png", data: boxPng(32, false) },
  rIdBoxChecked: { name: "box_checked.png", data: boxPng(32, true) },
  rIdLine: { name: "line.png", data: linePng(320, 24) },
  rIdLogo: { name: "logo.png", data: png(160, 60, (x, y) => ((x >> 3) + (y >> 3)) % 2 ? [31, 56, 100, 255] : [200, 210, 230, 255]) },
};
buildDocx("Mock3_Pictures_Shapes_LegacyFields.docx", [
  para(picture("rIdLogo", Math.round(1.6 * IN), Math.round(0.6 * IN), "School logo") + run("   ") + run("STUDENT ENROLLMENT FORM", '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr>')),
  para(run("Student: ") + run("<<STUDENTNAME>>", B) + run("   DOB: ") + run("<<STUDENTDOB>>") + run("   Campus: ") + run("<<CAMPUSNAME>>")),
  spacer(),
  para(run("The client pasted pictures of boxes instead of real checkboxes:", B)),
  para(picture("rIdBox", Math.round(0.2 * IN), Math.round(0.2 * IN), "") + run("  Enroll in free lunch program")),
  para(picture("rIdBox", Math.round(0.2 * IN), Math.round(0.2 * IN), "") + run("  Enroll in transportation")),
  para(picture("rIdBoxChecked", Math.round(0.2 * IN), Math.round(0.2 * IN), "checked") + run("  Photo release consent (client marked it checked)")),
  spacer(),
  para(run("…and a picture of a line for the text field:", B)),
  para(run("Guardian name: ") + picture("rIdLine", Math.round(3.3 * IN), Math.round(0.25 * IN), "")),
  spacer(),
  para(run("Old-style Word form fields (Legacy Tools):", B)),
  para(legacyCheckbox("Check1", false) + run("  Sibling attends this campus")),
  para(legacyCheckbox("Check2", true) + run("  Returning student (pre-checked)")),
  para(run("Home address: ") + legacyText("Text1")),
  para(run("Email: ") + legacyText("Text2")),
  spacer(),
  para(run("A drawn text box shape (Insert → Text Box):", B)),
  para(run("Comments: ") + textBoxShape(Math.round(3 * IN), Math.round(0.3 * IN), "")),
  spacer(),
  para(run("Markers the team already typed in Word — numbering must continue after these:", B)),
  para(run("<c1> Agree to policy     <c2c> Newsletter (checked)     Preferred name: <t1>")),
  spacer(),
  para(run("Prepared by <<ATTENDANCEOFFICERNAME>> for <<PARENTNAME>>.")),
].join(""), { media: media3 });

/* ═══════════════ Mock 4: everything mixed in a realistic 2-page form with tables ═══════════════ */
const media4 = { rIdBox: { name: "box.png", data: boxPng(28, false) } };
buildDocx("Mock4_Mixed_RealisticForm.docx", [
  heading("<<ACTIVEISD>> — STUDENT SERVICES REQUEST"),
  para(run("Student: <<STUDENTNAME>>   Grade: <<STUDENTGRADE>>   ID: <<STUDENTID>>   Campus: <<CAMPUSNAME>>")),
  para(run("Parent/Guardian: <<PARENTNAME>>   Phone: <<PARENTPHONE>>   Email: <<PARENTEMAIL>>")),
  spacer(),
  tbl([
    [para(run("Service", B)), para(run("Requested", B))],
    [para(run("Counseling")), para(ccCheckbox(false))],
    [para(run("Tutoring")), para(ccCheckbox(true))],
    [para(run("Speech therapy")), para(wingSym("F06F"))],
    [para(run("Transportation")), para(run("[ ]"))],
    [para(run("Free/reduced lunch")), para(picture("rIdBox", Math.round(0.18 * IN), Math.round(0.18 * IN), ""))],
  ]),
  spacer(),
  para(run("Details of request:", B)),
  para(ccText("Describe the request.")),
  para(run("Preferred contact time: ______________     Language: ") + ccText("Choose")),
  spacer(),
  para(run("Acknowledgement", B)),
  para(ccCheckbox(false) + run("  I understand services depend on availability.")),
  para(ccCheckbox(false) + run("  I consent to the release of records to <<CAMPUSNAME>> staff.")),
  spacer(),
  para(run("Parent signature: ") + legacyText("Sig") + run("     Date: ") + ccDate()),
  '<w:p><w:r><w:br w:type="page"/></w:r></w:p>',
  heading("OFFICE USE ONLY"),
  tbl([
    [para(run("Received by")), para(run("<<ATTENDANCEOFFICERNAME>>"))],
    [para(run("Reviewed")), para(legacyCheckbox("Rev", false))],
    [para(run("Approved")), para(legacyCheckbox("App", false))],
    [para(run("Notes")), para(run("________________________________________"))],
  ]),
  para(run("Principal: <<PRINCIPALNAME>>      Date: <<CURRENTDATE>>")),
].join(""), { media: media4 });

console.log("done");
