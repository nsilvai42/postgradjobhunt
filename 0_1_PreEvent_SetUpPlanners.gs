/**
 * 
 * Script Name: 0_1_PreEvent_SetUpPlanners.gs
 * 
 * /

/** 
 * ───────────────────────────────────────────────────────────────────────────────
 *  SET UP STUDENT PLANNERS  🗂️
 * ───────────────────────────────────────────────────────────────────────────────
 * QUICK SUMMARY (read this first)
 * -----------------------------------------------------------------------------
 *  • One‑click script you run *once* before the kickoff.
 *  • For every student who *doesn’t* yet have a planner link, it:
 *      ▹ copies the master template
 *      ▹ renames it "First Last – Career Planner"
 *      ▹ saves it inside the Drive folder listed in **Planner Setup**
 *      ▹ pastes the share‑URL back into **All Participants**
 *      ▹ fills the Welcome tab with the student’s name + email
 *  • Shows a pop‑up summary (✅ created, 🔄 skipped, ⚠️ errors).
 *
 * WHAT YOU NEED 📝
 *  • **Planner Setup** sheet             ↦ holds FOLDER ID (B1) & TEMPLATE ID (B2)
 *  • **All Participants** sheet          ↦ must include columns:
 *        "First Name", "Last Name", "Email", "Planner Link"
 *
 * HOW TO RUN ▶️
 * 1. Open the Admin spreadsheet.
 * 2. Admin Tools → 🚀 Set Up Student Planners
 * 3. Authorise the script if prompted (first run only).
 */

//  ╭───────────────────────────────╮
//  │  setupStudentPlanners()       │
//  ╰───────────────────────────────╯
/**
 * This function sets up student planners by making a copy of the template and personalizing it to each participant.
 * 
 * STEPS:
 *   1. Checks required sheets and IDs
 *   2. Maps header names → column numbers
 *   3. Reads First Name → Planner Link range
 *   4. Copies the template, personalises it, writes the URL
 *   5. Pops up a ✅ / 🔄 / ⚠️ summary
 */
function setupStudentPlanners() {
  const ui = SpreadsheetApp.getUi();                          // UI for alerts
  const ss = SpreadsheetApp.getActive();                      // Active spreadsheet

  // ────────────────────────── A. VALIDATE SHEETS & CONFIG ──────────────────────────
  
  // STEP 1️⃣  Are the two required sheets present?
  // 🚨 IMPORTANT : All sheets should be named the same exact text in the 0_GlobalConstants file!
  const setupSheet = ss.getSheetByName(PLANNER_SETUP_SHEET);          // Get Planner Setup sheet
  const partsSheet = ss.getSheetByName(ALL_PARTICIPANTS_SHEET);       // Get All Participants sheet 

  // Show error if either sheet is missing.
  if (!setupSheet || !partsSheet) {
    ui.alert("Error", "Missing 'Planner Setup' or 'All Participants' sheet. Did you change the name of either sheet?", 
    ui.ButtonSet.OK);
    return;
  }

  // STEP 2️⃣  Do we have the Folder ID and Template ID?
  // 🚨 IMPORTANT : Make sure both the file and folder IDs are correct in the Planner Setup sheet!
  const folderId = setupSheet.getRange(PLANNER_FOLDER_ID_CELL).getValue();      // Get Folder ID from cell
  const templateId = setupSheet.getRange(PLANNER_TEMPLATE_ID_CELL).getValue();  // Get Template ID from cell

  // Show error if either cell is missing a value.
  if (!folderId || !templateId) {
    ui.alert("Error", "No folder or template ID found. Make sure these are added to the Planner Setup sheet.", 
    ui.ButtonSet.OK);
    return;
  }

  // STEP 3️⃣  Can we open the template file and the folder?
  // 🚨 IMPORTANT : Both the template file and folder should be set to anyone with link can edit. 
  //.               Make sure both the file and folder IDs are correct in the Planner Setup sheet!
  let tplFile, tgtFolder;

  // Try to open the template, and show error if not able to.
  try { tplFile = DriveApp.getFileById(templateId); }  
  catch (e) { ui.alert("Error", 
              `Cannot open template: ${e}. 
              Check file access and make sure its ID is correct in the Planner Setup sheet.`); 
              return; } // Stop if error is shown.

  // Try to open the folder, and show error if not able to.
  try { tgtFolder = DriveApp.getFolderById(folderId); }
  catch (e) { ui.alert("Error", 
              `Cannot open folder: ${e}. 
              Check folder access and make sure its ID is correct in the Planner Setup sheet`); 
              return; } // Stop if error is shown.


  // ────────────────────────── B. MAP HEADERS → COLUMN NUMBERS ──────────────────────────
  
  // STEP 1️⃣  Locate all the columns based on header names
  // 🚨 IMPORTANT : All headers should be named the same exact text in the 0_GlobalConstants file!

  const headers = partsSheet.getRange(HEADER_ROW, 1, 1, partsSheet.getLastColumn()).getValues()[0];

  // Create a map to store header names and their column numbers (e.g., { "First Name": 2 }).
  const headerMap = {};
  headers.forEach((h, i) => headerMap[h] = i + 1);

  // STEP 2️⃣  Check for the required headers by name.
  const requiredHeaders = [COL_FIRST_NAME, COL_LAST_NAME, COL_EMAIL, COL_PLANNER];
  const missingHeaders = requiredHeaders.filter(header => !headerMap[header]);

  // STEP 3️⃣ If any required headers are missing, show an error and stop.
  if (missingHeaders.length > 0) {
    ui.alert("Error: Missing Headers",
             `The following headers were not found in the '${ALL_PARTICIPANTS_SHEET}' sheet: ` +
             `"${missingHeaders.join('", "')}".\n\nAre the headers spelled exactly like displayed above?.`);
    return;
  }

  // STEP 4️⃣ Store the column numbers for later use.
  const colFirst = headerMap[COL_FIRST_NAME];  // First Name column
  const colLast  = headerMap[COL_LAST_NAME];   // Last Name column
  const colEmail = headerMap[COL_EMAIL];       // Email column
  const colLink  = headerMap[COL_PLANNER];     // Planner Link column


  // ────────────────────────── C. FETCH PARTICIPANT DATA ──────────────────────────

  // STEP 1️⃣: Find the last row with content to determine how many participants there are.
  const lastRow = partsSheet.getLastRow();

  // If there are no participants listed after the header row, show an info message and stop.
  if (lastRow < DATA_START_ROW) {
    ui.alert("Info", `No participants found to process in the '${ALL_PARTICIPANTS_SHEET}' sheet.`, ui.ButtonSet.OK);
    return;
  }

  // STEP 2️⃣ Fetch only the block of data we need (from the first name column to the planner link column).
  const numRows = lastRow - DATA_START_ROW + 1; // Calculate how many rows of data to get.
  const width   = colLink - colFirst + 1;       // Calculate how many columns wide the data range is.
  const data    = partsSheet.getRange(DATA_START_ROW, 
                                      colFirst, 
                                      numRows, 
                                      width).getValues(); // Get all participant data in one efficient call.


  // ────────────────────────── D. COPY & PERSONALISE PLANNERS ──────────────────────────

  // Initialize counters to track the outcome for each student.
  let success = 0, 
      skipped = 0, 
      errors = 0;

  // STEP 1️⃣ Loop through each row of participant data.
  data.forEach((row, i) => {
    const rNum  = DATA_START_ROW + i;                     // Row's number in the sheet.
    const first = row[0];                                 // Get First Name.
    const last  = row[colLast - colFirst];                // Get Last Name.
    const email = row[colEmail - colFirst];               // Get Email.
    const linkCell= partsSheet.getRange(rNum, colLink);   // Get Planner Link cell.

    // STEP 2️⃣ Skip participant if missing key info or already have a planner.
    if (!first || !email)    { skipped++; return; }       // Skip if no first name or email.
    if (linkCell.getValue()) { skipped++; return; }       // Skip if planner link already exists.

    // STEP 3️⃣ Create a copy of the planner. 
    try {
      // Copy the template, name it with the student's full name, and save it to the target folder.
      const copy = tplFile.makeCopy(`${first} ${last} - Career Planner`, tgtFolder);
      const url = copy.getUrl();                         // Get  URL of the created planner.
      linkCell.setValue(url);                            // Write planner URL into student's row in Admin Panel.

      // STEP 4️⃣ Personalize the new planner's "Welcome" sheet.
      const stSS = SpreadsheetApp.openById(copy.getId());          // Open the new planner to edit it.

      const welcomeSheet = stSS.getSheetByName(PLANNER_WELCOME_SHEET); // Find the welcome sheet by name
      if (welcomeSheet) {
        welcomeSheet.getRange(PLANNER_FULLNAME_CELL).setValue(`${first} ${last}`);  // Write the student's full name 
        welcomeSheet.getRange(PLANNER_WELCOME_EMAIL).setValue(email);               // Write the student's email
      }
      success++; // Count successes.

      // If any errors occur:
    } catch (e) {
      errors++;                                                         // Count the error.
      linkCell.setValue(`ERROR: ${e.message.substring(0, 100)}...`);    // Add brief error to Planner Link cell.
      Logger.log(`Error on row ${rNum} for ${first} ${last}: ${e.stack || e}`);   // Log error here to debug.
    }
  });

  // ────────────────────────── E. SHOW SUMMARY ──────────────────────────

  // STEP 1️⃣ Display a final pop-up message to the user with the counts of successes, skips, and errors.
  ui.alert("Setup Complete",
    `✅ Success: ${success}  🔄 Skipped: ${skipped}  ⚠️ Errors: ${errors}`,
    ui.ButtonSet.OK
  );
}
