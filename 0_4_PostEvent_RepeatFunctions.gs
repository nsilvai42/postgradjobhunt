/**
 * 
 * Script Name: 0_4_PostEvent_RepeatFunctions.gs
 * 
 * 
/

//  ╭───────────────────────────────╮
//  │  fetchAllStudentTaskData()    │
//  ╰───────────────────────────────╯
/**
 * Aggregates task data from active student's planners in 'Task Data' and updates current progress in 'Active Participants'.
 *
 * WHAT YOU NEED 📝
 *    📊 "Active Participants" sheet with: PTID, Planner Link, Current Progress.
 *    📊 "Task Data" sheet (columns will be cleared and repopulated).
 *
 * HOW TO RUN ▶️
 *    1. Run right after the kick-off for baseline progress. This will be re-run automatically before each weekly nudge.
 *    2. Admin Tools → 📊 Update Student Task Data
 *    3. Copy and paste the Current Progress columns into S0_Progress column.
 *
 * STEPS:
 *    1. Checks for all required sheets and columns.
 *    2. Reads all active participant data.
 *    3. Loops through each participant, opens their planner, and scrapes all task data.
 *    4. Builds a master list of all tasks, deadlines and Done statuses and calculates progress percentages.
 *    5. Writes the master task list and the new percentages back to the sheets.
 */
function fetchAllStudentTaskData() {

  /* ───────────── Helper: detect UI vs trigger run ───────────── */
  let ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    ui = null;                       // trigger / headless execution
  }

  const ss = SpreadsheetApp.getActive();
  Logger.log('▶️ fetchAllStudentTaskData started ' + new Date());

  // ────────────────────────── A. VALIDATE SHEETS & CONFIG ──────────────────────────
  // STEP 1️⃣ Find the source and destination sheets.
  const partsSheet = ss.getSheetByName(ACTIVE_PARTICIPANTS_SHEET);
  const taskSheet  = ss.getSheetByName(TASK_DATA_SHEET);

  if (!partsSheet || !taskSheet) {
    Logger.log('❌ Missing sheet(s): ' + [ACTIVE_PARTICIPANTS_SHEET, TASK_DATA_SHEET].join(', '));
    if (ui) ui.alert('Error: Missing Sheet',
      `Could not find '${ACTIVE_PARTICIPANTS_SHEET}' or '${TASK_DATA_SHEET}'.`,
      ui.ButtonSet.OK);
    return;
  }
  Logger.log('✅ Sheets located OK');

  // ────────────────────────── B. MAP HEADERS → COLUMN NUMBERS ──────────────────────────
  // STEP 1️⃣ Locate columns in the 'Active Participants' sheet based on header names.
  const pHeaders = partsSheet.getRange(HEADER_ROW, 1, 1,
                    partsSheet.getLastColumn()).getValues()[0];
  const pIdx = {};
  pHeaders.forEach((h, i) => pIdx[h] = i);         // zero-based

  const requiredP = [COL_PTID, COL_FIRST_NAME, COL_LAST_NAME,
                     COL_EMAIL, COL_PLANNER, COL_CURRENT_PROGRESS];
  const missingP = requiredP.filter(h => pIdx[h] === undefined);
  if (missingP.length) {
    Logger.log('❌ Missing headers in Active Participants: ' + missingP.join(', '));
    if (ui) ui.alert('Error: Missing Headers',
      `The '${ACTIVE_PARTICIPANTS_SHEET}' sheet is missing: "${missingP.join('", "')}".`,
      ui.ButtonSet.OK);
    return;
  }

  // STEP 2️⃣ Locate columns in the 'Task Data' sheet.
  const taskHeaders = taskSheet.getRange(HEADER_ROW, 1, 1,
                       taskSheet.getLastColumn()).getValues()[0];
  const tIdx = {};
  taskHeaders.forEach((h, i) => tIdx[h] = i + 1);  // one-based for getRange

  const columnsToManage = [COL_PTID, COL_FIRST_NAME, COL_LAST_NAME, COL_EMAIL,
                           COL_TASK_ID, COL_TASK, COL_DEADLINE, COL_DONE, COL_LAST_UPDATE];
  const missingT = columnsToManage.filter(h => tIdx[h] === undefined);
  if (missingT.length) {
    Logger.log('❌ Missing headers in Task Data: ' + missingT.join(', '));
    if (ui) ui.alert('Error: Missing Headers',
      `The '${TASK_DATA_SHEET}' sheet is missing: "${missingT.join('", "')}".`,
      ui.ButtonSet.OK);
    return;
  }
  Logger.log('✅ Headers verified');

  // ────────────────── C. PREPARE FOR PROCESSING ────────────────────
  // STEP 1️⃣ Fetch all participant data.
  const participantData = partsSheet.getRange(
      DATA_START_ROW, 1,
      partsSheet.getLastRow() - DATA_START_ROW + 1,
      partsSheet.getLastColumn()).getValues();
  Logger.log(`➡️ Loaded ${participantData.length} participant rows`);

  // STEP 2️⃣ Prep data holders.
  const taskDataToWrite = {};
  columnsToManage.forEach(col => taskDataToWrite[col] = []);
  const progressMap = {};

  // ────────────────── D. COLLECT TASK DATA FROM PLANNERS ───────────────────
  participantData.forEach(row => {
    const ptid = row[pIdx[COL_PTID]];
    const url  = row[pIdx[COL_PLANNER]];
    if (!ptid || !url) return;

    progressMap[ptid] = { done: 0, total: 0 };

    try {
      const stSS = SpreadsheetApp.openByUrl(url);
      const sp   = stSS.getSheetByName(PLANNER_SPRINT_SHEET);
      if (!sp) return;

      const taskData = sp.getDataRange().getValues();
      const hRow = taskData.findIndex(r =>
        r.some(c => String(c).trim().toLowerCase() === PLANNER_COL_ID.toLowerCase()));
      if (hRow < 0) return;

      const idxMap = taskData[hRow].reduce(
        (m, c, i) => { m[String(c).trim().toLowerCase()] = i; return m; }, {});

      taskData.slice(hRow + 1).forEach(taskRow => {
        const taskId = taskRow[idxMap[PLANNER_COL_ID.toLowerCase()]];
        if (!taskId) return;

        const doneVal = taskRow[idxMap[PLANNER_COL_DONE.toLowerCase()]];
        const isDone  = (doneVal === true || String(doneVal).toLowerCase() === 'true');

        progressMap[ptid].total++;
        if (isDone) progressMap[ptid].done++;

        const deadlineVal = taskRow[idxMap[PLANNER_COL_DEADLINE.toLowerCase()]];

        taskDataToWrite[COL_PTID].push([ptid]);
        taskDataToWrite[COL_FIRST_NAME].push([row[pIdx[COL_FIRST_NAME]]]);
        taskDataToWrite[COL_LAST_NAME].push([row[pIdx[COL_LAST_NAME]]]);
        taskDataToWrite[COL_EMAIL].push([row[pIdx[COL_EMAIL]]]);
        taskDataToWrite[COL_TASK_ID].push([taskId]);
        taskDataToWrite[COL_TASK].push([taskRow[idxMap['task']]]);
        taskDataToWrite[COL_DEADLINE].push([deadlineVal]);
        taskDataToWrite[COL_DONE].push([isDone]);
        taskDataToWrite[COL_LAST_UPDATE].push([new Date()]);
      });

    } catch (e) {
      Logger.log(`Task fetch error for PTID ${ptid}: ${e.message}`);
    }
  });
  Logger.log(`✅ Collected ${taskDataToWrite[COL_PTID].length} task rows`);

  // ────────────────── E. WRITE AGGREGATED DATA & UPDATE PROGRESS ──────────────────
  // STEP 1️⃣ Clear old data in managed columns.
  if (taskSheet.getLastRow() > 1) {
    const numRowsToClear = taskSheet.getLastRow() - 1;
    columnsToManage.forEach(colName => {
      const colIndex = tIdx[colName];
      if (colIndex) {
        taskSheet.getRange(DATA_START_ROW, colIndex, numRowsToClear, 1).clearContent();
      }
    });
    Logger.log('🧹 Cleared old task data');
  }

  // STEP 2️⃣ Write fresh task data.
  if (taskDataToWrite[COL_PTID].length) {
    const numTasks = taskDataToWrite[COL_PTID].length;
    for (const colName in taskDataToWrite) {
      const colIndex = tIdx[colName];
      if (colIndex) {
        taskSheet.getRange(DATA_START_ROW, colIndex, numTasks, 1)
                 .setValues(taskDataToWrite[colName]);
      }
    }
    Logger.log(`✍️ Wrote ${numTasks} new task rows`);
  }

  // STEP 3️⃣ Compute progress percentages.
  const progressValues = participantData.map(row => {
    const ptid = row[pIdx[COL_PTID]];
    if (progressMap[ptid]) {
      const { done, total } = progressMap[ptid];
      return [ total ? done / total : 0 ];
    }
    return [ row[pIdx[COL_CURRENT_PROGRESS]] ];
  });

  // STEP 4️⃣ Write progress back.
  const progressRange = partsSheet.getRange(
      DATA_START_ROW, pIdx[COL_CURRENT_PROGRESS] + 1,
      progressValues.length, 1);
  progressRange.setValues(progressValues).setNumberFormat('0%');
  Logger.log('📊 Progress column updated');

  // STEP 5️⃣ Final toast / log.
  const summary = `Fetched ${taskDataToWrite[COL_PTID].length} tasks and updated ${participantData.length} students.`;
  if (ui){
    ss.toast(summary, 'Fetch Complete', 5);
  }
  Logger.log('✅ ' + summary);
}

//  ╭───────────────────────────────╮
//  │  matchJobFormResults()        │
//  ╰───────────────────────────────╯
/**
 * Matches job offer submissions from a Google Form into the 'Active Participants' sheet.
 * It uses a multi-stage matching logic: exact email, exact name, and finally "fuzzy" matching
 * for names and emails that are slightly different (e.g., typos, nicknames).
 *
 * WHAT YOU NEED 📝
 * 💼 "Job Form" sheet with form responses, including: Name, Email, Job Offer status.
 * 💼 "Active Participants" sheet with: PTID, First Name, Last Name, Email, Job Secured?, Job Secured Date.
 *
 * HOW TO RUN ▶️
 * 1. After students have submitted the "Job Secured" Google Form.
 * 2. Menu: Admin Tools → ✨ Match Job Form Results (You would add this to your onOpen function)
 *
 * STEPS:
 * 1. Validates all required sheets and columns exist.
 * 2. Reads all participant and form submission data into memory.
 * 3. Builds fast lookup maps to find a participant's row number by their email or name.
 * 4. Loops through each form submission that hasn't been matched yet.
 * 5. Tries to find a matching participant using a 4-step process:
 * a. Exact email match.
 * b. Exact full name match.
 * c. Fuzzy name match (allows for small typos).
 * d. Fuzzy email match (allows for small typos).
 * 6. When a match is found, it updates the 'Active Participants' sheet to TRUE for 'Job Secured?'
 * and sets the 'Job Secured Date' to the form submission timestamp.
 * 7. It also marks the row in the 'Job Form' sheet as matched to prevent re-processing.
 * 8. Writes all updates back to the sheets and provides a summary.
 */
function matchJobFormResults() {
  const ui = (function() { try { return SpreadsheetApp.getUi(); } catch (e) { return null; } })();
  const ss = SpreadsheetApp.getActive();

  // ────────────── A. VALIDATE SHEETS ──────────────
  const formSheet = ss.getSheetByName(JOB_FORM_SHEET);
  const partSheet = ss.getSheetByName(ACTIVE_PARTICIPANTS_SHEET);
  if (!formSheet || !partSheet) {
    const err = `Missing '${JOB_FORM_SHEET}' or '${ACTIVE_PARTICIPANTS_SHEET}'.`;
    if (ui) ui.alert('Error: Missing Sheet', err, ui.ButtonSet.OK);
    Logger.log('❌ ' + err);
    return;
  }

  // ────────────── B. HEADER MAPS ──────────────
  // STEP 1️⃣ Create maps of header names to column indices for easy data access.
  const mapHdr = sh => {
    const h = sh.getRange(HEADER_ROW, 1, 1, sh.getLastColumn()).getValues()[0];
    const m = {};
    h.forEach((c, i) => m[String(c).trim()] = i);
    return m;
  };
  const fIdx = mapHdr(formSheet);
  const pIdx = mapHdr(partSheet);

  // STEP 2️⃣ Verify all necessary columns are present in both sheets.
  const needForm = [COL_FORM_NAME, COL_FORM_EMAIL, COL_FORM_JOB, COL_FORM_MATCHED, COL_TIMESTAMP];
  const needPart = [COL_EMAIL, COL_FIRST_NAME, COL_LAST_NAME, COL_JOB_SECURED, COL_JOB_START];
  const missF = needForm.filter(h => fIdx[h] === undefined);
  const missP = needPart.filter(h => pIdx[h] === undefined);
  if (missF.length || missP.length) {
    const err = `JobForm missing: ${missF.join(', ')} | Parts missing: ${missP.join(', ')}`;
    if (ui) ui.alert('Error: Missing Headers', err, ui.ButtonSet.OK);
    Logger.log('❌ ' + err);
    return;
  }

  // ────────────── C. READ DATA ──────────────
  // STEP 1️⃣ Load all data from both sheets into memory at once for speed.
  const formData = formSheet.getRange(DATA_START_ROW, 1,
    formSheet.getLastRow() - DATA_START_ROW + 1, formSheet.getLastColumn()).getValues();
  const partData = partSheet.getRange(DATA_START_ROW, 1,
    partSheet.getLastRow() - DATA_START_ROW + 1, partSheet.getLastColumn()).getValues();

  // ────────────── D. BUILD LOOK-UP MAPS ──────────────
  // STEP 1️⃣ Helper functions to normalize names and emails for consistent matching.
  const normEmail = e => String(e || '').trim().toLowerCase();
  const normName = n => String(n || '').replace(/\s+/g, '').toLowerCase();

  // STEP 2️⃣ Create maps for instant lookups: {normalized_email: rowIndex, ...}
  const emailMap = {},
    nameMap = {};
  partData.forEach((r, i) => {
    const em = normEmail(r[pIdx[COL_EMAIL]]);
    if (em) emailMap[em] = i;
    const nm = normName(r[pIdx[COL_FIRST_NAME]] + r[pIdx[COL_LAST_NAME]]);
    if (nm) nameMap[nm] = i;
  });

  // ────────────── E. PROCESS FORM ROWS ──────────────
  let matched = 0,
    unmatched = 0;
  formData.forEach((r, i) => {
    // STEP 1️⃣ Skip rows that have already been successfully matched.
    if (String(r[fIdx[COL_FORM_MATCHED]]).toLowerCase() === 'true') return;

    // STEP 2️⃣ Normalize the name and email from the current form submission row.
    const fEmail = normEmail(r[fIdx[COL_FORM_EMAIL]]);
    const fName = normName(removeEmojis(r[fIdx[COL_FORM_NAME]]));

    // STEP 3️⃣ Attempt to find a match using a multi-step waterfall logic.
    let rowIdx = emailMap[fEmail]; // a. Try exact email match first.
    if (rowIdx === undefined) rowIdx = nameMap[fName]; // b. If no email match, try exact name match.

    // c. If still no match, try fuzzy name matching (allows for typos).
    if (rowIdx === undefined) {
      let best = 1e9,
        bestIdx = null;
      for (const k in nameMap) {
        const d = levenshtein(fName, k);
        if (d < best) { best = d; bestIdx = nameMap[k]; }
      }
      if (best <= 2) rowIdx = bestIdx; // Allow up to 2 character differences
    }
    // d. As a last resort, try fuzzy email matching.
    if (rowIdx === undefined && fEmail) {
      let best = 1e9,
        bestIdx = null;
      for (const k in emailMap) {
        const d = levenshtein(fEmail, k);
        if (d < best) { best = d; bestIdx = emailMap[k]; }
      }
      if (best <= 2) rowIdx = bestIdx;
    }

    // STEP 4️⃣ If a match was found, update the data arrays.
    if (rowIdx !== undefined && rowIdx !== null) {
      r[fIdx[COL_FORM_MATCHED]] = true; // Mark form row as matched
      partData[rowIdx][pIdx[COL_JOB_SECURED]] = true; // Mark participant as having a job
      partData[rowIdx][pIdx[COL_JOB_START]] = r[fIdx[COL_TIMESTAMP]]; // Copy timestamp as start date
      matched++;
    } else {
      unmatched++;
    }
  });

  // ────────────── F. WRITE BACK ──────────────
  // STEP 1️⃣ Write the "TRUE" values back to the "Matched" column in the form sheet.
  formSheet.getRange(DATA_START_ROW, fIdx[COL_FORM_MATCHED] + 1, formData.length, 1)
    .setValues(formData.map(r => [r[fIdx[COL_FORM_MATCHED]]]));
  
  // STEP 2️⃣ Write the entire updated participant data block back to its sheet.
  partSheet.getRange(DATA_START_ROW, 1, partData.length, partData[0].length)
    .setValues(partData);

  // STEP 3️⃣ Ensure the 'Job Secured Date' column is formatted as a date.
  const dateCol = pIdx[COL_JOB_START] + 1;
  partSheet.getRange(DATA_START_ROW, dateCol, partData.length, 1)
    .setNumberFormats(Array(partData.length).fill(['m/d/yyyy']));

  // ────────────── G. SUMMARY ──────────────
  // STEP 1️⃣ Show a final summary toast message to the user.
  const msg = `Matched ${matched} / ${formData.length} responses (${unmatched} unmatched).`;
  if (ui) ss.toast(msg, 'Job Form Match', 5);
  Logger.log('✅ ' + msg);
}


//  ╭───────────────────────────────╮
//  │  updateTrueLastModified()     │
//  ╰───────────────────────────────╯
/**
 * Finds the true last modified date for each student's planner by filtering out admin edits.
 * It inspects the detailed revision history of each file to find the last change made by a student.
 *
 * WHAT YOU NEED 📝
 * - The "Drive API" must be enabled in your Apps Script project (Services > Drive API > Add).
 * - "Active Participants" sheet with: 'PTID', 'Planner Link', and 'Last Update' columns.
 *
 * HOW TO RUN ▶️
 * 1. Ensure the Drive API is enabled (see above).
 * 2. Run from the menu: Admin Tools → 📊 Update 'True' Last Modified Date
 *
 * STEPS:
 * 1. Defines the list of admin/internal email addresses to ignore.
 * 2. Gets all student data from the 'Active Participants' sheet.
 * 3. Loops through each student one by one.
 * 4. For each student, it extracts the ID of their planner spreadsheet.
 * 5. It calls a helper function that uses the Drive API to fetch the file's entire revision history.
 * 6. The helper function searches backwards from the most recent edit, looking for the first change NOT made by an admin.
 * 7. If a student edit is found, its timestamp is written back into the 'Last Update' column for that student's row.
 * 8. After checking all students, it displays a summary report of how many were updated.
 */
function updateTrueLastModified() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // ────────────────────────── A. SETUP & CONFIGURATION ──────────────────────────
  
  // STEP 1️⃣ Find the target sheet.
  const sheet = ss.getSheetByName(TEST_ACTIVE_PARTICIPANTS_SHEET);
  if (!sheet) {
    ui.alert(`Error: Sheet named "${TEST_ACTIVE_PARTICIPANTS_SHEET}" not found.`);
    return;
  }

  // STEP 2️⃣ Define the list of admin/script emails to EXCLUDE from the search.
  const excludedEmails = [
    'nsilva@ideas42.org',
    'bottomlinesuccesscareers@gmail.com',
    'maddie@ideas42.org'
  ];

  // ────────────────────────── B. GET PARTICIPANT DATA ──────────────────────────
  
  // STEP 1️⃣ Read all data from the sheet into memory for fast processing.
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const headers = data.shift(); // Remove header row
  
  // STEP 2️⃣ Create a map of header names to column indices for easy access.
  const ptidCol           = headers.indexOf(COL_PTID);
  const plannerLinkCol    = headers.indexOf(COL_PLANNER);
  const lastModifiedCol   = headers.indexOf(COL_LAST_MODIFIED);
  
  // STEP 3️⃣ Validate that all required columns were found.
  if (plannerLinkCol === -1 || lastModifiedCol === -1) {
    ui.alert("Error: Missing 'Planner Link' or 'Last Update' column in the Active Participants sheet.");
    return;
  }
  
  ui.alert(
    'Starting Update',
    `This will check the revision history for ${data.length} student planners. Please don't close this sheet.`,
    ui.ButtonSet.OK
  );
  
  let updatedCount = 0;

  // ────────────────── C. PROCESS EACH PARTICIPANT ───────────────────
  data.forEach((row, idx) => {
    const ptid       = row[ptidCol];
    const plannerUrl = row[plannerLinkCol];
    
    if (!plannerUrl) {
      Logger.log(`Row ${idx+2} (PTID ${ptid}): no planner URL, skipping.`);
      return;
    }

    try {
      // Extract the spreadsheet ID
      const ssId = SpreadsheetApp.openByUrl(plannerUrl).getId();
      // Get the true last modified date (your helper)
      const trueLastModified = getTrueLastModified(ssId, excludedEmails);
      
      if (trueLastModified) {
        // Coerce to Date if it's a string
        const modDate = (trueLastModified instanceof Date)
          ? trueLastModified
          : new Date(trueLastModified);

        // Write it back into the sheet
        sheet
          .getRange(idx + 2, lastModifiedCol + 1) // +2 for header, +1 because getRange is 1-based
          .setValue(modDate);
        
        updatedCount++;
        Logger.log(`✅ PTID ${ptid} updated to ${modDate}`);
      } else {
        Logger.log(`⏩ PTID ${ptid}: no student edits found.`);
      }

    } catch (e) {
      Logger.log(`❌ Error at row ${idx+2}, PTID ${ptid}: ${e}`);
    }
  });

  // ────────────────── D. SHOW SUMMARY ───────────────────
  ui.alert(
    'Update Complete',
    `Checked all planners and updated Last Modified for ${updatedCount} students.`,
    ui.ButtonSet.OK
  );
}



//  ╭───────────────────────────────╮
//  │  getTrueLastModified()        │
//  ╰───────────────────────────────╯
/**
 * A helper function that uses the Drive API to find the last modification
 * to a file that was NOT made by one of the excluded admin users.
 *
 * @param {string} fileId The ID of the Google Sheet to check.
 * @param {string[]} excludedEmails An array of email addresses to ignore.
 * @return {Date|null} The date of the last student edit, or null if none is found.
 */
function getTrueLastModified(fileId, excludedEmails) {
  try {
    // STEP 1️⃣ Call the Drive API to get the list of all revisions for this file.
    const revisions = Drive.Revisions.list(fileId);
    
    if (!revisions.items || revisions.items.length === 0) {
      return null; // No revision history found for this file.
    }

    // STEP 2️⃣ The API returns revisions oldest-first, so we reverse the list to check the newest first.
    const reversedRevisions = revisions.items.reverse();

    // STEP 3️⃣ Loop through the revisions, starting with the most recent.
    for (const revision of reversedRevisions) {
      const editorEmail = revision.lastModifyingUser.emailAddress;

      // STEP 4️⃣ Check if the person who made this edit is on our exclusion list.
      if (excludedEmails.indexOf(editorEmail) === -1) {
        // If they are NOT on the list, we've found our target. This is the last edit made by a student.
        // Return the date of this revision and exit the function.
        return new Date(revision.modifiedDate);
      }
    }

    // STEP 5️⃣ If the loop finishes without finding a student edit, it means all edits were made by admins.
    return null;

  } catch (e) {
    Logger.log(`Could not retrieve revisions for file ID ${fileId}. Error: ${e.toString()}`);
    return null;
  }
}



// ────────────────────────── UTILITY FUNCTIONS ──────────────────────────

//  ╭──────────────────────────────────╮
//  │  Shared Helper Functions         │
//  ╰──────────────────────────────────╯

/**
 * A helper function that removes a wide range of emoji and special symbol
 * characters from a string of text.
 * @param {string} text The text that may contain emojis.
 * @return {string} The text with emojis removed, or an empty string if the input is invalid.
 */
function removeEmojis(text) {
  // Identify emojis by their Unicode and replace them with an empty string.
  return String(text || "").replace(
    /[\u{1F300}-\u{1F5FF}\u{1F600}-\u{1F64F}\u{1F680}-\u{1F6FF}\u{1F700}-\u{1F77F}\u{1F780}-\u{1F7FF}\u{1F800}-\u{1F8FF}\u{1F900}-\u{1F9FF}\u{1FA00}-\u{1FA6F}\u{1FA70}-\u{1FAFF}\u{2600}-\u{26FF}\u{2700}-\u{27BF}\u{1F1E6}-\u{1F1FF}]/gu,
    ""
  ).trim(); // Also remove leading/trailing whitespace.
}


/**
 * ───────────────────────────────────────────────────────────────────────────────
 * SECTION 2: KICK-OFF FOLLOW-UP CAMPAIGN (✉️)
 * ───────────────────────────────────────────────────────────────────────────────
 * QUICK SUMMARY (read this first)
 * -----------------------------------------------------------------------------
 * • Sends a follow-up email after the kickoff event.
 * • Sends one of three different emails based on whether a student:
 *    ▹ Attended and confirmed they reviewed their tasks.
 *    ▹ Attended but DID NOT confirm they reviewed tasks.
 *    ▹ Did not attend at all.
 * • Skips anyone who has already been sent this email.
 *
 * WHAT YOU NEED 📝
 * • All Participants sheet with columns: 
 *      First Name, Email, Planner Link, Attended, Checked Tasks, Couldn't Make It, and KickOff Sent.
 * • Outreach Plan sheet with columns:
 *      Template,	Send Date,	Subject Line
 *
 * HOW TO RUN ▶️
 *  1. Admin Tools → ✉️ Send Kick-Off Follow-Up
 */

//  ╭──────────────────────────────────╮
//  │  Kick-Off Follow-Up Functions    │
//  ╰──────────────────────────────────╯
/**
 * This is the function called when you click the menu item.
 * It runs the main logic and shows pop-ups.
 */
function sendKickoffFollowUp() { _sendKickoffFollowUp(true); }

/**
 * This is the function for an automated trigger or for running from the console.
 * It runs the main logic without showing pop-ups.
 */
function sendKickoffFollowUpTrigger() { _sendKickoffFollowUp(false); }

/**
 * This function contains the core logic for the Kick-Off Follow-Up.
 *
 * STEPS:
 *  A. Checks that the required sheets exist.
 *  B. Checks that the required columns exist in the sheets.
 *  C. Loads the email subject lines from the Outreach Plan.
 *  D. Goes through each student to determine which email to send (if any).
 *  E. Shows a final summary of how many emails were sent, skipped, or failed.
 */
/**
 * Sends Kick-Off follow-up emails and logs them in batch.
 */
function _sendKickoffFollowUp(showUi) {
  const ui = showUi ? SpreadsheetApp.getUi() : null;
  const ss = SpreadsheetApp.getActive();
  const partsSheet = ss.getSheetByName(TEST_ALL_PARTICIPANTS_SHEET);
  const planSheet = ss.getSheetByName(OUTREACH_PLAN_SHEET);
  const logSh = ss.getSheetByName(OUTREACH_LOG_SHEET);

  const stats = { sent: 0, skipped: 0, failed: 0, errors: [] };
  const newLogRows = [];

  // Load subjects from Outreach Plan
  const planHdr = planSheet.getRange(HEADER_ROW, 1, 1, planSheet.getLastColumn()).getValues()[0];
  const pIdx = Object.fromEntries(planHdr.map((h, i) => [h, i]));
  const planData = planSheet
    .getRange(DATA_START_ROW, 1, planSheet.getLastRow() - HEADER_ROW, planSheet.getLastColumn())
    .getValues();
  const subs = {};
  planData.forEach(r => {
    const tpl = r[pIdx[COL_TEMPLATE]];
    const sub = r[pIdx[COL_SUBJECT_LINE]];
    if (tpl && sub) subs[tpl] = sub;
  });

  // Iterate participants
  const headers = partsSheet.getRange(HEADER_ROW, 1, 1, partsSheet.getLastColumn()).getValues()[0];
  const idx = h => headers.indexOf(h);
  const rows = partsSheet
    .getRange(DATA_START_ROW, 1, partsSheet.getLastRow() - HEADER_ROW, partsSheet.getLastColumn())
    .getValues();

  rows.forEach((row, i) => {
    const rIdx = DATA_START_ROW + i;
    const sentFlag = row[idx(COL_KICKOFF_SENT)] === true;
    const didAttend = row[idx(COL_ATTENDED)] === true;
    const didCant = row[idx(COL_DIDNT_ATTEND)] === true;
    if (sentFlag || (!didAttend && !didCant)) {
      stats.skipped++;
      return;
    }

    const tplName = didAttend
      ? (row[idx(COL_PARTICIPATED)] ? 'KickOff_Participated' : 'KickOff_DidntParticipate')
      : 'KickOff_DidntAttend';
    const rawSub = subs[tplName];
    if (!rawSub) {
      stats.failed++;
      stats.errors.push(`No subject for ${tplName} @ row ${rIdx}`);
      return;
    }
    const subject = `=?utf-8?B?${Utilities.base64Encode(Utilities.newBlob(rawSub).getBytes())}?=`;
    const rawBitly = row[idx(bitlyCol)];
    const rawPlan = row[idx(plannerCol)];
    const finalLink = (rawBitly && String(rawBitly).trim()) ? rawBitly : rawPlan;

    // Prepare data
    const data = {
      ...LINK_PLACEHOLDERS,
      FirstName: row[idx(COL_FIRST_NAME)],
      PersonalPlannerLink: finalLink,
      title: 'Post-Grad Job Hunt Kickoff'
    };

    // Send
    const fileName = TEMPLATE_MAP[tplName] || tplName;
    const ok = _renderAndSendEmail(
      { to: row[idx(COL_EMAIL)], subject },
      { name: fileName, data }
    );

    if (ok) {
      stats.sent++;
      partsSheet.getRange(rIdx, idx(COL_KICKOFF_SENT) + 1).setValue(true);
      newLogRows.push([
        row[idx(COL_PTID)], // PTID
        tplName,            // Template
        '',                 // ThreadID (blank)
        rawSub,             // FullSubject
        row[idx(COL_EMAIL)],// To
        new Date()          // SentDate
      ]);
    } else {
      stats.failed++;
      stats.errors.push(`Send failure @ row ${rIdx}`);
    }
  });

  // Batch-log all new entries
  if (newLogRows.length) {
    logSh
      .getRange(logSh.getLastRow() + 1, 1, newLogRows.length, newLogRows[0].length)
      .setValues(newLogRows);
  }

  // Summary
  if (showUi) {
    ss.toast(`✉️ Sent: ${stats.sent}  🔄 Skipped: ${stats.skipped}  ⚠️ Failed: ${stats.failed}`, 'Kickoff Follow-Up', 5);
    if (stats.errors.length) ui.alert('Errors', stats.errors.join('\n'), ui.ButtonSet.OK);
  } else {
    Logger.log(`Kickoff complete: Sent=${stats.sent}, Skipped=${stats.skipped}, Failed=${stats.failed}`);
    if (stats.errors.length) Logger.log('Errors:\n' + stats.errors.join('\n'));
  }
}

//  ╭───────────────────────────────╮
//  │  syncSentMailToLog()          │
//  ╰───────────────────────────────╯
/**
 * Synchronizes sent emails from Gmail to the 'Outreach Log' sheet.
 * It finds emails matching subjects from the 'Outreach Plan', links them to a PTID,
 * and adds them to the log, ensuring no duplicates are created.
 *
 * WHAT YOU NEED 📝
 * ➡️ 'Outreach Plan' sheet with 'Template' and 'Subject Line' columns.
 * ➡️ 'Active Participants' sheet with 'PTID' and 'Email' columns.
 * ➡️ 'Outreach Log' sheet with headers for PTID, Template, ThreadID, etc.
 *
 * HOW TO RUN ▶️
 * - This function is a utility that can be run manually from the script editor
 * or set up on a timed trigger (e.g., daily) to keep the log updated.
 *
 * STEPS:
 * 1. Establishes connections to the Plan, Log, and Participants sheets.
 * 2. Creates a fast lookup map to find a student's PTID from their email address.
 * 3. Fetches all official email subject lines from the 'Outreach Plan'.
 * 4. Loads all existing entries from the 'Outreach Log' to prevent duplicates.
 * 5. Searches Gmail for sent messages matching the official subject lines.
 * 6. For each new, unlogged message found, it identifies the student PTID and prepares to add it to the log.
 * 7. Writes all new log entries to the sheet in a single, efficient batch operation.
 */
function syncSentMailToLog() {
  const ss = SpreadsheetApp.getActive();
  const planSh = ss.getSheetByName(OUTREACH_PLAN_SHEET);
  const logSh = ss.getSheetByName(OUTREACH_LOG_SHEET);
  const partSh = ss.getSheetByName(ACTIVE_PARTICIPANTS_SHEET);

  // A. Build email→PTID lookup
  const pHdr = partSh.getRange(1, 1, 1, partSh.getLastColumn()).getValues()[0];
  const pIdx = Object.fromEntries(pHdr.map((h, i) => [h, i]));
  const pData = partSh.getRange(2, 1, partSh.getLastRow() - 1, partSh.getLastColumn()).getValues();
  const emailToPt = {};
  pData.forEach(r => {
    const e = String(r[pIdx[COL_EMAIL]]).trim().toLowerCase();
    const id = String(r[pIdx[COL_PTID]]).trim();
    if (e && id) emailToPt[e] = id;
  });

  // B. Build dedupe set from existing log rows (only if there *are* any)
  const lastLogRow = logSh.getLastRow();
  let seen = new Set();
  if (lastLogRow > HEADER_ROW) {
    const numExisting = lastLogRow - HEADER_ROW;
    const existing = logSh
      .getRange(DATA_START_ROW, 1, numExisting, logSh.getLastColumn())
      .getValues();
    const logHdr = logSh.getRange(1, 1, 1, logSh.getLastColumn()).getValues()[0];
    const logIdx = Object.fromEntries(logHdr.map((h, i) => [h.replace(/\s+/g, ''), i]));
    existing.forEach(r => {
      seen.add(`${r[logIdx.PTID]}|${r[logIdx.ThreadID]}`);
    });
  }

  // C. Fetch subjects to look for
  const planHdr = planSh.getRange(HEADER_ROW, 1, 1, planSh.getLastColumn()).getValues()[0];
  const phIdx = Object.fromEntries(planHdr.map((h, i) => [h, i]));
  const planData = planSh.getRange(DATA_START_ROW, 1, planSh.getLastRow() - HEADER_ROW, planSh.getLastColumn()).getValues();
  const subjects = planData
    .map(r => ({ template: r[phIdx[COL_TEMPLATE]], subject: r[phIdx[COL_SUBJECT_LINE]].trim() }))
    .filter(x => x.subject);

  // D. Search Gmail & collect newRows
  const newRows = [];
  subjects.forEach(({ template, subject }) => {
    const threads = GmailApp.search(`in:sent subject:"${subject.replace(/"/g, '\\"')}"`);
    threads.forEach(thread => {
      thread.getMessages().forEach(msg => {
        if (msg.getSubject().trim() !== subject) return;
        const tos = msg.getTo().split(/\s*,\s*/).map(a => a.toLowerCase());
        const to = tos.find(a => !/nsilva@ideas42\.org|nikolas/i.test(a));
        if (!to) return;
        const ptid = emailToPt[to];
        if (!ptid) return;
        const threadId = thread.getId();
        const key = `${ptid}|${threadId}`;
        if (seen.has(key)) return;
        seen.add(key);

        // parse sender name & address
        const fromRaw = msg.getFrom();
        const match = fromRaw.match(/^(.*?)(?:\s*<(.+?)>)?$/);
        const senderName = match[1].trim();
        const senderAddress = match[2] || senderName;
        const replyToRaw = msg.getReplyTo() || '';

        newRows.push([
          ptid,            // PTID
          template,        // Template
          threadId,        // ThreadID
          subject,         // FullSubject
          msg.getTo(),     // To
          msg.getDate(),   // SentDate
          senderAddress,   // Sender Address
          senderName,      // Sender Name
          replyToRaw       // Reply-To Address
        ]);
      });
    });
  });

  // E. Batch-write newRows only if we actually have some
  if (newRows.length > 0) {
    const startRow = logSh.getLastRow() + 1;
    logSh
      .getRange(startRow, 1, newRows.length, newRows[0].length)
      .setValues(newRows);
    Logger.log(`syncSentMailToLog: appended ${newRows.length} row(s).`);
  } else {
    Logger.log('syncSentMailToLog: no new messages found.');
  }
}
//  ╭──────────────────────────────────╮
//  │  SECTION X: Deadline Setting     │
//  ╰──────────────────────────────────╯

//  ╭──────────────────────────────────╮
//  │  UI & Trigger Wrappers           │
//  ╰──────────────────────────────────╯
/**
 * Wrapper function to run the deadline setter with UI elements. Called from the menu.
 */
function runUiSetDeadlines() {
  _setAndSpaceDeadlines(true);
}

/**
 * Wrapper function to run the deadline setter without UI elements. For triggers or console runs.
 */
function runTriggerSetDeadlines() {
  _setAndSpaceDeadlines(false);
}

//  ╭──────────────────────────────────╮
//  │  _setAndSpaceDeadlines()         │
//  ╰──────────────────────────────────╯
/**
 * Core logic for proactively setting deadlines for students in at-risk clusters.
 * @param {boolean} showUi - If true, displays UI alerts. If false, only logs to the console.
 */
function _setAndSpaceDeadlines(showUi) {
  const ss = SpreadsheetApp.getActive();
  const ui = showUi ? ss.getUi() : null;
  Logger.log('▶️ Starting _setAndSpaceDeadlines...');

  // ────────────────── A. LOAD & PREPARE DATA ───────────────────
  const pSheet = ss.getSheetByName(ACTIVE_PARTICIPANTS_SHEET);
  const [pHdr, ...pRows] = pSheet.getDataRange().getValues();
  const pIdx = Object.fromEntries(pHdr.map((h, i) => [h, i]));

  const tdSheet = ss.getSheetByName(TASK_DATA_SHEET);
  const [tdHdr, ...tdRows] = tdSheet.getDataRange().getValues();
  const tIdx = Object.fromEntries(tdHdr.map((h, i) => [h, i]));

  const sprintDates = [
    {
      sprint: "1",
      start: new Date(ss.getRangeByName('Sprint1_Start').getValue()),
      end: new Date(new Date(ss.getRangeByName('Sprint1_End').getValue()).getTime() + 6*24*60*60*1000)
    },
    {
      sprint: "2",
      start: new Date(ss.getRangeByName('Sprint2_Start').getValue()),
      end: new Date(new Date(ss.getRangeByName('Sprint2_End').getValue()).getTime() + 6*24*60*60*1000)
    },
    {
      sprint: "3",
      start: new Date(ss.getRangeByName('Sprint3_Start').getValue()),
      end: new Date(new Date(ss.getRangeByName('Sprint3_End').getValue()).getTime() + 6*24*60*60*1000)
    }
  ];

  // ────────────────── B. FILTER FOR TARGET STUDENTS ───────────────────
  const targetStudents = pRows.filter(r => {
    const cluster = String(r[pIdx[COL_CLUSTER]] || '').trim();
    return ['Gatherers','Scouts'].includes(cluster);
  });

  if (!targetStudents.length) {
    if (ui) ui.alert('Status', 'No target students found to update.', ui.ButtonSet.OK);
    return;
  }
  if (showUi) {
    const confirm = ui.alert(
      'Confirm Action',
      `This will set default deadlines for ${targetStudents.length} students. Proceed?`,
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;
  }

  // ────────────────── C. PROCESS EACH TARGET STUDENT ───────────────────
  const stageOrder = [
    "I haven’t started yet",
    "I’ve browsed job listings and boards but haven’t started preparing materials",
    "I’ve created or updated my materials but haven’t applied yet",
    "I’ve submitted an application to at least one job",
    "I’ve applied to multiple jobs and have started hearing back"
  ];
  let updatedCount = 0;

  targetStudents.forEach(studentRow => {
    const ptid               = studentRow[pIdx[COL_PTID]];
    const plannerUrl         = studentRow[pIdx[COL_PLANNER]];
    const currentStage       = studentRow[pIdx[COL_S0_STAGE]];
    const currentStageIndex  = stageOrder.indexOf(currentStage);

    if (!plannerUrl || currentStageIndex === -1) {
      Logger.log(`Skipping PTID ${ptid}: Missing planner link or valid stage.`);
      return;
    }

    // STEP 1: Find tasks already done
    const doneTaskIds = new Set(
      tdRows
        .filter(r => r[tIdx[COL_PTID]] === ptid && r[tIdx[COL_DONE]] === true)
        .map(r => r[tIdx[COL_TASK_ID]])
    );

    // STEP 2: Build ideal package based on stage
    let idealTasks = [];
    const remainingStages = stageOrder.slice(currentStageIndex);
    remainingStages.forEach(stageName => {
      const arr = STAGE_TO_TASK_MAP[stageName] || [];
      idealTasks = idealTasks.concat(arr);
    });
    idealTasks = [...new Set(idealTasks)];

    // STEP 3: Remove already completed
    let tasksToSet = idealTasks.filter(id => !doneTaskIds.has(id));

    // STEP 4: Fallback if none left
    if (!tasksToSet.length) {
      Logger.log(`PTID ${ptid} has completed all high-priority tasks. Using next 3 undone tasks.`);
      tasksToSet = tdRows
        .filter(r => r[tIdx[COL_PTID]] === ptid && r[tIdx[COL_DONE]] !== true)
        .map(r => r[tIdx[COL_TASK_ID]])
        .slice(0, 3);
    }

    if (!tasksToSet.length) {
      Logger.log(`Skipping PTID ${ptid}: No uncompleted tasks.`);
      return;
    }

    const deadlines = _calculateSmartDeadlines(tasksToSet.length);
    const tasksWithDeadlines = tasksToSet.map((id, i) => ({
      id,
      deadline: deadlines[i],
      sprint: _getSprintForDate(deadlines[i], sprintDates)
    }));

    if (_setDeadlinesInPlanner(plannerUrl, tasksWithDeadlines)) {
      updatedCount++;
    }
  });

  const msg = `Successfully set deadlines for ${updatedCount} students.`;
  Logger.log(`✅ ${msg}`);
  if (ui) ui.alert('Process Complete', msg, ui.ButtonSet.OK);
}

//  ╭──────────────────────────────────╮
//  │  _getSprintForDate()             │
//  ╰──────────────────────────────────╯
/**
 * Determines which sprint a given date falls into based on start/end dates.
 */
function _getSprintForDate(date, sprintDates) {
  for (const s of sprintDates) {
    if (date >= s.start && date <= s.end) {
      return "Sprint " + s.sprint;
    }
  }
  return "";
}

//  ╭──────────────────────────────────╮
//  │  _calculateSmartDeadlines()      │
//  ╰──────────────────────────────────╯
/**
 * Calculates a series of deadlines: first Friday or Sunday, then Tue/Fri twice a week.
 */
function _calculateSmartDeadlines(numTasks) {
  const deadlines = [];
  if (!numTasks) return deadlines;

  let today = new Date();
  today.setHours(0, 0, 0, 0);

  // First deadline: this Friday or next Sunday
  let first = new Date(today);
  const daysToFri = (5 - today.getDay() + 7) % 7;
  if (daysToFri < 3) {
    const daysToSun = (7 - today.getDay() + 7) % 7 || 7;
    first.setDate(today.getDate() + daysToSun);
  } else {
    first.setDate(today.getDate() + daysToFri);
  }
  deadlines.push(first);

  // Subsequent: Tuesdays & Fridays
  let last = new Date(first);
  for (let i = 1; i < numTasks; i++) {
    const dow = last.getDay();
    let add;
    if (dow >= 5 || dow < 2) {       // Fri–Mon → next Tue
      add = (2 - dow + 7) % 7 || 7;
    } else {                         // Tue–Thu → next Fri
      add = 5 - dow;
    }
    last.setDate(last.getDate() + add);
    deadlines.push(new Date(last));
  }
  return deadlines;
}

//  ╭──────────────────────────────────╮
//  │  _setDeadlinesInPlanner()        │
//  ╰──────────────────────────────────╯
/**
 * Opens the student's planner and writes deadlines into both Sprint and Calendar sheets.
 */
function _setDeadlinesInPlanner(plannerUrl, tasksToSet) {
  try {
    const planner     = SpreadsheetApp.openByUrl(plannerUrl);
    const sprintSheet = planner.getSheetByName(PLANNER_SPRINT_SHEET);
    const calendar    = planner.getSheetByName(PLANNER_CALENDAR_SHEET);

    // Helper to update one sheet
    function updateSheet(sheet, colId, colDeadline, colSprint) {
      const data = sheet.getDataRange().getValues();
      const headerRow = data.findIndex(r => r.includes(colId));
      if (headerRow === -1) return;
      const hdr = data[headerRow];
      const idIdx = hdr.indexOf(colId);
      const dlIdx = hdr.indexOf(colDeadline);
      const spIdx = colSprint ? hdr.indexOf(colSprint) : -1;

      const mapRow = {};
      data.slice(headerRow+1).forEach((r,i) => {
        if (r[idIdx]) mapRow[String(r[idIdx]).trim()] = headerRow+2+i;
      });

      tasksToSet.forEach(t => {
        const row = mapRow[t.id.trim()];
        if (row) {
          sheet.getRange(row, dlIdx+1).setValue(t.deadline);
          if (spIdx >= 0) sheet.getRange(row, spIdx+1).setValue(t.sprint);
        }
      });
    }

    if (sprintSheet) updateSheet(sprintSheet, PLANNER_COL_ID, PLANNER_COL_DEADLINE, COL_SPRINT);
    if (calendar)    updateSheet(calendar,   PLANNER_COL_ID, COL_DEADLINE);

    SpreadsheetApp.flush();
    return true;
  } catch (e) {
    Logger.log(`Error setting deadlines for ${plannerUrl}: ${e.message}`);
    return false;
  }
}





/**
 * Calculates the "Follow-Through %" for each student and updates the 'Active Participants' sheet.
 * This function is designed to be run directly from the Apps Script editor console.
 */
function calculateAndApplyFollowThrough() {
  
  // Get a reference to the currently active spreadsheet.
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get the specific sheets we need to work with by their names.
  const participantSheet = ss.getSheetByName(ACTIVE_PARTICIPANTS_SHEET);
  const taskSheet = ss.getSheetByName(TASK_DATA_SHEET);

  // --- Pre-computation & Safety Checks ---

  // Stop execution if the required sheets don't exist, and log an error.
  if (!participantSheet || !taskSheet) {
    console.error("Script stopped: Could not find required sheets. Please check sheet names in the configuration.");
    return;
  }

  // Retrieve all data from both sheets into 2D arrays.
  // This is more efficient than reading from the sheet cell by cell.
  const participantData = participantSheet.getDataRange().getValues();
  const taskData = taskSheet.getDataRange().getValues();
  
  // Get today's date and reset the time to the beginning of the day.
  // This ensures that deadlines for "today" are handled correctly (deadline >= today).
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // --- Step 1: Process the Task Data Sheet Once ---
  // We will iterate through all tasks and aggregate the stats for each student.
  // This is much faster than repeatedly scanning the task sheet for every student.
  
  // 'taskStats' will store the calculated counts for each student.
  // The structure will be: { "PTID": { totalWithDeadline: 0, onTrackCount: 0 } }
  const taskStats = {}; 

  // Get the header row from the task data to find column indices by name.
  // This makes the script robust even if column order changes.
  const taskHeaders = taskData[0];
  const taskPtidIndex = taskHeaders.indexOf(COL_PTID);
  const deadlineIndex = taskHeaders.indexOf(COL_DEADLINE);
  const doneIndex = taskHeaders.indexOf(COL_DONE);

  // Iterate through each row of the task data, starting from the second row (index 1) to skip the headers.
  for (let i = 1; i < taskData.length; i++) {
    const row = taskData[i];
    const ptid = row[taskPtidIndex];
    const deadlineValue = row[deadlineIndex];
    const isDone = row[doneIndex] === true;

    // Skip any task that doesn't have a PTID or a deadline value, as it can't be calculated.
    if (!ptid || !deadlineValue) continue; 

    // Convert the deadline value from the sheet into a true Date object.
    const deadline = new Date(deadlineValue);
    // Skip if the deadline is not a valid date.
    if (isNaN(deadline.getTime())) continue; 
    
    // If we haven't seen this student before, initialize their stats object.
    if (!taskStats[ptid]) {
      taskStats[ptid] = { totalWithDeadline: 0, onTrackCount: 0 };
    }

    // Since this task has a valid deadline, increment the total count for this student.
    taskStats[ptid].totalWithDeadline++;

    // This is YOUR definition of "on track":
    // The task is considered on track if it's already marked as done, OR if its deadline has not yet passed.
    if (isDone || deadline >= today) {
      taskStats[ptid].onTrackCount++;
    }
  }

  console.log("Task data processed. Found stats for", Object.keys(taskStats).length, "students.");

  // --- Step 2: Update the 'Active Participants' Sheet ---

  // Get the header row from the participants sheet to find column indices.
  const participantHeaders = participantData[0];
  const participantPtidIndex = participantHeaders.indexOf(COL_PTID);
  const targetColIndex = participantHeaders.indexOf(COL_FOLLOW_THRU);
  
  // Stop execution if the target "Follow-Through %" column doesn't exist.
  if (targetColIndex === -1) {
    console.error(`Script stopped: Target column "${COL_FOLLOW_THRU}" not found in the '${SHEET_PARTICIPANTS}' sheet.`);
    return;
  }

  // Prepare an array to hold all the new percentage values.
  const percentagesToUpdate = [];
  
  // Iterate through each participant, starting from the second row to skip headers.
  for (let i = 1; i < participantData.length; i++) {
    const ptid = participantData[i][participantPtidIndex];
    const stats = taskStats[ptid];
    let percentage = null; // Default to null if we have no data for this student.

    // Check if we have stats for this student and if they have any tasks with deadlines.
    if (stats && stats.totalWithDeadline > 0) {
      // Calculate the follow-through proportion.
      percentage = stats.onTrackCount / stats.totalWithDeadline;
    }
    
    // Add the calculated percentage to our array. If it's null, add an empty string so the cell becomes blank.
    percentagesToUpdate.push([percentage === null ? "" : percentage]);
  }
  
  // Get the range in the target column that we need to update.
  // The range starts at row 2, in the target column, and is as tall as our array of percentages.
  const targetRange = participantSheet.getRange(2, targetColIndex + 1, percentagesToUpdate.length, 1);
  
  // Write all the calculated percentages to the spreadsheet in a single operation.
  // This is significantly more efficient than updating one cell at a time.
  targetRange.setValues(percentagesToUpdate);
  
  // Format the entire target column to display numbers as percentages.
  targetRange.setNumberFormat('0%');

  console.log(`Script complete. "${COL_FOLLOW_THRU}" column has been updated for ${percentagesToUpdate.length} students.`);
}


