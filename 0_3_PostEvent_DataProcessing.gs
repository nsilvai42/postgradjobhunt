/**
 * 
 * Script Name: 0_3_PostEvent_DataProcessing.gs
 * 
 * 
/

/**
 * ───────────────────────────────────────────────────────────────────────────────
 * POST-EVENT DATA PROCESSING ⚙️
 * ───────────────────────────────────────────────────────────────────────────────
 * QUICK SUMMARY (read this first)
 * -----------------------------------------------------------------------------
 * A collection of tools for processing data *after* the kick-off, including:
 *    ▹ Checking if students have completed any tasks at all.
 *    ▹ Collecting "purpose" statements from their planners.
 *    ▹ Matching Zoom poll answers to participants based on names.
 *    ▹ Aggregating task data from every student planner.
 */

//  ╭───────────────────────────────╮
//  │  markCheckedOffTasks()        │
//  ╰───────────────────────────────╯
/**
 * For each attended participant, checks if they marked *any* task in their checklist as Done.
 *  
 * WHAT YOU NEED 📝
 *  ✅ "All Participants" sheet with columns:
 *      PTID, Attended?, Planner Link, Checked Off Tasks
 *
 * HOW TO RUN ▶️
 *    1. Open the Admin Panel spreadsheet.
 *    2. Admin Tools → ✅ Checked Off Tasks
 *    3. Review All Participants sheet for students who attended but didn't check off any tasks.
 * 
 * STEPS:
 *    1. Checks that the required sheet and columns exist.
 *    2. Reads all participant data..
 *    3. Loops through each student's planner to check for any completed task.
 *    4. Writes true or false in the Checked Off Tasks column in All Participants sheet.
 *    5. Displays a toast ✅ summary.
 */
function markCheckedOffTasks() {
  const ui = SpreadsheetApp.getUi();                      // UI for alerts
  const ss = SpreadsheetApp.getActive();                  // Active spreadsheet

  // ────────────────────────── A. VALIDATE SHEETS & CONFIG ──────────────────────────

  // STEP 1️⃣  Is the main participants sheet present?
  // 🚨 IMPORTANT : All sheets should be named the same exact text in the 0_GlobalConstants file!
  const partsSheet = ss.getSheetByName(TEST_ALL_PARTICIPANTS_SHEET);
  if (!partsSheet) {
    ui.alert("Error: Missing Sheet", 
      `Could not find the '${TEST_ALL_PARTICIPANTS_SHEET}' sheet. Is the sheet's name written correctly?`, 
      ui.ButtonSet.OK);
    return;
  }

  // ────────────────────────── B. MAP HEADERS → COLUMN NUMBERS ──────────────────────────

  // STEP 1️⃣  Locate columns based on header names.
  // 🚨 IMPORTANT : All headers should be named the same exact text in the 0_GlobalConstants file!
  const headers = partsSheet.getRange(HEADER_ROW, 1, 1, partsSheet.getLastColumn()).getValues()[0];
  const headerMap = {};
  headers.forEach((h, i) => headerMap[h] = i + 1);

  // STEP 2️⃣  Check for required headers.
  const requiredHeaders = [COL_PLANNER, COL_ATTENDED, COL_PARTICIPATED];
  const missingHeaders = requiredHeaders.filter(h => !headerMap[h]);
  if (missingHeaders.length > 0) {
    ui.alert("Error: Missing Headers", 
      `The '${TEST_ALL_PARTICIPANTS_SHEET}' sheet is missing required headers: "${missingHeaders.join('", "')}". Are the headers spelled correctly?`, 
      ui.ButtonSet.OK);
    return;
  }
  
  // STEP 3️⃣ Store column locations for later.
  const colPlanner = headerMap[COL_PLANNER];            // Planner Link
  const colAttended = headerMap[COL_ATTENDED];          // Attended?
  const colCheckedTasks = headerMap[COL_PARTICIPATED]; // Checked Off Tasks

  // ────────────────────────── C. FETCH & PROCESS PARTICIPANTS ──────────────────────────

  // STEP 1️⃣: Find the last row with content to determine how many participants there are.
  const lastRow = partsSheet.getLastRow();

  // If there are no participants listed after the header row, show an error message and stop.
  if (lastRow < DATA_START_ROW) {
    ui.alert("Error: No Participants Found", 
      `No participants found in '${TEST_ALL_PARTICIPANTS_SHEET}'.`, 
      ui.ButtonSet.OK);
    return;
  }

  // STEP 2️⃣ Fetch all participant data.
  const dataRange = partsSheet.getRange(DATA_START_ROW, 1, 
                    lastRow - DATA_START_ROW + 1, 
                    partsSheet.getLastColumn());
  const data = dataRange.getValues();

  // STEP 3️⃣ Loop through each participant row to check their planner if they attended and have a planner link.
  const results = data.map((row, i) => {
    const attended = row[colAttended - 1];  // Attended?
    const url = row[colPlanner - 1];        // Planner Link
    let anyChecked = false; // Default to FALSE

    // Only process students who attended and have a planner link.
    if (attended === true && url) {
      try {
        const stSS = SpreadsheetApp.openByUrl(url);                   // Open Planner
        const clSheet = stSS.getSheetByName(PLANNER_CHECKLIST_SHEET); // Find checklist tab

        if (clSheet) {
          const vals = clSheet.getDataRange().getValues();
          const hRow = vals.findIndex(r => r.some(cell => String(cell).trim().toLowerCase() === PLANNER_COL_ID.toLowerCase()));
          const doneIdx = (hRow > -1) ? vals[hRow].indexOf(PLANNER_COL_DONE) : -1;  // Find Done column

          if (doneIdx > -1) {
            // Scan all checklist rows to see if *any* task is marked as done.
            anyChecked = vals.slice(hRow + 1).some(taskRow =>
              taskRow[doneIdx] === true || String(taskRow[doneIdx]).toLowerCase() === 'true'
            );
          }
        }
      } catch (e) {
        Logger.log(`Error checking tasks for row ${i + DATA_START_ROW}: ${e.stack || e}`); // Log error.
        anyChecked = false; // Default to FALSE if there's an error
      }
    }
    return [anyChecked]; // Return the results.
  });

  // ────────────────────────── D. WRITE RESULTS & SHOW SUMMARY ──────────────────────────

  // STEP 1️⃣ Write all TRUE/FALSE results to the Checked Off Tasks column in the All Participants sheet.
  partsSheet.getRange(DATA_START_ROW, 
                      colCheckedTasks, 
                      results.length, 1).setValues(results);

  // STEP 2️⃣ Show a final summary toast.
  const completedCount = results.filter(r => r[0] === true).length;
  ss.toast(
    `✅ Task Check Complete:\nMarked ${completedCount} students as having engaged with their task list.`,
    'Task Check Complete',
    5
  );
}


//  ╭───────────────────────────────╮
//  │  matchZoomResults()           │
//  ╰───────────────────────────────╯
/**
 * Imports poll results from Zoom sheets ('Intro_Poll', 'Closing_Poll') into 'All Participants'.
 * It uses a sophisticated "fuzzy" matching logic to link poll data (from Zoom user names)
 * to participants, even if names are slightly different.
 *
 * WHAT YOU NEED 📝
 *  💬 "All Participants" sheet with: First Name, Last Name, S0_Confidence_Pre, S0_Confidence_Post, S0_Stage
 * 
 *  💬 "Intro_Poll" sheet exported from Zoom with: 
 *                      User Name (name from Zoom),	
 *                      S0_Stage (job search poll answer from intro poll - rename header to match this),	
 *                      S0_Confidence_Pre (confidence poll answer from intro poll - rename header to match this)
 * 
 *  💬 "Closing_Poll" sheet exported from Zoom with:
 *                      User Name,	
 *                      S0_Confidence_Post (confidence poll answer from closing poll - rename header to match this)
 *
 * HOW TO RUN ▶️
 *    1. Export poll results from Zoom and paste them into an 'Intro_Poll' and 'Closing_Poll' sheet.
 *    2. Remove any unused columns and rows so that the sheet contains only the necessary columns specified above.
 *    3. Rename the poll answer columns to match User Name, S0_Stage, S0_Confidence_Pre and S0_Confidence_Post
 *    4. Admin Tools → 💬 Match Zoom Answers
 *
 * STEPS:
 *    1. Checks for all required sheets and columns.
 *    2. Reads all participant and poll data.
 *    3. Matches names from User Name columns in both poll sheets to First Name and Last Name in Active Participants.
 *    4. Writes the poll answer to the matched row in 'All Participants' sheet.
 */
function matchZoomResults() {
  const ui = SpreadsheetApp.getUi();     // UI for alerts
  const ss = SpreadsheetApp.getActive(); // Active spreadsheet

  // ────────────────────────── A. VALIDATE SHEETS & CONFIG ──────────────────────────

  const ps = ss.getSheetByName(TEST_ACTIVE_PARTICIPANTS_SHEET);
  const introSheet = ss.getSheetByName(ZOOM_INTRO_SHEET);
  const closingSheet = ss.getSheetByName(ZOOM_CLOSING_SHEET);

  if (!ps || !introSheet || !closingSheet) {
    ui.alert('Error: Missing Sheets', 
      "Missing 'All Participants', 'Intro_Poll', or 'Closing_Poll' sheet. Make sure you imported Zoom poll answers and named the sheets exactly like the names in this message.", 
      ui.ButtonSet.OK);
    return;
  }

  // ────────────────────────── B. MAP HEADERS → COLUMN NUMBERS ──────────────────────────

  const mapHdr = (s) => {
    const h = s.getRange(HEADER_ROW, 1, 1, s.getLastColumn()).getValues()[0];
    const m = {};
    h.forEach((c, i) => m[String(c || '').trim()] = i);
    return m;
  };
  const pIdx = mapHdr(ps);           // Participant sheet column indices (0-based)
  const iIdx = mapHdr(introSheet);   // Intro Poll column indices
  const cIdx = mapHdr(closingSheet); // Closing Poll column indices

  const requiredP = [COL_FIRST_NAME, COL_LAST_NAME, COL_S0_STAGE, COL_S0_CONF_PRE, COL_S0_CONF_POST];
  const missingP = requiredP.filter(h => pIdx[h] === undefined);
  if (missingP.length > 0) {
    ui.alert("Error: Missing Headers", 
      `The '${TEST_ALL_PARTICIPANTS_SHEET}' sheet is missing: "${missingP.join('", "')}". Make sure the headers are written exactly like in this message`, 
      ui.ButtonSet.OK);
    return;
  }

  // ────────────────── C. BUILD PARTICIPANT NAME MAP ────────────────────

  const lastPRow = ps.getLastRow();
  if (lastPRow < DATA_START_ROW) {
    ui.alert('Error: No Participants Found', 'No participants found to process.', ui.ButtonSet.OK);
    return;
  }

  // STEP 1️⃣ Create a map of full names to row numbers for matching.
  const names = ps.getRange(DATA_START_ROW, pIdx[COL_FIRST_NAME] + 1, lastPRow - DATA_START_ROW + 1, 2).getValues();
  const nameToRowIndex = {}; // 0-based index for our data array
  names.forEach((r, i) => {
    const key = (String(r[0] || '') + String(r[1] || '')).toLowerCase().replace(/\s+/g, '');
    if (key) nameToRowIndex[key] = i;
  });

  // STEP 2️⃣ Function to find the best matching participant for a given Zoom name.
  const findBestMatchIndex = (zoomName) => {
    const zoomKey = normalizeZoomKey(zoomName); 
    if (!zoomKey) return null;
    if (nameToRowIndex[zoomKey] !== undefined) return nameToRowIndex[zoomKey]; // Exact match found

    // Fuzzy match using Levenshtein distance to match names
    let bestDist = Infinity,
      bestKey = null;
    for (const pKey in nameToRowIndex) {
      const dist = levenshtein(zoomKey, pKey); 
      if (dist < bestDist) {
        bestDist = dist;
        bestKey = pKey;
      }
      if (dist === 0) break; // Perfect match
    }
    // Allow up to 2 character differences to count as a valid match.
    return bestDist <= 2 ? nameToRowIndex[bestKey] : null;
  };

  // ────────────────── D. PROCESS POLLS & UPDATE DATA ───────────────────

  // STEP 1️⃣ Fetch all participant and poll data.
  const participantData = ps.getRange(DATA_START_ROW, 1, ps.getLastRow() - DATA_START_ROW + 1, ps.getLastColumn()).getValues();
  const introData = introSheet.getRange(DATA_START_ROW, 1, introSheet.getLastRow() - DATA_START_ROW + 1, introSheet.getLastColumn()).getValues();
  const closeData = closingSheet.getRange(DATA_START_ROW, 1, closingSheet.getLastRow() - DATA_START_ROW + 1, closingSheet.getLastColumn()).getValues();
  let matchesFound = 0;

  // STEP 2️⃣ Process the intro poll.
  introData.forEach(r => {
    const rowIndex = findBestMatchIndex(r[iIdx[COL_USER_NAME]]);
    if (rowIndex !== null) {
      participantData[rowIndex][pIdx[COL_S0_STAGE]] = r[iIdx[COL_S0_STAGE]];
      participantData[rowIndex][pIdx[COL_S0_CONF_PRE]] = r[iIdx[COL_S0_CONF_PRE]];
      matchesFound++;
    }
  });

  // STEP 3️⃣ Process the closing poll.
  closeData.forEach(r => {
    const rowIndex = findBestMatchIndex(r[cIdx[COL_USER_NAME]]);
    if (rowIndex !== null) {
      participantData[rowIndex][pIdx[COL_S0_CONF_POST]] = r[cIdx[COL_S0_CONF_POST]];
    }
  });

  // ────────────────── E. WRITE RESULTS & SHOW SUMMARY ──────────────────

  // STEP 1️⃣ Write all the updated data back to the Participants sheet.
  ps.getRange(DATA_START_ROW, 1, participantData.length, participantData[0].length).setValues(participantData);

  // STEP 2️⃣ Show a final summary toast.
  ss.toast(
    `✅ Matched and imported poll results for ${matchesFound} participants.`,
    'Zoom Poll Import Complete',
    5
  );
}


//  ╭───────────────────────────────╮
//  │  collectStudentPurpose()      │
//  ╰───────────────────────────────╯
/**
 * Collects "purpose" statement from each attendee's planner and adds to the 'Fears and Purpose' sheet.
 *
 * WHAT YOU NEED 📝
 *    🕊️ "Active Participants" sheet with: PTID, Planner Link
 *    🕊️ "Fears and Purpose" sheet with: PTID, Purpose
 *
 * HOW TO RUN ▶️
 *     1. Open the Admin Panel spreadsheet.
 *     2. Admin Tools → 🕊️ Collect Student Purpose
 *
 * STEPS:
 *     1. Checks that the required sheets and columns exist.
 *     3. Loops through each active participant, opens their planner, and gets their purpose.
 *     4. Writes the collected purpose to the correct row in 'Fears and Purpose'.
 *     5. Displays a toast ✅ summary.
 */
function collectStudentPurpose() {
  const ui = SpreadsheetApp.getUi(); // UI for alerts
  const ss = SpreadsheetApp.getActive(); // Active spreadsheet

  // ────────────────────────── A. VALIDATE SHEETS & CONFIG ──────────────────────────

  // STEP 1️⃣ Find the Active Participants and Fears and Purpose sheets.
  const partsSheet = ss.getSheetByName(ACTIVE_PARTICIPANTS_SHEET);
  const outputSheet = ss.getSheetByName(FEARS_PURPOSE_SHEET);

  if (!partsSheet || !outputSheet) {
    ui.alert("Error: Missing Sheet", 
      `Could not find '${ACTIVE_PARTICIPANTS_SHEET}' or '${FEARS_PURPOSE_SHEET}'. Are the names spelled correctly?`, 
      ui.ButtonSet.OK);
    return;
  }

  // ────────────────────────── B. MAP HEADERS → COLUMN NUMBERS ──────────────────────────

  // STEP 1️⃣ Locate columns from header names for the Active Participants sheet.
  const pHeaders = partsSheet.getRange(HEADER_ROW, 1, 1, partsSheet.getLastColumn()).getValues()[0];
  const pHeaderMap = {};
  pHeaders.forEach((h, i) => pHeaderMap[h] = i + 1);

  // STEP 2️⃣ Locate columns from header names for the Fears and Purpose sheet.
  const oHeaders = outputSheet.getRange(HEADER_ROW, 1, 1, outputSheet.getLastColumn()).getValues()[0];
  const oHeaderMap = {};
  oHeaders.forEach((h, i) => oHeaderMap[h] = i + 1);

  // STEP 3️⃣ Validate required headers in both sheets or show error message.
  // Active Participants headers
  if (!pHeaderMap[COL_PTID] || !pHeaderMap[COL_PLANNER]) {
    ui.alert("Error: Missing Headers", 
      `The '${ACTIVE_PARTICIPANTS_SHEET}' sheet is missing 'PTID' or 'Planner Link' columns.`, 
      ui.ButtonSet.OK);
    return;
  }

  // Fears and Purpose headers
  if (!oHeaderMap[COL_PTID] || !oHeaderMap[COL_PURPOSE]) {
    ui.alert("Error: Missing Headers", 
      `The '${FEARS_PURPOSE_SHEET}' sheet is missing 'PTID' or 'Purpose' columns.`, 
      ui.ButtonSet.OK);
    return;
  }

  // STEP 4️⃣ Store column locations for later.
  // Active Participants sheet
  const colPtidP = pHeaderMap[COL_PTID];      // PTID 
  const colPlanner = pHeaderMap[COL_PLANNER]; // Planner Link
  
  // Fears and Purpose sheet
  const colPtidO = oHeaderMap[COL_PTID];      // PTID
  const colPurpose = oHeaderMap[COL_PURPOSE]; // Purpose

  // ────────────────── C. BUILD PTID → ROW LOOKUP MAP ───────────────────

  // STEP 1️⃣ Create a map of each participant's row number.
  const lastOutRow = outputSheet.getLastRow();  // Find last row with data
  const ptidToRow = {};

  if (lastOutRow >= DATA_START_ROW) {
    const ptidValues = outputSheet.getRange(DATA_START_ROW, colPtidO, lastOutRow - DATA_START_ROW + 1, 1).getValues();
    ptidValues.forEach((r, i) => {
      if (r[0]) ptidToRow[r[0]] = i + DATA_START_ROW; // Map the PTID to its sheet row number.
    });
  }

  // ────────────────── D. COLLECT & WRITE PURPOSE ──────────────────────

  const lastPRow = partsSheet.getLastRow();
  if (lastPRow < DATA_START_ROW) {
    ui.alert("Error: No Participants Found", 
      `No participants found in '${ACTIVE_PARTICIPANTS_SHEET}'.`, 
      ui.ButtonSet.OK);
    return;
  }

  // STEP 1️⃣ Fetch all active participant data.
  const participants = partsSheet.getRange(DATA_START_ROW, 1, lastPRow - DATA_START_ROW + 1, partsSheet.getLastColumn()).getValues();

  let updated = 0; // Track how many rows are updated.

  // STEP 2️⃣ Loop through participants, get purpose, and write to the correct row.
  participants.forEach(row => {
    const ptid = row[colPtidP - 1];
    const url = row[colPlanner - 1];
    if (!ptid || !url) return; // Skip if no PTID or URL.

    const targetRow = ptidToRow[ptid]; // Find the row to write to from our map.
    if (!targetRow) return;           // Skip if PTID isn't in the Fears and Purpose sheet.

    try {
      const stSS = SpreadsheetApp.openByUrl(url);
      const pSheet = stSS.getSheetByName(PLANNER_PURPOSE_SHEET);
      if (pSheet) {
        const raw = pSheet.getRange(PLANNER_PURPOSE_CELL).getValue();
        const clean = removeEmojis(raw); // Use global utility function
        if (clean) {
          outputSheet.getRange(targetRow, colPurpose).setValue(clean);
          updated++;
        }
      }
    } catch (e) {
      Logger.log(`Purpose collection error for PTID ${ptid}: ${e.stack || e}`);
    }
  });

  // STEP 5️⃣ Show a final summary toast.
  ss.toast(
    `✅ Collected and wrote Purpose for ${updated} students.`,
    'Purpose Collection Complete',
    5
  );
}


// ────────────────────────── UTILITY FUNCTIONS ──────────────────────────

//  ╭──────────────────────────────────╮
//  │  Shared Helper Functions         │
//  ╰──────────────────────────────────╯


/**
 * Normalizes a Zoom user name for matching by removing extra characters.
 * e.g., "John Doe (He/Him) #123" -> "johndoe"
 * @param {string} name The raw name from a Zoom report.
 * @return {string} The normalized name for matching.
 */
function normalizeZoomKey(name) {
  return name ? String(name).replace(/#.*/, '').replace(/\(.*?\)/g, '').replace(/\s+/g, '').toLowerCase() : null;
}

/**
 * Calculates the Levenshtein distance between two strings.
 * This measures the number of edits (insertions, deletions, substitutions)
 * needed to change one word into the other. A lower number means a closer match.
 * @param {string} a The first string.
 * @param {string} b The second string.
 * @return {number} The distance score.
 */
function levenshtein(a, b) {
  const dp = Array(a.length + 1).fill(0).map(() => Array(b.length + 1).fill(0));
  for (let i = 0; i <= a.length; i++) dp[i][0] = i;
  for (let j = 0; j <= b.length; j++) dp[0][j] = j;
  for (let i = 1; i <= a.length; i++) {
    for (let j = 1; j <= b.length; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      dp[i][j] = Math.min(dp[i - 1][j] + 1, dp[i][j - 1] + 1, dp[i - 1][j - 1] + cost);
    }
  }
  return dp[a.length][b.length];
}

