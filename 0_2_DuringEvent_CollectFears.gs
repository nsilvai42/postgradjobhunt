/**
 * 
 * Script Name: 0_2_DuringEvent_CollectFears.gs
 * 
 * 
/


/**
 * ───────────────────────────────────────────────────────────────────────────────
 * COLLECT STUDENT FEARS  😨
 * ───────────────────────────────────────────────────────────────────────────────
 * QUICK SUMMARY (read this first)
 * -----------------------------------------------------------------------------
 * One‑click script you run during the kickoff event.
 * 
 * For every student marked as Attended?, it:
 *      ▹ opens their individual planner using the "Planner Link"
 *      ▹ finds their Fears tab
 *      ▹ reads the fear they wrote
 *      ▹ writes their fear in the "Fears and Purpose" sheet
 * Shows a pop‑up summary of how many fears were collected.
 *
 * WHAT YOU NEED 📝
 *   ✅ "All Participants" sheet with columns:  
 *                                            PTID, Attended?, Planner Link
 *   ✅ "Fears and Purpose" sheet with columns: 
 *                                            PTID, Fear
 *
 * HOW TO RUN ▶️
 *    1. Open the Admin Panel spreadsheet.
 *    2. Admin Tools → 😣 Collect Student Fears
 *    3. Review the collected fears in the "Fears and Purpose" sheet.
 */

//  ╭───────────────────────────────╮
//  │  collectStudentFears()        │
//  ╰───────────────────────────────╯
/**
 * This function collects student fears by opening each attended participant's planner and reading a specific cell.
 *
 * STEPS:
 *    1. Checks that the required sheets exist
 *    2. Maps required header names → column numbers
 *    3. Reads participant data from Admin Panel
 *    4. Loops through attendees, opens their planner, and collects their fear
 *    5. Writes all collected fears to the output sheet and shows a ✅ summary
 */
//  ╭───────────────────────────────╮
//  │  collectStudentFears()        │
//  ╰───────────────────────────────╯
/**
 * Collects "Fear" and "Fear Reasons" from each attendee's planner and syncs
 * the data to the 'Fears and Purpose' sheet.
 *
 * WHAT YOU NEED 📝
 * - "All Participants" sheet with: PTID, Planner Link, Attended?
 * - "Fears and Purpose" sheet with: PTID, Fear, and the new 'Fear Reasons' column.
 *
 * HOW TO RUN ▶️
 * 1. Open the Admin Panel spreadsheet.
 * 2. Admin Tools → 😣 Collect Student Fears
 *
 * STEPS:
 * 1. Checks that the required sheets and columns exist.
 * 2. Reads ALL data from BOTH the 'All Participants' and 'Fears and Purpose' sheets into memory.
 * 3. Creates a lookup map to quickly find the correct row in 'Fears and Purpose' for each PTID.
 * 4. Loops through each participant marked as "Attended".
 * 5. For each, it opens their planner and scrapes both the 'Fear' and 'Fear Reasons' cells.
 * 6. It updates the corresponding data for that student IN MEMORY.
 * 7. After checking all students, it writes the entire updated data block back to the 'Fears and Purpose' sheet in a single batch operation.
 * 8. Displays a toast ✅ summary.
 */
//  ╭───────────────────────────────╮
//  │  collectStudentFears()        │
//  ╰───────────────────────────────╯
/**
 * Collects "Fear" and "Fear Reasons" from each attendee's planner and syncs
 * the data to the 'Fears and Purpose' sheet.
 */
function collectStudentFears() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  // ────────────────────────── A. VALIDATE & LOAD SHEETS ──────────────────────────
  
  const participantsSheet = ss.getSheetByName(ALL_PARTICIPANTS_SHEET);
  const outputSheet = ss.getSheetByName(FEARS_PURPOSE_SHEET);

  if (!participantsSheet || !outputSheet) {
    ui.alert("Error: Missing Sheet", "Could not find 'All Participants' or 'Fears and Purpose' sheet.", ui.ButtonSet.OK);
    return;
  }

  const pData = participantsSheet.getDataRange().getValues();
  const oData = outputSheet.getDataRange().getValues();

  // ────────────────────────── B. MAP HEADERS → COLUMN INDICES ──────────────────────────

  const pHeaders = pData.shift(); const pIdx = h => pHeaders.indexOf(h);
  const oHeaders = oData.shift(); const oIdx = h => oHeaders.indexOf(h);

  const requiredPHeaders = [COL_PTID, COL_PLANNER, COL_ATTENDED];
  if (requiredPHeaders.some(h => pIdx(h) === -1)) {
    // This alert was missing the third parameter
    ui.alert("Error: Missing Headers", `The '${ALL_PARTICIPANTS_SHEET}' sheet is missing one of these required headers: ${requiredPHeaders.join(', ')}.`, ui.ButtonSet.OK); // <--- FIX
    return;
  }
  
  const requiredOHeaders = [COL_PTID, COL_FEAR, COL_FEAR_REASONS];
  if (requiredOHeaders.some(h => oIdx(h) === -1)) {
    // This alert was missing the third parameter
    ui.alert("Error: Missing Headers", `The '${FEARS_PURPOSE_SHEET}' sheet is missing one of these required headers: ${requiredOHeaders.join(', ')}.`, ui.ButtonSet.OK); // <--- FIX
    return;
  }

  // ────────────────── C. BUILD PTID → ROW LOOKUP MAP ───────────────────
  
  const ptidToRowIndex = {};
  oData.forEach((row, i) => {
    const ptid = row[oIdx(COL_PTID)];
    if (ptid) {
      ptidToRowIndex[ptid] = i;
    }
  });

  // ────────────────── D. COLLECT & UPDATE FEARS DATA IN MEMORY ──────────────────────

  let updatedCount = 0;
  
  pData.forEach(row => {
    const ptid = row[pIdx(COL_PTID)];
    const attended = row[pIdx(COL_ATTENDED)];
    const plannerUrl = row[pIdx(COL_PLANNER)];

    if (attended !== true || !plannerUrl) {
      return;
    }

    const outputRowIndex = ptidToRowIndex[ptid];
    if (outputRowIndex === undefined) {
      Logger.log(`Skipping PTID ${ptid}: Not found in '${FEARS_PURPOSE_SHEET}'.`);
      return;
    }
    
    try {
      const studentSS = SpreadsheetApp.openByUrl(plannerUrl);
      const fearSheet = studentSS.getSheetByName(PLANNER_FEAR_SHEET);

      if (!fearSheet) {
        Logger.log(`Skipping PTID ${ptid}: Sheet '${PLANNER_FEAR_SHEET}' not found.`);
        return;
      }
      
      const rawFear = fearSheet.getRange(PLANNER_FEAR_CELL).getValue();
      const rawReasons = fearSheet.getRange(PLANNER_FEAR_REASONS_CELL).getValue();

      const cleanFear = removeEmojis(rawFear);
      const cleanReasons = removeEmojis(rawReasons);

      if (cleanFear || cleanReasons) {
        oData[outputRowIndex][oIdx(COL_FEAR)] = cleanFear;
        oData[outputRowIndex][oIdx(COL_FEAR_REASONS)] = cleanReasons;
        updatedCount++;
      }
      
    } catch (e) {
      Logger.log(`Error collecting fear for PTID ${ptid}: ${e.stack || e}`);
    }
  });

  // ────────────────── E. BATCH-WRITE ALL DATA & SHOW SUMMARY ───────────────────

  if (updatedCount > 0) {
    outputSheet.getRange(DATA_START_ROW, 1, oData.length, oHeaders.length).setValues(oData);
    // This alert was missing the third parameter
    ui.alert("✅ Collection Complete", `Collected and updated fears for ${updatedCount} attended students.`, ui.ButtonSet.OK); // <--- FIX
  } else {
    ui.alert("No New Data Collected", "No new fears were found to update. Make sure students are marked as 'Attended?' and have valid planner links.", ui.ButtonSet.OK);
  }
}

// ────────────────────────── UTILITY FUNCTIONS ──────────────────────────

//  ╭───────────────────────────────╮
//  │  removeEmojis()               │
//  ╰───────────────────────────────╯
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
