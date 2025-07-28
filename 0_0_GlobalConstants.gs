/**
 * Script Name: 0_0_GlobalConstants.gs
 *
 * Central configuration constants used across all phases of the Career Sprint Apps Script.
 * Includes sheet names, column headers, planner-template mappings, outreach plan columns,
 * task-data columns, and more.
 */

// ============================================================================
// I. SHEET NAMES
// ============================================================================
const PLANNER_SETUP_SHEET              = "Planner Setup";
const ALL_PARTICIPANTS_SHEET           = "All Participants";
const ACTIVE_PARTICIPANTS_SHEET        = "Active Participants";
const FEARS_PURPOSE_SHEET              = "Fears and Purpose";
const TASK_DATA_SHEET                  = "Task Data";
const JOB_FORM_SHEET                   = "Job Form"
const OUTREACH_PLAN_SHEET              = "Outreach Plan";
const OUTREACH_LOG_SHEET               = "Outreach Log";
const RESOURCE_MAP_SHEET               = "Resource Map";
const ZOOM_INTRO_SHEET                 = "Intro_Poll"; 
const ZOOM_CLOSING_SHEET               = "Closing_Poll"; 

const TEST_ALL_PARTICIPANTS_SHEET   = "Test All Participants";
const TEST_ACTIVE_PARTICIPANTS_SHEET   = "Test Active Participants";
const TEST_TASK_DATA_SHEET             = "Test Task Data";

// ============================================================================
// II. PLANNER SETUP CONFIGURATION
// ============================================================================
const PLANNER_FOLDER_ID_CELL    = "B1";  // Cell with Drive folder ID
const PLANNER_TEMPLATE_ID_CELL  = "B2";  // Cell with Planner template ID
const STATUS_CELL               = "B3";  // Cell to show setup status messages

// ============================================================================
// III. COMMON STRUCTURE
// ============================================================================
const HEADER_ROW     = 1;  // Header row number in all sheets
const DATA_START_ROW = 2;  // First data row number in all sheets

// ============================================================================
// IV. COMMON COLUMN HEADERS
// ============================================================================
const COL_PTID        = "PTID";
const COL_FIRST_NAME  = "First Name";
const COL_LAST_NAME   = "Last Name";
const COL_EMAIL       = "Email";

// ============================================================================
// V. PARTICIPANT COLUMNS
// ============================================================================
const COL_PLANNER        = "Planner Link";
const COL_BITLY_PLANNER = "Planner Bitly Link";

const COL_ATTENDED       = "Attended";
const COL_PARTICIPATED  = "Participated";
const COL_DIDNT_ATTEND = "Didn't Attend";

const COL_HAS_FEAR = "Has Fear";
const COL_HAS_FEAR_REASON = "Has Fear Reasons";
const COL_HAS_PURPOSE = "Has Purpose";
const COL_LAST_MODIFIED   = "Last Modified";
const COL_HAS_DEADLINE = "Has Deadline";
const COL_FOLLOW_THRU = "Follow Through %";
const COL_CLUSTER = "Cluster";

const COL_CURRENT_PROGRESS = "Current Progress";
const COL_S0_PROGRESS = "S0_Progress";
const COL_PCTCHANGE_LAST2WEEKS = "% Change Last 2 Weeks";
const COL_NUMCHANGE_LAST2WEEKS = "# Change Last 2 Weeks";

const COL_PCTCHANGE_S1 = "% Change Sprint 1";
const COL_NUMCHANGE_S1 = "# Change Sprint 1";
const COL_PCTCHANGE_S2 = "% Change Sprint 2";
const COL_NUMCHANGE_S2 = "# Change Sprint 2";
const COL_PCTCHANGE_S3 = "% Change Sprint 3";
const COL_NUMCHANGE_S3 = "# Change Sprint 3";

const COL_PCT_EXP_DONE = "% Explore Done";
const COL_EXP_LEFT = "Explore Left";
const COL_PCT_RES_DONE = "% Research Done";
const COL_RES_LEFT = "Research Left";
const COL_PCT_APP_DONE = "% Apply Done";
const COL_APP_LEFT = "Apply Left";
const COL_PCT_STR_DONE = "% Strategize Done";
const COL_STR_LEFT = "Strategize Left";

const COL_JOB_SECURED    = "Job Secured?";
const COL_JOB_START      = "Job Secured Date";


// Confidence Outcomes
const COL_S0_CONF_PRE = "S0_Confidence_Pre";    // Pre Kick-Off
const COL_S0_CONF_POST = "S0_Confidence_Post";  // Post Kick-Off
const COL_S1_CONF = "S1_Confidence";            // Post Sprint 1
const COL_S2_CONF = "S2_Confidence";            // Post Sprint 2
const COL_S3_CONF = "S3_Confidence";            // Post Sprint 3

// Job Search Stage Outcomes
const COL_S0_STAGE = "S0_Stage";                // Kick-Off
const COL_S1_STAGE = "S1_Stage";                // Post Sprint 1
const COL_S2_STAGE = "S2_Stage";                // Post Sprint 2
const COL_S3_STAGE = "S3_Stage";                // Post Sprint 3

const COL_RSVP_NOTE = "RSVP Note";

// Email sent flags:
const COL_KICKOFF_SENT = "FollowUp Sent"; 

const SENT_FLAG_COLUMNS = {
  "S1Wk1": "S1Week1 Sent",
  "S1Wk2": "S1Week2 Sent",
  "S1Wk3": "S1Week3 Sent",
  "S1Wk4": "S1Week4 Sent",
  "S2Wk1": "S2Week1 Sent",
  "S2Wk2": "S2Week2 Sent",
  "S2Wk3": "S2Week3 Sent",
  "S2Wk4": "S2Week4 Sent",
  "S3Wk1": "S3Week1 Sent",
  "S3Wk2": "S3Week2 Sent",
  "S3Wk3": "S3Week3 Sent",
  "S3Wk4": "S3Week4 Sent",
};

// ============================================================================
// VI. Zoom Poll Sheets
// ============================================================================
const COL_USER_NAME    = "User Name"

// ============================================================================
// VII. OUTREACH COLUMNS
// ============================================================================
const COL_OUTREACH_NUM = "Outreach #";
const COL_TEMPLATE     = "Template";
const COL_SPRINT       = "Sprint";
const COL_WEEK         = "Week";
const COL_SUBJECT_LINE = "Subject Line";
const COL_LINK_TO_DRAFT= "Link to Draft";
const COL_DOC_ID       = "Google Doc ID";
const COL_SEND_DATE    = "Send Date";
const COL_STATUS       = "Status";
const COL_TIMESTAMP = "Timestamp";

// ============================================================================
// VIII. TASK DATA SHEET COLUMNS
// ============================================================================
const COL_TASK_ID       = "Task ID";
const COL_TASK          = "Task";
const COL_DEADLINE      = "Deadline";
const COL_DUE_THIS_WEEK = "Due This Week?";
const COL_OVERDUE       = "Overdue?";
const COL_DONE          = "Done?";
const COL_LAST_UPDATE   = "Last Update";

// ============================================================================
// IX. RESOURCE MAP SHEET COLUMNS
// ============================================================================
const COL_BUTTON_TEXT  = "Button Text";
const COL_DESCRIPTION  = "Description";
const COL_RESOURCE_URL = "Resource URL";

// ============================================================================
// X. FEARS & PURPOSE SHEET COLUMNS
// ============================================================================
const COL_FEAR    = "Fear";
const COL_FEAR_REASON    = "Fear Reasons";
const COL_PURPOSE = "Purpose";

// ============================================================================
// XI. JOB FORM COLUMNS
// ============================================================================
const COL_FORM_NAME    = "What's your name?";
const COL_FORM_EMAIL   = "What's your email address? Use the one you normally use with Bottom Line so we can match your registration.";
const COL_FORM_JOB     = "Have you accepted a job offer for after graduation?";
const COL_FORM_MATCHED = "Matched";

// ============================================================================
// XII. STUDENT PLANNER TEMPLATE CONSTANTS
// ============================================================================
const PLANNER_WELCOME_SHEET   = "🌟 Welcome";
const PLANNER_FULLNAME_CELL   = "C2";
const PLANNER_WELCOME_EMAIL   = "C4";
const PLANNER_FEAR_SHEET      = "😣 Fears";
const PLANNER_FEAR_CELL       = "B3";
const PLANNER_FEAR_REASONS_CELL       = "B7";
const PLANNER_PURPOSE_SHEET   = "🎯 Purpose";
const PLANNER_PURPOSE_CELL    = "B13";
const PLANNER_CHECKLIST_SHEET = "✅ Checklist";
const PLANNER_SPRINT_SHEET    = "🏃🏼‍♂️ Sprint Planner";
const PLANNER_COL_ID          = "ID";
const PLANNER_COL_DONE        = "Done?";
const PLANNER_COL_DEADLINE    = "Deadline";
const PLANNER_CALENDAR_SHEET = "📆 Sprint Calendar";

// Mapping from Task Data columns to student header names in planner:
const STUDENT_HEADER_MAP = {
  "Task ID": PLANNER_COL_ID,
  Task:      "Task",
  Deadline:  PLANNER_COL_DEADLINE,
  "Done?":   PLANNER_COL_DONE
};

// Add this to your GlobalConstants.gs file

const STAGE_TO_TASK_MAP = {
  "I haven’t started yet":                                       ['EXP1', 'EXP4', 'EXP5'],
  "I’ve browsed job listings and boards but haven’t started preparing materials": ['EXP4', 'RES7', 'EXP2'],
  "I’ve created or updated my materials but haven’t applied yet":      ['APP1', 'RES5', 'RES3'],
  "I’ve submitted an application to at least one job":               ['APP3', 'APP7', 'APP4'],
  "I’ve applied to multiple jobs and have started hearing back":      ['APP8', 'STR1', 'STR3']
};

/**
 * Spreadsheet UI Set-Up
 *
 * Creates a custom menu in the Google Sheet UI when the spreadsheet is opened.
 * This allows users to run the script's functions easily without needing to open the script editor.
 */
function onOpen() {
  SpreadsheetApp.getUi() // Gets the user interface environment for the spreadsheet.
    .createMenu("Admin Tools") // Creates a new top-level menu named "Admin Tools".
    // -----------------------------------------------------------------------------
    .addItem("🚀 Set Up Student Planners", "setupStudentPlanners")      // Set Up Planners.
    .addSeparator() // Horizontal line to separate items.
    // -----------------------------------------------------------------------------
    .addItem("😣 Collect Student Fears", "collectStudentFears")         // Collect fears
    .addSeparator()
    // -----------------------------------------------------------------------------
    .addItem("✅ Checked Off Tasks", "markCheckedOffTasks")             // Mark students who checked off tasks
    .addItem("💬 Match Zoom Answers", "matchZoomResults")               // Match Poll answers to participants by name
    .addItem("🕊️ Collect Student Purpose", "collectStudentPurpose")     // Collect student purpose
    .addSeparator() 
    .addItem("📊 Update Student Task Data", "fetchAllStudentTaskData")  // Fetch student task data.
    .addItem("👀 Update Last Modified Date", "updateTrueLastModified")
    .addItem("💡 Set Deadlines for Students With No Plan", "runUiSetDeadlines")
    .addSeparator() 
    // -----------------------------------------------------------------------------
    .addItem("✉️ Send Kickoff Follow-Up", "sendKickoffFollowUp") // Adds a menu item to send the kickoff email.
    .addItem("✉️ Send Weekly Nudge", "sendWeeklyNudge") // Adds a menu item to send a weekly nudge email.
    .addToUi(); // Finalizes the menu and adds it to the spreadsheet's UI.
}
