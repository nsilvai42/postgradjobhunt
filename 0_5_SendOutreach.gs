/**
 * 
 * Script Name: 0_5_SendOutreach.gs
 * 
 * 
/

// ============================================================================
// SECTION 1: SETTINGS & GLOBALS
// ============================================================================

// Sender display name for outgoing emails
const SENDER_NAME = 'Career Sprint Team';
// Reply-to address for all emails
const REPLY_TO = 'bottomlinesuccesscareers@gmail.com';
// Addresses to ignore in logging
const IGNORED_LOGGING_REGEX = /nsilva@ideas42\.org|nikolas/i;


// ============================================================================
// SECTION 2: TEMPLATE & CLUSTER MAPS
// ============================================================================

// Mapping of logical template names to file IDs
const TEMPLATE_MAP = {
  EmailStyles: '1_01_EmailStyles',
  EmailHeader: '1_02_EmailHeader',
  EmailFooter: '1_03_EmailFooter',
  PlannerButton: '1_04_PlannerButton',
  KickOffLinks: '1_05_KickOffLinks',
  KickOffTasks: '1_06_KickOffTasks',
  SingleTaskItem: '1_07_SingleTaskItem',
  TaskGroup: '1_08_TaskGroup',
  UpcomingTasks: '1_09_UpcomingTasks',
  NoDeadlines: '1_10_NoDeadlines',
  OverdueTasks: '1_11_OverdueTasks',
  ERASStatement: '1_12_ERASStatement',
  ERASBullets: '1_13_ERASBullets',
  ProgressBar: '1_14_ProgressBar',
  KickOff_DidntAttend: '2_01_KO_DidntAttend',
  KickOff_DidntParticipate: '2_02_KO_DidntParticipate',
  KickOff_Participated: '2_03_KO_Participated',
  S1Wk1: '2_04_S1Wk1',
  S1Wk2: '2_05_S1Wk2',
  S1Wk3_OnTrack: '2_06_S1Wk3_OnTrack',
  S1Wk3_OffTrack: '2_07_S1Wk3_OffTrack',
  S1Wk3_Disengaged: '2_08_S1Wk3_Disengaged',
  S1Wk4: '2_09_S1Wk4',
  S2Wk1: '2_10_S2Wk1',
  S2Wk2: '2_11_S2Wk2',
  S2Wk3: '2_12_S2Wk3',
  OnTrack_Intro: '3_01_S1Wk3_OnTrack_Intro',
  OffTrack_Intro: '3_02_S1Wk3_OffTrack_Intro',
  OffTrack_PurposeFear: '3_03_S1Wk3_OffTrack_PurposeFear',
  OffTrack_Progress: '3_04_S1Wk3_OffTrack_Progress'
};

// Cluster-to-template mapping (duplicates kept as-is)
const CLUSTER_TO_TEMPLATE_MAP = {
  Hunters: 'S1Wk3_OnTrack',
  Trackers: 'S1Wk3_OffTrack',
  Trappers: 'S1Wk3_OffTrack',
  Scouts: 'S1Wk3_Disengaged',
  Gatherers: 'S1Wk3_Disengaged',
  Hunters: 'S2Wk2_Hunters',     // overwrites earlier Hunters
  Trackers: 'S2Wk2_Hunters',   // overwrites earlier Trackers
  Scouts: 'S2Wk2_Scouts'       // overwrites earlier Scouts
};

// Optional overrides for sender name or reply-to per template
const TEMPLATE_OVERRIDES = {
  S1Wk3_OffTrack: { fromName: 'Niko Silva', replyTo: 'nsilva@ideas42.org' },
  S1Wk3_Disengaged: { fromName: 'Niko Silva', replyTo: 'nsilva@ideas42.org' },
  S2Wk2_Trackers: { fromName: 'Niko Silva', replyTo: 'nsilva@ideas42.org' },
  S2Wk2_Scouts: { fromName: 'Niko Silva', replyTo: 'nsilva@ideas42.org' }
};


// ============================================================================
// SECTION 3: LINK PLACEHOLDERS
// ============================================================================

// Reusable URLs for all templates
const LINK_PLACEHOLDERS = {
  feedbackSurvey: 'https://forms.gle/fPuxrEbrmKFpZ4bk9',
  jobForm: 'https://bottomlinecareer.org/job-form',
  kickoffDeck: 'https://bottomlinecareer.org/kick-off-slides',
  kickoffRecording: 'https://bottomlinecareer.org/kick-off-recording',
  tealTracker: 'https://www.tealhq.com/tools/job-tracker',
  tealTrackerVideo: 'https://www.youtube.com/watch?v=PoCDQp2idHc',
  setupGuide: 'https://bottomlinecareer.org/planner-set-up-guide'
};


// ============================================================================
// SECTION 4: PARTIAL & TEMPLATE HELPERS
// ============================================================================

/**
 * include(filename, data)
 *   - Loads a named HTML template
 *   - Merges LINK_PLACEHOLDERS + provided data
 *   - Renders to HTML string
 */
function include(filename, data = {}) {
  // --- Step 1: Resolve template name ---
  const templateName = TEMPLATE_MAP[filename] || filename;
  Logger.log('include: loading template %s', templateName);

  // --- Step 2: Prepare template context ---
  const tpl = HtmlService.createTemplateFromFile(templateName);
  Object.assign(tpl, LINK_PLACEHOLDERS, data);

  // --- Step 3: Render HTML ---
  const htmlContent = tpl.evaluate().getContent();

  // --- Step 4: Log result ---
  Logger.log('include: rendered %s (%d chars)', templateName, htmlContent.length);
  return htmlContent;
}


// ============================================================================
// SECTION 5: EMOJI & LOGGING HELPERS
// ============================================================================

/**
 * replaceEmojis(html)
 *   - Swaps emoji characters for HTML entities
 *   - Logs before/after lengths
 */
function replaceEmojis(html) {
  Logger.log('replaceEmojis: start (len=%d)', html.length);

  // --- Perform replacements ---
  const mapped = html
    .replace(/🎯/g, '&#x1F3AF;')      // target emoji
    .replace(/💬/g, '&#x1F4AC;')      // speech bubble
    .replace(/🧠/g, '&#x1F9E0;')      // brain
    .replace(/💡/g, '&#x1F4A1;')      // light bulb
    .replace(/🏃/g, '&#x1F3C3;')      // running person
    .replace(/⏰/g, '&#x23F0;')       // alarm clock
    .replace(/🎉/g, '&#x1F389;')      // party popper
    .replace(/👏/g, '&#x1F44F;')      // clapping hands
    .replace(/🏅/g, '&#x1F3C5;')      // medal
    .replace(/👉/g, '&#x1F449;')      // pointing finger
    .replace(/🗺️/g, '&#x1F5FA;')     // map
    .replace(/🔬/g, '&#x1F52C;')      // microscope
    .replace(/📝/g, '&#x1F4DD;')      // memo
    .replace(/🛠️/g, '&#xFE0F;')      // hammer
    .replace(/🗓️/g, '&#x1F5D3;');     // calendar

  Logger.log('replaceEmojis: end (len=%d)', mapped.length);
  return mapped;
}



// ============================================================================
// SECTION 6: STATISTICAL HELPERS
// ============================================================================

/**
 * calculatePercentile(value, allValues)
 *   - Sorts `allValues` ascending
 *   - Finds the index where `value` fits
 *   - Returns percentile 0–100
 */
function calculatePercentile(value, allValues) {
  Logger.log('calculatePercentile: start (value=%d)', value);

  // Copy and sort the array
  const sorted = allValues.slice().sort((a, b) => a - b);

  // Find first index ≥ value
  const idx = sorted.findIndex(v => value <= v);
  if (idx < 0) {
    Logger.log('calculatePercentile: value above max → 100');
    return 100;
  }

  // Compute and log
  const pct = Math.round((idx / sorted.length) * 100);
  Logger.log('calculatePercentile: result=%d', pct);
  return pct;
}


/**
 * computeFollowThrough(tasks, today, tIdx)
 *   - Only looks at tasks with a deadline
 *   - Counts “on‐track” = done OR deadline ≥ today
 *   - Counts “overdue” = not done AND deadline < today
 *   - Returns fraction on‐track (0–1) and percent (0–100)
 */
function computeFollowThrough(tasks, today, tIdx) {
  // filter to only tasks the student has actually given a deadline
  const withDL = tasks.filter(t => t[tIdx[COL_DEADLINE]]);
  const total = withDL.length;
  if (!total) {
    // no plan = not on track
    return { onTrackProp: null, onTrackPct: null };
  }

  const onTrackCount = withDL.filter(t => {
    const dl = new Date(t[tIdx[COL_DEADLINE]]);
    // either already done, or still due in the future
    return t[tIdx[COL_DONE]] === true || dl >= today;
  }).length;

  const onTrackProp = onTrackCount / total;
  return {
    onTrackProp,
    onTrackPct: Math.round(onTrackProp * 100)
  };
}


/**
 * computeLeaderboardStats(aRows, tRows, aIdx, tIdx)
 *   - Aggregates totals and deltas for all participants
 *   - Calculates average tasks done and follow-through
 *   - Gathers arrays of stage-completion percentages
 */
function computeLeaderboardStats(aRows, tRows, aIdx, tIdx) {
  const deltas = [];
  let totalDone = 0;
  const allExplore = [], allResearch = [], allApply = [], allStrategize = [];

  aRows.forEach(row => {
    const ptid = String(row[aIdx[COL_PTID]]).trim();
    if (!ptid) return;
    const cur = (parseFloat(String(row[aIdx[COL_CURRENT_PROGRESS]]).replace('%', '')) || 0) / 100;
    const base = (parseFloat(String(row[aIdx[COL_S0_PROGRESS]]).replace('%', '')) || 0) / 100;
    deltas.push(cur - base);
    const done = tRows.filter(t =>
      String(t[tIdx[COL_PTID]]).trim() === ptid && t[tIdx[COL_DONE]] === true
    ).length;
    totalDone += done;
    allExplore.push(parseFloat(String(row[aIdx[COL_PCT_EXP_DONE]]).replace('%', '')) || 0);
    allResearch.push(parseFloat(String(row[aIdx[COL_PCT_RES_DONE]]).replace('%', '')) || 0);
    allApply.push(parseFloat(String(row[aIdx[COL_PCT_APP_DONE]]).replace('%', '')) || 0);
    allStrategize.push(parseFloat(String(row[aIdx[COL_PCT_STR_DONE]]).replace('%', '')) || 0);
  });

  // compute cohort average onTrackPct
  const allOnTrack = [];
  aRows.forEach(row => {
    const ptid = String(row[aIdx[COL_PTID]]).trim();
    const tasks = tRows.filter(t => String(t[tIdx[COL_PTID]]).trim() === ptid);
    const { onTrackPct } = computeFollowThrough(tasks, new Date(), tIdx);
    if (onTrackPct !== null) {
      allOnTrack.push(onTrackPct);
    }
  });

  const avgOnTrackPct = allOnTrack.length
    ? allOnTrack.reduce((s, v) => s + v, 0) / allOnTrack.length
    : 0;

  const n = aRows.length;
  return {
    deltas,
    avgDone: n ? totalDone / n : 0,
    avgOnTrackPct,
    allExploreScores: allExplore,
    allResearchScores: allResearch,
    allApplyScores: allApply,
    allStrategizeScores: allStrategize,
    avgExplore: allExplore.reduce((s, v) => s + v, 0) / allExplore.length,
    avgResearch: allResearch.reduce((s, v) => s + v, 0) / allResearch.length,
    avgApply: allApply.reduce((s, v) => s + v, 0) / allApply.length,
    avgStrategize: allStrategize.reduce((s, v) => s + v, 0) / allStrategize.length
  };
}


// ============================================================================
// SECTION 8: STUDENT PERCENTILES HELPER
// ============================================================================

/**
 * computeStudentPercentiles(student, deltas, allScores)
 *   - Calculates overall growth percentile
 *   - Calculates each ERAS stage percentile
 */
function computeStudentPercentiles(student, deltas, allScores) {
  Logger.log('computeStudentPercentiles: start (ptid=%s)', student.ptid);

  // 1) Growth metric
  const currentPct = parseFloat(String(student.currentProgress).replace('%', '')) || 0;
  const basePct = parseFloat(String(student.s0Progress).replace('%', '')) || 0;
  const growthDelta = currentPct - basePct;
  const lowerCount = deltas.filter(d => d < growthDelta).length;
  const growthPct = Math.round((lowerCount / deltas.length) * 100);
  const topRank = 100 - growthPct;

  // 2) ERAS stage percentiles
  const explorePct = calculatePercentile(
    parseFloat(String(student.pctExploreDone).replace('%', '')) || 0,
    allScores.explore
  );
  const researchPct = calculatePercentile(
    parseFloat(String(student.pctResearchDone).replace('%', '')) || 0,
    allScores.research
  );
  const applyPct = calculatePercentile(
    parseFloat(String(student.pctApplyDone).replace('%', '')) || 0,
    allScores.apply
  );
  const strategizePct = calculatePercentile(
    parseFloat(String(student.pctStrategizeDone).replace('%', '')) || 0,
    allScores.strategize
  );

  Logger.log('computeStudentPercentiles: end (ptid=%s)', student.ptid);
  return {
    growthPct,
    topRank,
    explorePercentile: explorePct,
    researchPercentile: researchPct,
    applyPercentile: applyPct,
    strategizePercentile: strategizePct
  };
}

// ============================================================================
// SECTION 8: EMAIL RENDER & SEND
// ============================================================================

/**
 * _renderAndSendEmail(mailOpts, tplOpts)
 *   - Renders HTML, applies emoji fixes, sends mail, logs summary
 */
function _renderAndSendEmail(mailOpts, tplOpts) {
  Logger.log('_renderAndSendEmail: start (to=%s, template=%s)', mailOpts.to, tplOpts.name);

  try {
    // --- Step 1: Render template ---
    const tpl = HtmlService.createTemplateFromFile(tplOpts.name);
    Object.assign(tpl, LINK_PLACEHOLDERS, tplOpts.data);
    const rawHtml = tpl.evaluate().getContent();
    Logger.log('_renderAndSendEmail: rendered HTML length=%d', rawHtml.length);

    // --- Step 2: Replace emojis ---
    const finalHtml = replaceEmojis(rawHtml);

    // --- Step 3: Send email via GmailApp ---
    GmailApp.sendEmail(mailOpts.to, mailOpts.subject, '', {
      htmlBody: finalHtml,
      name: mailOpts.fromName || SENDER_NAME,
      replyTo: mailOpts.replyTo || REPLY_TO
    });

    Logger.log('_renderAndSendEmail: success (to=%s)', mailOpts.to);
    return true;

  } catch (err) {
    Logger.log('_renderAndSendEmail: error (to=%s, error=%s)', mailOpts.to, err.toString());
    return false;
  }
}


// ============================================================================
// SECTION 9: ENTRY POINTS
// ============================================================================

// Manual trigger with UI
function sendWeeklyNudge() { _sendWeeklyNudge(true); }
// Time-driven trigger without UI
function sendWeeklyNudgeTrigger() { _sendWeeklyNudge(false); }


// ============================================================================
// SECTION 10: CORE LOGIC: _sendWeeklyNudge
// ============================================================================

/**
 * _sendWeeklyNudge(showUi)
 *   - Orchestrates loading data, computing stats, building student objects,
 *     and sending templated reminder emails
 */
function _sendWeeklyNudge(showUi) {
  Logger.log('_sendWeeklyNudge: start (UI=%s)', showUi);

  // --- 10.1 Setup & sheet references ---
  const ss = SpreadsheetApp.getActive();
  // const ui = showUi ? ss.getUi() : null;
  const tz = ss.getSpreadsheetTimeZone();
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const todayStr = Utilities.formatDate(today, tz, 'M/d/yyyy');

  const planSheet = ss.getSheetByName(OUTREACH_PLAN_SHEET);
  const partSheet = ss.getSheetByName(ACTIVE_PARTICIPANTS_SHEET);
  const taskSheet = ss.getSheetByName(TASK_DATA_SHEET);
  const fpSheet = ss.getSheetByName(FEARS_PURPOSE_SHEET);
  const resSheet = ss.getSheetByName(RESOURCE_MAP_SHEET);
  const logSheet = ss.getSheetByName(OUTREACH_LOG_SHEET);

  // --- 10.2 Load & filter today’s plan rows ---
  const [pHdr, ...pRows] = planSheet.getDataRange().getValues();
  const pIdx = Object.fromEntries(pHdr.map((h, i) => [h, i]));
  const todaysPlans = pRows.filter(r => {
    const cell = r[pIdx[COL_SEND_DATE]];
    const d = cell instanceof Date
      ? Utilities.formatDate(cell, tz, 'M/d/yyyy')
      : String(cell).trim();
    return d === todayStr;
  });
  if (!todaysPlans.length) {
    Logger.log('_sendWeeklyNudge: no plans for %s, exiting', todayStr);
    if (showUi) ui.alert('No plans for today');
    return;
  }

  // --- 10.3 Load all sheet data ---
  const [aHdr, ...aRows] = partSheet.getDataRange().getValues();
  const aIdx = Object.fromEntries(aHdr.map((h, i) => [h, i]));
  const [tHdr, ...tRows] = taskSheet.getDataRange().getValues();
  const tIdx = Object.fromEntries(tHdr.map((h, i) => [h, i]));
  const [rHdr, ...rRows] = resSheet.getDataRange().getValues();
  const rIdx = Object.fromEntries(rHdr.map((h, i) => [h, i]));
  const [fHdr, ...fRows] = fpSheet.getDataRange().getValues();
  const fIdx = Object.fromEntries(fHdr.map((h, i) => [h, i]));
  Logger.log('_sendWeeklyNudge: loaded sheet data (participants=%d, tasks=%d, resources=%d, fears=%d)',
    aRows.length, tRows.length, rRows.length, fRows.length);

  // --- 10.4 Compute global stats ---
  const {
    deltas,
    avgDone,
    avgOnTrackPct,
    allExploreScores,
    allResearchScores,
    allApplyScores,
    allStrategizeScores,
    avgExplore,
    avgResearch,
    avgApply,
    avgStrategize
  } = computeLeaderboardStats(aRows, tRows, aIdx, tIdx);

  Logger.log('_sendWeeklyNudge: leaderboard stats computed');

  // --- 10.5 Freeze key clusters ---
  const frozenClusters = {};
  aRows.forEach(r => {
    const id = String(r[aIdx[COL_PTID]]).trim();
    const cl = String(r[aIdx[COL_CLUSTER]]).trim();
    if (cl === 'Gatherers' || cl === 'Scouts') frozenClusters[id] = cl;
  });
  Logger.log('_sendWeeklyNudge: frozenClusters count=%d', Object.keys(frozenClusters).length);

  // Prepare cohort benchmarks & full-scores object
  const allScores = {
    explore: allExploreScores, research: allResearchScores,
    apply: allApplyScores, strategize: allStrategizeScores
  };
  const cohortBenchmarks = {
    avgDone,
    avgExplore,
    avgResearch,
    avgApply,
    avgStrategize,
    avgOnTrackPct
  };

  // --- 9.6 Build student objects ---
  const students = {};
  aRows.forEach((row, i) => {
    const ptid = String(row[aIdx[COL_PTID]]).trim();
    const email = String(row[aIdx[COL_EMAIL]]).toLowerCase().trim();
    if (!ptid || !email) return;

    // 1) Cluster
    const liveCluster = String(row[aIdx[COL_CLUSTER]]).trim();
    const finalCluster = frozenClusters[ptid] || liveCluster;

    // 2) Growth delta (0–1)
    const rawCur = String(row[aIdx[COL_CURRENT_PROGRESS]]).replace('%', '');
    const rawS0 = String(row[aIdx[COL_S0_PROGRESS]]).replace('%', '');
    const currentProgress = (parseFloat(rawCur) || 0) / 100;
    const s0Progress = (parseFloat(rawS0) || 0) / 100;
    const growthDelta = currentProgress - s0Progress;

    // 3) Task counts & follow-through
    const myTasks = tRows.filter(t => String(t[tIdx[COL_PTID]]).trim() === ptid);
    const doneCount = myTasks.filter(t => t[tIdx[COL_DONE]] === true).length;
    const { onTrackProp, onTrackPct } = computeFollowThrough(myTasks, today, tIdx);

    // 4) Raw ERAS stage % → numeric 0–100
    const explorePct = parseFloat(String(row[aIdx[COL_PCT_EXP_DONE]]).replace('%', '')) || 0;
    const researchPct = parseFloat(String(row[aIdx[COL_PCT_RES_DONE]]).replace('%', '')) || 0;
    const applyPct = parseFloat(String(row[aIdx[COL_PCT_APP_DONE]]).replace('%', '')) || 0;
    const strategizePct = parseFloat(String(row[aIdx[COL_PCT_STR_DONE]]).replace('%', '')) || 0;

    // 5) Percentiles & rank
    const growthPctile = calculatePercentile(growthDelta, deltas);
    const topRank = 100 - growthPctile;

    const explorePctile = calculatePercentile(explorePct, allExploreScores);
    const researchPctile = calculatePercentile(researchPct, allResearchScores);
    const applyPctile = calculatePercentile(applyPct, allApplyScores);
    const strategizePctile = calculatePercentile(strategizePct, allStrategizeScores);

    const percentiles = {
      growthPct: growthPctile,
      topRank,
      explorePctile,
      researchPctile,
      applyPctile,
      strategizePctile
    };

    // 6) Deviations vs cohort averages
    const diffs = {
      growthVsAvg: (growthDelta * 100) - (avgDone * 100),
      onTrackVsAvg: (onTrackPct || 0) - avgOnTrackPct,
      exploreVsAvg: explorePct - avgExplore,
      researchVsAvg: researchPct - avgResearch,
      applyVsAvg: applyPct - avgApply,
      strategizeVsAvg: strategizePct - avgStrategize
    };

    // 7) Assemble the final student object
    students[ptid] = {
      // identity
      ptid,
      email,
      name: row[aIdx[COL_FIRST_NAME]],
      cluster: finalCluster,
      row: i + 2,
      hasJob: row[aIdx[COL_JOB_SECURED]] === true,

      // progress
      currentProgress,
      s0Progress,
      growthDelta,

      // tasks & follow-through
      doneCount,
      onTrackProp,
      onTrackPct,

      // ERAS raw %
      explorePct,
      researchPct,
      applyPct,
      strategizePct,

      // percentiles & rank
      percentiles,

      // deviations
      diffs,

      // cohort & carry-through fields
      cohort: cohortBenchmarks,
      allScores,
      nChangePrevWeeks: row[aIdx[COL_NUMCHANGE_LAST2WEEKS]] || 0,
      pctChangePrevWeeks: row[aIdx[COL_PCTCHANGE_LAST2WEEKS]] || 0,
      kickoffJobStage: row[aIdx[COL_S0_STAGE]],
      rsvpNote: row[aIdx[COL_RSVP_NOTE]],
      purpose: '',
      fear: '',
      fearReason: '',
      highlightProgress: true
    };
  });
  Logger.log('_sendWeeklyNudge: built %d student objects', Object.keys(students).length);

  // Compute onTrack percentiles
  const allOnTrack = Object.values(students)
    .map(s => s.onTrackPct)
    .filter(v => v != null);

  Object.values(students).forEach(s => {
    s.onTrackPctile = s.onTrackPct == null
      ? null
      : calculatePercentile(s.onTrackPct, allOnTrack);
  });


  // --- 10.7 Merge fears & purpose data ---
  Object.values(students).forEach(s => {
    const fearRow = fRows.find(r => String(r[fIdx[COL_PTID]]).trim() === s.ptid);
    if (fearRow) {
      s.purpose = fearRow[fIdx[COL_PURPOSE]] || '';
      s.fear = fearRow[fIdx[COL_FEAR]] || '';
      s.fearReason = fearRow[fIdx[COL_FEAR_REASON]] || '';
    }
  });
  Logger.log('_sendWeeklyNudge: merged fears & purposes');

  // --- 10.8 Send Loop: prepare & dispatch emails ---
  let stats = { sent: 0, skipped: 0, failed: 0 };
  const logEntries = [];

  Object.values(students).forEach(student => {
    if (student.hasJob) { stats.skipped++; return; }




    // --- select plan row & template key (unchanged) ---
    const planRow = (todaysPlans.length === 1)
      ? todaysPlans[0]
      : todaysPlans.find(r => r[pIdx[COL_TEMPLATE]] === CLUSTER_TO_TEMPLATE_MAP[student.cluster]);
    if (!planRow) { stats.skipped++; return; }
    const tplKey = planRow[pIdx[COL_TEMPLATE]];

    // --- sent‑flag guard ---
    const sentFlagCol = SENT_FLAG_COLUMNS[tplKey] || SENT_FLAG_COLUMNS[tplKey.split('_')[0]];
    if (!sentFlagCol || !aIdx[sentFlagCol]) { stats.skipped++; return; }
    if (partSheet.getRange(student.row, aIdx[sentFlagCol] + 1).getValue()) { stats.skipped++; return; }

    // -------------------------------------------------
    // SUBJECT‑LINE CONSTRUCTION (RESTORED)
    // -------------------------------------------------
    let rawSub = String(planRow[pIdx[COL_SUBJECT_LINE]] || '');
    if (TEMPLATE_OVERRIDES[tplKey]?.subjectTransform) {
      rawSub = TEMPLATE_OVERRIDES[tplKey].subjectTransform(rawSub, student);
    }
    rawSub = rawSub.replace('<?!= firstName ?>', student.name);
    const emailSubject = '=?utf-8?B?' + Utilities.base64Encode(Utilities.newBlob(rawSub).getBytes()) + '?=';

    let hasDoneEXP4 = false;
    let hasDoneRES3 = false;
    let hasDoneAPP7 = false;

    // Only run this check if the student is in the Dormant cluster
    if (student.cluster === 'Dormant') {
      // Get all tasks for this specific student
      const studentTasks = tRows.filter(t => String(t[tIdx[COL_PTID]]).trim() === student.ptid);

      // Check if the specific tasks have been marked as done
      hasDoneEXP4 = studentTasks.some(t => t[tIdx[COL_TASK_ID]] === 'EXP4' && t[tIdx[COL_DONE]] === true);
      hasDoneRES3 = studentTasks.some(t => t[tIdx[COL_TASK_ID]] === 'RES3' && t[tIdx[COL_DONE]] === true);
      hasDoneAPP7 = studentTasks.some(t => t[tIdx[COL_TASK_ID]] === 'APP7' && t[tIdx[COL_DONE]] === true);
    }


    // -------------------------------------------------
    // TASK GROUPS & RESCHEDULE LIST (unchanged)
    // -------------------------------------------------
    const resourceMap = buildResourceMap(rRows, rIdx);
    const { TaskGroups, OverdueGroups, NextTasksByDate } = _buildTaskGroupsForStudent(
      student, tRows, tIdx, tz, today, resourceMap
    );
    let rescheduleList = [];
    if (Object.keys(OverdueGroups).length) {
      const overdueTasks = [].concat(...Object.values(OverdueGroups));
      const allDeadlines = [
        ...[].concat(...Object.values(TaskGroups)).map(o => new Date(o.deadline)),
        ...overdueTasks.map(o => new Date(o.deadline)),
        ...[].concat(...Object.values(NextTasksByDate)).map(o => new Date(o.deadline))
      ];
      rescheduleList = generateRescheduleDates(overdueTasks, allDeadlines, today, tz);
    }

    // -------------------------------------------------
    // PLANNER & RESOURCE LINK CONSTRUCTION (UPDATED)
    // -------------------------------------------------
    const prowRow = partSheet.getRange(student.row, 1, 1, partSheet.getLastColumn()).getValues()[0];

    // Use the Bitly link for the main planner button, as you currently do.
    const plannerLink = prowRow[aIdx[COL_BITLY_PLANNER]] || prowRow[aIdx[COL_PLANNER]];

    // **NEW:** Create the direct link to the Resource Center tab.
    // It takes the full planner URL and appends the specific GID.
    const fullPlannerUrl = prowRow[aIdx[COL_PLANNER]];
    const resourceCenterLink = `${fullPlannerUrl}#gid=915503576`;

    // -------------------------------------------------
    // TEMPLATE‑DATA OBJECT (PATCHED)
    // -------------------------------------------------
    // assemble template data
    const templateData = Object.assign(
      {},
      student,
      planRow.reduce((o, h, i) => (o[h] = planRow[i], o), {}),

      // overrides / additions
      {
        firstName: student.name,
        rawSubject: rawSub,
        planRow,
        careerPlanner: plannerLink,
        resourceLink: resourceCenterLink,

        // tasks & schedule
        TaskGroups,
        OverdueGroups,
        NextTasksByDate,
        rescheduleList,
        resourceMap,

        // counts
        nTasksDone: student.doneCount,
        avgTasksDone: student.avgTasksDoneWhole,

        // follow-through
        onTrackProp: student.onTrackProp,
        onTrackPct: student.onTrackPct,
        onTrackPctile: student.onTrackPctile,
        onTrackVsAvg: student.onTrackPct - cohortBenchmarks.avgOnTrackPct,
        // growth & percentiles
        growthDelta: student.growthDelta,
        growthPct: student.percentiles.growthPct,
        topRank: student.percentiles.topRank,
        growthVsAvg: student.diffs.growthVsAvg,

        // ERAS-stage metrics & percentiles
        explorePct: student.explorePct,
        explorePctile: student.percentiles.explorePctile,
        exploreVsAvg: student.diffs.exploreVsAvg,

        researchPct: student.researchPct,
        researchPctile: student.percentiles.researchPctile,
        researchVsAvg: student.diffs.researchVsAvg,

        applyPct: student.applyPct,
        applyPctile: student.percentiles.applyPctile,
        applyVsAvg: student.diffs.applyVsAvg,

        strategizePct: student.strategizePct,
        strategizePctile: student.percentiles.strategizePctile,
        strategizeVsAvg: student.diffs.strategizeVsAvg,

        // stageData (for any progress bar partials)
        stageData: {
          pctExploreDone: student.pctExploreDone,
          nExploreLeft: student.nExploreLeft,
          pctResearchDone: student.pctResearchDone,
          nResearchLeft: student.nResearchLeft,
          pctApplyDone: student.pctApplyDone,
          nApplyLeft: student.nApplyLeft,
          pctStrategizeDone: student.pctStrategizeDone,
          nStrategizeLeft: student.nStrategizeLeft
        },

        hasDoneEXP4: hasDoneEXP4,
        hasDoneRES3: hasDoneRES3,
        hasDoneAPP7: hasDoneAPP7,

        // fears & purpose with presence flags (RESTORED)
        purpose: student.purpose,
        fear: student.fear,
        fearReason: student.fearReason,
        hasPurpose: Boolean(student.purpose && student.purpose.trim()),
        hasFear: Boolean(student.fear && student.fear.trim() && student.fearReason.trim()),
        today,
        tz
      }
    );

    // -------------------------------------------------
    // MAIL OPTIONS & SEND
    // -------------------------------------------------
    const mailOpts = { to: student.email, subject: emailSubject };
    if (TEMPLATE_OVERRIDES[tplKey]?.fromName) mailOpts.fromName = TEMPLATE_OVERRIDES[tplKey].fromName;
    if (TEMPLATE_OVERRIDES[tplKey]?.replyTo) mailOpts.replyTo = TEMPLATE_OVERRIDES[tplKey].replyTo;

    if (_renderAndSendEmail(mailOpts, { name: TEMPLATE_MAP[tplKey], data: templateData })) {
      stats.sent++;

      partSheet.getRange(student.row, aIdx[sentFlagCol] + 1).setValue(true);
      logEntries.push([
        student.ptid, tplKey, '', rawSub, student.email,
        new Date(), Session.getEffectiveUser().getEmail(),
        mailOpts.fromName || SENDER_NAME,
        mailOpts.replyTo || REPLY_TO
      ]);
    } else {
      stats.failed++;
    }
  });

  // --- 10.9 Summary logging & write logs ---
  Logger.log('_sendWeeklyNudge: summary sent=%d skipped=%d failed=%d', stats.sent, stats.skipped, stats.failed);
  if (logEntries.length) {
    logSheet.getRange(logSheet.getLastRow() + 1, 1, logEntries.length, logEntries[0].length).setValues(logEntries);
  }
  if (showUi) ss.toast(`Sent:${stats.sent} Skipped:${stats.skipped} Failed:${stats.failed}`, 'Weekly Nudge', 5);
  Logger.log('_sendWeeklyNudge: end');
}


// ============================================================================
// SECTION 11: TASK GROUP BUILDER
// ============================================================================

/**
 * _buildTaskGroupsForStudent(student, allTasks, tIdx, tz, today, resourceMap)
 *   - Builds grouped due, overdue, upcoming tasks with labels
 */
function _buildTaskGroupsForStudent(student, allTasks, tIdx, tz, today, resourceMap) {
  Logger.log('_buildTaskGroupsForStudent: start (ptid=%s)', student.ptid);

  // Step 1: Filter incomplete tasks
  const myTasks = allTasks.filter(r =>
    String(r[tIdx[COL_PTID]]).trim() === student.ptid && r[tIdx[COL_DONE]] !== true
  );

  // Step 2: Classify by deadline. All past-due tasks are now considered overdue.
  const dueTasks = [];
  const overdueTasks = [];
  const upcomingRaw = [];
  myTasks.forEach(r => {
    const dl = r[tIdx[COL_DEADLINE]] ? new Date(r[tIdx[COL_DEADLINE]]) : null;
    if (!dl) return;

    if (dl >= today) {
      dueTasks.push(r);
    } else if (dl < today) { // Condition updated to include all past-due tasks
      overdueTasks.push(r);
    }
  });


  // Step 3: Take top 3 each
  const dueSlice = dueTasks.sort((a, b) => new Date(a[tIdx[COL_DEADLINE]]) - new Date(b[tIdx[COL_DEADLINE]])).slice(0, 3);
  const overdueSlice = overdueTasks.sort((a, b) => new Date(b[tIdx[COL_DEADLINE]]) - new Date(a[tIdx[COL_DEADLINE]])).slice(0, 3);
  const upcSlice = upcomingRaw.sort((a, b) => new Date(a[tIdx[COL_DEADLINE]]) - new Date(b[tIdx[COL_DEADLINE]])).slice(0, 3);

  // Step 4: Build group objects
  const TaskGroups = buildGroupObject(dueSlice, tIdx, tz, resourceMap, 'Due by ');
  const OverdueGroups = buildGroupObject(overdueSlice, tIdx, tz, resourceMap, 'Was Due ');
  let NextTasksByDate = {};
  if (!Object.keys(TaskGroups).length && !Object.keys(OverdueGroups).length) {
    NextTasksByDate = upcSlice.reduce((acc, r) => {
      const dl = new Date(r[tIdx[COL_DEADLINE]]);
      const lbl = 'Due by ' + Utilities.formatDate(dl, tz, 'EEEE, MMMM d');
      acc[lbl] = acc[lbl] || [];
      acc[lbl].push({
        task: String(r[tIdx[COL_TASK]]),
        resources: resourceMap[String(r[tIdx[COL_TASK_ID]]).toLowerCase()] || []
      });
      return acc;
    }, {});
  }

  Logger.log(
    '_buildTaskGroupsForStudent: groups due=%d overdue=%d upcoming=%d',
    Object.keys(TaskGroups).length,
    Object.keys(OverdueGroups).length,
    Object.keys(NextTasksByDate).length
  );

  return { TaskGroups, OverdueGroups, NextTasksByDate };
}



// ============================================================================
// SECTION 12: UTILITIES: Group Object & Resource Map
// ============================================================================

/**
 * buildGroupObject(tasks, tIdx, tz, resourceMap, prefix)
 *   - Converts task rows into labeled groups
 */
function buildGroupObject(tasks, tIdx, tz, resourceMap, prefix) {
  return tasks.reduce((acc, r) => {
    const dl = new Date(r[tIdx[COL_DEADLINE]]);
    const lbl = prefix + Utilities.formatDate(dl, tz, 'EEEE, MMMM d');
    acc[lbl] = acc[lbl] || [];
    acc[lbl].push({
      task: String(r[tIdx[COL_TASK]]),
      resources: resourceMap[String(r[tIdx[COL_TASK_ID]]).toLowerCase()] || [],
      deadline: dl
    });
    return acc;
  }, {});
}

/**
 * buildResourceMap(resRows, resIdx)
 *   - Groups resource entries by task ID
 */
function buildResourceMap(resRows, resIdx) {
  const map = {};
  resRows.forEach(rw => {
    const id = String(rw[resIdx[COL_TASK_ID]]).toLowerCase();
    if (!id) return;
    map[id] = map[id] || [];
    map[id].push({
      text: rw[resIdx[COL_BUTTON_TEXT]] || '',
      description: rw[resIdx[COL_DESCRIPTION]] || '',
      url: rw[resIdx[COL_RESOURCE_URL]] || ''
    });
  });
  Logger.log('buildResourceMap: keys=%d', Object.keys(map).length);
  return map;
}


// ============================================================================
// SECTION 13: SMART RESCHEDULING SUGGESTIONS
// ============================================================================

/**
 * generateRescheduleDates(overdueTasks, existingDeadlines, todayMidnight, tz)
 *   - Suggests new deadlines avoiding weekends & overload
 */
function generateRescheduleDates(overdueTasks, existingDeadlines, todayMidnight, tz) {
  const suggestions = [];
  const booked = existingDeadlines.map(d => new Date(d).setHours(0, 0, 0, 0));

  overdueTasks.forEach(taskItem => {
    let candidate = new Date(todayMidnight);
    candidate.setDate(candidate.getDate() + 1);

    while (true) {
      const day = candidate.getDay();
      const ts = candidate.getTime();

      // Skip weekends or already booked
      if (day === 0 || day === 6 || booked.includes(ts)) { candidate.setDate(candidate.getDate() + 1); continue; }

      // Week capacity check
      const weekStart = new Date(candidate);
      weekStart.setDate(candidate.getDate() - ((candidate.getDay() + 6) % 7)); weekStart.setHours(0, 0, 0, 0);
      const weekEnd = new Date(weekStart); weekEnd.setDate(weekStart.getDate() + 6);
      const tasksInWeek = booked.filter(dts => {
        const bd = new Date(dts);
        return bd >= weekStart && bd <= weekEnd;
      });
      if (tasksInWeek.length >= 2) { candidate.setDate(candidate.getDate() + 1); continue; }

      // Accept candidate
      booked.push(ts);
      suggestions.push({
        taskItem,
        originalDueDateString: Utilities.formatDate(new Date(taskItem.deadline), tz, 'EEEE, MMMM d'),
        newDueDate: Utilities.formatDate(candidate, tz, 'EEEE, MMMM d')
      });
      break;
    }
  });

  Logger.log('generateRescheduleDates: suggestions=%d', suggestions.length);
  return suggestions;
}



