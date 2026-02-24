// ============================================================
// TRAIN TICKET COLLECTOR - Google Apps Script
// ============================================================
// Collects TrainPal PDF receipts from Gmail and saves to Drive.
//
// Naming date logic:
//   1) PRIMARY: parse travel date from ATTACHMENT filename
//   2) FALLBACK: parse travel date from EMAIL body/subject
//   3) FINAL fallback: email sent date
//
// Backfill mode:
//   - Does NOT create any "Reconciled" folder.
//   - Saves into existing normal structure:
//       Train Tickets / YYYY / MonthName
//
// Logging:
//   - Emits per-thread + per-message summaries
//   - Emits per-attachment saved/skipped with reasons
//   - Emits final counters + a “saved map” of email -> files
// ============================================================

// ----- CONFIGURATION -----
var CONFIG = {
  senderEmail: "tp-accounts-noreply@trainpal.com",
  rootFolderName: "Train Tickets",
  processedLabelName: "TrainTicket-Processed",

  // Route filter: both must appear (subject/body)
  routeKeywordA: "leeds",
  routeKeywordB: "hull",

  // Daily collector: only save PDFs containing this word in filename.
  // Set to "" to save ALL PDFs from matching emails.
  receiptKeyword: "receipt",

  // -------------------------------
  // BACKFILL / REPAIR CONTROLS
  // -------------------------------
  // Choose ONE:
  //   "PREVIOUS_MONTH" -> only tickets whose parsed travel date falls in previous month
  //   "CURRENT_MONTH"  -> only tickets whose parsed travel date falls in current month
  backfillMode: "CURRENT_MONTH",

  // How far back Gmail should be searched (performance guard)
  backfillSearchDaysBack: 180,

  // Backfill attachment behaviour:
  //   "RECEIPTS_ONLY" -> honour receiptKeyword filter
  //   "ALL_PDFS"      -> ignore receiptKeyword and save all PDF attachments
backfillAttachmentMode: "RECEIPTS_ONLY",

  // Gmail search pagination
  pageSize: 100,

  // Set true to log the first ~300 chars of email body (useful when date parsing fails)
  logBodySnippet: false,
};

// ============================================================
// MAIN FUNCTION - Run manually or via daily trigger
// ============================================================
function collectTrainTickets() {
  var label = getOrCreateLabel_(CONFIG.processedLabelName);
  var rootFolder = getOrCreateFolderSmart_(null, CONFIG.rootFolderName);

  var query =
    "from:" +
    CONFIG.senderEmail +
    " -label:" +
    CONFIG.processedLabelName +
    " has:attachment";

  Logger.log("=== collectTrainTickets() starting ===");
  Logger.log("Gmail query: " + query);

  var start = 0;
  var totalThreads = 0;
  var totalMessages = 0;

  var counters = makeCounters_();
  var savedMap = {}; // messageId -> {subject, date, saved:[...], skipped:[...]}

  while (true) {
    var threads = GmailApp.search(query, start, CONFIG.pageSize);
    if (!threads || threads.length === 0) break;

    totalThreads += threads.length;
    Logger.log("Page: start=" + start + " threads=" + threads.length);

    for (var t = 0; t < threads.length; t++) {
      var thread = threads[t];
      var messages = thread.getMessages();
      counters.threadsScanned++;

      Logger.log("---- Thread #" + counters.threadsScanned + " messages=" + messages.length + " ----");

      for (var m = 0; m < messages.length; m++) {
        var message = messages[m];
        totalMessages++;
        counters.messagesScanned++;

        var subjectRaw = message.getSubject() || "";
        var subject = subjectRaw.toLowerCase();
        var bodyRaw = message.getPlainBody() || "";
        var body = bodyRaw.toLowerCase();
        var content = subject + " " + body;

        var msgDate = message.getDate();
        var msgId = safeMessageId_(message);

        ensureSavedMap_(savedMap, msgId, subjectRaw, msgDate);

        Logger.log(
          "[Message] id=" + msgId +
          " date=" + msgDate +
          " subject=\"" + subjectRaw + "\""
        );

        if (CONFIG.logBodySnippet) {
          Logger.log("[Body snippet] " + bodyRaw.substring(0, 300).replace(/\s+/g, " "));
        }

        // Route filter
        if (
          content.indexOf(CONFIG.routeKeywordA) === -1 ||
          content.indexOf(CONFIG.routeKeywordB) === -1
        ) {
          counters.skippedRoute++;
          savedMap[msgId].skipped.push("Skipped message (route filter)");
          Logger.log("  -> SKIP message: route keywords missing (" + CONFIG.routeKeywordA + ", " + CONFIG.routeKeywordB + ")");
          continue;
        }

        var attachments = message.getAttachments();
        Logger.log("  Attachments: " + attachments.length);

        for (var a = 0; a < attachments.length; a++) {
          var att = attachments[a];
          var attName = att.getName() || "(no name)";
          var attNameLower = attName.toLowerCase();
          var attType = att.getContentType();

          // Only PDF
          if (attType !== "application/pdf") {
            counters.skippedNonPdf++;
            savedMap[msgId].skipped.push(attName + " (non-PDF: " + attType + ")");
            Logger.log("  -> SKIP attachment: " + attName + " (non-PDF: " + attType + ")");
            continue;
          }

          // Receipt keyword filter (collector only)
          if (CONFIG.receiptKeyword && attNameLower.indexOf(CONFIG.receiptKeyword.toLowerCase()) === -1) {
            counters.skippedKeyword++;
            savedMap[msgId].skipped.push(attName + " (missing keyword: " + CONFIG.receiptKeyword + ")");
            Logger.log("  -> SKIP attachment: " + attName + " (missing keyword: " + CONFIG.receiptKeyword + ")");
            continue;
          }

          // Resolve travel date + record which method was used
          var travelInfo = resolveTravelDateWithSource_(attName, message, msgDate);
          var travelDate = travelInfo.date;
          var dateSource = travelInfo.source;

          // Folder path
          var yearFolder = getOrCreateFolderSmart_(rootFolder, travelDate.getFullYear().toString());
          var monthName = getMonthName_(travelDate.getMonth());
          var monthFolder = getOrCreateFolderSmart_(yearFolder, monthName);

          // File name (unique)
          var fileName = buildFileName_(travelDate);
          var uniqueNameInfo = getUniqueFileNameWithInfo_(monthFolder, fileName);

          // Save
          monthFolder.createFile(att.copyBlob().setName(uniqueNameInfo.name));
          counters.saved++;
          if (uniqueNameInfo.wasDuplicate) counters.duplicatesRenamed++;

          var saveLine =
            uniqueNameInfo.name +
            " -> " +
            CONFIG.rootFolderName + "/" + travelDate.getFullYear() + "/" + monthName +
            " (travelDateSource=" + dateSource + ", originalAttachment=\"" + attName + "\"" +
            (uniqueNameInfo.wasDuplicate ? ", dedupFrom=\"" + fileName + "\"" : "") +
            ")";

          savedMap[msgId].saved.push(saveLine);

          Logger.log("  -> SAVED: " + saveLine);
        }
      }

      // Mark the whole thread processed after scanning its messages
      thread.addLabel(label);
      counters.threadsLabelled++;
      Logger.log("---- Thread labelled: " + CONFIG.processedLabelName + " ----");
    }

    start += CONFIG.pageSize;
  }

  Logger.log("=== collectTrainTickets() finished ===");
  logCounters_(counters, totalThreads, totalMessages);
  logSavedMap_(savedMap);
}

// ============================================================
// BACKFILL FUNCTION (NON-DISRUPTIVE)
// ============================================================
// - Does NOT touch existing files.
// - Does NOT create special folders.
// - Re-saves PDFs into normal YYYY/Month folders based on parsed travel date.
// - Optional mode to save ALL PDFs during backfill (recommended).
// - Includes pagination + detailed logging.
// ============================================================
function backfillMonth() {
  var rootFolder = getOrCreateFolderSmart_(null, CONFIG.rootFolderName);
  backfillMode: "CURRENT_MONTH";
  // IMPORTANT: do NOT override CONFIG.backfillMode in code.
  // (Your earlier version forced CURRENT_MONTH, which stops February backfills later.)

  var targetInfo = getTargetMonthWindow_(CONFIG.backfillMode);
  var targetYear = targetInfo.year;
  var targetMonth = targetInfo.month;
  var targetMonthName = getMonthName_(targetMonth);

  var query =
    "from:" +
    CONFIG.senderEmail +
    " newer_than:" +
    CONFIG.backfillSearchDaysBack +
    "d has:attachment";

  Logger.log("=== backfillMonth() starting ===");
  Logger.log("Target month: " + targetMonthName + " " + targetYear + " (mode=" + CONFIG.backfillMode + ")");
  Logger.log("Attachment mode: " + CONFIG.backfillAttachmentMode);
  Logger.log("Gmail query: " + query);

  // Ensure target folders exist
  var yearFolder = getOrCreateFolderSmart_(rootFolder, targetYear.toString());
  var monthFolder = getOrCreateFolderSmart_(yearFolder, targetMonthName);

  var start = 0;
  var totalThreads = 0;
  var totalMessages = 0;

  var counters = makeCounters_();
  var savedMap = {}; // messageId -> {subject, date, saved:[...], skipped:[...]}

  while (true) {
    var threads = GmailApp.search(query, start, CONFIG.pageSize);
    if (!threads || threads.length === 0) break;

    totalThreads += threads.length;
    Logger.log("Page: start=" + start + " threads=" + threads.length);

    for (var t = 0; t < threads.length; t++) {
      var thread = threads[t];
      var messages = thread.getMessages();
      counters.threadsScanned++;

      Logger.log("---- Thread #" + counters.threadsScanned + " messages=" + messages.length + " ----");

      for (var m = 0; m < messages.length; m++) {
        var message = messages[m];
        totalMessages++;
        counters.messagesScanned++;

        var subjectRaw = message.getSubject() || "";
        var subject = subjectRaw.toLowerCase();
        var bodyRaw = message.getPlainBody() || "";
        var body = bodyRaw.toLowerCase();
        var content = subject + " " + body;

        var msgDate = message.getDate();
        var msgId = safeMessageId_(message);

        ensureSavedMap_(savedMap, msgId, subjectRaw, msgDate);

        Logger.log(
          "[Message] id=" + msgId +
          " date=" + msgDate +
          " subject=\"" + subjectRaw + "\""
        );

        if (CONFIG.logBodySnippet) {
          Logger.log("[Body snippet] " + bodyRaw.substring(0, 300).replace(/\s+/g, " "));
        }

        // Route filter
        if (
          content.indexOf(CONFIG.routeKeywordA) === -1 ||
          content.indexOf(CONFIG.routeKeywordB) === -1
        ) {
          counters.skippedRoute++;
          savedMap[msgId].skipped.push("Skipped message (route filter)");
          Logger.log("  -> SKIP message: route keywords missing (" + CONFIG.routeKeywordA + ", " + CONFIG.routeKeywordB + ")");
          continue;
        }

        var attachments = message.getAttachments();
        Logger.log("  Attachments: " + attachments.length);

        for (var a = 0; a < attachments.length; a++) {
          var att = attachments[a];
          var attName = att.getName() || "(no name)";
          var attNameLower = attName.toLowerCase();
          var attType = att.getContentType();

          if (attType !== "application/pdf") {
            counters.skippedNonPdf++;
            savedMap[msgId].skipped.push(attName + " (non-PDF: " + attType + ")");
            Logger.log("  -> SKIP attachment: " + attName + " (non-PDF: " + attType + ")");
            continue;
          }

          // Backfill attachment mode
          if (CONFIG.backfillAttachmentMode === "RECEIPTS_ONLY") {
            if (CONFIG.receiptKeyword && attNameLower.indexOf(CONFIG.receiptKeyword.toLowerCase()) === -1) {
              counters.skippedKeyword++;
              savedMap[msgId].skipped.push(attName + " (missing keyword: " + CONFIG.receiptKeyword + ")");
              Logger.log("  -> SKIP attachment: " + attName + " (missing keyword: " + CONFIG.receiptKeyword + ")");
              continue;
            }
          }

          // Parse travel date (no email-date fallback here; we want a deterministic month)
          var travelInfo = resolveTravelDateForBackfill_(attName, message, msgDate);
          if (!travelInfo.date) {
            counters.skippedNoDate++;
            savedMap[msgId].skipped.push(attName + " (could not parse travel date)");
            Logger.log("  -> SKIP attachment: " + attName + " (could not parse travel date)");
            continue;
          }

          var travelDate = travelInfo.date;
          var dateSource = travelInfo.source;

          // Month filter
          if (travelDate.getFullYear() !== targetYear || travelDate.getMonth() !== targetMonth) {
            counters.skippedMonthMismatch++;
            savedMap[msgId].skipped.push(attName + " (month mismatch: parsed " + (travelDate.getMonth()+1) + "/" + travelDate.getFullYear() + ")");
            Logger.log(
              "  -> SKIP attachment: " + attName +
              " (month mismatch: parsed " + (travelDate.getMonth()+1) + "/" + travelDate.getFullYear() + ")"
            );
            continue;
          }

          // Save into target month folder
          var fileName = buildFileName_(travelDate);
          var uniqueNameInfo = getUniqueFileNameWithInfo_(monthFolder, fileName);

          monthFolder.createFile(att.copyBlob().setName(uniqueNameInfo.name));
          counters.saved++;
          if (uniqueNameInfo.wasDuplicate) counters.duplicatesRenamed++;

          var saveLine =
            uniqueNameInfo.name +
            " -> " +
            CONFIG.rootFolderName + "/" + targetYear + "/" + targetMonthName +
            " (travelDateSource=" + dateSource + ", originalAttachment=\"" + attName + "\"" +
            (uniqueNameInfo.wasDuplicate ? ", dedupFrom=\"" + fileName + "\"" : "") +
            ")";

          savedMap[msgId].saved.push(saveLine);

          Logger.log("  -> SAVED: " + saveLine);
        }
      }
    }

    start += CONFIG.pageSize;
  }

  Logger.log("=== backfillMonth() finished ===");
  logCounters_(counters, totalThreads, totalMessages);
  logSavedMap_(savedMap);
}

// ============================================================
// DATE RESOLUTION (with source label for logging)
// ============================================================

function resolveTravelDateWithSource_(attachmentName, message, messageDate) {
  var d1 = extractTravelDateFromAttachmentName_(attachmentName, messageDate);
  if (d1) return { date: d1, source: "ATTACHMENT_NAME" };

  var d2 = extractTravelDateFromEmail_(message);
  if (d2) return { date: d2, source: "EMAIL_BODY" };

  return { date: messageDate, source: "EMAIL_SENT_DATE" };
}

// Backfill: prefer attachment then email; if neither parsed, return null.
// (We do NOT want to silently fall back to message sent date because it will mis-file across months.)
function resolveTravelDateForBackfill_(attachmentName, message, messageDate) {
  var d1 = extractTravelDateFromAttachmentName_(attachmentName, messageDate);
  if (d1) return { date: d1, source: "ATTACHMENT_NAME" };

  var d2 = extractTravelDateFromEmail_(message);
  if (d2) return { date: d2, source: "EMAIL_BODY" };

  return { date: null, source: "NONE" };
}

// ============================================================
// TARGET MONTH WINDOW
// ============================================================
function getTargetMonthWindow_(mode) {
  var now = new Date();
  var y = now.getFullYear();
  var m = now.getMonth();

  if (mode === "PREVIOUS_MONTH") {
    m = m - 1;
    if (m < 0) {
      m = 11;
      y = y - 1;
    }
    return { year: y, month: m };
  }

  // Default: CURRENT_MONTH
  return { year: y, month: m };
}

// ============================================================
// HELPERS (files/folders/naming)
// ============================================================
function buildFileName_(date) {
  var shortMonth = getShortMonthName_(date.getMonth());
  var shortDay = getShortDayName_(date.getDay());
  var dayNum = padZero_(date.getDate());
  var monthNum = padZero_(date.getMonth() + 1);
  var year = date.getFullYear();
  return shortMonth + " - " + shortDay + " - " + dayNum + "-" + monthNum + "-" + year + ".pdf";
}

function getUniqueFileNameWithInfo_(folder, fileName) {
  var baseName = fileName.replace(/\.pdf$/i, "");
  var candidate = fileName;
  var counter = 2;
  var dup = false;

  while (folder.getFilesByName(candidate).hasNext()) {
    dup = true;
    candidate = baseName + " (" + counter + ").pdf";
    counter++;
  }

  return { name: candidate, wasDuplicate: dup };
}

function getUniqueFileName_(folder, fileName) {
  return getUniqueFileNameWithInfo_(folder, fileName).name;
}

// PRIMARY date parsing: attachment filename
function extractTravelDateFromAttachmentName_(filename, messageDate) {
  if (!filename) return null;

  // Matches: "..._21_Feb_0655..." or "...-21-Feb-..." or "... 21 Feb ..."
  var m = filename.match(/\b(\d{1,2})[ _-](Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[ _-]/i);
  if (!m) return null;

  var day = parseInt(m[1], 10);
  var monStr = m[2].substr(0, 1).toUpperCase() + m[2].substr(1, 2).toLowerCase();
  var monthMap = { Jan:0, Feb:1, Mar:2, Apr:3, May:4, Jun:5, Jul:6, Aug:7, Sep:8, Oct:9, Nov:10, Dec:11 };
  var month = monthMap[monStr];
  if (month === undefined || isNaN(day)) return null;

  // Year heuristic from message date (handles year boundary)
  var year = messageDate.getFullYear();
  var msgMonth = messageDate.getMonth();
  if (msgMonth === 11 && month === 0) year += 1;
  if (msgMonth === 0 && month === 11) year -= 1;

  return new Date(year, month, day, 12, 0, 0);
}

// FALLBACK date parsing: email subject/body
function extractTravelDateFromEmail_(message) {
  var text = ((message.getSubject() || "") + " " + (message.getPlainBody() || ""))
    .replace(/\s+/g, " ")
    .trim();

  // Pattern 1: "Sat, Feb 21, 2026"
  var m1 = text.match(/\b(?:Sun|Mon|Tue|Wed|Thu|Fri|Sat),\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2}),\s+(\d{4})\b/);
  if (m1) return new Date(parseInt(m1[3], 10), monthIndex_(m1[1]), parseInt(m1[2], 10), 12, 0, 0);

  // Pattern 2 (extra): "21 Feb 2026"
  var m2 = text.match(/\b(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{4})\b/);
  if (m2) return new Date(parseInt(m2[3], 10), monthIndex_(m2[2]), parseInt(m2[1], 10), 12, 0, 0);

  return null;
}

function monthIndex_(monStr) {
  var m = monStr.substr(0, 1).toUpperCase() + monStr.substr(1, 2).toLowerCase();
  var map = { Jan:0, Feb:1, Mar:2, Apr:3, May:4, Jun:5, Jul:6, Aug:7, Sep:8, Oct:9, Nov:10, Dec:11 };
  return map[m];
}

// Gmail label helper
function getOrCreateLabel_(labelName) {
  var label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
    Logger.log('Created Gmail label: "' + labelName + '"');
  }
  return label;
}

// Folder getter that avoids duplicates due to casing/spacing differences.
function getOrCreateFolderSmart_(parentFolder, folderName) {
  var desired = normaliseFolderName_(folderName);

  var it = parentFolder ? parentFolder.getFolders() : DriveApp.getFolders();
  while (it.hasNext()) {
    var f = it.next();
    if (normaliseFolderName_(f.getName()) === desired) {
      return f;
    }
  }

  if (parentFolder) {
    Logger.log('Created folder: "' + folderName + '"');
    return parentFolder.createFolder(folderName);
  } else {
    Logger.log('Created root folder: "' + folderName + '"');
    return DriveApp.createFolder(folderName);
  }
}

function normaliseFolderName_(name) {
  return (name || "").toLowerCase().replace(/\s+/g, " ").trim();
}

// Date formatting helpers
function getMonthName_(monthIndex) {
  var months = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
  ];
  return months[monthIndex];
}
function getShortMonthName_(monthIndex) {
  var months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  return months[monthIndex];
}
function getShortDayName_(dayIndex) {
  var days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  return days[dayIndex];
}
function padZero_(num) {
  return num < 10 ? "0" + num : num.toString();
}

// ============================================================
// LOGGING HELPERS
// ============================================================
function makeCounters_() {
  return {
    threadsScanned: 0,
    threadsLabelled: 0,
    messagesScanned: 0,
    saved: 0,
    duplicatesRenamed: 0,
    skippedRoute: 0,
    skippedNonPdf: 0,
    skippedKeyword: 0,
    skippedNoDate: 0,
    skippedMonthMismatch: 0,
  };
}

function logCounters_(c, totalThreads, totalMessages) {
  Logger.log("=== Summary ===");
  Logger.log("Threads scanned: " + c.threadsScanned + " (approx total matched threads: " + totalThreads + ")");
  Logger.log("Threads labelled: " + c.threadsLabelled);
  Logger.log("Messages scanned: " + c.messagesScanned + " (approx total scanned messages: " + totalMessages + ")");
  Logger.log("Saved PDFs: " + c.saved);
  Logger.log("Duplicates renamed: " + c.duplicatesRenamed);
  Logger.log("Skipped messages (route): " + c.skippedRoute);
  Logger.log("Skipped attachments (non-PDF): " + c.skippedNonPdf);
  Logger.log("Skipped attachments (keyword): " + c.skippedKeyword);
  Logger.log("Skipped attachments (no travel date): " + c.skippedNoDate);
  Logger.log("Skipped attachments (month mismatch): " + c.skippedMonthMismatch);
}

function ensureSavedMap_(savedMap, msgId, subject, dateObj) {
  if (!savedMap[msgId]) {
    savedMap[msgId] = {
      subject: subject,
      date: dateObj,
      saved: [],
      skipped: [],
    };
  }
}

function logSavedMap_(savedMap) {
  Logger.log("=== Per-email saved/skipped map ===");
  var keys = Object.keys(savedMap);
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    var rec = savedMap[k];
    Logger.log("Email id=" + k + " date=" + rec.date + " subject=\"" + rec.subject + "\"");

    if (rec.saved.length > 0) {
      Logger.log("  SAVED (" + rec.saved.length + "):");
      for (var s = 0; s < rec.saved.length; s++) Logger.log("    - " + rec.saved[s]);
    } else {
      Logger.log("  SAVED: none");
    }

    if (rec.skipped.length > 0) {
      Logger.log("  SKIPPED (" + rec.skipped.length + "):");
      for (var x = 0; x < rec.skipped.length; x++) Logger.log("    - " + rec.skipped[x]);
    } else {
      Logger.log("  SKIPPED: none");
    }
  }
}

// Some Apps Script environments don't expose a stable message ID;
// this keeps logs readable without failing.
function safeMessageId_(message) {
  try {
    // message.getId() exists in GmailApp Message objects
    return message.getId();
  } catch (e) {
    return "unknown-id";
  }
}

// ============================================================
// SETUP: daily trigger
// ============================================================
function setupDailyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "collectTrainTickets") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger("collectTrainTickets")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log("Daily trigger created! Runs every day between 8-9 AM.");
}

// ============================================================
// OPTIONAL: reset processed label
// ============================================================
function resetProcessedEmails() {
  var label = GmailApp.getUserLabelByName(CONFIG.processedLabelName);
  if (!label) {
    Logger.log("No processed label found. Nothing to reset.");
    return;
  }

  var threads = label.getThreads();
  for (var t = 0; t < threads.length; t++) {
    threads[t].removeLabel(label);
  }

  Logger.log("Reset complete. Removed processed label from " + threads.length + " thread(s).");
}