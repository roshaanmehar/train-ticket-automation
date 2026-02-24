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
// Repair/Backfill mode:
//   - Does NOT create any "Reconciled" folder.
//   - Saves into your existing normal structure:
//       Train Tickets / YYYY / MonthName
//   - Can target:
//       A) previous calendar month, or
//       B) from the 1st of the current month
// ============================================================

// ----- CONFIGURATION -----
var CONFIG = {
  senderEmail: "tp-accounts-noreply@trainpal.com",
  rootFolderName: "Train Tickets",
  processedLabelName: "TrainTicket-Processed",

  // Route filter: both must appear (subject/body)
  routeKeywordA: "leeds",
  routeKeywordB: "hull",

  // Only save PDF attachments containing this word in filename.
  // Set to "" to save all PDF attachments from matching emails.
  receiptKeyword: "receipt",

  // -------------------------------
  // BACKFILL / REPAIR CONTROLS
  // -------------------------------
  // Choose ONE:
  //   "PREVIOUS_MONTH"  -> only tickets whose parsed travel date falls in previous month
  //   "CURRENT_MONTH"   -> only tickets whose parsed travel date falls in current month
  //
  // And: how far back Gmail should be searched (performance guard).
  backfillMode: "PREVIOUS_MONTH", // <-- change to "CURRENT_MONTH" if wanted
  backfillSearchDaysBack: 120,     // <-- adjust if you book further ahead
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

  var threads = GmailApp.search(query);

  if (threads.length === 0) {
    Logger.log("No new TrainPal emails found.");
    return;
  }

  Logger.log("Found " + threads.length + " new email thread(s) to process.");

  var savedCount = 0;

  for (var t = 0; t < threads.length; t++) {
    var messages = threads[t].getMessages();

    for (var m = 0; m < messages.length; m++) {
      var message = messages[m];
      var subject = (message.getSubject() || "").toLowerCase();
      var body = (message.getPlainBody() || "").toLowerCase();
      var content = subject + " " + body;

      // Route filter
      if (
        content.indexOf(CONFIG.routeKeywordA) === -1 ||
        content.indexOf(CONFIG.routeKeywordB) === -1
      ) {
        Logger.log("Skipping email (route filter): " + message.getSubject());
        continue;
      }

      var messageDate = message.getDate();
      var attachments = message.getAttachments();

      for (var a = 0; a < attachments.length; a++) {
        var attachment = attachments[a];
        var attachmentNameLower = (attachment.getName() || "").toLowerCase();

        if (attachment.getContentType() !== "application/pdf") continue;

        if (
          CONFIG.receiptKeyword &&
          attachmentNameLower.indexOf(CONFIG.receiptKeyword.toLowerCase()) === -1
        ) {
          Logger.log(
            'Skipping attachment (no "' +
              CONFIG.receiptKeyword +
              '" in name): ' +
              attachment.getName()
          );
          continue;
        }

        // Travel date resolution:
        var travelDate =
          extractTravelDateFromAttachmentName_(attachment.getName(), messageDate) ||
          extractTravelDateFromEmail_(message) ||
          messageDate;

        // Normal folder path: Train Tickets / YYYY / MonthName
        var yearFolder = getOrCreateFolderSmart_(
          rootFolder,
          travelDate.getFullYear().toString()
        );
        var monthName = getMonthName_(travelDate.getMonth());
        var monthFolder = getOrCreateFolderSmart_(yearFolder, monthName);

        // Normal file name
        var fileName = buildFileName_(travelDate);
        fileName = getUniqueFileName_(monthFolder, fileName);

        monthFolder.createFile(attachment.copyBlob().setName(fileName));
        savedCount++;
        Logger.log("Saved: " + fileName + " to " + monthName + "/");
      }
    }

    // Mark processed
    threads[t].addLabel(label);
  }

  Logger.log("Done! Saved " + savedCount + " receipt(s) to Google Drive.");
}

// ============================================================
// BACKFILL / REPAIR FUNCTION (NON-DISRUPTIVE)
// ============================================================
// - Does NOT touch your existing files.
// - Does NOT create any special folders.
// - Re-saves receipts (copies) into the normal YYYY/Month folders
//   based on the parsed travel date.
// - Uses a month filter based on CONFIG.backfillMode.
// ============================================================
function backfillMonth() {
  var rootFolder = getOrCreateFolderSmart_(null, CONFIG.rootFolderName);

  CONFIG.backfillMode = "CURRENT_MONTH";


  var targetInfo = getTargetMonthWindow_(CONFIG.backfillMode);
  var targetYear = targetInfo.year;
  var targetMonth = targetInfo.month; // 0-11
  var targetMonthName = getMonthName_(targetMonth);

  var query =
    "from:" +
    CONFIG.senderEmail +
    " newer_than:" +
    CONFIG.backfillSearchDaysBack +
    "d has:attachment";

  var threads = GmailApp.search(query);

  if (threads.length === 0) {
    Logger.log(
      "No TrainPal emails found in the last " +
        CONFIG.backfillSearchDaysBack +
        " days."
    );
    return;
  }

  Logger.log(
    "Backfill mode: " +
      CONFIG.backfillMode +
      " -> targeting " +
      targetMonthName +
      " " +
      targetYear +
      ". Scanning " +
      threads.length +
      " thread(s)..."
  );

  // Ensure target folders exist (normal structure)
  var yearFolder = getOrCreateFolderSmart_(rootFolder, targetYear.toString());
  var monthFolder = getOrCreateFolderSmart_(yearFolder, targetMonthName);

  var savedCount = 0;

  for (var t = 0; t < threads.length; t++) {
    var messages = threads[t].getMessages();

    for (var m = 0; m < messages.length; m++) {
      var message = messages[m];
      var subject = (message.getSubject() || "").toLowerCase();
      var body = (message.getPlainBody() || "").toLowerCase();
      var content = subject + " " + body;

      // Route filter still applies
      if (
        content.indexOf(CONFIG.routeKeywordA) === -1 ||
        content.indexOf(CONFIG.routeKeywordB) === -1
      ) {
        continue;
      }

      var messageDate = message.getDate();
      var attachments = message.getAttachments();

      for (var a = 0; a < attachments.length; a++) {
        var attachment = attachments[a];
        var attachmentNameLower = (attachment.getName() || "").toLowerCase();

        if (attachment.getContentType() !== "application/pdf") continue;

        if (
          CONFIG.receiptKeyword &&
          attachmentNameLower.indexOf(CONFIG.receiptKeyword.toLowerCase()) === -1
        ) {
          continue;
        }

        var travelDate =
          extractTravelDateFromAttachmentName_(attachment.getName(), messageDate) ||
          extractTravelDateFromEmail_(message) ||
          null;

        if (!travelDate) continue;

        // Month filter (strict)
        if (
          travelDate.getFullYear() !== targetYear ||
          travelDate.getMonth() !== targetMonth
        ) {
          continue;
        }

        // Save into the normal month folder (no special folder)
        var fileName = buildFileName_(travelDate);
        fileName = getUniqueFileName_(monthFolder, fileName);

        monthFolder.createFile(attachment.copyBlob().setName(fileName));
        savedCount++;
      }
    }
  }

  Logger.log(
    "Backfill complete. Saved " +
      savedCount +
      " receipt(s) into normal folder: " +
      CONFIG.rootFolderName +
      "/" +
      targetYear +
      "/" +
      targetMonthName
  );
}

// Returns the (year, month) pair for the target window.
function getTargetMonthWindow_(mode) {
  var now = new Date();
  var y = now.getFullYear();
  var m = now.getMonth(); // 0-11

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
// HELPERS
// ============================================================
function buildFileName_(date) {
  var shortMonth = getShortMonthName_(date.getMonth());
  var shortDay = getShortDayName_(date.getDay());
  var dayNum = padZero_(date.getDate());
  var monthNum = padZero_(date.getMonth() + 1);
  var year = date.getFullYear();
  return (
    shortMonth +
    " - " +
    shortDay +
    " - " +
    dayNum +
    "-" +
    monthNum +
    "-" +
    year +
    ".pdf"
  );
}

function getUniqueFileName_(folder, fileName) {
  var baseName = fileName.replace(/\.pdf$/i, "");
  var candidate = fileName;
  var counter = 2;

  while (folder.getFilesByName(candidate).hasNext()) {
    candidate = baseName + " (" + counter + ").pdf";
    counter++;
  }
  return candidate;
}

// PRIMARY date parsing: attachment filename
function extractTravelDateFromAttachmentName_(filename, messageDate) {
  if (!filename) return null;

  // Matches: "..._21_Feb_0655..." or "...-21-Feb-..."
  var m = filename.match(
    /\b(\d{1,2})[ _-](Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[ _-]/i
  );
  if (!m) return null;

  var day = parseInt(m[1], 10);
  var monStr =
    m[2].substr(0, 1).toUpperCase() + m[2].substr(1, 2).toLowerCase();
  var monthMap = {
    Jan: 0,
    Feb: 1,
    Mar: 2,
    Apr: 3,
    May: 4,
    Jun: 5,
    Jul: 6,
    Aug: 7,
    Sep: 8,
    Oct: 9,
    Nov: 10,
    Dec: 11,
  };
  var month = monthMap[monStr];
  if (month === undefined || isNaN(day)) return null;

  // Year heuristic from message date (handles year boundary)
  var year = messageDate.getFullYear();
  var msgMonth = messageDate.getMonth();

  if (msgMonth === 11 && month === 0) year += 1; // Dec email, Jan travel
  if (msgMonth === 0 && month === 11) year -= 1; // Jan email, Dec travel

  return new Date(year, month, day, 12, 0, 0);
}

// FALLBACK date parsing: email subject/body
function extractTravelDateFromEmail_(message) {
  var text = ((message.getSubject() || "") + " " + (message.getPlainBody() || ""))
    .replace(/\s+/g, " ")
    .trim();

  // Matches: "Sat, Feb 21, 2026"
  var m = text.match(
    /\b(?:Sun|Mon|Tue|Wed|Thu|Fri|Sat),\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2}),\s+(\d{4})\b/
  );
  if (!m) return null;

  var monthMap = {
    Jan: 0,
    Feb: 1,
    Mar: 2,
    Apr: 3,
    May: 4,
    Jun: 5,
    Jul: 6,
    Aug: 7,
    Sep: 8,
    Oct: 9,
    Nov: 10,
    Dec: 11,
  };
  var month = monthMap[m[1]];
  var day = parseInt(m[2], 10);
  var year = parseInt(m[3], 10);

  if (month === undefined || isNaN(day) || isNaN(year)) return null;
  return new Date(year, month, day, 12, 0, 0);
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
// It WON'T save you from an actual typo like "Februaru" already existing;
// but it WILL stop "february" vs "February " creating a second folder.
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

  Logger.log(
    "Reset complete. Removed processed label from " +
      threads.length +
      " thread(s)."
  );
}