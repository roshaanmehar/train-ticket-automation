// ============================================================
// TRAIN TICKET COLLECTOR - Google Apps Script (Depersonalised)
// ============================================================
// This script collects PDF e-receipts from a specified sender
// in Gmail and saves them to Google Drive.
//
// It optionally filters emails by two route keywords, and
// optionally filters attachments by a keyword in the filename.
// ============================================================

// ----- CONFIGURATION (change these if needed) -----

var CONFIG = {
  // Sender email address to filter on
  senderEmail: "sender@example.com",

  // Root folder name in Google Drive
  rootFolderName: "Email Attachments",

  // Gmail label applied to emails that have already been processed
  processedLabelName: "AttachmentCollector-Processed",

  // Keywords to match in the email (optional route/content filter)
  // If both are non-empty, BOTH must appear in subject or body.
  // Set either to "" to disable this filter.
  keywordA: "origin",
  keywordB: "destination",

  // Only save attachments that contain this word in the filename
  // Set to "" to save ALL PDF attachments from matching emails
  attachmentNameKeyword: "receipt",
};

// ============================================================
// MAIN FUNCTION - Run this manually or via a daily trigger
// ============================================================

function collectAttachments() {
  var label = getOrCreateLabel(CONFIG.processedLabelName);
  var rootFolder = getOrCreateFolder(null, CONFIG.rootFolderName);

  // Search for emails that haven't been processed yet
  var query =
    "from:" +
    CONFIG.senderEmail +
    " -label:" +
    CONFIG.processedLabelName +
    " has:attachment";

  var threads = GmailApp.search(query);

  if (threads.length === 0) {
    Logger.log("No new matching emails found.");
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

      // Optional keyword filter (both keywords must appear if enabled)
      if (isKeywordFilterEnabled()) {
        if (
          content.indexOf(CONFIG.keywordA.toLowerCase()) === -1 ||
          content.indexOf(CONFIG.keywordB.toLowerCase()) === -1
        ) {
          Logger.log(
            "Skipping email (keyword match failed): " + message.getSubject()
          );
          continue;
        }
      }

      // Get the email date for naming/foldering
      var emailDate = message.getDate();

      // Process attachments
      var attachments = message.getAttachments();

      for (var a = 0; a < attachments.length; a++) {
        var attachment = attachments[a];
        var attachmentName = (attachment.getName() || "").toLowerCase();

        // Only process PDF files
        if (attachment.getContentType() !== "application/pdf") {
          continue;
        }

        // Optional attachment name keyword filter
        if (
          CONFIG.attachmentNameKeyword &&
          attachmentName.indexOf(CONFIG.attachmentNameKeyword.toLowerCase()) === -1
        ) {
          Logger.log(
            'Skipping attachment (keyword not in name): ' + attachment.getName()
          );
          continue;
        }

        // Build the folder path: Root / YYYY / Month
        var yearFolder = getOrCreateFolder(
          rootFolder,
          emailDate.getFullYear().toString()
        );
        var monthName = getMonthName(emailDate.getMonth());
        var monthFolder = getOrCreateFolder(yearFolder, monthName);

        // Build the file name: Mon - Day - DD-MM-YYYY.pdf
        var fileName = buildFileName(emailDate);

        // Handle duplicates by adding (2), (3), etc.
        fileName = getUniqueFileName(monthFolder, fileName);

        // Save the file
        monthFolder.createFile(attachment.copyBlob().setName(fileName));
        savedCount++;
        Logger.log("Saved: " + fileName + " to " + monthName + "/");
      }
    }

    // Mark the thread as processed so we don't grab it again
    threads[t].addLabel(label);
  }

  Logger.log("Done! Saved " + savedCount + " PDF file(s) to Google Drive.");
}

// ============================================================
// HELPER FUNCTIONS
// ============================================================

/**
 * Returns true if both keywords are set (non-empty).
 */
function isKeywordFilterEnabled() {
  return (
    CONFIG.keywordA &&
    CONFIG.keywordB &&
    CONFIG.keywordA.toString().trim() !== "" &&
    CONFIG.keywordB.toString().trim() !== ""
  );
}

/**
 * Builds a file name like: Feb - Sat - 07-02-2026.pdf
 */
function buildFileName(date) {
  var shortMonth = getShortMonthName(date.getMonth());
  var shortDay = getShortDayName(date.getDay());
  var dayNum = padZero(date.getDate());
  var monthNum = padZero(date.getMonth() + 1);
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

/**
 * If a file with this name already exists in the folder,
 * appends (2), (3), etc. to make it unique.
 */
function getUniqueFileName(folder, fileName) {
  var baseName = fileName.replace(".pdf", "");
  var candidate = fileName;
  var counter = 2;

  while (folder.getFilesByName(candidate).hasNext()) {
    candidate = baseName + " (" + counter + ").pdf";
    counter++;
  }

  return candidate;
}

/**
 * Gets or creates a Gmail label.
 */
function getOrCreateLabel(labelName) {
  var label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
    Logger.log('Created Gmail label: "' + labelName + '"');
  }
  return label;
}

/**
 * Gets or creates a folder inside a parent folder (or root of Drive if parent is null).
 */
function getOrCreateFolder(parentFolder, folderName) {
  var folders;
  if (parentFolder) {
    folders = parentFolder.getFoldersByName(folderName);
  } else {
    folders = DriveApp.getFoldersByName(folderName);
  }

  if (folders.hasNext()) {
    return folders.next();
  }

  if (parentFolder) {
    Logger.log('Created folder: "' + folderName + '"');
    return parentFolder.createFolder(folderName);
  } else {
    Logger.log('Created root folder: "' + folderName + '"');
    return DriveApp.createFolder(folderName);
  }
}

function getMonthName(monthIndex) {
  var months = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
  ];
  return months[monthIndex];
}

function getShortMonthName(monthIndex) {
  var months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  return months[monthIndex];
}

function getShortDayName(dayIndex) {
  var days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  return days[dayIndex];
}

function padZero(num) {
  return num < 10 ? "0" + num : num.toString();
}

