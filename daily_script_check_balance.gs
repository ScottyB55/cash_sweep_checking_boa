// SEE BALANCE LOG HERE
// https://docs.google.com/spreadsheets/d/1e4Cjd7HYVJiZ_eHjveifmQ58KP4y4dcOke4ffJzniMU/edit?usp=sharing

// GOOGLE DOCS LINK WITH DETAILS
// https://docs.google.com/document/d/1RNZ5RyXFRPUHC_h_jFe92O7Fsygu_myjk27FanIywwg/edit?usp=sharing

// Balance Thresholds
//  See Doc:
// *** https://docs.google.com/spreadsheets/d/1e4Cjd7HYVJiZ_eHjveifmQ58KP4y4dcOke4ffJzniMU/edit#gid=1885744358 ***
//  We can tighten these if we are fine-tuning our weekly flow
//  Typically adjust these after adjusting the weekly flow
var HIGH_BALANCE_NOTIF_THRESH = 3590; // At paycheck - will temporarily go up and come down after RobinHood investment
var LOW_BALANCE_NOTIF_THRESH = 800; // Using Hysteresis - can narrow the range when tuning weekly flow

// Other Settings
var DAYS_OLD_THRESHOLD = 2;  // Number of days after which an email is considered old
var RECIPIENT_EMAIL = "srburnett111@gmail.com"; // Set the recipient email address as a constant
var ERROR_SUBJECT = "Error in Balance Checking Script"; // Subject for error notification emails
var DAYS_TO_KEEP_BEFORE_DELETING = 7; // Number of days to keep the emails
var SPREADSHEET_LOG_ID = "1e4Cjd7HYVJiZ_eHjveifmQ58KP4y4dcOke4ffJzniMU";


function sendEmailIfBalanceOutOfRange() {
  try {
    var emails = fetchBalanceEmailsAndDeleteOldOnes();
    if (emails.length === 0) {
      Logger.log("No balance emails found.");
      GmailApp.sendEmail(RECIPIENT_EMAIL, ERROR_SUBJECT, "No balance emails found.");
      Logger.log("Sent an email due to no balance emails found.");
      return;
    }
    
    // Sort the emails by date in descending order
    emails.sort(function(a, b) {
      return b.date - a.date;
    });
    
    var mostRecentEmail = emails[0];
    if (isEmailTooOld(mostRecentEmail.date)) {
      GmailApp.sendEmail(RECIPIENT_EMAIL, ERROR_SUBJECT, "Most recent balance email is older than " + DAYS_OLD_THRESHOLD + " days.");
      Logger.log("Sent an email due to the daily balance email being too old.");
      return;
    }
    
    var balanceInfo = parseEmailForBalance(mostRecentEmail.body);
    if (balanceInfo.found) {
      assessBalanceAndAct(balanceInfo.balance, mostRecentEmail);
    } else {
      Logger.log("Balance not found");
      GmailApp.sendEmail(RECIPIENT_EMAIL, ERROR_SUBJECT, "Balance not found in the most recent email.");
      Logger.log("Sent an email because balance was not found in the email.");
    }
  } catch (error) {
    Logger.log("Error encountered: " + error);
    GmailApp.sendEmail(RECIPIENT_EMAIL, ERROR_SUBJECT, "An error occurred: " + error);
    Logger.log("Sent an email due to an error encountered during script execution.");
  }
}

function isEmailTooOld(emailDate) {
  var currentDate = new Date();
  var thresholdDate = new Date(currentDate.getTime() - DAYS_OLD_THRESHOLD * 24 * 60 * 60 * 1000);
  return emailDate < thresholdDate;
}

function fetchBalanceEmailsAndDeleteOldOnes() {
  var label = GmailApp.getUserLabelByName("Daily Balance");
  var threads = label.getThreads();
  var messages = [];
  var expiryDate = new Date(new Date().getTime() - DAYS_TO_KEEP_BEFORE_DELETING * 24 * 60 * 60 * 1000);

  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    if (thread.getLastMessageDate() < expiryDate) {
      thread.moveToTrash();
      Logger.log("Moved thread to trash: " + thread.getFirstMessageSubject() + " (Last Message Date: " + thread.getLastMessageDate() + ")");
    } else {
      var threadMessages = thread.getMessages();
      for (var j = 0; j < threadMessages.length; j++) {
        var message = threadMessages[j];
        if (message.getDate() >= expiryDate) {
          messages.push({
            date: message.getDate(),
            subject: message.getSubject(),
            body: message.getPlainBody()
          });
        }
      }
    }
  }
  return messages;
}

function parseEmailForBalance(emailBody) {
  var balanceMatch = emailBody.match(/Balance: \$([\d,]+\.\d{2})/);
  if (balanceMatch) {
    var balance = parseFloat(balanceMatch[1].replace(/,/g, ''));
    return { found: true, balance: balance };
  }
  return { found: false };
}

function logBalanceToSheet(balance) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_LOG_ID).getActiveSheet();
    var currentDate = new Date();
    sheet.appendRow([currentDate, balance]);
  } catch (error) {
    Logger.log("Error logging balance to sheet: " + error.message);
    GmailApp.sendEmail(RECIPIENT_EMAIL, ERROR_SUBJECT, "An error occurred while logging balance to google sheets: " + error);
  }
}

function assessBalanceAndAct(balance, email) {

  logBalanceToSheet(balance)

  var emailSubject;
  // var emailBody = "Date: " + email.date + "\nSubject: " + email.subject + "\nBody: " + email.body;
  var emailBody; // Now will be the same as emailSubject

  if (balance < LOW_BALANCE_NOTIF_THRESH) {
    emailSubject = "ACTION REQUIRED: CHECKING ACCOUNT BALANCE IS LOW: $" + balance;
    emailBody = emailSubject; // Make body the same as subject
    GmailApp.sendEmail(RECIPIENT_EMAIL, emailSubject, emailBody);
    Logger.log("Sent an email due to balance too low.");
  } else if (balance > HIGH_BALANCE_NOTIF_THRESH) {
    emailSubject = "ACTION REQUIRED: CHECKING ACCOUNT BALANCE IS HIGH: $" + balance;
    emailBody = emailSubject; // Make body the same as subject
    GmailApp.sendEmail(RECIPIENT_EMAIL, emailSubject, emailBody);
    Logger.log("Sent an email due to balance too high.");
  } else {
    Logger.log("Current Balance: $" + balance); // Only logs, no email for normal balance
  }
}
