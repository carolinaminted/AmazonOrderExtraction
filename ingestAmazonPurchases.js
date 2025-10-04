/**
 * === CONFIG ===
 * Adjust these to match your Gmail label and Google Sheet tab name.
 */
// IMPORTANT: The Gmail label containing your Amazon order confirmations.
const LABEL_NAME_AMAZON = 'Amazon Orders';
// IMPORTANT: The Sheet tab where purchase data will be saved.
const SHEET_NAME_AMAZON = 'Amazon Orders';
const MAX_PER_RUN_AMAZON = 250; // Safety cap for emails to process per run.

/**
 * Main function to parse labeled Amazon purchase emails and append one row per order.
 * De-duplicates by Gmail message ID to prevent processing the same order twice.
 */
function ingestAmazonPurchases() {
  // LOGGING: Announce the start of the script execution.
  console.log('--- Script execution started: ingestAmazonPurchases ---');

  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_AMAZON);
  if (!sheet) {
    // LOGGING: Log critical error if the sheet is not found.
    console.error(`CRITICAL: Sheet named "${SHEET_NAME_AMAZON}" could not be found. Halting execution.`);
    throw new Error(`Missing sheet named "${SHEET_NAME_AMAZON}"`);
  }
  // LOGGING: Confirm the sheet was successfully accessed.
  console.log(`Successfully accessed sheet: "${SHEET_NAME_AMAZON}"`);

  const existingIds = loadExistingAmazonMessageIds_(sheet);
  const label = GmailApp.getUserLabelByName(LABEL_NAME_AMAZON);
  if (!label) {
    // LOGGING: Log critical error if the Gmail label is not found.
    console.error(`CRITICAL: Gmail label "${LABEL_NAME_AMAZON}" not found. Halting execution.`);
    throw new Error(`Gmail label "${LABEL_NAME_AMAZON}" not found`);
  }
  // LOGGING: Confirm the label was successfully accessed.
  console.log(`Successfully accessed Gmail label: "${LABEL_NAME_AMAZON}"`);

  let appended = 0;
  const threads = label.getThreads(0, Math.ceil(MAX_PER_RUN_AMAZON / 2));
  // LOGGING: Report how many email threads were found in the label.
  console.log(`Found ${threads.length} email threads to process.`);

  for (const thread of threads) {
    for (const msg of thread.getMessages()) {
      if (appended >= MAX_PER_RUN_AMAZON) {
        // LOGGING: Announce if the processing limit has been reached.
        console.log('Processing limit reached. Breaking loop.');
        break;
      }

      const msgId = msg.getId();
      const fromAddress = msg.getFrom();
      const subject = msg.getSubject();
      // LOGGING: Log details of the current email being checked.
      console.log(`\n--- Checking message | ID: ${msgId} | From: ${fromAddress} | Subject: "${subject}" ---`);

      if (existingIds.has(msgId)) {
        // LOGGING: State the reason for skipping this email.
        console.log(`  > RESULT: SKIPPED. Reason: Message ID already exists in the sheet.`);
        continue;
      }

      // LOGGING: Show the values being compared in the 'from' address check.
      const fromLower = fromAddress.toLowerCase();
      const fromCondition = 'auto-confirm@amazon.com';
      console.log(`  > CHECK 1 (From): Comparing "${fromLower}" to contains "${fromCondition}"`);
      if (!fromLower.includes(fromCondition)) {
        // LOGGING: State the reason for skipping this email.
        console.log(`  > RESULT: SKIPPED. Reason: 'From' address does not match condition.`);
        continue;
      }

      // LOGGING: Show the values being compared in the 'subject' check.
      const subjectLower = subject.toLowerCase();
      const subjectCondition = 'ordered';
      console.log(`  > CHECK 2 (Subject): Comparing "${subjectLower}" to contains "${subjectCondition}"`);
      if (!subjectLower.includes(subjectCondition)) {
        // LOGGING: State the reason for skipping this email.
        console.log(`  > RESULT: SKIPPED. Reason: 'Subject' does not match condition.`);
        continue;
      }
      
      // LOGGING: Announce that the email passed all checks and parsing will begin.
      console.log('  > All checks passed. Attempting to parse email content...');
      const parsed = parseAmazonEmail_(msg);

      if (!parsed) {
        // LOGGING: State that parsing failed for this email.
        console.log(`  > RESULT: SKIPPED. Reason: Parsing function returned no data.`);
        continue;
      }
      
      // LOGGING: Show the successfully parsed data.
      console.log(`  > PARSED DATA: ${JSON.stringify(parsed)}`);

      // LOGGING: Announce that the data is being appended to the sheet.
      console.log('  > Appending new row to the sheet...');
      sheet.appendRow([
        parsed.orderDate,
        parsed.orderNumber,
        parsed.itemTitle,
        parsed.orderTotal,
        msgId, // Add message ID for de-duplication
      ]);

      existingIds.add(msgId);
      appended++;
      // LOGGING: Confirm that the row was added.
      console.log('  > RESULT: SUCCESS. Row appended.');
    }
    if (appended >= MAX_PER_RUN_AMAZON) break;
  }
  
  // LOGGING: Final summary to be logged before the toast message appears.
  const summary = `Processed and appended ${appended} new Amazon purchases.`;
  console.log(`--- Script execution finished. Summary: ${summary} ---`);
  SpreadsheetApp.getActive().toast(summary);
}

/**
 * Extracts purchase details from the body of an Amazon order confirmation email.
 */
function parseAmazonEmail_(msg) {
  // LOGGING: Announce the start of the parsing function.
  console.log('    [parseAmazonEmail_]: Starting...');
  const bodyText = msg.getPlainBody() || '';
  const subject = msg.getSubject() || '';

  try {
    // LOGGING: Show the length of the email body being parsed.
    console.log(`    [parseAmazonEmail_]: Parsing body of length ${bodyText.length}`);
    const orderNumRegex = /Order #\s*(\d{3}-\d{7}-\d{7})/i;
    const orderNumMatch = bodyText.match(orderNumRegex);

    // LOGGING: Show whether the regular expression for Order # found a match or not.
    console.log(`    [parseAmazonEmail_]: Regex Match for Order #: ${orderNumMatch ? 'FOUND' : 'NOT FOUND'}`);
    const orderNumber = orderNumMatch ? orderNumMatch[1] : 'Not Found';

    // === NEW LOGIC FOR FINDING ORDER TOTAL ===
    let orderTotal = null;
    // Step 1: Find the index of a line starting with "Total". 'i' for case-insensitive, 'm' for multiline.
    const totalIndex = bodyText.search(/^Total/im);
    console.log(`    [parseAmazonEmail_]: Searching for line starting with "Total"... Status: ${totalIndex > -1 ? 'FOUND' : 'NOT FOUND'}`);

    if (totalIndex > -1) {
      // Step 2: If "Total" is found, search for the *next* dollar amount from that point forward.
      const substringAfterTotal = bodyText.substring(totalIndex);
      // UPDATED: Made the dollar sign optional (\$) with the question mark (?)
      const priceRegex = /\$?([0-9,.]+)/; 
      const priceMatch = substringAfterTotal.match(priceRegex);
      console.log(`    [parseAmazonEmail_]: Searching for price after "Total"... Status: ${priceMatch ? 'FOUND' : 'NOT FOUND'}`);
      
      if (priceMatch && priceMatch[1]) {
        orderTotal = Number(priceMatch[1].replace(/,/g, ''));
      }
    }
    // === END OF NEW LOGIC ===

    const orderDate = Utilities.formatDate(msg.getDate(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    let itemTitle = subject;
    if (subject.includes('Your Amazon.com order of "')) {
      itemTitle = subject.split('Your Amazon.com order of "')[1].replace(/"\./, '');
    } else if (subject.includes('Your Amazon.com order for "')) {
      itemTitle = subject.split('Your Amazon.com order for "')[1].replace(/"\./, '');
    }

    const result = {
      orderDate,
      orderNumber,
      itemTitle: itemTitle.trim(),
      orderTotal: orderTotal !== null ? round2_(orderTotal) : null,
    };
    // LOGGING: Show the final object being returned by the function.
    console.log(`    [parseAmazonEmail_]: Returning parsed object: ${JSON.stringify(result)}`);
    return result;
  } catch (e) {
    // LOGGING: Log any errors that occur during parsing.
    console.error(`    [parseAmazonEmail_]: ERROR during parsing for message ${msg.getId()}: ${e.message}`);
    return null;
  }
}

/**
 * Loads existing message IDs from the sheet to prevent duplicates.
 */
function loadExistingAmazonMessageIds_(sheet) {
  // LOGGING: Announce the start of this helper function.
  console.log('  [loadExistingAmazonMessageIds_]: Loading existing message IDs from sheet...');
  const idCol = 5;
  const lastRow = sheet.getLastRow();
  // LOGGING: Report how many rows are being checked for existing IDs.
  console.log(`  [loadExistingAmazonMessageIds_]: Checking up to row ${lastRow}.`);
  const set = new Set();
  if (lastRow >= 2) {
    sheet.getRange(2, idCol, lastRow - 1, 1).getValues().forEach(r => {
      const v = (r[0] || '').toString().trim();
      if (v) set.add(v);
    });
  }
  // LOGGING: Report the final count of unique IDs found.
  console.log(`  [loadExistingAmazonMessageIds_]: Found ${set.size} unique existing IDs.`);
  return set;
}

/**
 * Rounds a number to 2 decimal places.
 */
function round2_(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}
