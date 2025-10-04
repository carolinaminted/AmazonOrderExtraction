/***********************
 * CONFIG
 ***********************/
// IMPORTANT: The Gmail label containing your Amazon order confirmations.
const LABEL_NAME_PDF_AMAZON = 'Amazon Orders';
// IMPORTANT: The path in Google Drive where PDFs will be saved.
const DRIVE_FOLDER_PATH_AMAZON = 'Purchases/Amazon/Extracted PDFs';
const MAX_PER_RUN_PDF_AMAZON = 100; // Safety cap per run
const LOG_SHEET_AMAZON = 'Amazon PDFs'; // Helper sheet for dedupe

/***********************
 * ENTRY POINT
 ***********************/
function exportAmazonPurchasesToPDFs() {
  console.log('--- Starting Amazon PDF Export ---');
  console.log(`Searching for label: "${LABEL_NAME_PDF_AMAZON}"`);

  const label = GmailApp.getUserLabelByName(LABEL_NAME_PDF_AMAZON);
  if (!label) {
    const errorMsg = `Execution STOPPED: Label not found: "${LABEL_NAME_PDF_AMAZON}". Please check for typos or create the label in Gmail.`;
    console.error(errorMsg);
    throw new Error(errorMsg);
  }
  console.log('Label found successfully.');

  const folder = getOrCreateFolderByPath_(DRIVE_FOLDER_PATH_AMAZON);
  // CORRECTED LINE: Log the original path variable instead of calling a non-existent function.
  console.log(`Ensured Google Drive folder exists at path: "${DRIVE_FOLDER_PATH_AMAZON}"`);

  const processed = loadProcessedIds_(LOG_SHEET_AMAZON);
  console.log(`Loaded ${processed.size} previously processed message IDs from sheet: "${LOG_SHEET_AMAZON}".`);

  let exported = 0;
  let start = 0,
    pageSize = 50;
  let totalMessagesScanned = 0;

  while (exported < MAX_PER_RUN_PDF_AMAZON) {
    console.log(`Fetching up to ${pageSize} threads starting from index ${start}...`);
    const threads = label.getThreads(start, pageSize);

    if (!threads.length) {
      console.log('No more threads found in this label. Ending search.');
      break;
    }
    console.log(`Found ${threads.length} threads in this batch.`);

    for (const thread of threads) {
      console.log(`Processing thread: ${thread.getId()} with ${thread.getMessageCount()} message(s).`);
      for (const msg of thread.getMessages()) {
        if (exported >= MAX_PER_RUN_PDF_AMAZON) {
          console.log(`Hit processing limit of ${MAX_PER_RUN_PDF_AMAZON}. Stopping for this run.`);
          break;
        }

        const id = msg.getId();
        const subject = msg.getSubject();
        const from = msg.getFrom();
        totalMessagesScanned++;

        console.log(`- Checking message ID: ${id} | From: "${from}" | Subject: "${subject}"`);

        if (processed.has(id)) {
          console.log(`  -> SKIPPING: Message ID is already logged as processed.`);
          continue;
        }

        const fromLower = from.toLowerCase();
        if (!fromLower.includes('amazon.com')) {
          console.log(`  -> SKIPPING: Sender "${from}" does not contain 'amazon.com'.`);
          continue;
        }

        const subjectLower = subject.toLowerCase();
        if (!subjectLower.includes('ordered')) {
          console.log(`  -> SKIPPING: Subject does not contain 'ordered'.`);
          continue;
        }

        console.log('  -> QUALIFIED: Message passes all filters. Attempting to create PDF.');

        try {
          const pdfBlob = renderMessageToPDF_(msg);
          const filename = buildAmazonFileName_(msg);
          console.log(`  -> Generated filename: "${filename}"`);

          const file = folder.createFile(pdfBlob).setName(filename);
          console.log(`  -> SUCCESS: Created PDF "${file.getName()}" with ID ${file.getId()}`);

          processed.add(id);
          exported++;
        } catch (e) {
          console.error(`  -> ERROR on message ${id}: ${e && e.message ? e.message : e}`);
        }
      }
      if (exported >= MAX_PER_RUN_PDF_AMAZON) break;
    }

    if (threads.length < pageSize) {
      console.log('Last batch of threads was smaller than page size. Assuming all threads have been processed.');
      break;
    }
    start += pageSize;
  }

  console.log('--- Run Summary ---');
  console.log(`Total messages scanned: ${totalMessagesScanned}`);
  console.log(`Total new PDFs exported this run: ${exported}`);
  saveProcessedIds_(processed, LOG_SHEET_AMAZON);
  SpreadsheetApp.getActive().toast(`Exported ${exported} Amazon PDFs to: ${DRIVE_FOLDER_PATH_AMAZON}`);
  console.log('--- Script Finished ---');
}

/***********************
 * Smarter filename with Order Number
 ***********************/
function buildAmazonFileName_(msg) {
  const date = Utilities.formatDate(msg.getDate(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const subj = msg.getSubject() || 'No Subject';
  const body = msg.getPlainBody() || '';

  // Attempt to find the Order Number to make the filename more useful
  const orderNumRegex = /Order #\s*(\d{3}-\d{7}-\d{7})/i;
  const match = body.match(orderNumRegex);
  const orderNumber = match ? match[1] : null;

  const cleanedSubject = subj
    .replace(/[\\/:*?"<>|#]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 120);

  return orderNumber ?
    `${date} - Amazon Order ${orderNumber}.pdf` :
    `${date} - ${cleanedSubject}.pdf`;
}


/***********************
 * === GENERIC HELPER FUNCTIONS (No changes needed below) ===
 ***********************/

function renderMessageToPDF_(msg) {
  const meta = {
    from: msg.getFrom(),
    to: msg.getTo(),
    cc: msg.getCc(),
    date: Utilities.formatDate(msg.getDate(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    subject: msg.getSubject(),
    id: msg.getId()
  };
  let html = msg.getBody() || '';
  html = inlineCidImages_(html, msg);
  html = inlineExternalImages_(html);

  const wrapper = `
    <html><head><meta charset="UTF-8" /><style>@page{size:A4;margin:18mm;}body{font-family:Arial,sans-serif;font-size:12px;color:#222;}.meta{border-bottom:1px solid #ddd;margin-bottom:12px;padding-bottom:8px;}.meta div{margin:2px 0;}.subject{font-size:16px;font-weight:700;margin-bottom:6px;}img{max-width:100%;height:auto;}a{color:#1155cc;text-decoration:none;}table{border-collapse:collapse;}td,th{border:1px solid #e5e5e5;padding:4px 6px;vertical-align:top;}.email-body,p,table,div{page-break-inside:avoid;}</style></head>
    <body><div class="meta"><div class="subject">${escape_(meta.subject)}</div><div><b>From:</b> ${escape_(meta.from)}</div><div><b>To:</b> ${escape_(meta.to||'')}</div>${meta.cc?`<div><b>CC:</b> ${escape_(meta.cc)}</div>`:''}<div><b>Date:</b> ${escape_(meta.date)}</div><div><b>Message ID:</b> ${escape_(meta.id)}</div></div><div class="email-body">${html}</div></body></html>`;
  const htmlBlob = Utilities.newBlob(wrapper, 'text/html', 'email.html');
  return htmlBlob.getAs('application/pdf');
}

function getOrCreateFolderByPath_(path) {
  if (!path) throw new Error('DRIVE_FOLDER_PATH is empty');
  const parts = path.split('/').map(p => p.trim()).filter(Boolean);
  let folder = DriveApp.getRootFolder();
  for (const name of parts) {
    let next = null;
    const it = folder.getFoldersByName(name);
    next = it.hasNext() ? it.next() : folder.createFolder(name);
    folder = next;
  }
  return folder;
}

function inlineCidImages_(html, msg) {
  const atts = msg.getAttachments({
    includeInlineImages: true,
    includeAttachments: false
  }) || [];
  if (!atts.length) return html;
  const cidMap = {};
  for (const a of atts) {
    const cid = (a.getContentId() || '').replace(/[<>]/g, '').trim();
    if (!cid) continue;
    const contentType = a.getContentType() || 'application/octet-stream';
    const base64 = Utilities.base64Encode(a.getBytes());
    cidMap[cid.toLowerCase()] = `data:${contentType};base64,${base64}`;
  }
  html = html.replace(/src\s*=\s*(['"])cid:([^'"]+)\1/gi, (m, q, cid) => {
    const key = (cid || '').replace(/[<>]/g, '').trim().toLowerCase();
    const dataUri = cidMap[key];
    return dataUri ? `src=${q}${dataUri}${q}` : m;
  });
  return html;
}

function inlineExternalImages_(html) {
  if (!html) return html;
  html = html.replace(/\s(data-src|data-original)\s*=\s*(['"])(.*?)\2/gi, (m, attr, q, val) => ` src=${q}${val}${q}`);
  html = html.replace(/\ssrcset\s*=\s*(['"])[\s\S]*?\1/gi, '');
  html = html.replace(/src\s*=\s*(['"])(https?:\/\/[^'"]+)\1/gi, (m, q, rawUrl) => {
    try {
      const url = normalizeGoogleProxyUrl_(rawUrl);
      if (url.length > 2000) return m;
      const resp = UrlFetchApp.fetch(url, {
        followRedirects: true,
        muteHttpExceptions: true,
        headers: {
          'User-Agent': 'Mozilla/5.0 (AppsScript PDF embedder)'
        }
      });
      if (resp.getResponseCode() !== 200) return m;
      let ctype = resp.getHeaders()['Content-Type'] || '';
      if (!ctype) {
        if (/\.(png)(\?|$)/i.test(url)) ctype = 'image/png';
        else if (/\.(jpe?g)(\?|$)/i.test(url)) ctype = 'image/jpeg';
        else if (/\.(gif)(\?|$)/i.test(url)) ctype = 'image/gif';
        else if (/\.(webp)(\?|$)/i.test(url)) ctype = 'image/webp';
        else ctype = 'application/octet-stream';
      }
      const bytes = resp.getContent();
      if (bytes.length > 5 * 1024 * 1024) return m;
      const base64 = Utilities.base64Encode(bytes);
      const dataUri = `data:${ctype};base64,${base64}`;
      return `src=${q}${dataUri}${q}`;
    } catch (e) {
      return m;
    }
  });
  return html;
}

function normalizeGoogleProxyUrl_(u) {
  try {
    if (/googleusercontent\.com\/proxy\//i.test(u)) {
      const hash = u.indexOf('#');
      if (hash > -1) return u.substring(hash + 1);
    }
    return u;
  } catch (_) {
    return u;
  }
}

function loadProcessedIds_(logSheetName) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(logSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(logSheetName);
    sheet.hideSheet();
    sheet.getRange(1, 1).setValue('messageId');
  }
  const vals = sheet.getRange(2, 1, Math.max(0, sheet.getLastRow() - 1), 1).getValues();
  const set = new Set();
  vals.forEach(r => {
    if (r[0]) set.add(String(r[0]));
  });
  return set;
}

function saveProcessedIds_(set, logSheetName) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(logSheetName);
  if (!sheet) return; // Should not happen due to loadProcessedIds_
  
  console.log(`Saving ${set.size} total processed IDs to sheet "${logSheetName}".`);
  const ids = Array.from(set);
  sheet.clearContents();
  sheet.getRange(1, 1).setValue('messageId');
  if (ids.length) {
    sheet.getRange(2, 1, ids.length, 1).setValues(ids.map(id => [id]));
  }
}

function escape_(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
