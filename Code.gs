// =============================================================================
// SALT Transcript Processor — Google Apps Script
// =============================================================================
// Automatically routes Gemini meeting transcript segments to the correct
// client folders in Google Drive based on fuzzy name/address matching.
//
// Deploy: script.google.com → Extensions → Apps Script
// Trigger: processNewTranscripts() every 15 minutes (time-based)
// Setup:  Run setup() once to create the Gmail label and log sheet.
// =============================================================================

// ---------------------------------------------------------------------------
// CONFIGURATION — Edit these values to match your environment
// ---------------------------------------------------------------------------
var CONFIG = {
  // Google Drive folder ID for "Landscaping Projects" parent folder
  PARENT_FOLDER_ID: '1GS6Z6vPwqUmFK1LbYgaUFhB1BTRPdeyz',

  // Gmail account to search (used in the query)
  GMAIL_ACCOUNT: 'hello@saltlandscaping.com.au',

  // Sender of Gemini meeting notes
  GEMINI_SENDER: 'gemini-notes@google.com',

  // Meeting title matching (case-insensitive). The subject format from
  // Gemini is: "Notes: '{Meeting Title}' {Date}"
  // ALL of these keywords must appear in the title (fuzzy match).
  // e.g. "Warren + Billy - Quoting", "Billy + Warren - Quick Quote", etc.
  MEETING_TITLE_REQUIRED_WORDS: ['Warren', 'Billy'],
  MEETING_TITLE_REQUIRED_ANY: ['Quote', 'Quoting'],  // at least ONE of these must appear

  // Gmail label used to mark processed emails
  LABEL_NAME: 'Processed-Transcript',

  // Subfolder inside each client folder where docs are saved
  CLIENT_SUBFOLDER: '06 - Quoting Context',

  // Name prefix for saved Google Docs
  DOC_NAME_PREFIX: 'Warren-Billy Quoting Notes',

  // Name of the log spreadsheet (created in Drive root)
  LOG_SHEET_NAME: 'Transcript Processing Log',

  // Name of the Drive folder for unmatched transcripts
  UNMATCHED_FOLDER_NAME: 'Unmatched Transcripts',

  // Maximum emails to process per run (safety valve)
  MAX_EMAILS_PER_RUN: 10,

  // Email to notify when a client mention can't be matched to a Drive folder
  NOTIFICATION_EMAIL: 'billy@cactusculture.com.au',

  // Address noise words to strip when doing address-only matching
  ADDRESS_NOISE: [
    'st', 'street', 'rd', 'road', 'ave', 'avenue', 'dr', 'drive', 'ct',
    'court', 'pl', 'place', 'cres', 'crescent', 'blvd', 'boulevard',
    'ln', 'lane', 'way', 'pde', 'parade', 'tce', 'terrace', 'cct',
    'circuit', 'hwy', 'highway', 'close', 'cl', 'qld', 'nsw', 'vic'
  ]
};


// ---------------------------------------------------------------------------
// SETUP — Run once to create label + log sheet
// ---------------------------------------------------------------------------
function setup() {
  Logger.log('--- SETUP START ---');

  // 1. Create Gmail label if it does not exist
  var label = GmailApp.getUserLabelByName(CONFIG.LABEL_NAME);
  if (!label) {
    label = GmailApp.createLabel(CONFIG.LABEL_NAME);
    Logger.log('Created Gmail label: ' + CONFIG.LABEL_NAME);
  } else {
    Logger.log('Gmail label already exists: ' + CONFIG.LABEL_NAME);
  }

  // 2. Create log spreadsheet if it does not exist
  var sheet = findLogSheet_();
  if (!sheet) {
    var ss = SpreadsheetApp.create(CONFIG.LOG_SHEET_NAME);
    var ws = ss.getActiveSheet();
    ws.setName('Log');
    ws.appendRow([
      'Timestamp', 'Email Subject', 'Message ID', 'Action',
      'Client Matched', 'Folder Path', 'Doc URL', 'Notes'
    ]);
    ws.setFrozenRows(1);
    ws.getRange('1:1').setFontWeight('bold');
    Logger.log('Created log spreadsheet: ' + ss.getUrl());
  } else {
    Logger.log('Log spreadsheet already exists.');
  }

  // 3. Create Unmatched Transcripts folder if needed
  var unmatchedFolders = DriveApp.getFoldersByName(CONFIG.UNMATCHED_FOLDER_NAME);
  if (!unmatchedFolders.hasNext()) {
    DriveApp.createFolder(CONFIG.UNMATCHED_FOLDER_NAME);
    Logger.log('Created folder: ' + CONFIG.UNMATCHED_FOLDER_NAME);
  } else {
    Logger.log('Unmatched folder already exists.');
  }

  // 4. Create the time-based trigger if one does not already exist
  var triggers = ScriptApp.getProjectTriggers();
  var hasTranscriptTrigger = triggers.some(function(t) {
    return t.getHandlerFunction() === 'processNewTranscripts';
  });
  if (!hasTranscriptTrigger) {
    ScriptApp.newTrigger('processNewTranscripts')
      .timeBased()
      .everyMinutes(15)
      .create();
    Logger.log('Created 15-minute trigger for processNewTranscripts.');
  } else {
    Logger.log('Trigger already exists.');
  }

  // 5. Remind about Web App deployment
  Logger.log('');
  Logger.log('IMPORTANT: To enable the "Assign Client" feature for unmatched transcripts:');
  Logger.log('1. Click Deploy > New Deployment');
  Logger.log('2. Type: Web App');
  Logger.log('3. Execute as: Me (hello@saltlandscaping.com.au)');
  Logger.log('4. Who has access: Anyone within Salt Landscaping');
  Logger.log('5. Click Deploy and copy the URL');
  Logger.log('The notification emails will automatically use this URL.');

  Logger.log('--- SETUP COMPLETE ---');
}


// ---------------------------------------------------------------------------
// MAIN ENTRY POINT — Called by the 15-minute trigger
// ---------------------------------------------------------------------------
function processNewTranscripts() {
  var startTime = new Date();
  logToSheet_('', '', 'RUN_START', '', '', '', 'Batch run started');

  try {
    var threads = findUnprocessedThreads_();
    if (threads.length === 0) {
      logToSheet_('', '', 'NO_NEW', '', '', '', 'No unprocessed transcripts found');
      return;
    }

    // Load all client folders once per run (expensive call, do it once)
    var clientFolders = loadClientFolders_();
    Logger.log('Loaded ' + clientFolders.length + ' client folders');

    var processed = 0;
    for (var t = 0; t < threads.length && processed < CONFIG.MAX_EMAILS_PER_RUN; t++) {
      var messages = threads[t].getMessages();
      for (var m = 0; m < messages.length && processed < CONFIG.MAX_EMAILS_PER_RUN; m++) {
        var msg = messages[m];
        if (isFromGemini_(msg) && isMatchingSubject_(msg)) {
          processOneMessage_(msg, clientFolders);
          processed++;
        }
      }
      // Label the entire thread as processed
      applyProcessedLabel_(threads[t]);
    }

    var elapsed = ((new Date() - startTime) / 1000).toFixed(1);
    logToSheet_('', '', 'RUN_COMPLETE', '', '', '',
      'Processed ' + processed + ' message(s) in ' + elapsed + 's');

  } catch (err) {
    Logger.log('FATAL: ' + err.message + '\n' + err.stack);
    logToSheet_('', '', 'FATAL_ERROR', '', '', '', err.message);
  }
}


// ---------------------------------------------------------------------------
// MANUAL PROCESSING — For testing a single message by ID
// ---------------------------------------------------------------------------
function processManually(messageId) {
  if (!messageId) {
    Logger.log('Usage: processManually("message-id-here")');
    return;
  }

  var msg = GmailApp.getMessageById(messageId);
  if (!msg) {
    Logger.log('Message not found: ' + messageId);
    return;
  }

  Logger.log('Processing message: ' + msg.getSubject());
  var clientFolders = loadClientFolders_();
  Logger.log('Loaded ' + clientFolders.length + ' client folders');
  processOneMessage_(msg, clientFolders);

  // Label the thread
  applyProcessedLabel_(msg.getThread());
  Logger.log('Done. Check the log sheet for results.');
}


// ---------------------------------------------------------------------------
// TEST FUNCTION — Dry run that shows matching without saving docs
// ---------------------------------------------------------------------------
function testMatching() {
  var threads = findUnprocessedThreads_();
  if (threads.length === 0) {
    Logger.log('No unprocessed transcripts found.');
    return;
  }

  var clientFolders = loadClientFolders_();
  Logger.log('Client folders loaded: ' + clientFolders.length);

  var msg = threads[0].getMessages()[0];
  var body = extractTranscriptFromDoc_(msg) || msg.getPlainBody();
  Logger.log('Subject: ' + msg.getSubject());
  Logger.log('Body length: ' + body.length + ' chars');

  var matches = findMatchingClients_(body, clientFolders);
  Logger.log('Matched ' + matches.length + ' client(s):');
  matches.forEach(function(match) {
    Logger.log('  - ' + match.clientName + ' (score: ' + match.score +
      ', tokens: ' + match.matchedTokens.join(', ') + ')');
  });
}


// ===========================================================================
// CORE PROCESSING
// ===========================================================================

/**
 * Processes a single Gemini notes email: finds client matches, splits the
 * transcript, and saves segments to the appropriate Drive folders.
 */
function processOneMessage_(msg, clientFolders) {
  var subject = msg.getSubject();
  var messageId = msg.getId();

  Logger.log('Processing: ' + subject + ' (ID: ' + messageId + ')');
  logToSheet_(subject, messageId, 'PROCESSING', '', '', '', '');

  // Extract the meeting date from the subject line
  var meetingDate = extractDateFromSubject_(subject);

  // Get the full transcript — first try to extract it from the linked Google Doc
  // (the email body is just Gemini's summary, the real transcript is in the Doc's Transcript tab)
  var body = extractTranscriptFromDoc_(msg);
  if (!body) {
    // Fallback to email body if we can't get the doc transcript
    body = msg.getPlainBody();
    Logger.log('Using email body (could not extract doc transcript). Length: ' + body.length);
  } else {
    Logger.log('Using Google Doc transcript. Length: ' + body.length);
  }

  // Find all client folders mentioned in the transcript
  var matches = findMatchingClients_(body, clientFolders);

  if (matches.length === 0) {
    // No clients matched at all — save full transcript to Unmatched folder
    // and notify Billy to manually assign it
    Logger.log('No client matches found. Saving to Unmatched folder.');
    var unmatchedResult = saveToUnmatchedFolder_(body, subject, meetingDate);
    logToSheet_(subject, messageId, 'NO_MATCH', '', 'Unmatched Transcripts', unmatchedResult.url,
      'Full transcript saved to Unmatched folder');
    sendUnmatchedNotification_(subject, meetingDate, [
      { text: body, reason: 'No client names or addresses could be matched to any project folder.', docId: unmatchedResult.docId }
    ], unmatchedResult.url);
    return;
  }

  Logger.log('Found ' + matches.length + ' client match(es)');

  // Split transcript into client-specific segments
  var segments = splitTranscriptByClient_(body, matches);

  // Save each segment to the appropriate client folder
  var unmatchedSegments = [];
  for (var i = 0; i < segments.length; i++) {
    var seg = segments[i];
    try {
      var docUrl = saveSegmentToClientFolder_(seg, meetingDate);
      logToSheet_(subject, messageId, 'SAVED', seg.clientName,
        seg.folderPath, docUrl, seg.matchType + ' match | ' + seg.text.length + ' chars');
    } catch (err) {
      Logger.log('Error saving segment for ' + seg.clientName + ': ' + err.message);
      logToSheet_(subject, messageId, 'SAVE_ERROR', seg.clientName,
        '', '', err.message);
    }
  }
}


// ===========================================================================
// GMAIL SEARCH & FILTERING
// ===========================================================================

/**
 * Finds Gmail threads that match our criteria and have NOT been labelled.
 */
function findUnprocessedThreads_() {
  var query = 'from:' + CONFIG.GEMINI_SENDER +
    ' subject:(Notes:) -label:' + CONFIG.LABEL_NAME;

  // If searching a delegated/shared mailbox, add to: filter
  if (CONFIG.GMAIL_ACCOUNT) {
    query += ' to:' + CONFIG.GMAIL_ACCOUNT;
  }

  Logger.log('Gmail query: ' + query);
  var threads = GmailApp.search(query, 0, CONFIG.MAX_EMAILS_PER_RUN);
  Logger.log('Found ' + threads.length + ' unprocessed thread(s)');
  return threads;
}

/**
 * Returns true if the message sender is the Gemini notes address.
 */
function isFromGemini_(msg) {
  var from = msg.getFrom().toLowerCase();
  return from.indexOf(CONFIG.GEMINI_SENDER.toLowerCase()) !== -1;
}

/**
 * Returns true if the meeting title contains ALL required words AND at least
 * one of the quoting keywords. Handles any word order.
 * e.g. "Warren + Billy - Quoting", "Billy + Warren - Quick Quote",
 *      "Warren and Billy quote review" all match.
 * Subject format from Gemini: "Notes: 'Warren + Billy - Quoting' 10 Apr 2026"
 */
function isMatchingSubject_(msg) {
  var subject = msg.getSubject();

  // Extract the meeting title from between quotes (handles straight + smart quotes)
  var titleToCheck = '';
  var match = subject.match(/^Notes:\s*['\u2018\u2019\u201C\u201D](.+?)['\u2018\u2019\u201C\u201D]/i);
  if (match) {
    titleToCheck = match[1].trim().toLowerCase();
  } else {
    // Fallback: use full subject
    titleToCheck = subject.toLowerCase();
  }

  // ALL required words must be present (e.g. "Warren" AND "Billy")
  var hasAllRequired = CONFIG.MEETING_TITLE_REQUIRED_WORDS.every(function(word) {
    return titleToCheck.indexOf(word.toLowerCase()) !== -1;
  });

  if (!hasAllRequired) return false;

  // At least ONE quoting keyword must be present (e.g. "Quote" OR "Quoting")
  var hasQuotingWord = CONFIG.MEETING_TITLE_REQUIRED_ANY.some(function(word) {
    return titleToCheck.indexOf(word.toLowerCase()) !== -1;
  });

  return hasQuotingWord;
}

/**
 * Extracts a date from the subject line. Falls back to today's date.
 * Subject format: "Notes: 'Meeting Title' April 10, 2026"
 */
function extractDateFromSubject_(subject) {
  // Try to find a date at the end of the subject
  // Pattern: month day, year  OR  YYYY-MM-DD  OR  DD/MM/YYYY
  var patterns = [
    /(\w+ \d{1,2},?\s*\d{4})\s*$/,           // "April 10, 2026"
    /(\d{4}-\d{2}-\d{2})\s*$/,                // "2026-04-10"
    /(\d{1,2}\/\d{1,2}\/\d{2,4})\s*$/,        // "10/04/2026"
    /(\d{1,2}\s+\w+\s+\d{4})\s*$/             // "10 April 2026"
  ];

  for (var i = 0; i < patterns.length; i++) {
    var match = subject.match(patterns[i]);
    if (match) {
      var parsed = new Date(match[1]);
      if (!isNaN(parsed.getTime())) {
        return formatDate_(parsed);
      }
    }
  }

  // Fallback: use today's date
  return formatDate_(new Date());
}

/**
 * Extracts the full transcript text from the Google Doc linked in the Gemini email.
 * The email HTML contains a link to a Google Doc. That doc has a "Transcript" tab
 * with the verbatim meeting transcription. We read that tab's content.
 *
 * Returns the transcript text, or null if extraction fails.
 */
function extractTranscriptFromDoc_(msg) {
  try {
    // Get the HTML body to find the Google Doc link
    var htmlBody = msg.getBody();

    // Extract Google Doc ID from the link
    // Pattern: docs.google.com/document/d/{DOC_ID}/edit
    var docMatch = htmlBody.match(/docs\.google\.com\/document\/d\/([a-zA-Z0-9_-]+)\//);
    if (!docMatch) {
      Logger.log('No Google Doc link found in email');
      return null;
    }

    var docId = docMatch[1];
    Logger.log('Found Google Doc ID: ' + docId);

    // Open the document and read all tabs
    var doc = DocumentApp.openById(docId);

    // Apps Script doesn't have direct tab API, but the document body
    // contains all tab content. Search for the transcript section.
    var docBody = doc.getBody();
    var fullText = docBody.getText();

    // The transcript section typically starts after "Transcript" heading
    // and contains timestamped dialogue like "00:00:00"
    var transcriptStart = fullText.indexOf('Transcript');
    if (transcriptStart === -1) {
      // Try looking for timestamp pattern that indicates transcript content
      var tsMatch = fullText.match(/\d{2}:\d{2}:\d{2}/);
      if (tsMatch) {
        transcriptStart = fullText.indexOf(tsMatch[0]);
      }
    }

    if (transcriptStart !== -1) {
      var transcript = fullText.substring(transcriptStart);
      // Clean up: remove the "Transcription ended" footer and boilerplate
      var endMarker = transcript.indexOf('This editable transcript was computer generated');
      if (endMarker !== -1) {
        transcript = transcript.substring(0, endMarker);
      }
      Logger.log('Extracted transcript from Doc tab: ' + transcript.length + ' chars');
      return transcript;
    }

    // If no transcript section found, the doc might only have notes/summary
    Logger.log('No transcript section found in Google Doc — falling back to email body');
    return null;

  } catch (err) {
    Logger.log('Error extracting transcript from Doc: ' + err.message);
    return null;
  }
}

/**
 * Applies the "Processed-Transcript" label to a Gmail thread.
 */
function applyProcessedLabel_(thread) {
  var label = GmailApp.getUserLabelByName(CONFIG.LABEL_NAME);
  if (!label) {
    label = GmailApp.createLabel(CONFIG.LABEL_NAME);
  }
  thread.addLabel(label);
}


// ===========================================================================
// CLIENT FOLDER LOADING & TOKEN GENERATION
// ===========================================================================

/**
 * Loads all client folders from the parent "Landscaping Projects" folder.
 * Returns an array of objects with folder metadata and search tokens.
 */
function loadClientFolders_() {
  var parent = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);
  var folders = parent.getFolders();
  var result = [];

  while (folders.hasNext()) {
    var folder = folders.next();
    var name = folder.getName();

    // Expected format: "Client Name - Address"
    var tokens = generateSearchTokens_(name);

    result.push({
      folder: folder,
      name: name,
      tokens: tokens,
      id: folder.getId()
    });
  }

  // Sort by name for deterministic processing
  result.sort(function(a, b) { return a.name.localeCompare(b.name); });

  return result;
}

/**
 * Generates fuzzy search tokens from a client folder name.
 * Input:  "Sean Bryant - 42 Marine Pde Capalaba"
 * Output: { firstName: "sean", lastName: "bryant", fullName: "sean bryant",
 *           addressWords: ["marine", "capalaba"], streetNumber: "42",
 *           allTokens: ["sean", "bryant", "marine", "capalaba"] }
 */
function generateSearchTokens_(folderName) {
  var parts = folderName.split(/\s*-\s*/);
  var namePart = (parts[0] || '').trim();
  var addressPart = (parts.slice(1).join(' - ') || '').trim();

  // Parse the client name
  var nameWords = namePart.split(/\s+/).filter(Boolean);
  var firstName = (nameWords[0] || '').toLowerCase();
  var lastName = (nameWords[nameWords.length - 1] || '').toLowerCase();
  var fullName = nameWords.map(function(w) { return w.toLowerCase(); }).join(' ');

  // Parse the address — strip noise words and numbers
  var addressWords = addressPart.split(/\s+/)
    .map(function(w) { return w.toLowerCase().replace(/[^a-z]/g, ''); })
    .filter(function(w) {
      return w.length > 2 &&
        CONFIG.ADDRESS_NOISE.indexOf(w) === -1 &&
        !/^\d+$/.test(w);
    });

  // Extract street number if present
  var streetNumMatch = addressPart.match(/^(\d+[A-Za-z]?)\s/);
  var streetNumber = streetNumMatch ? streetNumMatch[1] : '';

  // Build the combined token list (deduplicated)
  var allTokens = [];
  var seen = {};
  [firstName, lastName].concat(addressWords).forEach(function(t) {
    if (t && t.length > 1 && !seen[t]) {
      seen[t] = true;
      allTokens.push(t);
    }
  });

  return {
    firstName: firstName,
    lastName: lastName,
    fullName: fullName,
    nameWords: nameWords.map(function(w) { return w.toLowerCase(); }),
    addressWords: addressWords,
    streetNumber: streetNumber,
    allTokens: allTokens
  };
}


// ===========================================================================
// CLIENT MATCHING — Exact "Name - Address" detection
// ===========================================================================

/**
 * Scans the transcript body for exact client references that Billy announces
 * during the call. Billy says "{Name} - {Address}" before discussing each
 * client project. We match these announcements against the Drive folder names.
 *
 * Matching cascade (in priority order):
 * 1. EXACT: full folder name appears verbatim in the transcript
 * 2. NAME_ONLY: the client name (before the dash) matches exactly
 * 3. ADDRESS_ONLY: key address words (suburb, street name) match a folder
 *
 * If none of these work, the segment is flagged as unmatched and Billy
 * receives an email notification to manually assign it.
 *
 * Returns an array of match objects sorted by position in the transcript.
 */
function findMatchingClients_(body, clientFolders) {
  var bodyLower = normaliseWhitespace_(body.toLowerCase());
  var matches = [];

  for (var i = 0; i < clientFolders.length; i++) {
    var cf = clientFolders[i];
    var folderNameNorm = normaliseWhitespace_(cf.name.toLowerCase());
    var matchType = '';
    var position = -1;

    // --- Strategy 1: EXACT full folder name match ---
    var exactPos = bodyLower.indexOf(folderNameNorm);
    if (exactPos !== -1) {
      matchType = 'EXACT';
      position = exactPos;
    }

    // --- Strategy 2: NAME_ONLY (client name before the dash) ---
    if (position === -1) {
      var namePart = cf.tokens.fullName; // already lowercase
      if (namePart && namePart.length > 3) {
        var nameRegex = new RegExp('\\b' + escapeRegex_(namePart) + '\\b');
        var nameMatch = nameRegex.exec(bodyLower);
        if (nameMatch) {
          matchType = 'NAME_ONLY';
          position = nameMatch.index;
        }
      }
    }

    // --- Strategy 3: ADDRESS_ONLY (suburb or street name) ---
    if (position === -1) {
      var addrWords = cf.tokens.addressWords;
      // Need at least 2 address tokens to match (avoids false positives on
      // common words). Suburb names are the strongest signal.
      if (addrWords && addrWords.length > 0) {
        var addrHits = 0;
        var firstHitPos = -1;
        for (var a = 0; a < addrWords.length; a++) {
          var aw = addrWords[a];
          if (aw.length < 4) continue; // skip short words
          var awRegex = new RegExp('\\b' + escapeRegex_(aw) + '\\b');
          var awMatch = awRegex.exec(bodyLower);
          if (awMatch) {
            addrHits++;
            if (firstHitPos === -1) firstHitPos = awMatch.index;
          }
        }
        // Require 2+ address words to match, OR 1 long suburb name (6+ chars)
        var hasStrongSuburb = addrWords.some(function(w) {
          return w.length >= 6 && new RegExp('\\b' + escapeRegex_(w) + '\\b').test(bodyLower);
        });
        if (addrHits >= 2 || hasStrongSuburb) {
          matchType = 'ADDRESS_ONLY';
          position = firstHitPos;
        }
      }
    }

    if (position !== -1) {
      var scores = { 'EXACT': 10, 'NAME_ONLY': 7, 'ADDRESS_ONLY': 4 };
      matches.push({
        clientFolder: cf,
        clientName: cf.name,
        score: scores[matchType] || 1,
        matchType: matchType,
        matchedTokens: [matchType + ':' + cf.name],
        position: position
      });
    }
  }

  // Sort by position in the transcript (order Billy discussed them)
  matches.sort(function(a, b) { return a.position - b.position; });

  return matches;
}

/**
 * Normalises whitespace: collapses multiple spaces/tabs into single space,
 * normalises different dash types to a simple hyphen with spaces.
 */
function normaliseWhitespace_(str) {
  return str
    .replace(/[\u2013\u2014\u2015]/g, '-')   // em/en dashes → hyphen
    .replace(/\s+/g, ' ')                      // collapse whitespace
    .trim();
}


// ===========================================================================
// TRANSCRIPT SPLITTING
// ===========================================================================

/**
 * Splits the transcript into segments, one per matched client.
 * Strategy:
 *   1. Matches are already sorted by position (order Billy discussed them)
 *   2. Everything between mention[i] and mention[i+1] belongs to client[i]
 *   3. Any text before the first mention is "preamble" — prepended to all segments
 *   4. Any text after the last mention to end-of-body goes to the last client
 */
function splitTranscriptByClient_(body, matches) {
  if (matches.length === 0) return [];

  // Matches are already sorted by position from findMatchingClients_
  // Map them to the mentions format
  var mentions = matches.map(function(m) {
    return { position: m.position, match: m };
  });

  // Deduplicate — if a client appears at multiple positions, keep the first
  var seen = {};
  var uniqueMentions = [];
  for (var j = 0; j < mentions.length; j++) {
    var key = mentions[j].match.clientFolder.id;
    if (!seen[key]) {
      seen[key] = true;
      uniqueMentions.push(mentions[j]);
    }
  }
  mentions = uniqueMentions;

  // Extract preamble (text before first client mention)
  // Walk back to find the start of the line/paragraph containing the first mention
  var firstPos = mentions[0].position;
  var preambleEnd = findParagraphStart_(body, firstPos);
  var preamble = body.substring(0, preambleEnd).trim();

  // Build segments
  var segments = [];
  for (var k = 0; k < mentions.length; k++) {
    var segStart = findParagraphStart_(body, mentions[k].position);
    var segEnd;

    if (k < mentions.length - 1) {
      segEnd = findParagraphStart_(body, mentions[k + 1].position);
    } else {
      segEnd = body.length;
    }

    var segText = body.substring(segStart, segEnd).trim();

    // Prepend preamble if it exists and this is not the first segment
    // (first segment already includes preamble naturally if it starts at 0)
    var fullText = '';
    if (preamble && segStart > preambleEnd) {
      fullText = '--- Meeting Context ---\n\n' + preamble +
        '\n\n--- ' + mentions[k].match.clientName + ' Discussion ---\n\n' + segText;
    } else {
      fullText = segText;
    }

    segments.push({
      clientName: mentions[k].match.clientName,
      clientFolder: mentions[k].match.clientFolder,
      folderPath: mentions[k].match.clientName + '/' + CONFIG.CLIENT_SUBFOLDER,
      matchType: mentions[k].match.matchType,
      text: fullText
    });
  }

  return segments;
}

/**
 * Finds the position of a word in text with word-boundary awareness.
 * Returns the character index or -1 if not found.
 */
function findWordBoundaryPosition_(text, word) {
  if (!word || word.length === 0) return -1;
  var regex = new RegExp('\\b' + escapeRegex_(word) + '\\b');
  var match = regex.exec(text);
  return match ? match.index : -1;
}

/**
 * Given a position in the body, walks backwards to find the start of
 * the paragraph (double newline) or line, so we don't cut mid-sentence.
 */
function findParagraphStart_(body, pos) {
  // Look for double newline (paragraph break) before this position
  var searchFrom = Math.max(0, pos - 500);
  var chunk = body.substring(searchFrom, pos);

  var paraBreak = chunk.lastIndexOf('\n\n');
  if (paraBreak !== -1) {
    return searchFrom + paraBreak + 2; // skip the double newline
  }

  // Fall back to single newline
  var lineBreak = chunk.lastIndexOf('\n');
  if (lineBreak !== -1) {
    return searchFrom + lineBreak + 1;
  }

  return searchFrom;
}


// ===========================================================================
// GOOGLE DRIVE — SAVE DOCUMENTS
// ===========================================================================

/**
 * Saves a transcript segment as a Google Doc in the client's subfolder.
 * Creates the "03 - Site Information" subfolder if it does not exist.
 * Returns the URL of the created document.
 */
function saveSegmentToClientFolder_(segment, meetingDate) {
  var clientFolder = segment.clientFolder.folder;

  // Find or create the "03 - Site Information" subfolder
  var subFolder = findOrCreateSubfolder_(clientFolder, CONFIG.CLIENT_SUBFOLDER);

  // Build the document name
  var docName = CONFIG.DOC_NAME_PREFIX + ' - ' + meetingDate;

  // Check if a doc with this name already exists (avoid duplicates)
  var existingFiles = subFolder.getFilesByName(docName);
  if (existingFiles.hasNext()) {
    var existing = existingFiles.next();
    Logger.log('Doc already exists: ' + docName + ' — skipping');
    return existing.getUrl();
  }

  // Create the Google Doc
  var doc = DocumentApp.create(docName);
  var docBody = doc.getBody();

  // Add header
  docBody.appendParagraph(CONFIG.DOC_NAME_PREFIX)
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  docBody.appendParagraph('Date: ' + meetingDate)
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  docBody.appendParagraph('Client: ' + segment.clientName)
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  docBody.appendParagraph(''); // spacer

  // Add transcript content — split by line for readability
  var lines = segment.text.split('\n');
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    // Detect section headers (lines starting with ---)
    if (/^---\s*.+\s*---$/.test(line.trim())) {
      docBody.appendParagraph(line.replace(/---/g, '').trim())
        .setHeading(DocumentApp.ParagraphHeading.HEADING3);
    } else {
      docBody.appendParagraph(line);
    }
  }

  // Add metadata footer
  docBody.appendParagraph('');
  docBody.appendHorizontalRule();
  docBody.appendParagraph(
    'Auto-generated by SALT Transcript Processor on ' +
    Utilities.formatDate(new Date(), 'Australia/Brisbane', 'yyyy-MM-dd HH:mm:ss')
  ).setItalic(true);

  doc.saveAndClose();

  // Move the doc to the target subfolder (it starts in Drive root)
  var docFile = DriveApp.getFileById(doc.getId());
  subFolder.addFile(docFile);

  // Remove from root (Drive root is the default parent)
  var parents = docFile.getParents();
  while (parents.hasNext()) {
    var parent = parents.next();
    if (parent.getId() !== subFolder.getId()) {
      parent.removeFile(docFile);
    }
  }

  Logger.log('Saved: ' + docName + ' → ' + segment.folderPath);
  return doc.getUrl();
}

/**
 * Saves the full transcript to the "Unmatched Transcripts" folder.
 */
function saveToUnmatchedFolder_(body, subject, meetingDate) {
  var folders = DriveApp.getFoldersByName(CONFIG.UNMATCHED_FOLDER_NAME);
  var folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(CONFIG.UNMATCHED_FOLDER_NAME);
  }

  var docName = 'Unmatched - ' + subject + ' - ' + meetingDate;
  // Sanitise the name (Drive doesn't like certain characters)
  docName = docName.replace(/[\/\\:*?"<>|]/g, '-').substring(0, 200);

  var doc = DocumentApp.create(docName);
  var docBody = doc.getBody();

  docBody.appendParagraph('Unmatched Meeting Transcript')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  docBody.appendParagraph('Original Subject: ' + subject);
  docBody.appendParagraph('Date: ' + meetingDate);
  docBody.appendParagraph('');
  docBody.appendHorizontalRule();
  docBody.appendParagraph('');

  var lines = body.split('\n');
  for (var i = 0; i < lines.length; i++) {
    docBody.appendParagraph(lines[i]);
  }

  docBody.appendParagraph('');
  docBody.appendHorizontalRule();
  docBody.appendParagraph(
    'No client matches found. Auto-saved by SALT Transcript Processor on ' +
    Utilities.formatDate(new Date(), 'Australia/Brisbane', 'yyyy-MM-dd HH:mm:ss')
  ).setItalic(true);

  doc.saveAndClose();

  var docFile = DriveApp.getFileById(doc.getId());
  folder.addFile(docFile);
  var parents = docFile.getParents();
  while (parents.hasNext()) {
    var parent = parents.next();
    if (parent.getId() !== folder.getId()) {
      parent.removeFile(docFile);
    }
  }

  Logger.log('Saved unmatched transcript: ' + docName);
  return { url: doc.getUrl(), docId: doc.getId() };
}


// ===========================================================================
// EMAIL NOTIFICATIONS
// ===========================================================================

/**
 * Sends Billy an email when transcript segments couldn't be matched to any
 * client folder. Includes the unmatched text so he can identify which client
 * it belongs to and take action.
 */
function sendUnmatchedNotification_(subject, meetingDate, unmatchedItems, docUrl) {
  var to = CONFIG.NOTIFICATION_EMAIL;
  if (!to) return;

  // Get the web app URL for the assignment page
  var webAppUrl = ScriptApp.getService().getUrl();

  var htmlBody = '<div style="font-family:DM Sans,Arial,sans-serif;color:#2C2C2A;max-width:600px;">'
    + '<div style="background:#2C2C2A;padding:16px 20px;border-radius:8px 8px 0 0;">'
    + '<span style="color:#8A9B8E;font-size:11px;letter-spacing:0.15em;text-transform:uppercase;">Salt Landscaping</span>'
    + '<br><span style="color:#F5F2ED;font-size:16px;">Unmatched Transcript — Assign Client</span>'
    + '</div>'
    + '<div style="background:#FFFFFF;padding:20px;border:1px solid #E2D9CC;border-top:none;border-radius:0 0 8px 8px;">'
    + '<p style="margin:0 0 12px;color:#5A6068;font-size:14px;">The transcript processor couldn\'t match some discussion from the '
    + '<strong>' + subject + '</strong> meeting (' + meetingDate + ') to a client folder.</p>'
    + '<p style="margin:0 0 16px;color:#5A6068;font-size:14px;">Tap <strong>Assign Client</strong> below to select the correct project — the transcript will be saved automatically.</p>';

  for (var i = 0; i < unmatchedItems.length; i++) {
    var item = unmatchedItems[i];
    // Extract the doc ID from the URL for the assignment link
    var itemDocId = '';
    if (item.docId) {
      itemDocId = item.docId;
    } else if (docUrl) {
      var docIdMatch = docUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (docIdMatch) itemDocId = docIdMatch[1];
    }

    htmlBody += '<div style="background:#F5F2ED;padding:14px 16px;border-radius:8px;margin:12px 0;border-left:3px solid #C4953A;">'
      + '<div style="font-size:11px;font-weight:600;color:#C4953A;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:8px;">Unmatched Segment ' + (i + 1) + '</div>'
      + '<div style="font-size:12px;color:#5A6068;margin-bottom:10px;">' + escapeHtml_(item.reason) + '</div>'
      + '<div style="font-family:monospace;font-size:11px;color:#4A4A47;white-space:pre-wrap;max-height:150px;overflow-y:auto;background:#fff;padding:10px;border-radius:6px;border:1px solid #E2D9CC;">'
      + escapeHtml_(item.text.substring(0, 400))
      + (item.text.length > 400 ? '\n\n[... truncated]' : '')
      + '</div>';

    if (itemDocId && webAppUrl) {
      var assignUrl = webAppUrl + '?docId=' + itemDocId;
      htmlBody += '<div style="margin-top:10px;">'
        + '<a href="' + assignUrl + '" style="display:inline-block;padding:10px 20px;'
        + 'background:#A07850;color:#F5F2ED;text-decoration:none;border-radius:6px;font-size:13px;font-weight:500;">'
        + 'Assign Client</a></div>';
    }

    htmlBody += '</div>';
  }

  htmlBody += '<p style="margin:16px 0 0;font-size:11px;color:#C8BFB0;">Auto-generated by SALT Transcript Processor</p>'
    + '</div></div>';

  GmailApp.sendEmail(to, 'Assign Client: Unmatched quoting transcript — ' + meetingDate,
    'Unmatched transcript from ' + subject + '. Open this email to assign it to a client.',
    { htmlBody: htmlBody, name: 'SALT Transcript Processor' }
  );

  Logger.log('Sent unmatched notification to ' + to);
}

/**
 * Escapes HTML entities for safe inclusion in email body.
 */
function escapeHtml_(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Finds a subfolder by name within a parent folder, or creates it.
 */
function findOrCreateSubfolder_(parentFolder, subfolderName) {
  var subs = parentFolder.getFoldersByName(subfolderName);
  if (subs.hasNext()) {
    return subs.next();
  }
  Logger.log('Creating subfolder: ' + subfolderName + ' in ' + parentFolder.getName());
  return parentFolder.createFolder(subfolderName);
}


// ===========================================================================
// WEB APP — Client assignment for unmatched segments
// ===========================================================================
// Deploy as: Web App → Execute as "Me" → Access "Anyone within organisation"
// This serves a page where Billy can select which client an unmatched
// transcript belongs to. The doc is then auto-moved to the correct folder.

/**
 * Serves the client selection page when Billy clicks the "Assign Client"
 * link in the unmatched notification email.
 * URL parameter: ?docId={Google Doc ID of the unmatched transcript}
 */
function doGet(e) {
  var docId = (e && e.parameter && e.parameter.docId) ? e.parameter.docId : '';

  if (!docId) {
    return HtmlService.createHtmlOutput(
      '<html><body style="font-family:DM Sans,Arial,sans-serif;padding:40px;text-align:center;">'
      + '<h2 style="color:#9E4B4B;">Missing document ID</h2>'
      + '<p>This link is invalid. Please use the link from the notification email.</p>'
      + '</body></html>'
    ).setTitle('SALT — Error');
  }

  // Verify the doc exists
  var docName = '';
  try {
    var file = DriveApp.getFileById(docId);
    docName = file.getName();
  } catch (err) {
    return HtmlService.createHtmlOutput(
      '<html><body style="font-family:DM Sans,Arial,sans-serif;padding:40px;text-align:center;">'
      + '<h2 style="color:#9E4B4B;">Document not found</h2>'
      + '<p>The document may have been moved or deleted.</p>'
      + '</body></html>'
    ).setTitle('SALT — Error');
  }

  // Load all client folders
  var parent = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);
  var folders = parent.getFolders();
  var folderOptions = [];
  while (folders.hasNext()) {
    var f = folders.next();
    folderOptions.push({ id: f.getId(), name: f.getName() });
  }
  folderOptions.sort(function(a, b) { return a.name.localeCompare(b.name); });

  // Build the HTML page
  var html = '<!DOCTYPE html><html><head>'
    + '<meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;500&family=DM+Sans:wght@300;400;500;600&display=swap" rel="stylesheet">'
    + '<style>'
    + ':root{--charcoal:#2C2C2A;--salt:#F5F2ED;--sage:#8A9B8E;--timber:#A07850;--timber-dark:#7A5830;--steel:#5A6068;--sand:#E2D9CC;--success:#6B8F71;--error:#9E4B4B;}'
    + '*{margin:0;padding:0;box-sizing:border-box;}'
    + 'body{background:var(--salt);font-family:"DM Sans",Arial,sans-serif;color:var(--charcoal);}'
    + '.header{background:var(--charcoal);padding:16px 20px;}'
    + '.header-brand{font-size:10px;letter-spacing:0.2em;text-transform:uppercase;color:var(--sage);font-weight:600;}'
    + '.header-title{font-size:18px;color:var(--salt);font-family:"Playfair Display",Georgia,serif;margin-top:4px;}'
    + '.container{max-width:600px;margin:24px auto;padding:0 16px;}'
    + '.doc-info{background:#fff;border:1px solid var(--sand);border-radius:10px;padding:16px;margin-bottom:20px;}'
    + '.doc-info h3{font-size:14px;color:var(--steel);font-weight:400;margin-bottom:4px;}'
    + '.doc-info p{font-size:16px;font-weight:500;}'
    + '.search{width:100%;padding:10px 14px;border:1.5px solid var(--sand);border-radius:8px;font-size:14px;font-family:inherit;margin-bottom:12px;outline:none;}'
    + '.search:focus{border-color:var(--timber);box-shadow:0 0 0 3px rgba(160,120,80,0.15);}'
    + '.folder-list{list-style:none;max-height:50vh;overflow-y:auto;}'
    + '.folder-item{padding:12px 14px;border:1.5px solid var(--sand);border-radius:8px;margin-bottom:6px;cursor:pointer;background:#fff;transition:all 0.15s;display:flex;justify-content:space-between;align-items:center;}'
    + '.folder-item:hover{border-color:var(--timber);background:rgba(160,120,80,0.04);}'
    + '.folder-item .name{font-size:14px;font-weight:500;}'
    + '.folder-item .arrow{color:var(--timber);font-size:18px;opacity:0;transition:opacity 0.15s;}'
    + '.folder-item:hover .arrow{opacity:1;}'
    + '.folder-item.hidden{display:none;}'
    + '.status{text-align:center;padding:40px 20px;display:none;}'
    + '.status.show{display:block;}'
    + '.status.success h2{color:var(--success);}'
    + '.status.error h2{color:var(--error);}'
    + '.spinner{width:24px;height:24px;border:3px solid var(--sand);border-top-color:var(--timber);border-radius:50%;animation:spin 0.6s linear infinite;margin:0 auto 12px;}'
    + '@keyframes spin{to{transform:rotate(360deg);}}'
    + '</style></head><body>'
    + '<div class="header">'
    + '<div class="header-brand">Salt Landscaping</div>'
    + '<div class="header-title">Assign Unmatched Transcript</div>'
    + '</div>'
    + '<div class="container">'
    + '<div class="doc-info">'
    + '<h3>Unmatched Document</h3>'
    + '<p>' + escapeHtml_(docName) + '</p>'
    + '</div>'
    + '<input type="text" class="search" id="search" placeholder="Search clients..." oninput="filterFolders()">'
    + '<ul class="folder-list" id="folderList">';

  for (var i = 0; i < folderOptions.length; i++) {
    var fo = folderOptions[i];
    html += '<li class="folder-item" data-name="' + escapeHtml_(fo.name).toLowerCase() + '" onclick="assignClient(\'' + fo.id + '\',\'' + escapeHtml_(fo.name).replace(/'/g, "\\'") + '\')">'
      + '<span class="name">' + escapeHtml_(fo.name) + '</span>'
      + '<span class="arrow">&rarr;</span>'
      + '</li>';
  }

  html += '</ul>'
    + '<div class="status" id="loading"><div class="spinner"></div><p>Saving to client folder...</p></div>'
    + '<div class="status success" id="success"><h2>Saved</h2><p id="successMsg"></p></div>'
    + '<div class="status error" id="error"><h2>Error</h2><p id="errorMsg"></p></div>'
    + '</div>'
    + '<script>'
    + 'function filterFolders(){'
    + '  var q=document.getElementById("search").value.toLowerCase();'
    + '  document.querySelectorAll(".folder-item").forEach(function(el){'
    + '    el.classList.toggle("hidden",q&&el.dataset.name.indexOf(q)===-1);'
    + '  });'
    + '}'
    + 'function assignClient(folderId,folderName){'
    + '  document.getElementById("folderList").style.display="none";'
    + '  document.getElementById("search").style.display="none";'
    + '  document.getElementById("loading").classList.add("show");'
    + '  google.script.run'
    + '    .withSuccessHandler(function(result){'
    + '      document.getElementById("loading").classList.remove("show");'
    + '      if(result.success){'
    + '        document.getElementById("successMsg").textContent="Saved to "+folderName+" / ' + CONFIG.CLIENT_SUBFOLDER + '";'
    + '        document.getElementById("success").classList.add("show");'
    + '      }else{'
    + '        document.getElementById("errorMsg").textContent=result.error||"Unknown error";'
    + '        document.getElementById("error").classList.add("show");'
    + '      }'
    + '    })'
    + '    .withFailureHandler(function(err){'
    + '      document.getElementById("loading").classList.remove("show");'
    + '      document.getElementById("errorMsg").textContent=err.message||"Unknown error";'
    + '      document.getElementById("error").classList.add("show");'
    + '    })'
    + '    .assignToClient("' + docId + '",folderId,folderName);'
    + '}'
    + '</script>'
    + '</body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle('SALT — Assign Transcript')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Called by the web app when Billy selects a client folder.
 * Moves the unmatched doc into the client's 06 - Quoting Context subfolder.
 */
function assignToClient(docId, targetFolderId, folderName) {
  try {
    var docFile = DriveApp.getFileById(docId);
    var targetFolder = DriveApp.getFolderById(targetFolderId);

    // Find or create the quoting context subfolder
    var subFolder = findOrCreateSubfolder_(targetFolder, CONFIG.CLIENT_SUBFOLDER);

    // Move the doc: add to new folder, remove from old
    subFolder.addFile(docFile);
    var parents = docFile.getParents();
    while (parents.hasNext()) {
      var parent = parents.next();
      if (parent.getId() !== subFolder.getId()) {
        parent.removeFile(docFile);
      }
    }

    // Log it
    logToSheet_('', '', 'MANUAL_ASSIGN', folderName,
      folderName + '/' + CONFIG.CLIENT_SUBFOLDER, docFile.getUrl(),
      'Billy manually assigned unmatched transcript');

    Logger.log('Assigned doc ' + docId + ' to ' + folderName + '/' + CONFIG.CLIENT_SUBFOLDER);
    return { success: true };

  } catch (err) {
    Logger.log('Error assigning doc: ' + err.message);
    return { success: false, error: err.message };
  }
}


// ===========================================================================
// LOGGING
// ===========================================================================

/**
 * Appends a row to the "Transcript Processing Log" spreadsheet.
 */
function logToSheet_(subject, messageId, action, client, folderPath, docUrl, notes) {
  try {
    var sheet = findLogSheet_();
    if (!sheet) {
      // Log sheet doesn't exist yet — create it on the fly
      var ss = SpreadsheetApp.create(CONFIG.LOG_SHEET_NAME);
      var ws = ss.getActiveSheet();
      ws.setName('Log');
      ws.appendRow([
        'Timestamp', 'Email Subject', 'Message ID', 'Action',
        'Client Matched', 'Folder Path', 'Doc URL', 'Notes'
      ]);
      ws.setFrozenRows(1);
      sheet = ws;
    }

    sheet.appendRow([
      Utilities.formatDate(new Date(), 'Australia/Brisbane', 'yyyy-MM-dd HH:mm:ss'),
      subject,
      messageId,
      action,
      client,
      folderPath,
      docUrl,
      notes
    ]);
  } catch (err) {
    // If we can't log, at least write to the execution log
    Logger.log('LOG_ERROR: Could not write to sheet — ' + err.message);
  }
}

/**
 * Finds the log spreadsheet by name. Returns the first sheet, or null.
 */
function findLogSheet_() {
  var files = DriveApp.getFilesByName(CONFIG.LOG_SHEET_NAME);
  while (files.hasNext()) {
    var file = files.next();
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      var ss = SpreadsheetApp.openById(file.getId());
      return ss.getSheets()[0];
    }
  }
  return null;
}


// ===========================================================================
// UTILITY FUNCTIONS
// ===========================================================================

/**
 * Formats a Date object as YYYY-MM-DD.
 */
function formatDate_(date) {
  return Utilities.formatDate(date, 'Australia/Brisbane', 'yyyy-MM-dd');
}

/**
 * Escapes special regex characters in a string.
 */
function escapeRegex_(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
