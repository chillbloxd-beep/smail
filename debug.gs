function testGeminiConnection() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY');
  const digestTo = props.getProperty('DIGEST_TO');

  if (!apiKey) throw new Error('Missing Script Property: GEMINI_API_KEY');
  if (!digestTo) throw new Error('Missing Script Property: DIGEST_TO');

  const timezone = (typeof CONFIG !== 'undefined' && CONFIG.TIMEZONE) ? CONFIG.TIMEZONE : 'Asia/Singapore';
  const model = (typeof CONFIG !== 'undefined' && CONFIG.GEMINI_MODEL) ? CONFIG.GEMINI_MODEL : 'gemini-2.5-flash';

  const prompt = `
Reply with ONLY valid JSON in this exact format:
{
  "ok": true,
  "message": "short success message"
}
  `.trim();

  const url =
    'https://generativelanguage.googleapis.com/v1beta/models/' +
    encodeURIComponent(model) +
    ':generateContent?key=' +
    encodeURIComponent(apiKey);

  let subject = '';
  let body = '';
  let html = '';

  try {
    const started = Date.now();

    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: {
          temperature: 0,
          responseMimeType: 'application/json'
        }
      }),
      muteHttpExceptions: true
    });

    const elapsedMs = Date.now() - started;
    const code = response.getResponseCode();
    const raw = response.getContentText();
    const nowStr = Utilities.formatDate(new Date(), timezone, 'EEE, d MMM yyyy, h:mm a');

    if (code < 200 || code >= 300) {
      subject = `Gemini Connection Test FAILED (${code})`;
      body =
        `Gemini connection test failed.\n\n` +
        `Time: ${nowStr}\n` +
        `Model: ${model}\n` +
        `HTTP Code: ${code}\n` +
        `Elapsed: ${elapsedMs} ms\n\n` +
        `Response:\n${raw}`;

      html =
        `<div style="font-family:Arial,sans-serif;max-width:800px;margin:0 auto;padding:20px;">` +
        `<h2 style="color:#d93025;">Gemini Connection Test FAILED</h2>` +
        `<p><strong>Time:</strong> ${escapeHtmlSafe_(nowStr)}<br>` +
        `<strong>Model:</strong> ${escapeHtmlSafe_(model)}<br>` +
        `<strong>HTTP Code:</strong> ${code}<br>` +
        `<strong>Elapsed:</strong> ${elapsedMs} ms</p>` +
        `<pre style="white-space:pre-wrap;">${escapeHtmlSafe_(raw)}</pre>` +
        `</div>`;
    } else {
      const parsed = JSON.parse(raw);
      const text = extractGeminiTextSafe_(parsed);

      let json;
      try {
        json = JSON.parse(text);
      } catch (e) {
        const match = String(text || '').match(/\{[\s\S]*\}/);
        if (!match) throw new Error('Could not parse Gemini JSON output');
        json = JSON.parse(match[0]);
      }

      subject = `Gemini Connection Test OK`;
      body =
        `Gemini connection test succeeded.\n\n` +
        `Time: ${nowStr}\n` +
        `Model: ${model}\n` +
        `Elapsed: ${elapsedMs} ms\n` +
        `Message: ${json.message || 'OK'}`;

      html =
        `<div style="font-family:Arial,sans-serif;max-width:800px;margin:0 auto;padding:20px;">` +
        `<h2 style="color:#188038;">Gemini Connection Test OK</h2>` +
        `<p><strong>Time:</strong> ${escapeHtmlSafe_(nowStr)}<br>` +
        `<strong>Model:</strong> ${escapeHtmlSafe_(model)}<br>` +
        `<strong>Elapsed:</strong> ${elapsedMs} ms<br>` +
        `<strong>Message:</strong> ${escapeHtmlSafe_(json.message || 'OK')}</p>` +
        `</div>`;
    }
  } catch (err) {
    const nowStr = Utilities.formatDate(new Date(), timezone, 'EEE, d MMM yyyy, h:mm a');
    subject = `Gemini Connection Test ERROR`;
    body =
      `Gemini connection test crashed.\n\n` +
      `Time: ${nowStr}\n` +
      `Model: ${model}\n` +
      `Error: ${err.message}`;

    html =
      `<div style="font-family:Arial,sans-serif;max-width:800px;margin:0 auto;padding:20px;">` +
      `<h2 style="color:#d93025;">Gemini Connection Test ERROR</h2>` +
      `<p><strong>Time:</strong> ${escapeHtmlSafe_(nowStr)}<br>` +
      `<strong>Model:</strong> ${escapeHtmlSafe_(model)}<br>` +
      `<strong>Error:</strong> ${escapeHtmlSafe_(err.message)}</p>` +
      `</div>`;
  }

  MailApp.sendEmail({
    to: digestTo,
    subject: subject,
    body: body,
    htmlBody: html
  });
}


function debug() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY');
  const digestTo = props.getProperty('DIGEST_TO');

  const timezone = (typeof CONFIG !== 'undefined' && CONFIG.TIMEZONE) ? CONFIG.TIMEZONE : 'Asia/Singapore';
  const query = (typeof CONFIG !== 'undefined' && CONFIG.GMAIL_QUERY) ? CONFIG.GMAIL_QUERY : 'in:inbox is:unread -in:chats';
  const processedLabelName = (typeof CONFIG !== 'undefined' && CONFIG.PROCESSED_LABEL) ? CONFIG.PROCESSED_LABEL : 'AI-DIGESTED';
  const maxThreads = (typeof CONFIG !== 'undefined' && CONFIG.MAX_THREADS_PER_RUN) ? CONFIG.MAX_THREADS_PER_RUN : '(not found)';
  const batchSize = (typeof CONFIG !== 'undefined' && CONFIG.BATCH_SIZE) ? CONFIG.BATCH_SIZE : '(not found)';
  const model = (typeof CONFIG !== 'undefined' && CONFIG.GEMINI_MODEL) ? CONFIG.GEMINI_MODEL : 'gemini-2.5-flash';
  const maxEmailSummaries = (typeof CONFIG !== 'undefined' && CONFIG.MAX_EMAIL_SUMMARIES) ? CONFIG.MAX_EMAIL_SUMMARIES : '(not found)';
  const maxTxtSummaries = (typeof CONFIG !== 'undefined' && CONFIG.MAX_TXT_SUMMARIES) ? CONFIG.MAX_TXT_SUMMARIES : '(not found)';
  const sleepMs = (typeof CONFIG !== 'undefined' && CONFIG.SLEEP_BETWEEN_BATCHES_MS) ? CONFIG.SLEEP_BETWEEN_BATCHES_MS : '(not found)';

  const now = new Date();
  const nowStr = Utilities.formatDate(now, timezone, 'EEE, d MMM yyyy, h:mm a');

  const processedLabel = GmailApp.getUserLabelByName(processedLabelName);
  const unlabeledQuery = `${query} -label:${processedLabelName}`;

  const unlabeledThreads = GmailApp.search(unlabeledQuery, 0, 100);
  const unreadAnyThreads = GmailApp.search(query, 0, 100);
  const processedThreads = processedLabel ? GmailApp.search(`label:${processedLabelName}`, 0, 100) : [];

  const totalSlots = (Number(maxEmailSummaries) || 0) + (Number(maxTxtSummaries) || 0);
  const scheduledWouldProcess = Math.min(Number(maxThreads) || 0, totalSlots || Number(maxThreads) || 0);
  const estimatedBatchCalls = batchSize && scheduledWouldProcess
    ? Math.ceil(scheduledWouldProcess / Number(batchSize))
    : '(unknown)';

  const triggers = ScriptApp.getProjectTriggers().map(function (t) {
    return {
      fn: t.getHandlerFunction(),
      type: String(t.getEventType())
    };
  });

  const domainCounts = {};
  unlabeledThreads.slice(0, 30).forEach(function (thread) {
    try {
      const messages = thread.getMessages();
      const msg = messages[messages.length - 1];
      const email = extractSenderEmailSafe_(msg.getFrom() || '');
      const domain = (email.split('@')[1] || '(unknown)').toLowerCase();
      domainCounts[domain] = (domainCounts[domain] || 0) + 1;
    } catch (err) {
      domainCounts['(parse-error)'] = (domainCounts['(parse-error)'] || 0) + 1;
    }
  });

  const domainSummary = Object.keys(domainCounts)
    .sort(function (a, b) { return domainCounts[b] - domainCounts[a]; })
    .slice(0, 10)
    .map(function (d) { return `${d}: ${domainCounts[d]}`; })
    .join('\n') || 'No domain sample data';

  const sampleLines = [];
  unlabeledThreads.slice(0, 8).forEach(function (thread, i) {
    try {
      const messages = thread.getMessages();
      const msg = messages[messages.length - 1];
      const fromRaw = msg.getFrom() || '(Unknown sender)';
      const senderEmail = extractSenderEmailSafe_(fromRaw);
      const subject = msg.getSubject() || '(No subject)';
      const time = Utilities.formatDate(msg.getDate(), timezone, 'EEE, d MMM yyyy, h:mm a');

      sampleLines.push(
        `#${i + 1}\n` +
        `Subject: ${subject}\n` +
        `Sender: ${fromRaw}\n` +
        `Sender Email: ${senderEmail}\n` +
        `Time: ${time}`
      );
    } catch (err) {
      sampleLines.push(`#${i + 1}\nCould not inspect thread: ${err.message}`);
    }
  });

  const triggerText = triggers.length
    ? triggers.map(function (t, i) {
        return `${i + 1}. ${t.fn} (${t.type})`;
      }).join('\n')
    : 'No triggers found';

  const body =
    `AI Email Debug Report\n\n` +
    `Time: ${nowStr}\n` +
    `Timezone: ${timezone}\n` +
    `Gemini API Key Present: ${apiKey ? 'Yes' : 'No'}\n` +
    `DIGEST_TO Present: ${digestTo ? 'Yes' : 'No'}\n` +
    `DIGEST_TO Value: ${digestTo || '(missing)'}\n` +
    `Gemini Model: ${model}\n` +
    `Gmail Query: ${query}\n` +
    `Unlabeled Query: ${unlabeledQuery}\n` +
    `Processed Label Name: ${processedLabelName}\n` +
    `Processed Label Exists: ${processedLabel ? 'Yes' : 'No'}\n` +
    `MAX_THREADS_PER_RUN: ${maxThreads}\n` +
    `MAX_EMAIL_SUMMARIES: ${maxEmailSummaries}\n` +
    `MAX_TXT_SUMMARIES: ${maxTxtSummaries}\n` +
    `BATCH_SIZE: ${batchSize}\n` +
    `SLEEP_BETWEEN_BATCHES_MS: ${sleepMs}\n` +
    `Unread Any Threads Sample Count (up to 100 checked): ${unreadAnyThreads.length}\n` +
    `Unread Unprocessed Threads Sample Count (up to 100 checked): ${unlabeledThreads.length}\n` +
    `Processed Label Sample Count (up to 100 checked): ${processedThreads.length}\n` +
    `Would Process Per Scheduled Run: ${scheduledWouldProcess}\n` +
    `Estimated Gemini Batch Calls Per Full Run: ${estimatedBatchCalls}\n\n` +
    `Top Sender Domains in Unprocessed Sample:\n${domainSummary}\n\n` +
    `Triggers:\n${triggerText}\n\n` +
    `Sample Unprocessed Threads:\n${sampleLines.join('\n\n') || 'No unprocessed sample threads found.'}`;

  const html =
    `<div style="font-family:Arial,sans-serif;max-width:980px;margin:0 auto;padding:20px;">` +
    `<h1>AI Email Debug Report</h1>` +
    `<p><strong>Time:</strong> ${escapeHtmlSafe_(nowStr)}<br>` +
    `<strong>Timezone:</strong> ${escapeHtmlSafe_(timezone)}<br>` +
    `<strong>Gemini API Key Present:</strong> ${apiKey ? 'Yes' : 'No'}<br>` +
    `<strong>DIGEST_TO Present:</strong> ${digestTo ? 'Yes' : 'No'}<br>` +
    `<strong>DIGEST_TO Value:</strong> ${escapeHtmlSafe_(digestTo || '(missing)')}<br>` +
    `<strong>Gemini Model:</strong> ${escapeHtmlSafe_(model)}<br>` +
    `<strong>Gmail Query:</strong> ${escapeHtmlSafe_(query)}<br>` +
    `<strong>Unlabeled Query:</strong> ${escapeHtmlSafe_(unlabeledQuery)}<br>` +
    `<strong>Processed Label Name:</strong> ${escapeHtmlSafe_(processedLabelName)}<br>` +
    `<strong>Processed Label Exists:</strong> ${processedLabel ? 'Yes' : 'No'}<br>` +
    `<strong>MAX_THREADS_PER_RUN:</strong> ${escapeHtmlSafe_(String(maxThreads))}<br>` +
    `<strong>MAX_EMAIL_SUMMARIES:</strong> ${escapeHtmlSafe_(String(maxEmailSummaries))}<br>` +
    `<strong>MAX_TXT_SUMMARIES:</strong> ${escapeHtmlSafe_(String(maxTxtSummaries))}<br>` +
    `<strong>BATCH_SIZE:</strong> ${escapeHtmlSafe_(String(batchSize))}<br>` +
    `<strong>SLEEP_BETWEEN_BATCHES_MS:</strong> ${escapeHtmlSafe_(String(sleepMs))}<br>` +
    `<strong>Unread Any Threads Sample Count:</strong> ${escapeHtmlSafe_(String(unreadAnyThreads.length))}<br>` +
    `<strong>Unread Unprocessed Threads Sample Count:</strong> ${escapeHtmlSafe_(String(unlabeledThreads.length))}<br>` +
    `<strong>Processed Label Sample Count:</strong> ${escapeHtmlSafe_(String(processedThreads.length))}<br>` +
    `<strong>Would Process Per Scheduled Run:</strong> ${escapeHtmlSafe_(String(scheduledWouldProcess))}<br>` +
    `<strong>Estimated Gemini Batch Calls Per Full Run:</strong> ${escapeHtmlSafe_(String(estimatedBatchCalls))}</p>` +
    `<h2>Top Sender Domains in Unprocessed Sample</h2>` +
    `<pre style="white-space:pre-wrap;">${escapeHtmlSafe_(domainSummary)}</pre>` +
    `<h2>Triggers</h2>` +
    `<pre style="white-space:pre-wrap;">${escapeHtmlSafe_(triggerText)}</pre>` +
    `<h2>Sample Unprocessed Threads</h2>` +
    `<pre style="white-space:pre-wrap;">${escapeHtmlSafe_(sampleLines.join('\n\n') || 'No unprocessed sample threads found.')}</pre>` +
    `</div>`;

  if (!digestTo) {
    throw new Error('Missing Script Property: DIGEST_TO');
  }

  MailApp.sendEmail({
    to: digestTo,
    subject: 'AI Email Debug Report',
    body: body,
    htmlBody: html
  });
}


function createMidnightGeminiTestTrigger() {
  const timezone = (typeof CONFIG !== 'undefined' && CONFIG.TIMEZONE) ? CONFIG.TIMEZONE : 'Asia/Singapore';
  const triggers = ScriptApp.getProjectTriggers();

  for (const t of triggers) {
    if (t.getHandlerFunction() === 'testGeminiConnection') {
      ScriptApp.deleteTrigger(t);
    }
  }

  ScriptApp.newTrigger('testGeminiConnection')
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .nearMinute(0)
    .inTimezone(timezone)
    .create();
}


/* SAFE HELPERS */

function extractGeminiTextSafe_(resp) {
  const candidates = resp && resp.candidates;
  if (!Array.isArray(candidates) || !candidates.length) return '';

  const parts = candidates[0] &&
                candidates[0].content &&
                candidates[0].content.parts;

  if (!Array.isArray(parts)) return '';

  return parts.map(function (p) {
    return p.text || '';
  }).join('\n').trim();
}

function extractSenderEmailSafe_(fromRaw) {
  const angle = String(fromRaw || '').match(/<([^>]+)>/);
  if (angle) return angle[1].trim();

  const bare = String(fromRaw || '').match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return bare ? bare[0].trim() : '';
}

function escapeHtmlSafe_(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
