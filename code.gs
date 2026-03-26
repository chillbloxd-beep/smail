const CONFIG = {
  TIMEZONE: 'Asia/Singapore',

  // Processing limits
  MAX_EMAIL_SUMMARIES: 50,      // max items shown in email body
  MAX_TXT_SUMMARIES: 50,        // max items shown in txt attachment
  MAX_THREADS_PER_RUN: 100,     // total max actually processed per run
  BATCH_SIZE: 20,               // 20 emails per Gemini call
  SLEEP_BETWEEN_BATCHES_MS: 15000,

  // Gmail
  GMAIL_QUERY: 'in:inbox is:unread -in:chats',
  PROCESSED_LABEL: 'AI-DIGESTED',

  // Gemini
  GEMINI_MODEL: 'gemini-2.5-flash',
  MAX_BODY_CHARS_PER_EMAIL: 3500,

  // Email
  SUBJECT_PREFIX: 'AI Email Digest',
  SEND_EMPTY_DIGEST: true
};


/**
 * ===== SETUP =====
 */
function setup() {
  getOrCreateProcessedLabel_();
}


/**
 * ===== TRIGGERS =====
 * Run once.
 */
function createDigestTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  for (const t of triggers) {
    const fn = t.getHandlerFunction();
    if (
      fn === 'scheduledMorningDigest' ||
      fn === 'scheduledEveningDigest'
    ) {
      ScriptApp.deleteTrigger(t);
    }
  }

  ScriptApp.newTrigger('scheduledMorningDigest')
    .timeBased()
    .everyDays(1)
    .atHour(5)
    .nearMinute(45)
    .inTimezone(CONFIG.TIMEZONE)
    .create();

  ScriptApp.newTrigger('scheduledEveningDigest')
    .timeBased()
    .everyDays(1)
    .atHour(17)
    .nearMinute(45)
    .inTimezone(CONFIG.TIMEZONE)
    .create();
}


function scheduledMorningDigest() {
  runDigest_({
    mode: 'scheduled',
    bucket: 'AM',
    dryRun: false
  });
}


function scheduledEveningDigest() {
  runDigest_({
    mode: 'scheduled',
    bucket: 'PM',
    dryRun: false
  });
}


/**
 * Preview only. Safe.
 */
function manualPreviewDigestNow() {
  runDigest_({
    mode: 'manual-preview',
    bucket: 'MANUAL',
    dryRun: true
  });
}


/**
 * Real manual processing.
 */
function manualProcessDigestNow() {
  runDigest_({
    mode: 'manual-process',
    bucket: 'MANUAL',
    dryRun: false
  });
}


/**
 * ===== MAIN =====
 */
function runDigest_(options) {
  const opts = options || {};
  const dryRun = Boolean(opts.dryRun);
  const bucket = opts.bucket || 'MANUAL';
  const mode = opts.mode || 'manual';

  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY');
  const digestTo = props.getProperty('DIGEST_TO');

  if (!apiKey) throw new Error('Missing Script Property: GEMINI_API_KEY');
  if (!digestTo) throw new Error('Missing Script Property: DIGEST_TO');

  const processedLabel = getOrCreateProcessedLabel_();

  let query = CONFIG.GMAIL_QUERY;
  if (!dryRun) {
    query += ` -label:${CONFIG.PROCESSED_LABEL}`;
  }

  const totalSlots = CONFIG.MAX_EMAIL_SUMMARIES + CONFIG.MAX_TXT_SUMMARIES;
  const maxToProcess = Math.min(CONFIG.MAX_THREADS_PER_RUN, totalSlots);

  const threads = GmailApp.search(query, 0, maxToProcess);
  const totalUnreadUnprocessed = dryRun
    ? GmailApp.search(CONFIG.GMAIL_QUERY, 0, 500).length
    : GmailApp.search(`${CONFIG.GMAIL_QUERY} -label:${CONFIG.PROCESSED_LABEL}`, 0, 500).length;

  if (!threads.length) {
    if (CONFIG.SEND_EMPTY_DIGEST) {
      sendEmptyDigest_(digestTo, mode, bucket, dryRun, totalUnreadUnprocessed);
    }
    return;
  }

  const emailItems = [];
  const threadRefs = [];

  for (const thread of threads) {
    try {
      const messages = thread.getMessages();
      const msg = messages[messages.length - 1];

      const fromRaw = safeText_(msg.getFrom(), '(Unknown sender)');
      const senderEmail = extractSenderEmail_(fromRaw);
      const senderName = extractSenderName_(fromRaw);
      const subject = safeText_(msg.getSubject(), '(No subject)');
      const dateObj = msg.getDate();
      const timeStr = Utilities.formatDate(dateObj, CONFIG.TIMEZONE, 'EEE, d MMM yyyy, h:mm a');

      const bodyPlain = msg.getPlainBody() || stripHtml_(msg.getBody() || '');
      const cleanedBody = cleanEmailBody_(bodyPlain).slice(0, CONFIG.MAX_BODY_CHARS_PER_EMAIL);

      emailItems.push({
        id: String(emailItems.length + 1),
        subject: subject,
        senderName: senderName,
        senderEmail: senderEmail,
        time: timeStr,
        body: cleanedBody
      });

      threadRefs.push({
        thread: thread,
        senderEmail: senderEmail
      });
    } catch (err) {
      emailItems.push({
        id: String(emailItems.length + 1),
        subject: '(Processing failed before AI)',
        senderName: '(Unknown)',
        senderEmail: '',
        time: Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'EEE, d MMM yyyy, h:mm a'),
        body: `Could not read this email: ${err.message}`
      });

      threadRefs.push({
        thread: thread,
        senderEmail: ''
      });
    }
  }

  const batches = chunkArray_(emailItems, CONFIG.BATCH_SIZE);
  let aiResults = [];

  for (let i = 0; i < batches.length; i++) {
    const batch = batches[i];
    const batchResults = summarizeBatchWithGemini_(batch, apiKey);
    aiResults = aiResults.concat(batchResults);

    if (i < batches.length - 1) {
      Utilities.sleep(CONFIG.SLEEP_BETWEEN_BATCHES_MS);
    }
  }

  const merged = mergeResultsWithOriginal_(emailItems, aiResults);

  const emailDigestItems = merged.slice(0, CONFIG.MAX_EMAIL_SUMMARIES);
  const txtDigestItems = merged.slice(CONFIG.MAX_EMAIL_SUMMARIES, CONFIG.MAX_EMAIL_SUMMARIES + CONFIG.MAX_TXT_SUMMARIES);

  if (!dryRun) {
    for (let i = 0; i < threadRefs.length; i++) {
      try {
        threadRefs[i].thread.addLabel(processedLabel);
      } catch (err) {
        // ignore single-thread label failure
      }
    }
  }

  const groupedEmail = groupResults_(emailDigestItems);
  const txtContent = buildTxtAttachment_(txtDigestItems, mode, bucket, dryRun);
  const remainingCount = Math.max(0, totalUnreadUnprocessed - merged.length);

  sendDigestEmail_({
    groupedEmail: groupedEmail,
    txtItems: txtDigestItems,
    txtContent: txtContent,
    digestTo: digestTo,
    mode: mode,
    bucket: bucket,
    dryRun: dryRun,
    totalProcessed: merged.length,
    totalInEmail: emailDigestItems.length,
    totalInTxt: txtDigestItems.length,
    remainingCount: remainingCount
  });
}


/**
 * ===== GEMINI =====
 */
function summarizeBatchWithGemini_(emails, apiKey) {
  const prompt = `
You are summarizing inbox emails for one user.

You will receive an array of emails.
For EACH email, return:
- id
- summary (2 to 5 short lines)
- spam (true/false)
- important (true/false)
- category ("normal" or "homework")
- reason (very short reason)

Rules:
1. Be accurate. Do not invent facts.
2. Summary must mention the main content clearly.
3. Flag obvious scams, junk, fake urgency, suspicious offers, suspicious links, mass promos, or irrelevant ad-like emails as spam=true.
4. Mark important=true if the user likely needs to notice or act on it.
5. If the sender is related to school/classwork, mention deadlines or tasks if present.
6. category should be:
   - "homework" when the email is clearly assignment/classwork related
   - otherwise "normal"

Return ONLY valid JSON in this exact shape:
{
  "results": [
    {
      "id": "1",
      "summary": "Line 1\\nLine 2\\nLine 3",
      "spam": false,
      "important": true,
      "category": "normal",
      "reason": "action needed"
    }
  ]
}

Emails JSON:
${JSON.stringify(emails)}
  `.trim();

  const url =
    'https://generativelanguage.googleapis.com/v1beta/models/' +
    encodeURIComponent(CONFIG.GEMINI_MODEL) +
    ':generateContent?key=' +
    encodeURIComponent(apiKey);

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.2,
      responseMimeType: 'application/json'
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const raw = response.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error(`Gemini API error ${code}: ${raw}`);
  }

  const parsed = JSON.parse(raw);
  const text = extractGeminiText_(parsed);

  if (!text) {
    throw new Error('Gemini returned empty output');
  }

  let json;
  try {
    json = JSON.parse(text);
  } catch (e) {
    const match = text.match(/\{[\s\S]*\}/);
    if (!match) throw new Error('Could not parse Gemini JSON output');
    json = JSON.parse(match[0]);
  }

  if (!json || !Array.isArray(json.results)) {
    throw new Error('Gemini output missing results array');
  }

  return json.results.map(function (r) {
    return {
      id: String(r.id || ''),
      summary: normalizeSummaryLines_(String(r.summary || 'No summary available.')),
      spam: Boolean(r.spam),
      important: Boolean(r.important),
      category: String(r.category || 'normal'),
      reason: String(r.reason || '')
    };
  });
}


/**
 * ===== POST-PROCESSING =====
 */
function mergeResultsWithOriginal_(originalEmails, aiResults) {
  const aiMap = {};
  for (const r of aiResults) {
    aiMap[String(r.id)] = r;
  }

  return originalEmails.map(function (email) {
    const ai = aiMap[email.id] || {
      summary: 'No summary available.',
      spam: false,
      important: false,
      category: 'normal',
      reason: 'No AI result'
    };

    const senderEmail = (email.senderEmail || '').toLowerCase();
    const forced = applyForcedRules_(senderEmail, ai);

    return {
      id: email.id,
      subject: email.subject,
      senderName: email.senderName,
      senderEmail: email.senderEmail,
      time: email.time,
      summary: forced.summary || ai.summary,
      spam: forced.spam,
      important: forced.important,
      category: forced.category,
      reason: forced.reason
    };
  });
}


function applyForcedRules_(senderEmail, ai) {
  const out = {
    summary: ai.summary,
    spam: ai.spam,
    important: ai.important,
    category: ai.category,
    reason: ai.reason
  };

  if (endsWithDomain_(senderEmail, 'classroom.google.com')) {
    out.category = 'homework';
    out.important = true;
    out.reason = appendReason_(out.reason, 'forced homework by classroom.google.com');
  }

  if (
    endsWithDomain_(senderEmail, 'students.edu.sg') ||
    endsWithDomain_(senderEmail, 'moe.edu.sg') ||
    endsWithDomain_(senderEmail, 'gmail.com')
  ) {
    out.important = true;
    out.reason = appendReason_(out.reason, 'forced important by sender domain');
  }

  return out;
}


function groupResults_(items) {
  const important = [];
  const homework = [];
  const normal = [];
  const spam = [];

  for (const item of items) {
    if (item.spam) {
      spam.push(item);
    } else if (item.category === 'homework') {
      homework.push(item);
    } else if (item.important) {
      important.push(item);
    } else {
      normal.push(item);
    }
  }

  return {
    important: important,
    homework: homework,
    normal: normal,
    spam: spam
  };
}


/**
 * ===== EMAIL + TXT OUTPUT =====
 */
function sendDigestEmail_(args) {
  const groupedEmail = args.groupedEmail;
  const txtItems = args.txtItems;
  const txtContent = args.txtContent;
  const digestTo = args.digestTo;
  const mode = args.mode;
  const bucket = args.bucket;
  const dryRun = args.dryRun;
  const totalProcessed = args.totalProcessed;
  const totalInEmail = args.totalInEmail;
  const totalInTxt = args.totalInTxt;
  const remainingCount = args.remainingCount;

  const emailBodyCount =
    groupedEmail.important.length +
    groupedEmail.homework.length +
    groupedEmail.normal.length +
    groupedEmail.spam.length;

  const tag = dryRun ? '[MANUAL PREVIEW] ' : '';
  const subject =
    `${tag}${CONFIG.SUBJECT_PREFIX} — ${bucket} — Email:${totalInEmail} TXT:${totalInTxt} Remaining:${remainingCount}`;

  const nowStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'EEE, d MMM yyyy, h:mm a');

  const html = `
    <div style="font-family:Arial,sans-serif;max-width:900px;margin:0 auto;padding:20px;">
      <h1 style="margin-bottom:8px;">${escapeHtml_(tag + 'AI Email Digest')}</h1>
      <div style="font-size:14px;color:#666;margin-bottom:16px;">
        Generated: ${escapeHtml_(nowStr)}<br>
        Mode: ${escapeHtml_(mode)}<br>
        Bucket: ${escapeHtml_(bucket)}<br>
        ${dryRun ? '<strong>This was a preview only. No labels were changed.</strong><br>' : ''}
        <strong>Processed this run:</strong> ${totalProcessed}<br>
        <strong>Shown in email:</strong> ${totalInEmail}<br>
        <strong>Shown in TXT:</strong> ${totalInTxt}<br>
        <strong>Still left for later:</strong> ${remainingCount}
      </div>

      <div style="font-size:14px;margin-bottom:20px;">
        <strong>⚠ Important:</strong> ${groupedEmail.important.length} &nbsp;|&nbsp;
        <strong>📚 Homework:</strong> ${groupedEmail.homework.length} &nbsp;|&nbsp;
        <strong>Normal:</strong> ${groupedEmail.normal.length} &nbsp;|&nbsp;
        <strong>Likely Spam:</strong> ${groupedEmail.spam.length}
      </div>

      ${txtItems.length ? `<div style="background:#f8f9fa;border:1px solid #ddd;padding:12px;border-radius:10px;margin-bottom:20px;">
        <strong>TXT attachment included:</strong> ${txtItems.length} more summarized email(s)
      </div>` : ''}

      ${buildSectionHtml_('⚠ Important', groupedEmail.important, '#d93025')}
      ${buildSectionHtml_('📚 Homework', groupedEmail.homework, '#f29900')}
      ${buildSectionHtml_('Normal', groupedEmail.normal, '#1a73e8')}
      ${buildSectionHtml_('Likely Spam', groupedEmail.spam, '#5f6368')}
    </div>
  `;

  const text = buildDigestText_(groupedEmail, mode, bucket, dryRun, totalProcessed, totalInEmail, totalInTxt, remainingCount);

  const mailOptions = {
    to: digestTo,
    subject: subject,
    body: text,
    htmlBody: html
  };

  if (txtItems.length) {
    const filename = `unsummarized_overflow_${bucket.toLowerCase()}_${Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd_HHmm')}.txt`;
    mailOptions.attachments = [
      Utilities.newBlob(txtContent, 'text/plain', filename)
    ];
  }

  MailApp.sendEmail(mailOptions);
}


function sendEmptyDigest_(digestTo, mode, bucket, dryRun, totalUnreadUnprocessed) {
  const subject = `${CONFIG.SUBJECT_PREFIX} — ${bucket} — No unread emails`;
  const body =
    `No unread emails matched your filter.\n` +
    `Mode: ${mode}\n` +
    `Bucket: ${bucket}\n` +
    `Preview only: ${dryRun ? 'Yes' : 'No'}\n` +
    `Unread checked pool estimate: ${totalUnreadUnprocessed}`;

  const html = `
    <div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;padding:20px;">
      <h2>No unread emails</h2>
      <p>No unread emails matched your filter.</p>
      <p><strong>Mode:</strong> ${escapeHtml_(mode)}<br>
      <strong>Bucket:</strong> ${escapeHtml_(bucket)}<br>
      <strong>Preview only:</strong> ${dryRun ? 'Yes' : 'No'}<br>
      <strong>Unread checked pool estimate:</strong> ${totalUnreadUnprocessed}</p>
    </div>
  `;

  MailApp.sendEmail({
    to: digestTo,
    subject: subject,
    body: body,
    htmlBody: html
  });
}


function buildTxtAttachment_(items, mode, bucket, dryRun) {
  const nowStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'EEE, d MMM yyyy, h:mm a');

  const lines = [];
  lines.push('AI EMAIL DIGEST OVERFLOW TXT');
  lines.push('============================');
  lines.push(`Generated: ${nowStr}`);
  lines.push(`Mode: ${mode}`);
  lines.push(`Bucket: ${bucket}`);
  lines.push(`Preview only: ${dryRun ? 'Yes' : 'No'}`);
  lines.push(`Items in this TXT: ${items.length}`);
  lines.push('');

  if (!items.length) {
    lines.push('No overflow items.');
    return lines.join('\n');
  }

  items.forEach(function (e, i) {
    lines.push(`EMAIL ${i + 1}`);
    lines.push('----------------------------------------');
    lines.push(`Subject: ${e.subject}`);
    lines.push(`Sender: ${e.senderName}${e.senderEmail ? ` <${e.senderEmail}>` : ''}`);
    lines.push(`Time: ${e.time}`);
    lines.push(`Important: ${e.important ? 'Yes' : 'No'}`);
    lines.push(`Homework: ${e.category === 'homework' ? 'Yes' : 'No'}`);
    lines.push(`Likely Spam: ${e.spam ? 'Yes' : 'No'}`);
    if (e.reason) lines.push(`AI Note: ${e.reason}`);
    lines.push('Summary:');
    lines.push(e.summary);
    lines.push('');
  });

  return lines.join('\n');
}


function buildSectionHtml_(title, items, color) {
  if (!items.length) {
    return `
      <div style="margin-bottom:28px;">
        <h2 style="color:${color};margin-bottom:8px;">${escapeHtml_(title)} (0)</h2>
        <div style="font-size:14px;color:#777;">No emails in this section.</div>
      </div>
    `;
  }

  const blocks = items.map(function (e, i) {
    const badges = [];
    if (title === '⚠ Important' || e.important) {
      badges.push('<span style="display:inline-block;background:#d93025;color:#fff;padding:4px 8px;border-radius:999px;font-size:12px;font-weight:700;margin-right:8px;">⚠ IMPORTANT</span>');
    }
    if (e.category === 'homework') {
      badges.push('<span style="display:inline-block;background:#f29900;color:#fff;padding:4px 8px;border-radius:999px;font-size:12px;font-weight:700;margin-right:8px;">📚 HOMEWORK</span>');
    }
    if (e.spam) {
      badges.push('<span style="display:inline-block;background:#5f6368;color:#fff;padding:4px 8px;border-radius:999px;font-size:12px;font-weight:700;">LIKELY SPAM</span>');
    }

    return `
      <div style="border:1px solid #e0e0e0;border-radius:14px;padding:16px;margin-bottom:14px;">
        <div style="font-size:12px;color:#777;margin-bottom:8px;">${escapeHtml_(title)} #${i + 1}</div>
        <div style="font-size:18px;font-weight:700;margin-bottom:8px;">${escapeHtml_(e.subject)}</div>
        <div style="margin-bottom:10px;">${badges.join('')}</div>
        <div style="font-size:14px;line-height:1.6;color:#333;margin-bottom:10px;">
          <strong>Sender:</strong> ${escapeHtml_(e.senderName)}${e.senderEmail ? ` &lt;${escapeHtml_(e.senderEmail)}&gt;` : ''}<br>
          <strong>Time:</strong> ${escapeHtml_(e.time)}
        </div>
        <div style="font-size:14px;line-height:1.7;white-space:pre-wrap;margin-bottom:8px;">${escapeHtml_(e.summary)}</div>
        ${e.reason ? `<div style="font-size:13px;color:#666;"><strong>AI note:</strong> ${escapeHtml_(e.reason)}</div>` : ''}
      </div>
    `;
  }).join('');

  return `
    <div style="margin-bottom:28px;">
      <h2 style="color:${color};margin-bottom:12px;">${escapeHtml_(title)} (${items.length})</h2>
      ${blocks}
    </div>
  `;
}


function buildDigestText_(grouped, mode, bucket, dryRun, totalProcessed, totalInEmail, totalInTxt, remainingCount) {
  const lines = [];
  lines.push(`${dryRun ? '[MANUAL PREVIEW] ' : ''}AI Email Digest`);
  lines.push(`Mode: ${mode}`);
  lines.push(`Bucket: ${bucket}`);
  if (dryRun) lines.push('Preview only. No labels were changed.');
  lines.push(`Processed this run: ${totalProcessed}`);
  lines.push(`Shown in email: ${totalInEmail}`);
  lines.push(`Shown in TXT: ${totalInTxt}`);
  lines.push(`Still left for later: ${remainingCount}`);
  lines.push('');
  lines.push(`Important: ${grouped.important.length}`);
  lines.push(`Homework: ${grouped.homework.length}`);
  lines.push(`Normal: ${grouped.normal.length}`);
  lines.push(`Likely Spam: ${grouped.spam.length}`);
  lines.push('');

  appendTextSection_(lines, '⚠ Important', grouped.important);
  appendTextSection_(lines, '📚 Homework', grouped.homework);
  appendTextSection_(lines, 'Normal', grouped.normal);
  appendTextSection_(lines, 'Likely Spam', grouped.spam);

  return lines.join('\n');
}


function appendTextSection_(lines, title, items) {
  lines.push(`${title} (${items.length})`);
  lines.push('');

  if (!items.length) {
    lines.push('No emails in this section.');
    lines.push('');
    return;
  }

  items.forEach(function (e, i) {
    lines.push(`${title} #${i + 1}`);
    lines.push(`Subject: ${e.subject}`);
    lines.push(`Sender: ${e.senderName}${e.senderEmail ? ` <${e.senderEmail}>` : ''}`);
    lines.push(`Time: ${e.time}`);
    lines.push(`Summary:\n${e.summary}`);
    if (e.reason) lines.push(`AI note: ${e.reason}`);
    lines.push('--------------------------------------------------');
    lines.push('');
  });
}


/**
 * ===== HELPERS =====
 */
function getOrCreateProcessedLabel_() {
  return GmailApp.getUserLabelByName(CONFIG.PROCESSED_LABEL) || GmailApp.createLabel(CONFIG.PROCESSED_LABEL);
}


function chunkArray_(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) {
    out.push(arr.slice(i, i + size));
  }
  return out;
}


function extractGeminiText_(resp) {
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


function cleanEmailBody_(text) {
  if (!text) return '';

  return String(text)
    .replace(/\r/g, '')
    .replace(/On .*wrote:\n[\s\S]*$/i, '')
    .replace(/From:.*\nSent:.*\nTo:.*\nSubject:.*\n[\s\S]*$/i, '')
    .replace(/[-_]{5,}[\s\S]*$/g, '')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]{2,}/g, ' ')
    .trim();
}


function stripHtml_(html) {
  return String(html || '')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/\s{2,}/g, ' ')
    .trim();
}


function normalizeSummaryLines_(text) {
  return String(text || '')
    .split('\n')
    .map(function (s) { return s.trim(); })
    .filter(Boolean)
    .slice(0, 5)
    .join('\n');
}


function extractSenderName_(fromRaw) {
  const angle = String(fromRaw).match(/^(.*)<.+>$/);
  if (angle) return angle[1].replace(/"/g, '').trim();
  return String(fromRaw).trim();
}


function extractSenderEmail_(fromRaw) {
  const angle = String(fromRaw).match(/<([^>]+)>/);
  if (angle) return angle[1].trim();

  const bare = String(fromRaw).match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return bare ? bare[0].trim() : '';
}


function safeText_(value, fallback) {
  return value ? String(value).trim() : fallback;
}


function endsWithDomain_(email, domain) {
  const e = String(email || '').toLowerCase();
  return e.endsWith('@' + domain.toLowerCase());
}


function appendReason_(base, extra) {
  if (!base) return extra;
  if (!extra) return base;
  return base + '; ' + extra;
}


function escapeHtml_(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
