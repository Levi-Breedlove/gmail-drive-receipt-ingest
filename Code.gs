/**
 * Receipt ingestion MVP for Gmail -> Google Drive
 *
 * What this version does:
 * - Uses Advanced Gmail service for richer MIME/body access
 * - Saves a PDF converted from a self-contained HTML snapshot of the email body
 * - Saves regular attachments separately
 * - Embeds cid:inline images into the HTML snapshot
 * - Best-effort fetches remote images used by <img src="..."> and CSS url(...)
 * - Skips remote image failures and keeps moving
 * - Limits ingestion to March 2026 and onward by default
 * - Prevents duplicate crawling with both Gmail message ID tracking and receipt-level dedupe
 * - Ignores obvious non-receipt threads even if they were labeled by mistake
 * - Sends borderline messages to review instead of ingesting them
 *
 * Important:
 * Gmail labels are thread-level. Because of that, this script does NOT exclude
 * processed/review/ignored labels from search. It always scans candidate expense
 * threads, then decides at the message level what to do. That avoids hiding a
 * later real receipt inside a thread that was previously reviewed or ignored.
 *
 * Required services:
 * - Built-in: GmailApp, DriveApp, PropertiesService, LockService, ScriptApp, UrlFetchApp
 * - Advanced Gmail service: Gmail
 */

/* =========================
 * SETUP / CONFIG
 * ========================= */

function setupConfig() {
  const props = PropertiesService.getScriptProperties();

  props.setProperties(
    {
      ROOT_FOLDER_NAME: 'Receipts Archive',
      TEST_SUBROOT_NAME: '_TEST',
      PROCESSED_LABEL: 'receipt-ingested-test',
      REVIEW_LABEL: 'receipt-needs-review-test',
      IGNORED_LABEL: 'receipt-ignored-nonreceipt-test',
      LABELS: 'archana-expenses,dan-expenses,rhea-expenses',
      TIMEZONE: 'America/Los_Angeles',
      DRY_RUN: 'true',
      MIN_EMAIL_DATE: '2026-03-01',

      // Classification thresholds
      RECEIPT_THRESHOLD: '4',
      REVIEW_THRESHOLD: '2',

      // Artifact controls
      SAVE_HTML: 'false',
      SAVE_PDF: 'true',
      SAVE_MANIFEST: 'false',
      SAVE_ATTACHMENTS: 'true',
      SAVE_INLINE_IMAGE_FILES: 'false',

      // Fetch remote images for better PDF fidelity, best-effort only.
      FETCH_REMOTE_IMAGES: 'true',

      // New: do not save .eml attachments
      IGNORE_EML_ATTACHMENTS: 'true'
    },
    true
  );

  Logger.log('Config saved to Script Properties.');
}

function getConfig_() {
  const props = PropertiesService.getScriptProperties();

  return {
    rootFolderName: props.getProperty('ROOT_FOLDER_NAME') || 'Receipts Archive',
    testSubrootName: props.getProperty('TEST_SUBROOT_NAME') || '_TEST',
    processedLabel: props.getProperty('PROCESSED_LABEL') || 'receipt-ingested-test',
    reviewLabel: props.getProperty('REVIEW_LABEL') || 'receipt-needs-review-test',
    ignoredLabel: props.getProperty('IGNORED_LABEL') || 'receipt-ignored-nonreceipt-test',
    labels: (props.getProperty('LABELS') || '')
      .split(',')
      .map(function (s) { return s.trim(); })
      .filter(Boolean),
    timezone: props.getProperty('TIMEZONE') || 'America/Los_Angeles',
    dryRun: String(props.getProperty('DRY_RUN')).toLowerCase() === 'true',
    minEmailDate: props.getProperty('MIN_EMAIL_DATE') || '2026-03-01',
    receiptThreshold: Number(props.getProperty('RECEIPT_THRESHOLD') || '4'),
    reviewThreshold: Number(props.getProperty('REVIEW_THRESHOLD') || '2'),

    saveHtml: String(props.getProperty('SAVE_HTML')).toLowerCase() === 'true',
    savePdf: String(props.getProperty('SAVE_PDF')).toLowerCase() === 'true',
    saveManifest: String(props.getProperty('SAVE_MANIFEST')).toLowerCase() === 'true',
    saveAttachments: String(props.getProperty('SAVE_ATTACHMENTS')).toLowerCase() === 'true',
    saveInlineImageFiles:
      String(props.getProperty('SAVE_INLINE_IMAGE_FILES')).toLowerCase() === 'true',
    fetchRemoteImages:
      String(props.getProperty('FETCH_REMOTE_IMAGES')).toLowerCase() === 'true'
  };
}

/* =========================
 * DATE FILTER HELPERS
 * ========================= */

function getAfterQueryDate_(minEmailDate) {
  // Gmail "after:" is exclusive. To include the configured date,
  // search after the previous day.
  const minDate = new Date(minEmailDate + 'T00:00:00');
  const previousDay = new Date(minDate.getTime() - 24 * 60 * 60 * 1000);

  return Utilities.formatDate(previousDay, Session.getScriptTimeZone(), 'yyyy/MM/dd');
}

function isMessageOnOrAfterMinDate_(message, minEmailDate) {
  const boundary = new Date(minEmailDate + 'T00:00:00');
  return message.getDate().getTime() >= boundary.getTime();
}

function setDryRunTrue() {
  PropertiesService.getScriptProperties().setProperty('DRY_RUN', 'true');
  Logger.log('DRY_RUN set to true');
}

function setDryRunFalse() {
  PropertiesService.getScriptProperties().setProperty('DRY_RUN', 'false');
  Logger.log('DRY_RUN set to false');
}

/* =========================
 * VERIFY / LABEL HELPERS
 * ========================= */

function verifySetup() {
  const config = getConfig_();

  Logger.log('--- VERIFY SETUP START ---');

  assertAdvancedGmailEnabled_();
  Logger.log('Advanced Gmail service OK');

  config.labels.forEach(function (labelName) {
    const label = GmailApp.getUserLabelByName(labelName);
    Logger.log(label ? 'Label OK: ' + labelName : 'Label MISSING: ' + labelName);
  });

  [config.processedLabel, config.reviewLabel, config.ignoredLabel].forEach(function (name) {
    const label = GmailApp.getUserLabelByName(name);
    Logger.log(label ? 'Label OK: ' + name : 'Label MISSING: ' + name);
  });

  const root = findFolderByName_(config.rootFolderName);
  if (!root) {
    Logger.log('Folder MISSING: ' + config.rootFolderName);
    Logger.log('--- VERIFY SETUP END ---');
    return;
  }
  Logger.log('Folder OK: ' + config.rootFolderName);

  const testRoot = findChildFolderByName_(root, config.testSubrootName);
  if (!testRoot) {
    Logger.log('Folder MISSING: ' + config.rootFolderName + '/' + config.testSubrootName);
    Logger.log('--- VERIFY SETUP END ---');
    return;
  }
  Logger.log('Folder OK: ' + config.rootFolderName + '/' + config.testSubrootName);

  ['archana', 'dan', 'rhea', '_Needs Review', '_Logs'].forEach(function (name) {
    const child = findChildFolderByName_(testRoot, name);
    Logger.log(
      child
        ? 'Folder OK: ' + config.rootFolderName + '/' + config.testSubrootName + '/' + name
        : 'Folder MISSING: ' + config.rootFolderName + '/' + config.testSubrootName + '/' + name
    );
  });

  Logger.log('--- VERIFY SETUP END ---');
}

function createMissingLabels() {
  const config = getConfig_();
  const needed = [
    config.processedLabel,
    config.reviewLabel,
    config.ignoredLabel
  ].concat(config.labels);

  needed.forEach(function (name) {
    if (!GmailApp.getUserLabelByName(name)) {
      GmailApp.createLabel(name);
      Logger.log('Created label: ' + name);
    } else {
      Logger.log('Already exists: ' + name);
    }
  });
}

function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

/* =========================
 * DRY RUN
 * ========================= */

function dryRunScan() {
  const config = getConfig_();
  assertAdvancedGmailEnabled_();

  Logger.log('--- DRY RUN SCAN START ---');
  Logger.log('DRY_RUN = ' + config.dryRun);
  Logger.log('MIN_EMAIL_DATE = ' + config.minEmailDate);
  Logger.log('FETCH_REMOTE_IMAGES = ' + config.fetchRemoteImages);
  Logger.log('RECEIPT_THRESHOLD = ' + config.receiptThreshold);
  Logger.log('REVIEW_THRESHOLD = ' + config.reviewThreshold);

  config.labels.forEach(function (labelName) {
    const afterDate = getAfterQueryDate_(config.minEmailDate);
    const query = 'label:' + labelName + ' after:' + afterDate;
    const threads = GmailApp.search(query, 0, 50);

    Logger.log('Label: ' + labelName);
    Logger.log('Matching threads: ' + threads.length);

    threads.forEach(function (thread) {
      const selection = inspectThreadForReceipt_(thread, config);

      if (!selection) {
        Logger.log('No eligible message found in thread ' + thread.getId());
        return;
      }

      const message = selection.message;
      const apiBundle = selection.apiBundle;
      const receiptCheck = selection.receiptCheck;
      const context = buildMessageContext_(message, labelName, config, apiBundle.parsed);
      const receiptDedupKey = buildReceiptDedupKey_(context, apiBundle.parsed);

      const info = {
        label: labelName,
        person: context.person,
        threadId: thread.getId(),
        messageId: message.getId(),
        decision: selection.decision,
        score: receiptCheck.score,
        reasons: receiptCheck.reasons,
        from: message.getFrom(),
        subject: message.getSubject(),
        date: message.getDate().toISOString(),
        orderNumber: context.orderNumber,
        vendorGuess: context.vendor,
        receiptDedupKey: receiptDedupKey,
        alreadyProcessedMessage: isMessageAlreadyProcessed_(message.getId()),
        alreadyReviewedMessage: isMessageReviewed_(message.getId()),
        alreadyIgnoredMessage: isMessageIgnoredNonReceipt_(message.getId()),
        duplicateReceiptKey: isReceiptAlreadyProcessed_(receiptDedupKey),
        htmlLength: apiBundle.parsed.htmlBody.length,
        plainLength: apiBundle.parsed.plainBody.length,
        inlineImagesEmbedded: apiBundle.parsed.inlineImages.length,
        regularAttachments: apiBundle.parsed.regularAttachments.length,
        inspectedMessages: selection.inspectedMessages,
        actions: buildPlannedActions_(config, apiBundle.parsed, selection.decision)
      };

      Logger.log(JSON.stringify(info));
    });
  });

  Logger.log('--- DRY RUN SCAN END ---');
}

function buildPlannedActions_(config, parsed, decision) {
  const actions = [];

  if (decision === 'process') {
    if (config.savePdf) actions.push('save-pdf');
    if (config.saveAttachments && parsed.regularAttachments.length > 0) {
      actions.push('save-attachments');
    }
    if (config.fetchRemoteImages) {
      actions.push('best-effort-inline-remote-images');
    }
    actions.push('mark-processed');
  } else if (decision === 'review') {
    actions.push('route-to-review');
    actions.push('apply-review-label');
  } else if (decision === 'ignore') {
    actions.push('apply-ignored-nonreceipt-label');
  }

  actions.push('full-email-png-not-supported-in-apps-script-only');

  return actions;
}

/* =========================
 * MAIN INGESTION
 * ========================= */

function runReceiptIngestion() {
  const config = getConfig_();
  assertAdvancedGmailEnabled_();

  const lock = LockService.getScriptLock();
  const summary = {
    runId: new Date().toISOString(),
    dryRun: config.dryRun,
    scannedThreads: 0,
    inspectedMessages: 0,
    processed: 0,
    reviewed: 0,
    ignoredNonReceipt: 0,
    skipped: 0,
    skippedNoEligibleMessage: 0,
    skippedDuplicateMessage: 0,
    skippedDuplicateReceipt: 0,
    skippedAlreadyReviewed: 0,
    skippedAlreadyIgnored: 0,
    failed: 0,
    needsReview: 0,
    savedFiles: 0,
    savedHtml: 0,
    savedPdf: 0,
    savedManifest: 0,
    savedAttachments: 0,
    savedInlineImages: 0,
    pngSkipped: 0
  };

  lock.waitLock(30000);

  try {
    Logger.log('--- RECEIPT INGESTION START ---');
    Logger.log(JSON.stringify(summary));

    config.labels.forEach(function (labelName) {
      processLabel_(labelName, config, summary);
    });

    writeRunLog_(summary, config);

    Logger.log('--- RECEIPT INGESTION END ---');
    Logger.log(JSON.stringify(summary, null, 2));
  } finally {
    lock.releaseLock();
  }
}

function processLabel_(labelName, config, summary) {
  const afterDate = getAfterQueryDate_(config.minEmailDate);
  const query = 'label:' + labelName + ' after:' + afterDate;
  const threads = GmailApp.search(query, 0, 100);

  Logger.log(
    'Processing label ' +
      labelName +
      ' with ' +
      threads.length +
      ' matching threads from ' +
      config.minEmailDate +
      ' onward'
  );

  threads.forEach(function (thread) {
    summary.scannedThreads += 1;

    const selection = inspectThreadForReceipt_(thread, config);
    if (!selection) {
      Logger.log('Skipping thread ' + thread.getId() + ' because no eligible message was found');
      summary.skipped += 1;
      summary.skippedNoEligibleMessage += 1;
      return;
    }

    summary.inspectedMessages += selection.inspectedMessages;

    const message = selection.message;
    const apiBundle = selection.apiBundle;
    const receiptCheck = selection.receiptCheck;
    const messageId = message.getId();

    Logger.log(
      'Selected message ' +
        messageId +
        ' from thread ' +
        thread.getId() +
        ' with decision=' +
        selection.decision +
        ' score=' +
        receiptCheck.score
    );

    if (selection.decision === 'process') {
      if (isMessageAlreadyProcessed_(messageId)) {
        Logger.log('Skipping already-processed message ID ' + messageId);
        summary.skipped += 1;
        summary.skippedDuplicateMessage += 1;
        return;
      }

      try {
        const context = buildMessageContext_(message, labelName, config, apiBundle.parsed);
        const receiptDedupKey = buildReceiptDedupKey_(context, apiBundle.parsed);
        context.receiptDedupKey = receiptDedupKey;

        if (isReceiptAlreadyProcessed_(receiptDedupKey)) {
          Logger.log('Skipping duplicate receipt key ' + receiptDedupKey + ' for message ' + messageId);
          if (!config.dryRun) {
            markProcessed_(thread, messageId, receiptDedupKey, config, true);
          }
          summary.skipped += 1;
          summary.skippedDuplicateReceipt += 1;
          return;
        }

        if (config.dryRun) {
          Logger.log('[DRY RUN] Would process ' + messageId + ' for ' + context.person);
          summary.processed += 1;
          return;
        }

        const saveResult = saveEmailPackage_(message, apiBundle, context, config);

        markProcessed_(thread, messageId, receiptDedupKey, config);

        summary.processed += 1;
        summary.savedFiles += saveResult.savedFiles;
        summary.savedHtml += saveResult.savedHtml;
        summary.savedPdf += saveResult.savedPdf;
        summary.savedManifest += saveResult.savedManifest;
        summary.savedAttachments += saveResult.savedAttachments;
        summary.savedInlineImages += saveResult.savedInlineImages;
        summary.pngSkipped += 1;
      } catch (err) {
        Logger.log('FAILED for message ' + messageId + ': ' + err.message);

        routeToNeedsReview_(message, labelName, err, config, receiptCheck);
        addReviewLabel_(thread, config);
        markMessageReviewed_(messageId);

        summary.failed += 1;
        summary.needsReview += 1;
      }

      return;
    }

    if (selection.decision === 'review') {
      if (isMessageReviewed_(messageId)) {
        Logger.log('Skipping already-reviewed borderline message ID ' + messageId);
        summary.skipped += 1;
        summary.skippedAlreadyReviewed += 1;
        return;
      }

      if (config.dryRun) {
        Logger.log('[DRY RUN] Would send message ' + messageId + ' to review');
        summary.reviewed += 1;
        summary.needsReview += 1;
        return;
      }

      routeToNeedsReview_(
        message,
        labelName,
        new Error('Borderline receipt classification'),
        config,
        receiptCheck
      );
      addReviewLabel_(thread, config);
      markMessageReviewed_(messageId);

      summary.reviewed += 1;
      summary.needsReview += 1;
      return;
    }

    if (isMessageIgnoredNonReceipt_(messageId)) {
      Logger.log('Skipping already-ignored non-receipt message ID ' + messageId);
      summary.skipped += 1;
      summary.skippedAlreadyIgnored += 1;
      return;
    }

    if (config.dryRun) {
      Logger.log('[DRY RUN] Would ignore non-receipt message ' + messageId);
      summary.ignoredNonReceipt += 1;
      return;
    }

    addIgnoredLabel_(thread, config);
    markMessageIgnoredNonReceipt_(messageId);
    summary.ignoredNonReceipt += 1;
  });
}

/* =========================
 * RECEIPT CLASSIFICATION
 * ========================= */

function inspectThreadForReceipt_(thread, config) {
  const messages = thread.getMessages();
  let inspectedCount = 0;
  let bestReview = null;
  let bestIgnore = null;

  for (let i = messages.length - 1; i >= 0; i -= 1) {
    const message = messages[i];

    if (!isMessageOnOrAfterMinDate_(message, config.minEmailDate)) {
      continue;
    }

    inspectedCount += 1;

    const apiBundle = getApiMessageBundle_(message.getId());
    const receiptCheck = classifyReceiptMessage_(message, apiBundle.parsed, config);

    const candidate = {
      decision: receiptCheck.decision,
      message: message,
      apiBundle: apiBundle,
      receiptCheck: receiptCheck,
      inspectedMessages: inspectedCount
    };

    if (receiptCheck.decision === 'process') {
      return candidate;
    }

    if (receiptCheck.decision === 'review' && !bestReview) {
      bestReview = candidate;
      continue;
    }

    if (receiptCheck.decision === 'ignore' && !bestIgnore) {
      bestIgnore = candidate;
    }
  }

  if (bestReview) {
    bestReview.inspectedMessages = inspectedCount;
    return bestReview;
  }

  if (bestIgnore) {
    bestIgnore.inspectedMessages = inspectedCount;
    return bestIgnore;
  }

  return null;
}

function classifyReceiptMessage_(message, parsed, config) {
  const subject = stripReplyForwardPrefixes_(message.getSubject() || '');
  const from = String(message.getFrom() || '');
  const plainBody = String(parsed.plainBody || '');
  const htmlText = htmlToPlainText_(parsed.htmlBody || '');
  const attachmentNames = (parsed.regularAttachments || [])
    .map(function (att) { return String(att.filename || ''); })
    .join(' ');

  const combined = [subject, from, plainBody, htmlText, attachmentNames].join('\n');
  const combinedLower = combined.toLowerCase();

  let score = 0;
  const reasons = [];

  const strongPositivePatterns = [
    /\breceipt\b/i,
    /\binvoice\b/i,
    /\border confirmation\b/i,
    /\bpayment received\b/i,
    /\bbooking confirmation\b/i,
    /\bitinerary\b/i,
    /\btrip receipt\b/i,
    /\byour order\b/i,
    /\bthank you for your order\b/i,
    /\border number\b/i,
    /\binvoice number\b/i,
    /\btransaction id\b/i,
    /\bconfirmation number\b/i,
    /\bpayment summary\b/i,
    /\bcharged\b/i,
    /\bamount paid\b/i
  ];

  strongPositivePatterns.forEach(function (pattern) {
    if (pattern.test(combined)) {
      score += 2;
      reasons.push('positive:' + pattern);
    }
  });

  const attachments = parsed.regularAttachments || [];
  if (attachments.length > 0) {
    score += 2;
    reasons.push('has-attachment');
  }

  const hasPdf = attachments.some(function (att) {
    const mimeType = String(att.mimeType || '').toLowerCase();
    const filename = String(att.filename || '');
    return mimeType.indexOf('pdf') > -1 || /\.pdf$/i.test(filename);
  });
  if (hasPdf) {
    score += 2;
    reasons.push('has-pdf');
  }

  const weakPositivePatterns = [
    /\border\b/i,
    /\bconfirmation\b/i,
    /\btotal\b/i,
    /\bsubtotal\b/i,
    /\btax\b/i,
    /\bbilling\b/i,
    /\bamount\b/i,
    /\bpayment\b/i,
    /\bmerchant\b/i,
    /\bvisa\b/i,
    /\bmastercard\b/i,
    /\bcard ending\b/i,
    /\blast 4\b/i
  ];

  weakPositivePatterns.forEach(function (pattern) {
    if (pattern.test(combined)) {
      score += 1;
      reasons.push('weak-positive:' + pattern);
    }
  });

  if (containsCurrencyAmount_(combinedLower)) {
    score += 1;
    reasons.push('currency-amount');
  }

  const negativePatterns = [
    /\bunsubscribe\b/i,
    /\bmanage preferences\b/i,
    /\bview in browser\b/i,
    /\bshop now\b/i,
    /\bbuy now\b/i,
    /\blimited time\b/i,
    /\bsale\b/i,
    /\bpromo\b/i,
    /\bpromotion\b/i,
    /\bnewsletter\b/i,
    /\bnew arrivals\b/i,
    /\btrending\b/i,
    /\bfollow us\b/i,
    /\bon instagram\b/i,
    /\bon facebook\b/i,
    /\bon tiktok\b/i,
    /\byou may also like\b/i,
    /\brecommended for you\b/i,
    /\bflash sale\b/i,
    /\bexclusive offer\b/i,
    /\buse code\b/i,
    /\bdiscount code\b/i
  ];

  negativePatterns.forEach(function (pattern) {
    if (pattern.test(combined)) {
      score -= 2;
      reasons.push('negative:' + pattern);
    }
  });

  const shippingOnlyPatterns = [
    /\bshipped\b/i,
    /\bout for delivery\b/i,
    /\btracking number\b/i,
    /\btrack your package\b/i,
    /\bdelivery update\b/i,
    /\bshipment\b/i,
    /\bestimated delivery\b/i
  ];

  let shippingHits = 0;
  shippingOnlyPatterns.forEach(function (pattern) {
    if (pattern.test(combined)) {
      shippingHits += 1;
    }
  });

  if (shippingHits > 0) {
    score -= 1;
    reasons.push('shipping-language');
  }

  const rawSubject = String(message.getSubject() || '').trim();
  if (/^fwd?:/i.test(rawSubject) || /^re:/i.test(rawSubject)) {
    score -= 1;
    reasons.push('reply-or-forward');
  }

  const decision =
    score >= config.receiptThreshold
      ? 'process'
      : score >= config.reviewThreshold
      ? 'review'
      : 'ignore';

  return {
    isReceipt: decision === 'process',
    decision: decision,
    score: score,
    reasons: reasons
  };
}

function containsCurrencyAmount_(text) {
  const value = String(text || '');
  return /(\$\s?\d[\d,]*(?:\.\d{2})?)|(usd\s?\d[\d,]*(?:\.\d{2})?)/i.test(value);
}

/* =========================
 * GMAIL API PARSING
 * ========================= */

function getApiMessageBundle_(messageId) {
  const fullMessage = Gmail.Users.Messages.get('me', messageId, { format: 'full' });
  const parsed = parseApiMessage_(fullMessage);

  return {
    fullMessage: fullMessage,
    parsed: parsed
  };
}

function parseApiMessage_(apiMessage) {
  const state = {
    htmlCandidates: [],
    plainCandidates: [],
    inlineImages: [],
    regularAttachments: [],
    topHeaders: headersArrayToMap_(
      apiMessage.payload && apiMessage.payload.headers ? apiMessage.payload.headers : []
    ),
    messageId: apiMessage.id
  };

  walkMimePart_(apiMessage.payload, apiMessage.id, state);

  const bestHtml = chooseBestCandidate_(state.htmlCandidates);
  const bestPlain = chooseBestCandidate_(state.plainCandidates);

  let htmlBody = bestHtml ? bestHtml.content : '';
  const plainBody = bestPlain ? bestPlain.content : '';

  htmlBody = inlineCidImages_(htmlBody, state.inlineImages);

  return {
    htmlBody: htmlBody,
    plainBody: plainBody,
    inlineImages: state.inlineImages,
    regularAttachments: state.regularAttachments,
    headers: state.topHeaders
  };
}

function walkMimePart_(part, messageId, state) {
  if (!part) return;

  const mimeType = String(part.mimeType || '').toLowerCase();
  const headers = headersArrayToMap_(part.headers || []);
  const disposition = String(headers['content-disposition'] || '').toLowerCase();
  const contentId = normalizeContentId_(headers['content-id'] || '');
  const filename = part.filename || '';

  const hasChildParts = part.parts && part.parts.length > 0;
  const hasBodyData = part.body && (part.body.data || part.body.attachmentId);

  if (hasChildParts) {
    part.parts.forEach(function (child) {
      walkMimePart_(child, messageId, state);
    });
  }

  if (!hasBodyData) return;

  if (mimeType === 'text/html') {
    const html = getDecodedPartString_(messageId, part);
    if (html) {
      state.htmlCandidates.push({
        mimeType: mimeType,
        content: stripScriptsOnly_(html)
      });
    }
    return;
  }

  if (mimeType === 'text/plain') {
    const text = getDecodedPartString_(messageId, part);
    if (text) {
      state.plainCandidates.push({
        mimeType: mimeType,
        content: text
      });
    }
    return;
  }

  if (mimeType.indexOf('image/') === 0 && (contentId || disposition.indexOf('inline') > -1)) {
    const bytes = getPartBytes_(messageId, part);
    if (bytes.length > 0) {
      state.inlineImages.push({
        contentId: contentId,
        filename:
          filename ||
          'inline-' +
            (part.partId || 'image') +
            '.' +
            (guessExtensionFromMime_(mimeType) || 'bin'),
        mimeType: mimeType,
        bytes: bytes,
        dataUri: buildDataUri_(mimeType, bytes)
      });
    }
    return;
  }

  if (filename || disposition.indexOf('attachment') > -1 || isLikelyAttachmentMime_(mimeType)) {
    const bytes = getPartBytes_(messageId, part);
    if (bytes.length > 0) {
      state.regularAttachments.push({
        filename: filename || buildFallbackAttachmentName_(part),
        mimeType: mimeType || 'application/octet-stream',
        bytes: bytes
      });
    }
  }
}

function chooseBestCandidate_(candidates) {
  if (!candidates || candidates.length === 0) return null;

  let best = candidates[0];
  for (let i = 1; i < candidates.length; i += 1) {
    if ((candidates[i].content || '').length > (best.content || '').length) {
      best = candidates[i];
    }
  }
  return best;
}

function getDecodedPartString_(messageId, part) {
  const bytes = getPartBytes_(messageId, part);
  if (!bytes || bytes.length === 0) return '';

  const charset = extractCharsetFromHeaders_(part.headers || []) || 'UTF-8';

  try {
    return Utilities.newBlob(bytes).getDataAsString(charset);
  } catch (err) {
    try {
      return Utilities.newBlob(bytes).getDataAsString();
    } catch (err2) {
      Logger.log(
        'WARN getDecodedPartString_ failed for message ' +
          messageId +
          ', part ' +
          (part.partId || 'unknown')
      );
      return '';
    }
  }
}

function getPartBytes_(messageId, part) {
  if (!part || !part.body) return [];

  try {
    if (part.body.data !== null && part.body.data !== undefined) {
      return decodeBase64UrlToBytes_(part.body.data);
    }

    if (part.body.attachmentId) {
      const body = Gmail.Users.Messages.Attachments.get('me', messageId, part.body.attachmentId);
      return decodeBase64UrlToBytes_(body.data || '');
    }

    return [];
  } catch (err) {
    Logger.log(
      'WARN getPartBytes_ failed for message ' +
        messageId +
        ', part ' +
        (part.partId || 'unknown') +
        ': ' +
        err.message
    );
    return [];
  }
}

/* =========================
 * SAVE PACKAGE
 * ========================= */

function saveEmailPackage_(message, apiBundle, context, config) {
  const folder = getTargetFolder_(context, config);
  const result = {
    savedFiles: 0,
    savedHtml: 0,
    savedPdf: 0,
    savedManifest: 0,
    savedAttachments: 0,
    savedInlineImages: 0
  };

  const renderHtml = buildRenderableEmailHtml_(message, apiBundle.parsed, config);

  if (config.savePdf) {
    const htmlBlobForPdf = Utilities.newBlob(
      renderHtml,
      'text/html',
      context.datePart + '-' + context.vendor + '-' + context.shortId + '-email-body.html'
    );

    const pdfFilename = buildEmailBodyPdfFilename_(context);
    const pdfBlob = htmlBlobForPdf.getAs('application/pdf').setName(pdfFilename);
    const pdfFile = folder.createFile(pdfBlob).setName(pdfFilename);

    addFileDescription_(
      pdfFile,
      mergeObjects_(context, {
        artifactType: 'email-body-pdf',
        status: 'success'
      })
    );

    result.savedFiles += 1;
    result.savedPdf += 1;
  }

  if (config.saveAttachments && apiBundle.parsed.regularAttachments.length > 0) {
    const attachFolder = getOrCreateChildFolder_(folder, '_attachments');

    apiBundle.parsed.regularAttachments.forEach(function (att) {
      const safeName = buildAttachmentOutputName_(context, att);

      const blob = Utilities.newBlob(
        att.bytes,
        att.mimeType || 'application/octet-stream',
        safeName
      );

      const file = attachFolder.createFile(blob).setName(safeName);

      addFileDescription_(
        file,
        mergeObjects_(context, {
          artifactType: 'attachment',
          originalAttachmentName: att.filename,
          contentType: att.mimeType,
          status: 'success'
        })
      );

      result.savedFiles += 1;
      result.savedAttachments += 1;
    });
  }

  return result;
}

function buildRenderableEmailHtml_(message, parsed, config) {
  let html = parsed.htmlBody || '';

  if (!html || !html.trim()) {
    html = stripScriptsOnly_(message.getBody() || '');
  }

  if (!html || !html.trim()) {
    html = '<pre>' + escapeHtml_(parsed.plainBody || message.getPlainBody() || '') + '</pre>';
  }

  html = inlineCidImages_(html, parsed.inlineImages);

  if (config.fetchRemoteImages) {
    html = inlineRemoteImagesAsDataUris_(html);
  }

  return ensureHtmlDocument_(html, message.getSubject() || '(no subject)');
}

function buildEmailBodyPdfFilename_(context) {
  const ref = context.orderNumber || context.shortId;
  return context.datePart + '-' + context.vendor + '-' + ref + '-email-body.pdf';
}

function normalizeFetchableUrl_(url) {
  let value = String(url || '').trim();

  if (!value) {
    return '';
  }

  if (value.indexOf('//') === 0) {
    value = 'https:' + value;
  }

  value = value.replace(/&amp;/g, '&');

  return value;
}

function inlineRemoteImagesAsDataUris_(html) {
  let value = String(html || '');
  if (!value) return value;

  const urls = collectRemoteImageUrls_(value);
  if (urls.length === 0) return value;

  const replacements = fetchImageDataUris_(urls);

  value = value.replace(
    /src=(['"])(https?:\/\/[^'"]+|\/\/[^'"]+)\1/gi,
    function (match, quote, url) {
      const normalized = normalizeFetchableUrl_(url);
      if (replacements[normalized]) {
        return 'src=' + quote + replacements[normalized] + quote;
      }
      return match;
    }
  );

  value = value.replace(
    /url\((['"]?)(https?:\/\/[^'"()]+|\/\/[^'"()]+)\1\)/gi,
    function (match, quote, url) {
      const normalized = normalizeFetchableUrl_(url);
      if (replacements[normalized]) {
        return 'url(' + quote + replacements[normalized] + quote + ')';
      }
      return match;
    }
  );

  return value;
}

function collectRemoteImageUrls_(html) {
  const found = {};
  const urls = [];
  const MAX_REMOTE_IMAGES = 20;

  function addUrl(url) {
    const normalized = normalizeFetchableUrl_(url);

    if (!shouldFetchRemoteImageUrl_(normalized)) {
      return;
    }

    if (!found[normalized]) {
      found[normalized] = true;
      urls.push(normalized);
    }
  }

  html.replace(
    /<img[^>]+src=(['"])(https?:\/\/[^'"]+|\/\/[^'"]+)\1/gi,
    function (_match, _quote, url) {
      addUrl(url);
      return _match;
    }
  );

  html.replace(
    /url\((['"]?)(https?:\/\/[^'"()]+|\/\/[^'"()]+)\1\)/gi,
    function (_match, _quote, url) {
      addUrl(url);
      return _match;
    }
  );

  return urls.slice(0, MAX_REMOTE_IMAGES);
}

function fetchImageDataUris_(urls) {
  const map = {};
  if (!urls || urls.length === 0) return map;

  const startedAt = Date.now();
  const MAX_TOTAL_FETCH_MS = 20000;
  const MAX_IMAGE_BYTES = 3 * 1024 * 1024;

  for (let i = 0; i < urls.length; i += 1) {
    if (Date.now() - startedAt > MAX_TOTAL_FETCH_MS) {
      Logger.log('WARN image fetch budget exceeded, skipping remaining images');
      break;
    }

    const url = urls[i];

    try {
      const response = UrlFetchApp.fetch(url, {
        method: 'get',
        muteHttpExceptions: true,
        followRedirects: true,
        headers: {
          'User-Agent': 'Mozilla/5.0',
          Accept: 'image/avif,image/webp,image/apng,image/*,*/*;q=0.8'
        }
      });

      const code = response.getResponseCode();
      if (code < 200 || code >= 300) {
        Logger.log('WARN image fetch failed ' + code + ' for ' + url);
        continue;
      }

      const blob = response.getBlob();
      const contentType = String(blob.getContentType() || '').toLowerCase();

      if (contentType.indexOf('image/') !== 0) {
        Logger.log('WARN non-image content for ' + url + ': ' + contentType);
        continue;
      }

      const bytes = blob.getBytes();
      if (!bytes || bytes.length === 0) {
        Logger.log('WARN empty image response for ' + url);
        continue;
      }

      if (bytes.length > MAX_IMAGE_BYTES) {
        Logger.log('WARN oversized image skipped (' + bytes.length + ' bytes) for ' + url);
        continue;
      }

      map[url] = buildDataUri_(contentType, bytes);
    } catch (err) {
      Logger.log('WARN remote image fetch exception for ' + url + ': ' + err.message);
    }
  }

  return map;
}

/* =========================
 * HTML / MIME HELPERS
 * ========================= */

function shouldFetchRemoteImageUrl_(url) {
  const value = String(url || '').trim().toLowerCase();
  if (!value) return false;

  if (value.length > 1800) return false;

  if (/pixel|tracking|beacon|analytics|openrate|impression/.test(value)) {
    return false;
  }

  if (/static-?map|maps\.uber|mapbox|marker=|markers=|waypoint/.test(value)) {
    return false;
  }

  if (
    value.indexOf('data:') === 0 ||
    value.indexOf('cid:') === 0 ||
    value.indexOf('blob:') === 0
  ) {
    return false;
  }

  return true;
}

function ensureHtmlDocument_(html, title) {
  let value = stripScriptsOnly_(String(html || '').trim());

  if (!value) {
    return (
      '<!DOCTYPE html>\n' +
      '<html>\n' +
      '  <head>\n' +
      '    <meta charset="utf-8" />\n' +
      '    <title>' +
      escapeHtml_(title || '(no subject)') +
      '</title>\n' +
      '  </head>\n' +
      '  <body></body>\n' +
      '</html>'
    );
  }

  if (/<!doctype/i.test(value) || /<html[\s>]/i.test(value)) {
    return value;
  }

  if (/<head[\s>]/i.test(value) || /<body[\s>]/i.test(value)) {
    return (
      '<!DOCTYPE html>\n' +
      '<html>\n' +
      '  <head>\n' +
      '    <meta charset="utf-8" />\n' +
      '    <title>' +
      escapeHtml_(title || '(no subject)') +
      '</title>\n' +
      '  </head>\n' +
      value +
      '\n</html>'
    );
  }

  return (
    '<!DOCTYPE html>\n' +
    '<html>\n' +
    '  <head>\n' +
    '    <meta charset="utf-8" />\n' +
    '    <title>' +
    escapeHtml_(title || '(no subject)') +
    '</title>\n' +
    '    <style>\n' +
    '      html, body { margin: 0; padding: 0; background: #ffffff; }\n' +
    '      pre { white-space: pre-wrap; word-wrap: break-word; }\n' +
    '      * { box-sizing: border-box; }\n' +
    '    </style>\n' +
    '  </head>\n' +
    '  <body>\n' +
    value +
    '\n  </body>\n' +
    '</html>'
  );
}

function stripScriptsOnly_(html) {
  return String(html || '').replace(/<script[\s\S]*?<\/script>/gi, '').trim();
}

function inlineCidImages_(html, inlineImages) {
  let value = String(html || '');
  if (!value || !inlineImages || inlineImages.length === 0) return value;

  const map = {};
  inlineImages.forEach(function (img) {
    if (img.contentId) {
      map[img.contentId.toLowerCase()] = img.dataUri;
    }
  });

  value = value.replace(/src=(['"])cid:([^'"]+)\1/gi, function (match, quote, cid) {
    const normalized = normalizeContentId_(cid).toLowerCase();
    if (map[normalized]) {
      return 'src=' + quote + map[normalized] + quote;
    }
    return match;
  });

  value = value.replace(/url\((['"]?)cid:([^"')]+)\1\)/gi, function (match, _quote, cid) {
    const normalized = normalizeContentId_(cid).toLowerCase();
    if (map[normalized]) {
      return 'url(' + map[normalized] + ')';
    }
    return match;
  });

  return value;
}

function headersArrayToMap_(headers) {
  const map = {};
  (headers || []).forEach(function (header) {
    map[String(header.name || '').toLowerCase()] = String(header.value || '');
  });
  return map;
}

function extractCharsetFromHeaders_(headers) {
  const map = headersArrayToMap_(headers || []);
  const contentType = map['content-type'] || '';
  const match = contentType.match(/charset="?([^";]+)"?/i);
  return match ? match[1] : '';
}

function decodeBase64UrlToBytes_(data) {
  if (data === null || data === undefined) {
    return [];
  }

  if (Array.isArray(data)) {
    return data;
  }

  if (typeof data === 'object') {
    try {
      if (typeof data.length === 'number') {
        const arr = [];
        for (let i = 0; i < data.length; i += 1) {
          arr.push(Number(data[i]));
        }
        if (arr.length > 0 && arr.every(function (n) { return !isNaN(n); })) {
          return arr;
        }
      }
    } catch (err) {
      // Fall through to string path.
    }
  }

  const rawString = String(data).trim();
  if (!rawString) {
    return [];
  }

  if (/^\d+(,\d+)+$/.test(rawString)) {
    return rawString.split(',').map(function (n) { return Number(n); });
  }

  const cleaned = rawString.replace(/\s+/g, '');
  let padded = cleaned;
  while (padded.length % 4 !== 0) {
    padded += '=';
  }

  try {
    return Utilities.base64DecodeWebSafe(padded);
  } catch (err) {
    const fallback = padded.replace(/-/g, '+').replace(/_/g, '/');
    try {
      return Utilities.base64Decode(fallback);
    } catch (err2) {
      throw new Error('Could not decode binary field. Prefix=' + rawString.slice(0, 60) + '...');
    }
  }
}

function buildDataUri_(mimeType, bytes) {
  if (!bytes || bytes.length === 0) return '';
  const base64 = Utilities.base64Encode(bytes);
  return 'data:' + mimeType + ';base64,' + base64;
}

function normalizeContentId_(contentId) {
  return String(contentId || '').replace(/[<>]/g, '').trim();
}

function isLikelyAttachmentMime_(mimeType) {
  const mime = String(mimeType || '').toLowerCase();
  if (!mime) return false;
  if (mime === 'text/html' || mime === 'text/plain') return false;
  if (mime.indexOf('multipart/') === 0) return false;
  return true;
}

function buildFallbackAttachmentName_(part) {
  const ext = guessExtensionFromMime_(part.mimeType || '') || 'bin';
  return 'part-' + (part.partId || 'attachment') + '.' + ext;
}

function guessExtensionFromMime_(mime) {
  const value = String(mime || '').toLowerCase();
  if (value.indexOf('pdf') > -1) return 'pdf';
  if (value.indexOf('png') > -1) return 'png';
  if (value.indexOf('jpeg') > -1 || value.indexOf('jpg') > -1) return 'jpg';
  if (value.indexOf('gif') > -1) return 'gif';
  if (value.indexOf('webp') > -1) return 'webp';
  if (value.indexOf('html') > -1) return 'html';
  if (value.indexOf('plain') > -1) return 'txt';
  if (value.indexOf('csv') > -1) return 'csv';
  if (value.indexOf('json') > -1) return 'json';
  if (value.indexOf('zip') > -1) return 'zip';
  return '';
}

function sanitizeFilename_(name) {
  return String(name || '')
    .replace(/[\\/:*?"<>|]+/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}

function escapeHtml_(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/* =========================
 * FOLDER HELPERS
 * ========================= */

function getTargetFolder_(context, config) {
  const root = getOrCreateTopFolder_(config.rootFolderName);
  const testRoot = getOrCreateChildFolder_(root, config.testSubrootName);
  const personFolder = getOrCreateChildFolder_(testRoot, context.person);
  const yearFolder = getOrCreateChildFolder_(personFolder, context.year);
  return getOrCreateChildFolder_(yearFolder, context.month);
}

function getNeedsReviewFolder_(person, config) {
  const root = getOrCreateTopFolder_(config.rootFolderName);
  const testRoot = getOrCreateChildFolder_(root, config.testSubrootName);
  const reviewRoot = getOrCreateChildFolder_(testRoot, '_Needs Review');
  return getOrCreateChildFolder_(reviewRoot, person);
}

function getLogsFolder_(config) {
  const root = getOrCreateTopFolder_(config.rootFolderName);
  const testRoot = getOrCreateChildFolder_(root, config.testSubrootName);
  return getOrCreateChildFolder_(testRoot, '_Logs');
}

function getOrCreateTopFolder_(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function getOrCreateChildFolder_(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function findFolderByName_(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : null;
}

function findChildFolderByName_(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : null;
}

/* =========================
 * LOGGING / REVIEW
 * ========================= */

function writeRunLog_(summary, config) {
  const logsFolder = getLogsFolder_(config);
  const timestamp = Utilities.formatDate(new Date(), config.timezone, 'yyyy-MM-dd_HH-mm-ss');
  const filename = timestamp + '_run-log.json';

  logsFolder.createFile(filename, JSON.stringify(summary, null, 2), MimeType.PLAIN_TEXT);
}

function routeToNeedsReview_(message, labelName, err, config, receiptCheck) {
  const person = mapLabelToPerson_(labelName);
  const reviewFolder = getNeedsReviewFolder_(person, config);
  const now = new Date();

  const name = [
    Utilities.formatDate(now, config.timezone, 'yyyy-MM-dd_HH-mm-ss'),
    person,
    message.getId().slice(-8),
    'needs-review.txt'
  ].join('_');

  const contents = [
    'Person: ' + person,
    'Label: ' + labelName,
    'Message ID: ' + message.getId(),
    'Thread ID: ' + message.getThread().getId(),
    'From: ' + message.getFrom(),
    'Subject: ' + message.getSubject(),
    'Received: ' + message.getDate().toISOString(),
    'Error: ' + err.message,
    'Classifier decision: ' + (receiptCheck ? receiptCheck.decision : ''),
    'Classifier score: ' + (receiptCheck ? receiptCheck.score : ''),
    'Classifier reasons: ' + (receiptCheck ? receiptCheck.reasons.join(' | ') : ''),
    '',
    'Plain body preview:',
    '------------------',
    truncate_(message.getPlainBody() || '', 4000)
  ].join('\n');

  reviewFolder.createFile(name, contents, MimeType.PLAIN_TEXT);
}

function addFileDescription_(file, metadata) {
  file.setDescription(JSON.stringify(metadata, null, 2));
}

/* =========================
 * PROCESSED / REVIEW / IGNORED STATE
 * ========================= */

function markProcessed_(thread, messageId, receiptDedupKey, config, skipReceiptMark) {
  const processedLabel = getOrCreateLabel_(config.processedLabel);
  thread.addLabel(processedLabel);
  markMessageProcessed_(messageId);

  if (!skipReceiptMark && receiptDedupKey) {
    markReceiptProcessed_(receiptDedupKey);
  }
}

function addReviewLabel_(thread, config) {
  const reviewLabel = getOrCreateLabel_(config.reviewLabel);
  thread.addLabel(reviewLabel);
}

function addIgnoredLabel_(thread, config) {
  const ignoredLabel = getOrCreateLabel_(config.ignoredLabel);
  thread.addLabel(ignoredLabel);
}

function isMessageAlreadyProcessed_(messageId) {
  const map = readJsonPropertyObject_('PROCESSED_MESSAGE_IDS');
  return Boolean(map[messageId]);
}

function markMessageProcessed_(messageId) {
  const map = readJsonPropertyObject_('PROCESSED_MESSAGE_IDS');
  map[messageId] = new Date().toISOString();
  writeJsonPropertyObject_('PROCESSED_MESSAGE_IDS', map);
}

function isMessageReviewed_(messageId) {
  const map = readJsonPropertyObject_('REVIEWED_MESSAGE_IDS');
  return Boolean(map[messageId]);
}

function markMessageReviewed_(messageId) {
  const map = readJsonPropertyObject_('REVIEWED_MESSAGE_IDS');
  map[messageId] = new Date().toISOString();
  writeJsonPropertyObject_('REVIEWED_MESSAGE_IDS', map);
}

function isMessageIgnoredNonReceipt_(messageId) {
  const map = readJsonPropertyObject_('IGNORED_NON_RECEIPT_MESSAGE_IDS');
  return Boolean(map[messageId]);
}

function markMessageIgnoredNonReceipt_(messageId) {
  const map = readJsonPropertyObject_('IGNORED_NON_RECEIPT_MESSAGE_IDS');
  map[messageId] = new Date().toISOString();
  writeJsonPropertyObject_('IGNORED_NON_RECEIPT_MESSAGE_IDS', map);
}

function isReceiptAlreadyProcessed_(receiptDedupKey) {
  if (!receiptDedupKey) return false;
  const map = readJsonPropertyObject_('PROCESSED_RECEIPT_KEYS');
  return Boolean(map[receiptDedupKey]);
}

function markReceiptProcessed_(receiptDedupKey) {
  if (!receiptDedupKey) return;
  const map = readJsonPropertyObject_('PROCESSED_RECEIPT_KEYS');
  map[receiptDedupKey] = new Date().toISOString();
  writeJsonPropertyObject_('PROCESSED_RECEIPT_KEYS', map);
}

function readJsonPropertyObject_(key) {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(key) || '{}';

  try {
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch (err) {
    Logger.log('WARN invalid JSON in Script Property ' + key + '. Resetting to empty object.');
    return {};
  }
}

function writeJsonPropertyObject_(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(value || {}));
}

function buildReceiptDedupKey_(context, parsed) {
  const person = sanitizeKeyPart_(context.person || 'unknown');
  const vendor = sanitizeKeyPart_(context.vendor || 'unknownvendor');
  const orderNumber = sanitizeKeyPart_(context.orderNumber || '');

  if (orderNumber) {
    return ('v2|' + person + '|' + vendor + '|order|' + orderNumber).toLowerCase();
  }

  const attachmentNames = (parsed.regularAttachments || [])
    .map(function (att) { return sanitizeKeyPart_(att.filename || ''); })
    .filter(Boolean)
    .sort()
    .join('|');

  const bodyText = normalizeKeyText_(
    parsed.plainBody || htmlToPlainText_(parsed.htmlBody || '')
  ).slice(0, 2000);

  const source = [
    person,
    vendor,
    normalizeKeyText_(context.from || ''),
    normalizeKeyText_(context.subject || ''),
    attachmentNames,
    bodyText
  ].join('|');

  return (
    'v2|' +
    person +
    '|' +
    vendor +
    '|hash|' +
    computeDigestHex_(source).slice(0, 32)
  ).toLowerCase();
}

function sanitizeKeyPart_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^\w-]+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '')
    .trim();
}

function normalizeKeyText_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function computeDigestHex_(value) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(value || ''),
    Utilities.Charset.UTF_8
  );

  return bytes
    .map(function (b) {
      const normalized = b < 0 ? b + 256 : b;
      const hex = normalized.toString(16);
      return hex.length === 1 ? '0' + hex : hex;
    })
    .join('');
}

/* =========================
 * MESSAGE CONTEXT / VENDOR HELPERS
 * ========================= */

function buildMessageContext_(message, labelName, config, parsed) {
  const receivedDate = message.getDate();
  const person = mapLabelToPerson_(labelName);
  const year = Utilities.formatDate(receivedDate, config.timezone, 'yyyy');
  const month = Utilities.formatDate(receivedDate, config.timezone, 'yyyy-MM');

  const vendor = detectVendor_(message, parsed);
  const orderNumber = extractOrderNumber_(message, parsed);

  return {
    messageId: message.getId(),
    threadId: message.getThread().getId(),
    labelName: labelName,
    person: person,
    from: message.getFrom(),
    subject: message.getSubject(),
    receivedAtIso: receivedDate.toISOString(),
    year: year,
    month: month,
    vendor: vendor || 'UnknownVendor',
    orderNumber: orderNumber || '',
    shortId: message.getId().slice(-8),
    datePart: Utilities.formatDate(receivedDate, config.timezone, 'ddMMM')
  };
}

function mapLabelToPerson_(labelName) {
  if (labelName === 'archana-expenses') return 'archana';
  if (labelName === 'dan-expenses') return 'dan';
  if (labelName === 'rhea-expenses') return 'rhea';
  return 'unknown';
}

function detectVendor_(message, parsed) {
  const attachments = parsed && parsed.regularAttachments ? parsed.regularAttachments : [];
  const attachmentName = getFirstAttachmentName_(attachments);
  const subject = stripReplyForwardPrefixes_(message.getSubject() || '');
  const plainBody = parsed && parsed.plainBody ? parsed.plainBody : '';
  const htmlBody = parsed && parsed.htmlBody ? parsed.htmlBody : '';

  const knownFromAttachment = guessVendorFromKnownPatterns_(attachmentName);
  if (knownFromAttachment) return knownFromAttachment;

  const forwardedDisplayName = extractForwardedFromDisplayName_(plainBody);
  if (forwardedDisplayName) {
    return sanitizeVendor_(forwardedDisplayName);
  }

  const htmlText = htmlToPlainText_(htmlBody);
  const titleText = extractHtmlTitle_(htmlBody);
  const combined = [titleText, plainBody, htmlText, subject, attachmentName].join('\n');

  const knownFromBody = guessVendorFromKnownPatterns_(combined);
  if (knownFromBody) return knownFromBody;

  const brandFromHtml = extractLikelyBrandFromHtml_(htmlBody);
  if (brandFromHtml) return brandFromHtml;

  const forwardedDomainVendor = extractForwardedDomainVendor_(plainBody);
  if (forwardedDomainVendor) return forwardedDomainVendor;

  const subjectVendor = cleanLikelyVendor_(subject);
  if (subjectVendor) return subjectVendor;

  return 'UnknownVendor';
}

function getFirstAttachmentName_(attachments) {
  if (!attachments || attachments.length === 0) return '';
  const first = attachments[0] || {};
  return String(first.filename || first.name || '');
}

function sanitizeVendor_(value) {
  const cleaned = String(value || '')
    .replace(/[^\w]+/g, ' ')
    .trim();

  if (!cleaned) return 'UnknownVendor';

  return cleaned
    .split(/\s+/)
    .map(function (part) {
      return part.charAt(0).toUpperCase() + part.slice(1).toLowerCase();
    })
    .join('');
}

function stripReplyForwardPrefixes_(text) {
  let value = String(text || '').trim();

  while (/^(?:fwd?|fw|re)\s*:\s*/i.test(value)) {
    value = value.replace(/^(?:fwd?|fw|re)\s*:\s*/i, '').trim();
  }

  return value;
}

function extractForwardedFromDisplayName_(plainBody) {
  const value = String(plainBody || '');
  const match = value.match(/^From:\s*"?(.*?)"?\s*<[^>]+>/mi);

  if (!match || !match[1]) return '';

  const name = match[1].replace(/^["']|["']$/g, '').trim();

  if (/^rhea amante$/i.test(name)) return '';

  return name;
}

function extractForwardedDomainVendor_(plainBody) {
  const value = String(plainBody || '');
  const match = value.match(/^From:\s.*?@([A-Za-z0-9.-]+\.[A-Za-z]{2,})/mi);

  if (!match || !match[1]) return '';

  return normalizeVendorFromDomain_(match[1]);
}

function normalizeVendorFromDomain_(domain) {
  const value = String(domain || '').toLowerCase();

  const stripped = value
    .replace(/^mail\./, '')
    .replace(/^accounts\./, '')
    .replace(/^invoices\./, '')
    .replace(/^notifications\./, '');

  const parts = stripped.split('.');
  if (parts.length < 2) return null;

  const core = parts[parts.length - 2];
  if (!core) return null;

  return sanitizeVendor_(core);
}

function guessVendorFromKnownPatterns_(text) {
  const value = String(text || '').trim();
  if (!value) return '';

  const knownPatterns = [
    { pattern: /hampers\s*&\s*co/i, vendor: 'HampersAndCo' },
    { pattern: /pop\s*mart/i, vendor: 'PopMart' },
    { pattern: /devine'?s/i, vendor: 'Devines' },
    { pattern: /amazon/i, vendor: 'Amazon' },
    { pattern: /uber/i, vendor: 'Uber' },
    { pattern: /lyft/i, vendor: 'Lyft' },
    { pattern: /starbucks/i, vendor: 'Starbucks' },
    { pattern: /walmart/i, vendor: 'Walmart' },
    { pattern: /target/i, vendor: 'Target' },
    { pattern: /united/i, vendor: 'United' }
  ];

  for (let i = 0; i < knownPatterns.length; i += 1) {
    if (knownPatterns[i].pattern.test(value)) {
      return knownPatterns[i].vendor;
    }
  }

  return '';
}

function extractHtmlTitle_(html) {
  const value = String(html || '');
  const match = value.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
  if (!match || !match[1]) return '';
  return htmlToPlainText_(match[1]);
}

function extractLikelyBrandFromHtml_(html) {
  const title = cleanLikelyVendor_(extractHtmlTitle_(html));
  if (title) return title;

  const value = String(html || '');
  const headingMatch = value.match(/<(h1|h2|h3)[^>]*>([\s\S]*?)<\/\1>/i);

  if (!headingMatch || !headingMatch[2]) return '';

  return cleanLikelyVendor_(htmlToPlainText_(headingMatch[2]));
}

function htmlToPlainText_(html) {
  return String(html || '')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&#39;/gi, "'")
    .replace(/&quot;/gi, '"')
    .replace(/\s+/g, ' ')
    .trim();
}

function cleanLikelyVendor_(text) {
  let value = stripReplyForwardPrefixes_(String(text || ''));

  value = value
    .replace(/\b(order|receipt|invoice|confirmed|summary|trip|your|for|from|re|fwd|fw)\b/gi, ' ')
    .replace(/[#:]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (!value) return '';

  if (/^(thank you|view your order|customer information)$/i.test(value)) {
    return '';
  }

  return sanitizeVendor_(value);
}

function extractOrderNumber_(message, parsed) {
  const subject = stripReplyForwardPrefixes_(message.getSubject() || '');
  const plainBody = parsed && parsed.plainBody ? parsed.plainBody : '';
  const htmlText = parsed && parsed.htmlBody ? htmlToPlainText_(parsed.htmlBody) : '';
  const attachmentName = getFirstAttachmentName_(
    parsed && parsed.regularAttachments ? parsed.regularAttachments : []
  );

  const sources = [subject, plainBody, htmlText, attachmentName];

  const patterns = [
    /\b(?:order|receipt|invoice|booking|reference|ref)\s*(?:number|no\.?|#)?\s*[:\-]?\s*([A-Z]?\d[\w*-]{3,})\b/i,
    /\border\s*#\s*([A-Z0-9*-]{3,})\b/i,
    /\b(W\d{4,})\b/,
    /\b(O\d{6,})\b/,
    /\b#([A-Z0-9*-]{4,})\b/,
    /\b(\d{6,}\*\d+)\b/
  ];

  for (let s = 0; s < sources.length; s += 1) {
    const source = String(sources[s] || '');
    if (!source) continue;

    for (let p = 0; p < patterns.length; p += 1) {
      const match = source.match(patterns[p]);
      if (match && match[1]) {
        return normalizeReferenceForFilename_(match[1]);
      }
    }
  }

  return '';
}

function normalizeReferenceForFilename_(value) {
  return String(value || '')
    .replace(/^#/, '')
    .replace(/\*/g, 'x')
    .replace(/[^\w-]+/g, '')
    .trim();
}

function buildAttachmentOutputName_(context, attachment) {
  const originalName =
    sanitizeFilename_(attachment.filename) ||
    'attachment.' + (guessExtensionFromMime_(attachment.mimeType) || 'bin');

  const ref = context.orderNumber || context.shortId;
  return context.datePart + '-' + context.vendor + '-' + ref + '-' + originalName;
}

/* =========================
 * GENERAL UTILITIES
 * ========================= */

function truncate_(text, maxLength) {
  const value = String(text || '');
  if (value.length <= maxLength) return value;
  return value.slice(0, maxLength) + '\n\n...[truncated]';
}

function assertAdvancedGmailEnabled_() {
  if (
    typeof Gmail === 'undefined' ||
    !Gmail ||
    !Gmail.Users ||
    !Gmail.Users.Messages ||
    typeof Gmail.Users.Messages.get !== 'function'
  ) {
    throw new Error(
      'Advanced Gmail service is not enabled. In Apps Script, open Services and add Gmail.'
    );
  }
}

function mergeObjects_(base, extra) {
  const result = {};
  Object.keys(base || {}).forEach(function (key) {
    result[key] = base[key];
  });
  Object.keys(extra || {}).forEach(function (key) {
    result[key] = extra[key];
  });
  return result;
}

/* =========================
 * TRIGGER / RESET
 * ========================= */

function installWeeklyTrigger() {
  // Remove older copies of the same trigger first.
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'runReceiptIngestion') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Install a new weekly trigger for Thursday at 5 PM.
  ScriptApp.newTrigger('runReceiptIngestion')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.THURSDAY)
    .atHour(17)
    .create();

  Logger.log('Weekly trigger installed for Thursday around 5 PM.');
}

function resetTestProcessedState() {
  const config = getConfig_();

  const processedLabel = GmailApp.getUserLabelByName(config.processedLabel);
  const reviewLabel = GmailApp.getUserLabelByName(config.reviewLabel);
  const ignoredLabel = GmailApp.getUserLabelByName(config.ignoredLabel);

  let updatedThreads = 0;

  config.labels.forEach(function (labelName) {
    const query = 'label:' + labelName;
    const threads = GmailApp.search(query, 0, 200);

    threads.forEach(function (thread) {
      if (processedLabel) thread.removeLabel(processedLabel);
      if (reviewLabel) thread.removeLabel(reviewLabel);
      if (ignoredLabel) thread.removeLabel(ignoredLabel);
      updatedThreads += 1;
    });
  });

  PropertiesService.getScriptProperties().deleteProperty('PROCESSED_MESSAGE_IDS');
  PropertiesService.getScriptProperties().deleteProperty('PROCESSED_RECEIPT_KEYS');
  PropertiesService.getScriptProperties().deleteProperty('REVIEWED_MESSAGE_IDS');
  PropertiesService.getScriptProperties().deleteProperty('IGNORED_NON_RECEIPT_MESSAGE_IDS');

  Logger.log('Reset complete. Updated threads: ' + updatedThreads);
  Logger.log(
    'Removed processed/review/ignored labels and cleared PROCESSED_MESSAGE_IDS, PROCESSED_RECEIPT_KEYS, REVIEWED_MESSAGE_IDS, and IGNORED_NON_RECEIPT_MESSAGE_IDS.'
  );
}
