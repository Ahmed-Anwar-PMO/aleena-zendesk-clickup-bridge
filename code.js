/* ===== Zendesk ↔ ClickUp Bridge (Complete Sync) =====
   Script Properties (required):
     CLICKUP_TOKEN
     ZENDESK_EMAIL
     ZENDESK_API_TOKEN
     GDRIVE_FOLDER_ID
   Optional:
     SHARED_KEY, DEBUG_SHEET_ID, CLICKUP_REOPEN_STATUS
*/

const SP = PropertiesService.getScriptProperties();
const CLICKUP_TOKEN = SP.getProperty('CLICKUP_TOKEN');
const ZD_EMAIL = SP.getProperty('ZENDESK_EMAIL');
const ZD_API_TOKEN = SP.getProperty('ZENDESK_API_TOKEN');
const SHARED_KEY = SP.getProperty('SHARED_KEY') || '';
const DEFAULT_DEBUG_SHEET_ID = '1jCs0MWqY8VW2vRREYc22Rxjg0X_GzFBA2EtOFhPd-yY';
const DEBUG_SHEET_ID =
  SP.getProperty('DEBUG_SHEET_ID') || DEFAULT_DEBUG_SHEET_ID;
const GDRIVE_FOLDER_ID = SP.getProperty('GDRIVE_FOLDER_ID');

const CLICKUP_LIST_FORWARD = '901811203589';
const CLICKUP_LIST_REVERSE = '901811777405';
const ASSIGNEE_EMAIL = 'qrasool@tgc-ksa.com';

// "Reopened" status in ClickUp (hard default: REOPENED)
const CLICKUP_REOPEN_STATUS =
  SP.getProperty('CLICKUP_REOPEN_STATUS') || 'REOPENED';

const CACHE_TTL_SECONDS = 6 * 60 * 60;
const CREATED_FLAG_TTL = 24 * 60 * 60;

const STATUS_STATE_PREFIX = 'CLICKUP_STATUS::';
const ZD_LINK_PROPERTY_PREFIX = 'CLICKUP_ZD_LINK::';
const ZD_LINK_CACHE_PREFIX = 'cu:task:zd_link:';

let cachedDebugSheet = null;
let cachedDebugSheetId = '';

/* ================= Logging ================= */
function logToSheet(eventType, status, details) {
  try {
    if (!DEBUG_SHEET_ID) return;
    if (!cachedDebugSheet || cachedDebugSheetId !== DEBUG_SHEET_ID) {
      cachedDebugSheet = SpreadsheetApp.openById(
        DEBUG_SHEET_ID
      ).getSheets()[0];
      cachedDebugSheetId = DEBUG_SHEET_ID;
    }
    const ts = new Date().toISOString();
    const d = typeof details === 'object' ? JSON.stringify(details) : String(details);
    cachedDebugSheet.appendRow([ts, eventType, status, d]);
  } catch (e) {
    const msg = 'sheet log fail: ' + e;
    Logger.log(msg);
    try {
      if (Array.isArray(executionLogs)) executionLogs.push(msg);
    } catch (_) {}
    try {
      SP.setProperty('LAST_SHEET_LOG_ERROR', `${new Date().toISOString()} ${msg}`);
    } catch (_) {}
  }
}

var executionLogs = [];
function logBoth(msg) {
  executionLogs.push(msg);
  Logger.log(msg);
  console.log(msg);
}

/* ================= Entry ================= */
function doPost(e) {
  executionLogs = [];
  try {
    if (!e || !e.postData || !e.postData.contents) {
      logToSheet('ERROR', 'FAILED', 'No postData.contents');
      return respond_({ ok: false, reason: 'No postData' });
    }

    let payload;
    try {
      payload = JSON.parse(e.postData.contents);
    } catch (err) {
      return respond_({ ok: false, reason: 'invalid JSON', error: String(err) });
    }

    logToSheet('WEBHOOK_RECEIVED', 'INFO', {
      payload: JSON.stringify(payload).substring(0, 500),
    });

    // ClickUp v1
    if (payload && payload.event && payload.task) {
      return handleClickUpWebhook_(payload);
    }

    // ClickUp v2 / long_running
    if (payload && payload.trigger_id && payload.payload) {
      const v2 = payload.payload || {};
      const task = normalizeClickUpV2Task_(v2);
      const history = v2.history_items || payload.history_items || [];
      const hasComment =
        Array.isArray(history) &&
        history.some(h => h && (h.field === 'comment' || h.type === 'comment' || h.history_type === 'comment'));
      const eventName =
        hasComment
          ? 'taskCommentPosted'
          : (payload.event || payload.type || inferEventFromV2_(v2) || 'taskUpdated');
      return handleClickUpWebhook_({
        event: eventName,
        task,
        history_items: history,
        payload: v2,
      });
    }

    // Zendesk direct (recommended)
    // Expecting payload to ALREADY contain:
    // ticket_id, ticket_url, ops_reason, comment_body, comment_id, agent_name, account
    if (payload && payload.ticket_id) {
      return handleZendeskWebhook_(payload);
    }

    // Zendesk wrapped (fallback / legacy)
    if (payload && payload.ticket) {
      const t = payload.ticket || {};
      const unwrapped = {
        ticket_id: t.id || payload.ticket_id,
        ticket_url:
          t.url ||
          payload.ticket_url ||
          (t.id ? `https://shopaleena.zendesk.com/agent/tickets/${t.id}` : ''),
        ops_reason:
          t.ops_reason || payload.ops_reason || t.custom_fields_ops_reason || '',
        comment_body:
          t.latest_comment && t.latest_comment.body
            ? t.latest_comment.body
            : (t.latest_comment || t.comment_body || payload.comment_body || ''),
        comment_id:
          (t.latest_comment && t.latest_comment.id) ||
          payload.comment_id ||
          '',
        agent_name:
          (t.latest_comment && t.latest_comment.author && t.latest_comment.author.name) ||
          t.agent_name ||
          payload.agent_name ||
          t.requester_name ||
          '',
        account: t.account || payload.account || 'shopaleena',
      };
      return handleZendeskWebhook_(unwrapped);
    }

    logToSheet('UNRECOGNIZED_PAYLOAD', 'FAILED', {
      keys: Object.keys(payload || {}),
    });
    return respond_({
      ok: false,
      reason: 'unrecognized payload',
      keys: Object.keys(payload || {}),
    });
  } catch (err) {
    logToSheet('FATAL_ERROR', 'FAILED', { error: String(err), stack: err.stack });
    return respond_({ ok: false, error: String(err) });
  }
}

/* =============== Zendesk → ClickUp ================== */
function handleZendeskWebhook_(p) {
  if (p.ticket_url && !p.ticket_url.startsWith('http')) {
    p.ticket_url = 'https://' + p.ticket_url;
  }

  // comment_id is REQUIRED so each note is treated as a unique event
  const need = ['ticket_id', 'ticket_url', 'ops_reason', 'comment_body', 'comment_id'];
  const missing = need.filter(
    k => !(k in p) || p[k] === undefined || p[k] === null || p[k] === ''
  );
  if (missing.length) {
    return respond_({
      ok: false,
      side: 'zd->cu',
      reason: 'missing',
      missing,
    });
  }

  const noteBody = String(p.comment_body || '');
  const commentId = String(p.comment_id || '').trim();

  // Loop guard: ignore our own notes
  if (isOurOwnZendeskNote_(noteBody)) {
    logToSheet('ZENDESK_NOTE_SKIPPED', 'INFO', {
      preview: noteBody.slice(0, 120),
    });
    return respond_({ ok: true, side: 'zd->cu', skipped: true });
  }

  // Must match "clickup - <ORDER> - <note>"
  const m = noteBody.match(/^\s*clickup\s*-\s*([^-\n]+?)\s*-\s*([\s\S]+)$/i);
  if (!m) {
    return respond_({
      ok: false,
      side: 'zd->cu',
      reason: 'bad note format',
      comment_body: noteBody,
    });
  }

  const orderNumber = normalizeOrderToken_(m[1]);
  const noteText = m[2].trim();
  const subdomain = inferSubdomain_(p);
  if (!subdomain) {
    return respond_({
      ok: false,
      side: 'zd->cu',
      reason: 'no subdomain',
    });
  }

  const opsReason = String(p.ops_reason || '');
  let listId = null;
  if (/^fwd[\s_-]/i.test(opsReason)) {
    listId = CLICKUP_LIST_FORWARD;
  } else if (/^rev[\s_-]/i.test(opsReason)) {
    listId = CLICKUP_LIST_REVERSE;
  }
  if (!listId) {
    return respond_({
      ok: false,
      side: 'zd->cu',
      reason: 'invalid ops_reason prefix',
      ops_reason: opsReason,
    });
  }

  // Pull author + attachments for THIS SPECIFIC comment (by comment_id)
  let attachments = [];
  let agentName =
    normalizeFirstLast_(p.agent_name || '') || 'Unknown Agent';
  let zdCommentId = commentId || null;

  if (zdCommentId) {
    try {
      const c = getLastInternalClickupComment_(
        subdomain,
        p.ticket_id,
        zdCommentId
      );
      if (c) {
        attachments = c.attachments || [];
        agentName =
          normalizeFirstLast_(c.author_name || p.agent_name || '') ||
          agentName ||
          'Unknown Agent';
      }
    } catch (err) {
      // If Zendesk comments API fails, we just continue without attachments
      logToSheet('ZD_COMMENT_FETCH_FAILED', 'FAILED', {
        ticketId: p.ticket_id,
        commentId: zdCommentId,
        error: String(err),
      });
    }
  }

  // Dedupe this Zendesk comment (per comment_id) via cache
  const cache = CacheService.getScriptCache();
  if (zdCommentId) {
    const dk = `zd:${p.ticket_id}:c:${zdCommentId}`;
    if (cache.get(dk)) {
      return respond_({
        ok: true,
        side: 'zd->cu',
        deduped: true,
      });
    }
  }

  /* === Find or create ClickUp task === */

  let task;
  let isNewTask = false;

  const description = [
    `🔗 Zendesk Ticket: ${p.ticket_url}`,
    `🎯 Ops Escalation Reason: ${opsReason}`,
    `👤 Sent by: ${agentName || 'Unknown Agent'}`,
    `📝 Notes:`,
    `- ${noteText}`,
  ].join('\n');

  try {
    const existing = findTaskByNameInLists_(orderNumber, [
      CLICKUP_LIST_FORWARD,
      CLICKUP_LIST_REVERSE,
    ]);

    if (!existing) {
      // New task
      task = createClickUpTask_({
        listId,
        name: orderNumber,
        description,
        tags: [opsReason],
        assigneeEmail: ASSIGNEE_EMAIL,
      });
      isNewTask = true;
    } else {
      task = existing;
      try {
        const cur = getClickUpTask_(task.id);

        if ((cur.description || '') !== description) {
          updateClickUpTaskDescription_(task.id, description);
        }

        const tagNames = (cur.tags || []).map(t => t.name);
        if (!(tagNames.length === 1 && tagNames[0] === opsReason)) {
          replaceClickUpTaskTags_(task.id, opsReason);
        }
      } catch (_) {}
    }
  } catch (err) {
    return respond_({
      ok: false,
      side: 'zd->cu',
      reason: 'create/find failed',
      error: String(err),
    });
  }

  // Remember link task ↔ ticket
  rememberZendeskLinkForTask_(task.id, {
    ticketId: p.ticket_id,
    subdomain,
    ticketUrl:
      p.ticket_url || buildZendeskTicketUrl_(subdomain, p.ticket_id),
  });

  /* === Add ClickUp context comment (created / updated) === */

  const ctxHeader = isNewTask ? 'ClickUp task created' : 'ClickUp task updated';

  const ctx = [
    `${ctxHeader}`,
    `From Zendesk #${p.ticket_id}`,
    `By: ${agentName || 'Unknown Agent'}`,
    `Ops Reason: ${opsReason}`,
    `Note: ${noteText}`,
    p.ticket_url ? `Link: ${p.ticket_url}` : '',
    `Zendesk Comment ID: ${zdCommentId || 'n/a'}`,
  ]
    .filter(Boolean)
    .join('\n');

  try {
    addClickUpTaskComment_(task.id, ctx);
  } catch (_) {}

  /* === Reopen Zendesk ticket + ClickUp task if ClickUp was COMPLETE === */

  try {
    // Prefer live status from ClickUp API; fall back to cached last known
    let lastStatus = '';
    try {
      const live = getClickUpTask_(task.id);
      const st = live.status || {};
      lastStatus =
        typeof st === 'string' ? st : (st.status || '');
    } catch (_) {
      lastStatus = getLastKnownClickUpStatus_(task.id);
    }

    const norm = normalizeStatusLabel_(lastStatus);

    if (norm && /COMPLETE|COMPLETED|DONE|RESOLVED|CLOSED/.test(norm)) {
      // 1) Reopen Zendesk ticket
      try {
        updateZendeskStatus_({ subdomain }, p.ticket_id, 'new');
        logToSheet('ZD_REOPEN_AFTER_NEW_COMMENT', 'SUCCESS', {
          ticketId: p.ticket_id,
          taskId: task.id,
          lastClickUpStatus: lastStatus,
        });
      } catch (errZ) {
        logToSheet('ZD_REOPEN_AFTER_NEW_COMMENT', 'FAILED', {
          ticketId: p.ticket_id,
          taskId: task.id,
          error: String(errZ),
        });
      }

      // 2) Reopen ClickUp task → REOPENED
      try {
        updateClickUpTaskStatus_(task.id, CLICKUP_REOPEN_STATUS);
        rememberClickUpStatus_(task.id, CLICKUP_REOPEN_STATUS);
        logToSheet('CU_REOPEN_AFTER_ZD_NOTE', 'SUCCESS', {
          taskId: task.id,
          fromStatus: lastStatus,
          toStatus: CLICKUP_REOPEN_STATUS,
        });
      } catch (errCU) {
        logToSheet('CU_REOPEN_AFTER_ZD_NOTE', 'FAILED', {
          taskId: task.id,
          fromStatus: lastStatus,
          toStatus: CLICKUP_REOPEN_STATUS,
          error: String(errCU),
        });
      }
    }
  } catch (err) {
    logToSheet('REOPEN_FLOW_ERROR', 'FAILED', {
      ticketId: p.ticket_id,
      taskId: task.id,
      error: String(err),
    });
  }

  /* === Attachments handling (for THIS comment only) === */

  for (const att of attachments || []) {
    try {
      const blob = downloadZendeskAttachment_(
        subdomain,
        att.content_url,
        att.file_name
      );
      uploadAttachmentToClickUp_(task.id, blob);
    } catch (err) {
      try {
        addClickUpTaskComment_(
          task.id,
          `Attachment (fallback link): ${att.file_name}\n${att.content_url}`
        );
      } catch (_) {}
    }
  }

  if (zdCommentId) {
    cache.put(`zd:${p.ticket_id}:c:${zdCommentId}`, '1', CACHE_TTL_SECONDS);
  }

  return respond_({
    ok: true,
    side: 'zd->cu',
    task,
    isNewTask,
  });
}

/* =============== ClickUp → Zendesk ================== */
function handleClickUpWebhook_(p) {
  const event = p.event || '';
  const task = p.task || {};
  const statusObj = task.status || {};
  const status =
    typeof statusObj === 'string'
      ? statusObj
      : (statusObj.status || '');
  const taskId = task.id;
  const taskUrl = task.url || '';
  const name = task.name || '';
  const description = task.description || '';

  logToSheet('CU_WEBHOOK_START', 'INFO', {
    event,
    taskId,
    status,
  });

  // Extract Zendesk ticket id + subdomain from description or stored mapping
  let link = extractZendeskInfo_(description);
  if (!link && taskId) {
    const stored = getZendeskLinkForTask_(taskId);
    if (stored) {
      link = stored;
      logToSheet('CU_ZD_LINK_RECOVERED', 'INFO', {
        taskId,
        source: stored.source || 'property',
      });
    }
  }

  if (!link) {
    logToSheet('CU_NO_ZD_LINK', 'FAILED', {
      taskId,
      has_description: !!description,
      desc_length: (description || '').length,
    });
    return respond_({
      ok: false,
      side: 'cu->zd',
      reason: 'no zendesk link',
      taskId,
    });
  }

  const { ticketId, subdomain, ticketUrl } = link;
  const cache = CacheService.getScriptCache();

  rememberZendeskLinkForTask_(taskId, {
    ticketId,
    subdomain,
    ticketUrl: ticketUrl || buildZendeskTicketUrl_(subdomain, ticketId),
  });

  // Extract user who made the change
  let userName = 'Operations Agent';
  try {
    const hist = p.history_items || [];
    for (let i = hist.length - 1; i >= 0; i--) {
      const h = hist[i] || {};
      const u = h.user || h.author || h.member || {};
      if (u.username || u.name || u.email) {
        userName = u.username || u.name || u.email;
        break;
      }
    }
  } catch (_) {}

  const normalizedStatus = normalizeStatusLabel_(status);
  const statusChange = detectClickUpStatusChange_({
    payload: p,
    taskId,
    event,
    normalizedStatus,
    rawStatus: status,
  });

  if (statusChange) {
    const statusResult = processClickUpStatusChange_({
      taskId,
      ticketId,
      subdomain,
      taskUrl,
      userName,
      change: statusChange,
    });
    return respond_(statusResult);
  } else if (normalizedStatus) {
    ensureStatusCacheBaseline_(taskId, normalizedStatus);
  }

  // Handle task creation note (only once)
  const createdKey = `cu:task:${taskId}:created_note_sent`;
  if (event === 'taskCreated' && !cache.get(createdKey)) {
    const body = [
      `✅ ClickUp task created`,
      `Task #${taskId} — ${name}`,
      taskUrl || '',
      status ? `Status: ${status}` : '',
    ]
      .filter(Boolean)
      .join('\n');

    try {
      addZendeskInternalNote_({ subdomain }, ticketId, body);
      logToSheet('TASK_CREATED', 'SUCCESS', {
        taskId,
        ticketId,
      });
    } catch (err) {
      logToSheet('TASK_CREATED', 'FAILED', {
        taskId,
        ticketId,
        error: String(err),
      });
    }
    cache.put(createdKey, '1', CREATED_FLAG_TTL);
  }

  // Mirror comments
  const fromWebhook = extractLatestCommentFromWebhook_(p);
  const latest = fromWebhook || fetchLatestHumanClickUpComment_(taskId);
  if (latest) {
    const dedupeKey = `cu:task:${taskId}:comment:${latest.id}`;
    if (!cache.get(dedupeKey)) {
      const safeText = latest.text && latest.text.trim() ? latest.text : '(no text)';
      const lines = [
        `ClickUp update | ${name}`,
        `Comment added by: ${latest.userName}`,
        `Comment: ${safeText}`,
      ];
      if (taskUrl) lines.push(`Task: ${taskUrl}`);
      (latest.attachments || []).forEach(a => {
        if (a.url) {
          lines.push(`Attachment URL: ${a.url}`);
        }
      });

      try {
        addZendeskInternalNote_({ subdomain }, ticketId, lines.join('\n'));
        logToSheet('COMMENT_ADDED', 'SUCCESS', {
          taskId,
          ticketId,
          commentId: latest.id,
          user: latest.userName,
        });
      } catch (err) {
        logToSheet('COMMENT_FAILED', 'FAILED', {
          taskId,
          ticketId,
          error: String(err),
        });
      }
      cache.put(dedupeKey, '1', CACHE_TTL_SECONDS);
    }
  }

  return respond_({
    ok: true,
    side: 'cu->zd',
    event,
    ticketId,
  });
}

/* =============== Status Change Parsing =============== */
function extractStatusChangeFromWebhook_(p) {
  try {
    const hist = p.history_items || [];
    for (let i = hist.length - 1; i >= 0; i--) {
      const h = hist[i] || {};
      const isStatus =
        h.field === 'status' ||
        h.type === 'status' ||
        h.history_type === 'status';
      if (!isStatus) continue;

      const from =
        (h.before && (h.before.status || h.before)) ||
        h.from ||
        h.old ||
        h.value_before ||
        '';
      const to =
        (h.after && (h.after.status || h.after)) ||
        h.to ||
        h.new ||
        h.value_after ||
        '';

      const id = String(h.id || h.history_id || Date.now());
      const fromStr = typeof from === 'string' ? from : (from && from.status) || '';
      const toStr = typeof to === 'string' ? to : (to && to.status) || '';
      if (toStr) {
        return { id, from: fromStr, to: toStr };
      }
    }
    return null;
  } catch (e) {
    logToSheet('EXTRACT_STATUS_ERROR', 'FAILED', { error: String(e) });
    return null;
  }
}

function detectClickUpStatusChange_({
  payload,
  taskId,
  event,
  normalizedStatus,
  rawStatus,
}) {
  try {
    const cache = CacheService.getScriptCache();
    const historyChange = extractStatusChangeFromWebhook_(payload);
    if (historyChange && taskId) {
      const toNormalized = normalizeStatusLabel_(historyChange.to);
      if (toNormalized) {
        const histId = historyChange.id
          ? statusHistoryCacheKey_(taskId, historyChange.id)
          : '';
        if (histId) {
          if (cache.get(histId)) return null;
          cache.put(histId, '1', CACHE_TTL_SECONDS);
        }
        const fromNormalized =
          normalizeStatusLabel_(historyChange.from) ||
          getLastKnownClickUpStatus_(taskId) ||
          'UNKNOWN';
        return {
          source: 'history',
          from: fromNormalized,
          to: toNormalized,
          rawFrom: historyChange.from || '',
          rawTo: historyChange.to || '',
        };
      }
    }

    if (!taskId || !normalizedStatus) return null;

    const lastKnown = getLastKnownClickUpStatus_(taskId);

    if (lastKnown) {
      if (lastKnown === normalizedStatus) {
        return null;
      }
      return {
        source: event || 'status',
        from: lastKnown,
        to: normalizedStatus,
        rawFrom: lastKnown,
        rawTo: rawStatus || '',
      };
    }

    const eventImpliesStatusUpdate =
      event === 'taskStatusUpdated' ||
      event === 'taskUpdated' ||
      event === 'taskStatusChanged';

    if (!eventImpliesStatusUpdate && normalizedStatus !== 'COMPLETE') {
      return null;
    }

    return {
      source: event || 'status',
      from: 'UNKNOWN',
      to: normalizedStatus,
      rawFrom: '',
      rawTo: rawStatus || '',
    };
  } catch (err) {
    logToSheet('STATUS_DETECT_FAILED', 'FAILED', {
      error: String(err),
    });
    return null;
  }
}

function processClickUpStatusChange_({
  taskId,
  ticketId,
  subdomain,
  taskUrl,
  userName,
  change,
}) {
  const readableFrom = change.rawFrom || change.from || 'UNKNOWN';
  const readableTo = change.rawTo || change.to;
  const result = {
    ok: true,
    side: 'cu->zd',
    action: 'status_updated',
    taskId,
    ticketId,
    from: readableFrom,
    to: readableTo,
    source: change.source,
  };

  const lines = [
    `ClickUp status changed (${change.source})`,
    `From: ${readableFrom}`,
    `To: ${readableTo}`,
    `Updated by: ${userName}`,
  ];
  if (taskUrl) lines.push(`Task: ${taskUrl}`);
  const statusNote = lines.join('\n');
  const errors = [];

  try {
    addZendeskInternalNote_({ subdomain }, ticketId, statusNote);
    logToSheet('STATUS_COMMENT_ADDED', 'SUCCESS', {
      taskId,
      ticketId,
      from: readableFrom,
      to: readableTo,
      user: userName,
    });
  } catch (err) {
    errors.push(`note:${err}`);
    logToSheet('STATUS_COMMENT_FAILED', 'FAILED', {
      taskId,
      ticketId,
      error: String(err),
    });
  }

  const targetZendeskStatus = mapClickUpStatusToZendesk_(change.to);
  if (targetZendeskStatus) {
    try {
      updateZendeskStatus_({ subdomain }, ticketId, targetZendeskStatus);
      logToSheet('ZD_STATUS_UPDATED', 'SUCCESS', {
        taskId,
        ticketId,
        zdStatus: targetZendeskStatus,
      });
      result.zendesk_status = targetZendeskStatus;
    } catch (err) {
      errors.push(`status:${err}`);
      logToSheet('ZD_STATUS_UPDATE_FAILED', 'FAILED', {
        taskId,
        ticketId,
        error: String(err),
      });
    }
  } else {
    logToSheet('ZD_STATUS_SKIPPED', 'INFO', {
      taskId,
      ticketId,
      clickUpStatus: change.to,
    });
  }

  rememberClickUpStatus_(taskId, change.to);

  if (errors.length) {
    result.ok = false;
    result.errors = errors;
  }
  return result;
}

/* ================= Comment Extraction ================= */
function extractLatestCommentFromWebhook_(p) {
  try {
    const hist = p.history_items || [];
    for (let i = hist.length - 1; i >= 0; i--) {
      const h = hist[i] || {};
      const isComment =
        h.field === 'comment' ||
        h.type === 'comment' ||
        h.history_type === 'comment';
      if (!isComment) continue;

      const raw =
        (h.comment_text != null ? h.comment_text : '') ||
        (h.comment != null ? h.comment : '') ||
        (h.value != null ? h.value : '') ||
        (h.text != null ? h.text : '') ||
        (h.text_content != null ? h.text_content : '') ||
        (h.html_text != null ? h.html_text : '');
      let text = String(raw)
        .replace(/<br\s*\/?>/gi, '\n')
        .replace(/<\/p>/gi, '\n')
        .replace(/<[^>]+>/g, '')
        .replace(/\u00A0/g, ' ')
        .trim();

      const u = h.user || h.author || h.member || {};
      const userName = u.username || u.name || u.email || 'Operations Agent';
      const id = String(h.id || h.history_id || h.comment_id || Date.now());

      let atts = []
        .concat(h.attachments || [])
        .concat(h.comment_attachments || [])
        .concat(h.files || [])
        .map(normalizeAttachmentLoose_)
        .filter(Boolean);

      const task = p.task || p.payload || {};
      if (task.attachments && Array.isArray(task.attachments)) {
        task.attachments.forEach(att => {
          if (att.parent_id === id || att.parent_id === h.id) {
            const url = att.url || att.url_w_query || att.url_w_host;
            if (url) atts.push({ url });
          }
        });
      }

      const markdownRegex = /!\[[^\]]*]\((.*?)\)/g;
      let m;
      while ((m = markdownRegex.exec(text)) !== null) {
        if (m[1]) atts.push({ url: m[1] });
      }
      text = text.replace(markdownRegex, '').trim();

      text = text
        .replace(/\S*%[0-9A-Fa-f]{2}\S*/g, '')
        .replace(
          /\b\S+\.(png|jpg|jpeg|gif|pdf|doc|docx|xls|xlsx|zip|rar|txt|mp4|mov|avi|csv)\b/gi,
          ''
        )
        .replace(/\s+/g, ' ')
        .trim();

      return {
        id,
        text,
        userName,
        attachments: dedupeUrls_(atts),
      };
    }
    return null;
  } catch (err) {
    logBoth(`ERROR in extractLatestCommentFromWebhook_: ${err}`);
    return null;
  }
}

function fetchLatestHumanClickUpComment_(taskId) {
  try {
    const taskData = getClickUpTask_(taskId);
    const taskAttachments = taskData.attachments || [];

    const url = `https://api.clickup.com/api/v2/task/${encodeURIComponent(
      taskId
    )}/comment`;
    const resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: CLICKUP_TOKEN },
      muteHttpExceptions: true,
    });
    if (resp.getResponseCode() >= 300) return null;

    const data = JSON.parse(resp.getContentText());
    const comments = data.comments || [];
    if (!comments.length) return null;

    for (let i = 0; i < comments.length; i++) {
      const c = comments[i] || {};
      const commentId = String(c.id || c.comment_id || '');
      const rawText =
        c.comment_text || c.comment || c.text || c.text_content || c.html_text || '';
      let text = String(rawText)
        .replace(/<br\s*\/?>/gi, ' ')
        .replace(/<\/p>/gi, ' ')
        .replace(/<[^>]+>/g, '')
        .replace(/\u00A0/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();

      const attachments = normalizeClickUpAttachments_(
        (c.attachments || []).concat(c.comment_attachments || []).concat(c.files || [])
      );

      if (commentId && taskAttachments.length) {
        taskAttachments.forEach(att => {
          if (att.parent_id === commentId) {
            const url = att.url || att.url_w_query || att.url_w_host;
            if (url) attachments.push({ url });
          }
        });
      }

      const markdownRegex = /!\[[^\]]*]\((.*?)\)/g;
      let m;
      while ((m = markdownRegex.exec(text)) !== null) {
        if (m[1]) attachments.push({ url: m[1] });
      }
      text = text.replace(markdownRegex, '').trim();

      text = text
        .replace(/\S*%[0-9A-Fa-f]{2}\S*/g, '')
        .replace(
          /\b\S+\.(png|jpg|jpeg|gif|pdf|doc|docx|xls|xlsx|zip|rar|txt|mp4|mov|avi|csv)\b/gi,
          ''
        )
        .replace(/\s+/g, ' ')
        .trim();

      if (/^\s*From Zendesk #/i.test(text)) continue;
      if (/^\s*Attachment \(fallback link\):/i.test(text)) continue;

      const user = c.user || c.author || {};
      const userName = user.username || user.name || user.email || 'Operations Agent';
      const created = Number(c.date || c.date_created || c.created || 0) || Date.now();
      const id = commentId || String(created);

      return {
        id,
        text,
        userName,
        created,
        attachments: dedupeUrls_(attachments),
      };
    }
    return null;
  } catch (err) {
    logBoth(`ERROR in fetchLatestHumanClickUpComment_: ${err}`);
    return null;
  }
}

/* ================= Attachments Helpers ================= */
function normalizeAttachmentLoose_(a) {
  try {
    if (!a) return null;
    const url =
      a.download_url ||
      a.url ||
      a.preview_url ||
      a.link ||
      a.href ||
      (a.file && (a.file.url || a.file.download_url)) ||
      (a.attachment && (a.attachment.url || a.attachment.download_url)) ||
      (typeof a === 'string' ? a : '');
    return url ? { url } : null;
  } catch (_) {
    return null;
  }
}

function normalizeClickUpAttachments_(arr) {
  if (!Array.isArray(arr)) return [];
  return arr
    .map(a => {
      const url =
        a.download_url || a.url || a.preview_url || a.link || a.href || '';
      return { url };
    })
    .filter(x => x.url);
}

function dedupeUrls_(arr) {
  const seen = Object.create(null);
  const out = [];
  for (const a of arr) {
    if (!a || !a.url) continue;
    if (seen[a.url]) continue;
    seen[a.url] = true;
    out.push({ url: a.url });
  }
  return out;
}

function downloadClickUpAttachment_(clickUpUrl) {
  const resp = UrlFetchApp.fetch(clickUpUrl, {
    method: 'get',
    headers: { Authorization: CLICKUP_TOKEN },
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() >= 300) {
    throw new Error(
      `ClickUp attachment download failed: ${resp.getResponseCode()} ${resp.getContentText()}`
    );
  }
  let fileName = 'file';
  try {
    const parts = clickUpUrl.split('?')[0].split('/');
    fileName = decodeURIComponent(parts[parts.length - 1] || 'file');
  } catch (e) {}
  const ct = resp.getHeaders()['Content-Type'] || 'application/octet-stream';
  return Utilities.newBlob(resp.getContent(), ct, fileName);
}

function downloadZendeskAttachment_(subdomain, contentUrl, fileName) {
  const resp = UrlFetchApp.fetch(contentUrl, {
    method: 'get',
    headers: zendeskAuthHeader_(),
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() >= 300) {
    throw new Error(
      `Zendesk attachment fetch failed: ${resp.getResponseCode()} ${resp.getContentText()}`
    );
  }
  const ct = resp.getHeaders()['Content-Type'] || 'application/octet-stream';
  return Utilities.newBlob(resp.getContent(), ct, fileName || 'file');
}

function uploadAttachmentToClickUp_(taskId, blob) {
  const url = `https://api.clickup.com/api/v2/task/${encodeURIComponent(
    taskId
  )}/attachment`;
  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: { Authorization: CLICKUP_TOKEN },
    payload: { attachment: blob },
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() >= 300) {
    throw new Error(
      `ClickUp attachment upload failed: ${resp.getResponseCode()} ${resp.getContentText()}`
    );
  }
}

/* ================= Zendesk API Helpers ================= */
function addZendeskInternalNote_({ subdomain }, ticketId, body) {
  const url = `https://${subdomain}.zendesk.com/api/v2/tickets/${encodeURIComponent(
    ticketId
  )}.json`;
  const payload = { ticket: { comment: { public: false, body } } };
  const resp = UrlFetchApp.fetch(url, {
    method: 'put',
    headers: zendeskAuthHeader_(),
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() >= 300) {
    throw new Error(
      `Zendesk note failed: ${resp.getResponseCode()} ${resp.getContentText()}`
    );
  }
}

function updateZendeskStatus_({ subdomain }, ticketId, status) {
  const url = `https://${subdomain}.zendesk.com/api/v2/tickets/${encodeURIComponent(
    ticketId
  )}.json`;
  const payload = { ticket: { status } };
  const resp = UrlFetchApp.fetch(url, {
    method: 'put',
    headers: zendeskAuthHeader_(),
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() >= 300) {
    throw new Error(
      `Zendesk status update failed: ${resp.getResponseCode()} ${resp.getContentText()}`
    );
  }
}

function zendeskAuthHeader_() {
  const tokenPair = `${ZD_EMAIL}/token:${ZD_API_TOKEN}`;
  const auth = Utilities.base64Encode(tokenPair);
  return { Authorization: `Basic ${auth}` };
}

/* ================= ClickUp Status Helpers ================= */
function updateClickUpTaskStatus_(taskId, newStatus) {
  if (!taskId || !newStatus) return;
  const url = `https://api.clickup.com/api/v2/task/${encodeURIComponent(
    taskId
  )}`;
  const payload = { status: newStatus };
  const resp = UrlFetchApp.fetch(url, {
    method: 'put',
    headers: {
      Authorization: CLICKUP_TOKEN,
      'Content-Type': 'application/json',
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() >= 300) {
    throw new Error(
      `ClickUp status update failed: ${resp.getResponseCode()} ${resp.getContentText()}`
    );
  }
}

function mapClickUpStatusToZendesk_(status) {
  const normalized = normalizeStatusLabel_(status);
  if (!normalized) return null;

  const explicit = {
    'COMPLETE': 'solved',
    'COMPLETED': 'solved',
    'DONE': 'solved',
    'RESOLVED': 'solved',
    'CLOSED': 'solved',
    'READY FOR QA': 'pending',
    'READY FOR REVIEW': 'pending',
    'QA REVIEW': 'pending',
    'IN REVIEW': 'pending',
    'IN QA': 'pending',
    'IN PROGRESS': 'open',
    'ACTIVE': 'open',
    'TO DO': 'open',
    'BACKLOG': 'new',
    'ON HOLD': 'hold',
    'HOLD': 'hold',
    'BLOCKED': 'hold',
  };

  if (explicit[normalized]) return explicit[normalized];
  if (/COMPLETE|DONE|RESOLVED/.test(normalized)) return 'solved';
  if (/HOLD|BLOCK/.test(normalized)) return 'hold';
  if (/REVIEW|QA/.test(normalized)) return 'pending';
  return 'open';
}

function normalizeStatusLabel_(status) {
  if (!status) return '';
  return String(status)
    .replace(/\u00A0/g, ' ')
    .replace(/[_-]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

function rememberClickUpStatus_(taskId, status) {
  if (!taskId) return;
  const normalized = normalizeStatusLabel_(status);
  const cacheKey = statusCacheKey_(taskId);
  CacheService.getScriptCache().put(cacheKey, normalized, CACHE_TTL_SECONDS);
  if (normalized) {
    SP.setProperty(statusStatePropKey_(taskId), normalized);
  } else {
    SP.deleteProperty(statusStatePropKey_(taskId));
  }
}

function ensureStatusCacheBaseline_(taskId, normalizedStatus) {
  if (!taskId || !normalizedStatus) return;
  if (!getLastKnownClickUpStatus_(taskId)) {
    rememberClickUpStatus_(taskId, normalizedStatus);
  }
}

function getLastKnownClickUpStatus_(taskId) {
  if (!taskId) return '';
  const cacheKey = statusCacheKey_(taskId);
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if (cached !== null) return cached;
  const stored = SP.getProperty(statusStatePropKey_(taskId)) || '';
  if (stored) {
    cache.put(cacheKey, stored, CACHE_TTL_SECONDS);
  }
  return stored;
}

function clearKnownClickUpStatusForQa_(taskId) {
  if (!taskId) return;
  CacheService.getScriptCache().put(statusCacheKey_(taskId), '', 1);
  SP.deleteProperty(statusStatePropKey_(taskId));
}

function statusCacheKey_(taskId) {
  return `cu:task:${taskId}:last_status`;
}

function statusStatePropKey_(taskId) {
  return `${STATUS_STATE_PREFIX}${taskId}`;
}

function statusHistoryCacheKey_(taskId, historyId) {
  return `cu:task:${taskId}:status_hist:${historyId}`;
}

/* ================= Zendesk Link Helpers ================= */
function rememberZendeskLinkForTask_(taskId, data) {
  try {
    if (!taskId || !data || !data.ticketId) return;
    const payload = {
      ticketId: String(data.ticketId),
      subdomain: data.subdomain || '',
      ticketUrl:
        data.ticketUrl || buildZendeskTicketUrl_(data.subdomain, data.ticketId),
    };
    const serialized = JSON.stringify(payload);
    CacheService.getScriptCache().put(
      zdLinkCacheKey_(taskId),
      serialized,
      CACHE_TTL_SECONDS
    );
    SP.setProperty(zdLinkPropKey_(taskId), serialized);
  } catch (err) {
    logToSheet('ZD_LINK_STORE_FAILED', 'FAILED', {
      taskId,
      error: String(err),
    });
  }
}

function getZendeskLinkForTask_(taskId) {
  if (!taskId) return null;
  const cacheKey = zdLinkCacheKey_(taskId);
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if (cached) {
    const parsed = parseZendeskLinkPayload_(cached);
    if (parsed) {
      parsed.source = 'cache';
      return parsed;
    }
  }
  const stored = SP.getProperty(zdLinkPropKey_(taskId));
  if (!stored) return null;
  cache.put(cacheKey, stored, CACHE_TTL_SECONDS);
  const parsed = parseZendeskLinkPayload_(stored);
  if (parsed) {
    parsed.source = parsed.source || 'property';
    return parsed;
  }
  return null;
}

function clearZendeskLinkForTask_(taskId) {
  if (!taskId) return;
  CacheService.getScriptCache().put(zdLinkCacheKey_(taskId), '', 1);
  SP.deleteProperty(zdLinkPropKey_(taskId));
}

function parseZendeskLinkPayload_(value) {
  try {
    if (!value) return null;
    const parsed = typeof value === 'string' ? JSON.parse(value) : value;
    if (!parsed || !parsed.ticketId) {
      return null;
    }
    const subdomain =
      parsed.subdomain || inferSubdomain_({ ticket_url: parsed.ticketUrl });
    const ticketUrl =
      parsed.ticketUrl || buildZendeskTicketUrl_(subdomain, parsed.ticketId);
    return {
      ticketId: String(parsed.ticketId),
      subdomain: subdomain || '',
      ticketUrl,
      source: parsed.source || '',
    };
  } catch (err) {
    logToSheet('ZD_LINK_PARSE_FAILED', 'FAILED', { error: String(err) });
    return null;
  }
}

function zdLinkCacheKey_(taskId) {
  return `${ZD_LINK_CACHE_PREFIX}${taskId}`;
}

function zdLinkPropKey_(taskId) {
  return `${ZD_LINK_PROPERTY_PREFIX}${taskId}`;
}

/* ================= Misc Helpers ================= */
function inferEventFromV2_(v2) {
  const dc = Number(v2.date_created || 0);
  const du = Number(v2.date_updated || 0);
  if (dc && du && Math.abs(du - dc) <= 5000) {
    return 'taskCreated';
  }
  return null;
}

function normalizeClickUpV2Task_(v2) {
  const id =
    v2.id ||
    v2.task_id ||
    (v2.task && (v2.task.id || v2.task.task_id)) ||
    '';
  const name = v2.name || (v2.task && v2.task.name) || '';
  const description =
    v2.description ||
    v2.text_content ||
    (v2.task && (v2.task.description || v2.task.text_content)) ||
    '';
  const url = v2.url || (v2.task && v2.task.url) || '';
  let status = v2.status || (v2.task && v2.task.status) || '';
  if (typeof status === 'string') status = { status };
  return { id, name, description, url, status };
}

function inferSubdomain_(p) {
  if (p.account) return String(p.account);
  if (p.ticket_url) {
    const m = String(p.ticket_url).match(
      /^https:\/\/([a-z0-9-]+)\.zendesk\.com\//i
    );
    if (m) return m[1];
  }
  return '';
}

function buildZendeskTicketUrl_(subdomain, ticketId, fallbackUrl) {
  if (subdomain && ticketId) {
    return `https://${subdomain}.zendesk.com/agent/tickets/${ticketId}`;
  }
  return fallbackUrl || '';
}

function extractZendeskInfo_(text) {
  const m = String(text || '').match(
    /https:\/\/([a-z0-9-]+)\.zendesk\.com\/agent\/tickets\/(\d+)/i
  );
  if (!m) return null;
  const ticketUrl = `https://${m[1]}.zendesk.com/agent/tickets/${m[2]}`;
  return {
    subdomain: m[1],
    ticketId: m[2],
    ticketUrl,
    source: 'description',
  };
}

function normalizeFirstLast_(fullName) {
  if (!fullName) return 'Unknown Agent';
  const parts = String(fullName).trim().split(/\s+/);
  return parts.length >= 2 ? `${parts[0]} ${parts[1]}` : parts[0];
}

function isOurOwnZendeskNote_(body) {
  const s = String(body || '').trim();
  return (
    /^✅\s*ClickUp task created/i.test(s) ||
    /^ClickUp task created/i.test(s) ||
    /^ClickUp task updated/i.test(s) ||
    /^ℹ️?\s*ClickUp task updated/i.test(s) ||
    /^ℹ️?\s*ClickUp task status:/i.test(s) ||
    /^Task update\s*\|/i.test(s) ||
    /^Clicup update\s*\|/i.test(s) ||
    /^Clickup update\s*\|/i.test(s) ||
    /^ClickUp update\s*\|/i.test(s) ||
    /^Comment added by:/i.test(s)
  );
}

function normalizeOrderToken_(raw) {
  let s = String(raw || '').replace(/<[^>]*>/g, '');
  s = s.replace(/[\[\]\(\)\{\}<>【】]/g, ' ');
  s = s.replace(/&[a-z]+;/gi, ' ');
  s = s.replace(/[^\w-]+/g, ' ').trim();
  s = s.split(/\s+/)[0] || '';
  return s;
}

/* ================= ClickUp API Helpers ================= */
function listTasks_(listId) {
  const url = `https://api.clickup.com/api/v2/list/${encodeURIComponent(
    listId
  )}/task?include_closed=true&page=0`;
  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: CLICKUP_TOKEN },
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() >= 300) {
    throw new Error(
      `ClickUp list failed: ${resp.getResponseCode()} ${resp.getContentText()}`
    );
  }
  const data = JSON.parse(resp.getContentText());
  return data.tasks || [];
}

function findTaskByNameInLists_(name, listIds) {
  const needle = normalizeOrderToken_((name || '').trim());
  for (const id of listIds) {
    const tasks = listTasks_(id);
    const f = tasks.find(t => normalizeOrderToken_(t.name || '') === needle);
    if (f) {
      return { id: f.id, url: f.url, name: f.name };
    }
  }
  return null;
}

function getClickUpTask_(taskId) {
  const url = `https://api.clickup.com/api/v2/task/${encodeURIComponent(
    taskId
  )}`;
  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: CLICKUP_TOKEN },
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() >= 300) {
    throw new Error(
      `ClickUp task get failed: ${resp.getResponseCode()} ${resp.getContentText()}`
    );
  }
  return JSON.parse(resp.getContentText());
}

function createClickUpTask_({ listId, name, description, tags, assigneeEmail }) {
  const url = `https://api.clickup.com/api/v2/list/${encodeURIComponent(
    listId
  )}/task`;
  const assigneeId = getClickUpUserIdByEmail_(assigneeEmail);
  const payload = {
    name,
    description,
    priority: 3,
    tags: tags || [],
    assignees: assigneeId ? [assigneeId] : [],
  };
  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      Authorization: CLICKUP_TOKEN,
      'Content-Type': 'application/json',
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() >= 300) {
    throw new Error(
      `ClickUp create failed: ${resp.getResponseCode()} ${resp.getContentText()}`
    );
  }
  const d = JSON.parse(resp.getContentText());
  return { id: d.id, url: d.url, name: d.name };
}

function getClickUpUserIdByEmail_(email) {
  const url = `https://api.clickup.com/api/v2/team`;
  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: CLICKUP_TOKEN },
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() >= 300) return null;
  const data = JSON.parse(resp.getContentText());
  for (const team of data.teams || []) {
    for (const member of team.members || []) {
      const u = member.user || {};
      if ((u.email || '').toLowerCase() === String(email).toLowerCase()) {
        return u.id;
      }
    }
  }
  return null;
}

function updateClickUpTaskDescription_(taskId, description) {
  const url = `https://api.clickup.com/api/v2/task/${encodeURIComponent(
    taskId
  )}`;
  const resp = UrlFetchApp.fetch(url, {
    method: 'put',
    headers: {
      Authorization: CLICKUP_TOKEN,
      'Content-Type': 'application/json',
    },
    payload: JSON.stringify({ description }),
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() >= 300) {
    throw new Error(
      `ClickUp update description failed: ${resp.getResponseCode()} ${resp.getContentText()}`
    );
  }
}

function replaceClickUpTaskTags_(taskId, newTag) {
  const getUrl = `https://api.clickup.com/api/v2/task/${encodeURIComponent(
    taskId
  )}`;
  const getResp = UrlFetchApp.fetch(getUrl, {
    method: 'get',
    headers: { Authorization: CLICKUP_TOKEN },
    muteHttpExceptions: true,
  });

  if (getResp.getResponseCode() < 300) {
    const cur = JSON.parse(getResp.getContentText());
    for (const tag of cur.tags || []) {
      try {
        const rm = `https://api.clickup.com/api/v2/task/${encodeURIComponent(
          taskId
        )}/tag/${encodeURIComponent(tag.name)}`;
        UrlFetchApp.fetch(rm, {
          method: 'delete',
          headers: { Authorization: CLICKUP_TOKEN },
          muteHttpExceptions: true,
        });
      } catch (_) {}
    }
  }

  const add = `https://api.clickup.com/api/v2/task/${encodeURIComponent(
    taskId
  )}/tag/${encodeURIComponent(newTag)}`;
  const addResp = UrlFetchApp.fetch(add, {
    method: 'post',
    headers: { Authorization: CLICKUP_TOKEN },
    muteHttpExceptions: true,
  });
  if (addResp.getResponseCode() >= 300) {
    throw new Error(
      `ClickUp add tag failed: ${addResp.getResponseCode()} ${addResp.getContentText()}`
    );
  }
}

function addClickUpTaskComment_(taskId, comment) {
  const url = `https://api.clickup.com/api/v2/task/${encodeURIComponent(
    taskId
  )}/comment`;
  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      Authorization: CLICKUP_TOKEN,
      'Content-Type': 'application/json',
    },
    payload: JSON.stringify({ comment_text: comment }),
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() >= 300) {
    throw new Error(
      `ClickUp comment failed: ${resp.getResponseCode()} ${resp.getContentText()}`
    );
  }
}

/* ================= Zendesk Comment Fetch (per-comment) ================= */
function getLastInternalClickupComment_(subdomain, ticketId, commentId) {
  const url = `https://${subdomain}.zendesk.com/api/v2/tickets/${encodeURIComponent(
    ticketId
  )}/comments.json?include=users`;
  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: zendeskAuthHeader_(),
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() >= 300) {
    throw new Error(
      `Zendesk comments failed: ${resp.getResponseCode()} ${resp.getContentText()}`
    );
  }

  const data = JSON.parse(resp.getContentText());
  const comments = data.comments || [];
  const users = data.users || [];
  const byId = Object.create(null);
  users.forEach(u => (byId[u.id] = u));

  const targetId = String(commentId || '').trim();
  if (!targetId) return null;

  for (let i = comments.length - 1; i >= 0; i--) {
    const c = comments[i];
    if (!c) continue;
    const cid = String(c.id || '').trim();
    if (!cid) continue;
    if (cid !== targetId) continue;

    if (c.public === false) {
      const author = byId[c.author_id] || {};
      c.author_name = author.name || author.email || 'Unknown';
      return c;
    }
  }
  return null;
}

/* ================= HTTP Helpers ================= */
function doGet(e) {
  const info = {
    ok: false,
    method: 'GET',
    has_params:
      !!(e && e.parameter && Object.keys(e.parameter).length),
    params: e && e.parameter ? e.parameter : {},
    hint: 'Use POST + JSON',
  };
  return ContentService.createTextOutput(JSON.stringify(info)).setMimeType(
    ContentService.MimeType.JSON
  );
}

function respond_(obj) {
  obj.execution_logs = executionLogs;
  return ContentService.createTextOutput(
    JSON.stringify(obj, null, 2)
  ).setMimeType(ContentService.MimeType.JSON);
}
