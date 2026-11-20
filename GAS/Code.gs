/**
 * Code.gs - Production Enterprise Edition
 * - Dual independent daily emails (RED & GREEN)
 * - MS Graph API (Shared Mailbox) with automatic MailApp fallback
 * - Proactive Client Secret expiration alert
 * - Dynamic IT_Config driven names, permissions, and recipients
 */

const MAIL_TEST_MODE = false; // Set true only when testing without CC
const SECRET_WARN_DAYS = 30;

function sendDailyFilterEmails() {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const redSs = SpreadsheetApp.getActiveSpreadsheet();
    const redConfig = readConfigMap_(redSs);

    checkSecretExpiration_(redSs, redConfig);
    processSingleSheetReminder_(redSs, redConfig, "RED");

    const greenId = String(redConfig["GREEN_SHEET_ID"] || "").trim();
    if (greenId) {
      const greenSs = SpreadsheetApp.openById(greenId);
      const greenConfig = readConfigMap_(greenSs);
      processSingleSheetReminder_(greenSs, greenConfig, "GREEN");
    }
  } finally {
    lock.releaseLock();
  }
}

function checkSecretExpiration_(ss, config) {
  const props = PropertiesService.getScriptProperties();
  const expRaw = String(props.getProperty("MS_SECRET_EXPIRES_AT") || "").trim();
  if (!expRaw) return;

  const expDate = parseValidDate_(expRaw);
  if (!expDate) return;

  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const expStart = new Date(expDate.getFullYear(), expDate.getMonth(), expDate.getDate()).getTime();
  const daysLeft = Math.floor((expStart - todayStart) / 86400000);

  if (daysLeft <= SECRET_WARN_DAYS) {
    const itRecipients = splitEmails_(config["ROLE_IT"] || config["TO_EMAIL"]);
    if (itRecipients.length === 0) return;

    const expStr = Utilities.formatDate(expDate, tz, "yyyy-MM-dd");
    const subject = daysLeft < 0
      ? `URGENT: MS Graph Client Secret EXPIRED (${expStr})`
      : `Action Required: MS Graph Client Secret expires in ${daysLeft} day(s) (${expStr})`;

    const body = `
      <html>
        <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif; color: #222;">
          <h3>MS Graph Client Secret Renewal Reminder</h3>
          <p><b>Expiration Date:</b> ${escapeHtml_(expStr)}</p>
          <p><b>Status:</b> ${daysLeft < 0 ? `Expired ${Math.abs(daysLeft)} day(s) ago` : `${daysLeft} day(s) remaining`}</p>
          <p>Please generate a new client secret in Microsoft Entra ID and update <code>MS_CLIENT_SECRET</code> and <code>MS_SECRET_EXPIRES_AT</code> in GAS Script Properties.</p>
        </body>
      </html>
    `;
    dispatchEmailWithFallback_(ss, itRecipients, [], subject, body, "SECRET_ALERT");
  }
}

function processSingleSheetReminder_(ss, config, fallbackCode) {
  const sheet = ss.getSheetByName("Sheet 1");
  if (!sheet) throw new Error(`Sheet 1 not found in ${ss.getName()}`);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
  const colMap = buildHeaderIndexMap_(headers);
  const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();

  const sheetCode = String(config["SHEET_CODE"] || fallbackCode).trim().toUpperCase();
  const sheetLabel = String(
    config["SHEET_LABEL"] || (sheetCode === "RED" ? "Water Filter Service" : "Itsaso Filter Service")
  ).trim();
  const webAppUrl = ScriptApp.getService().getUrl();

  const pendingItems = [];

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const rowNum = i + 2;

    const name = String(getVal_(r, colMap, ["Name"]) || "").trim();
    const rawStatus = String(getVal_(r, colMap, ["Status"]) || "").trim();
    const notifiedVal = getVal_(r, colMap, ["Notification Status"]);
    const notified = (notifiedVal === true || String(notifiedVal).toUpperCase() === "TRUE");

    if (!name || !rawStatus) continue;
    if (rawStatus.toUpperCase() === "ARCHIVED") continue;

    if (sheetCode === "RED") {
      const usageVal = getVal_(r, colMap, ["Machine usage", "Machine Usage"]);
      const inUse = (usageVal === true || String(usageVal).toUpperCase() === "TRUE");
      if (!inUse) continue;
    }

    const lastChangedDate = parseValidDate_(getVal_(r, colMap, ["Last Changed Date"]));
    const cycleDays = Number(getVal_(r, colMap, ["Change Cycle (Days)"]));
    let nextDateObj = parseValidDate_(getVal_(r, colMap, ["Next Change Date"]));

    if (!nextDateObj && lastChangedDate && !isNaN(cycleDays) && cycleDays > 0) {
      nextDateObj = computeNextChangeDate_(lastChangedDate, cycleDays);
    }

    const effectiveStatus = computeEffectiveStatus_(rawStatus, nextDateObj, todayStart);
    const isTarget = (effectiveStatus === "Pending Notification" || effectiveStatus === "REPLACEMENT REQUIRED");

    if (isTarget && !notified) {
      const idVal = getVal_(r, colMap, ["第 1 列", "ID", "No."]);
      const extraInfo = String(getVal_(r, colMap, ["Filter / Service Name", "Phone Number"]) || "").trim();

      let dueStr = "-";
      let sortTs = Number.MAX_SAFE_INTEGER;
      let overdueDays = 0;

      if (nextDateObj) {
        dueStr = Utilities.formatDate(nextDateObj, tz, "dd/MM/yy");
        const dueStart = new Date(nextDateObj.getFullYear(), nextDateObj.getMonth(), nextDateObj.getDate()).getTime();
        sortTs = dueStart;
        if (effectiveStatus === "REPLACEMENT REQUIRED" && todayStart > dueStart) {
          overdueDays = Math.floor((todayStart - dueStart) / 86400000);
        }
      }

      pendingItems.push({
        id: (idVal !== "" && idVal !== null) ? idVal : (rowNum - 1),
        name: name,
        extraInfo: extraInfo,
        dueStr: dueStr,
        sortTs: sortTs,
        status: effectiveStatus,
        overdueDays: overdueDays,
        isUrgent: (effectiveStatus === "REPLACEMENT REQUIRED")
      });
    }
  }

  pendingItems.sort((a, b) => {
    const urgDiff = Number(b.isUrgent) - Number(a.isUrgent);
    if (urgDiff !== 0) return urgDiff;
    return a.sortTs - b.sortTs;
  });

  const extraColHeader = (sheetCode === "RED") ? "Filter / Service" : "Phone Number";
  const dashboardUrl = String(config["DASHBOARD_URL"] || webAppUrl || "").trim();
  const bossViewUrl = String(config["BOSS_VIEW_URL"] || dashboardUrl).trim();

  const subject = (pendingItems.length === 0)
    ? `[${sheetLabel}] Daily Summary: No Pending Replacements`
    : `[${sheetLabel}] Daily Summary: ${pendingItems.length} Pending`;

  const htmlBody = buildEnglishReminderHtml_({
    sheetCode: sheetCode,
    sheetLabel: sheetLabel,
    extraColHeader: extraColHeader,
    pendingItems: pendingItems,
    dashboardUrl: dashboardUrl,
    bossViewUrl: bossViewUrl,
    webAppUrl: webAppUrl
  });

  const toList = splitEmails_(config["TO_EMAIL"]);
  const ccList = MAIL_TEST_MODE ? [] : splitEmails_(config["CC_EMAILS"]);
  if (toList.length === 0) {
    throw new Error(`No TO_EMAIL configured in IT_Config for ${sheetCode}`);
  }

  dispatchEmailWithFallback_(ss, toList, ccList, subject, htmlBody, sheetCode);
}

function buildEnglishReminderHtml_(ctx) {
  const topButtons = `
    <div style="margin-bottom: 14px;">
      ${ctx.dashboardUrl ? `<a href="${escapeHtml_(ctx.dashboardUrl)}" style="display:inline-block;background:#0d6efd;color:#fff;padding:8px 14px;text-decoration:none;border-radius:6px;font-size:13px;margin-right:8px;">Open Dashboard</a>` : ""}
      ${ctx.bossViewUrl ? `<a href="${escapeHtml_(ctx.bossViewUrl)}" style="display:inline-block;background:#343a40;color:#fff;padding:8px 14px;text-decoration:none;border-radius:6px;font-size:13px;">Boss Overdue View</a>` : ""}
    </div>
  `;

  if (ctx.pendingItems.length === 0) {
    return `
      <html>
        <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif; color: #222;">
          <h2 style="color: #198754; margin-bottom: 8px;">${escapeHtml_(ctx.sheetLabel)} - All Caught Up</h2>
          ${topButtons}
          <p>No active records require replacement notification today.</p>
        </body>
      </html>
    `;
  }

  let rowsHtml = "";
  for (let i = 0; i < ctx.pendingItems.length; i++) {
    const item = ctx.pendingItems[i];
    const rowBg = item.isUrgent ? "background-color: #f8d7da;" : "background-color: #fff3cd;";
    const badgeColor = item.isUrgent ? "#842029" : "#664d03";
    const statusText = item.overdueDays > 0
      ? `${item.status} (${item.overdueDays}d overdue)`
      : item.status;

    const confirmUrl = ctx.webAppUrl
      ? `${ctx.webAppUrl}?action=confirmDone&sheet=${encodeURIComponent(ctx.sheetCode)}&id=${encodeURIComponent(item.id)}`
      : "#";

    rowsHtml += `
      <tr style="${rowBg}">
        <td style="padding: 8px; border: 1px solid #ddd; text-align: center;">${escapeHtml_(String(item.id))}</td>
        <td style="padding: 8px; border: 1px solid #ddd;"><b>${escapeHtml_(item.name)}</b></td>
        <td style="padding: 8px; border: 1px solid #ddd;">${escapeHtml_(item.extraInfo)}</td>
        <td style="padding: 8px; border: 1px solid #ddd; text-align: center;">${escapeHtml_(item.dueStr)}</td>
        <td style="padding: 8px; border: 1px solid #ddd; color: ${badgeColor}; font-weight: bold;">${escapeHtml_(statusText)}</td>
        <td style="padding: 8px; border: 1px solid #ddd; text-align: center;">
          <a href="${confirmUrl}" style="display:inline-block;background:#0d6efd;color:#fff;padding:5px 10px;text-decoration:none;border-radius:4px;font-size:12px;">Confirm Done</a>
        </td>
      </tr>
    `;
  }

  return `
    <html>
      <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif; color: #222;">
        <h2 style="margin-bottom: 8px;">${escapeHtml_(ctx.sheetLabel)} Daily Summary (${ctx.pendingItems.length} Pending)</h2>
        ${topButtons}
        <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
          <thead>
            <tr style="background-color: #f1f3f5;">
              <th style="padding: 8px; border: 1px solid #ddd; width: 48px;">ID</th>
              <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Customer</th>
              <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">${escapeHtml_(ctx.extraColHeader)}</th>
              <th style="padding: 8px; border: 1px solid #ddd; width: 90px;">Next Date</th>
              <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Status</th>
              <th style="padding: 8px; border: 1px solid #ddd; width: 105px;">Action</th>
            </tr>
          </thead>
          <tbody>
            ${rowsHtml}
          </tbody>
        </table>
        <p style="margin-top: 12px; font-size: 12px; color: #6c757d;">
          Clicking "Confirm Done" sets Last Changed Date to today and automatically calculates the next cycle.
        </p>
      </body>
    </html>
  `;
}

function dispatchEmailWithFallback_(ss, toList, ccList, subject, htmlBody, tag) {
  const config = readConfigMap_(ss);
  const enableGraph = String(config["ENABLE_MS_GRAPH"] || "TRUE").trim().toUpperCase() === "TRUE";

  if (enableGraph) {
    try {
      sendViaMsGraph_(toList, ccList, subject, htmlBody);
      Logger.log(`[${tag}] Sent via MS Graph to ${toList.join(", ")}`);
      return;
    } catch (err) {
      Logger.log(`[${tag}] MS Graph failed, falling back to MailApp: ${err.message}`);
      writeAuditLog_(
        ss,
        "SYSTEM",
        "IT",
        "-",
        tag,
        "GRAPH_FALLBACK",
        err.message,
        "MS_GRAPH",
        "MAIL_APP",
        "-"
      );
    }
  }

  const mimeSubject = "=?UTF-8?B?" + Utilities.base64Encode(subject, Utilities.Charset.UTF_8) + "?=";
  MailApp.sendEmail({
    to: toList.join(","),
    cc: ccList.join(","),
    subject: mimeSubject,
    htmlBody: htmlBody
  });
  Logger.log(`[${tag}] Sent via MailApp fallback to ${toList.join(", ")}`);
}

function sendViaMsGraph_(toList, ccList, subject, htmlBody) {
  const props = PropertiesService.getScriptProperties();
  const tenantId = String(props.getProperty("MS_TENANT_ID") || "").trim();
  const clientId = String(props.getProperty("MS_CLIENT_ID") || "").trim();
  const clientSecret = String(props.getProperty("MS_CLIENT_SECRET") || "").trim();
  const senderEmail = String(props.getProperty("MS_SENDER_EMAIL") || "").trim();

  if (!tenantId || !clientId || !clientSecret || !senderEmail) {
    throw new Error("Missing MS Graph Script Properties");
  }

  const tokenUrl = `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`;
  const tokenResp = UrlFetchApp.fetch(tokenUrl, {
    method: "post",
    payload: {
      client_id: clientId,
      client_secret: clientSecret,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials"
    },
    muteHttpExceptions: true
  });

  if (tokenResp.getResponseCode() !== 200) {
    throw new Error(`Token HTTP ${tokenResp.getResponseCode()}: ${tokenResp.getContentText()}`);
  }

  const accessToken = JSON.parse(tokenResp.getContentText()).access_token;
  const sendUrl = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(senderEmail)}/sendMail`;

  const payload = {
    message: {
      subject: subject,
      body: {
        contentType: "HTML",
        content: htmlBody
      },
      toRecipients: toList.map(addr => ({ emailAddress: { address: addr } })),
      ccRecipients: ccList.map(addr => ({ emailAddress: { address: addr } }))
    },
    saveToSentItems: true
  };

  const mailResp = UrlFetchApp.fetch(sendUrl, {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + accessToken },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = mailResp.getResponseCode();
  if (code !== 202 && code !== 200) {
    throw new Error(`Graph sendMail HTTP ${code}: ${mailResp.getContentText()}`);
  }
}

/* ================= Shared Helpers ================= */

function computeNextChangeDate_(lastChangedDate, cycleDays) {
  const d = new Date(lastChangedDate.getFullYear(), lastChangedDate.getMonth(), lastChangedDate.getDate());
  d.setDate(d.getDate() + Number(cycleDays));
  const day = d.getDay();
  const isoDay = (day === 0) ? 7 : day;
  if (isoDay > 5) {
    d.setDate(d.getDate() - (isoDay - 5));
  }
  return d;
}

function computeEffectiveStatus_(rawStatus, nextDateObj, todayStart) {
  const upper = String(rawStatus || "").trim().toUpperCase();
  if (upper === "ARCHIVED") return "Archived";
  if (!nextDateObj) return rawStatus || "Normal";

  const dueStart = new Date(nextDateObj.getFullYear(), nextDateObj.getMonth(), nextDateObj.getDate()).getTime();
  const notifStart = dueStart - (14 * 86400000);

  if (todayStart >= notifStart) {
    return (todayStart <= dueStart) ? "Pending Notification" : "REPLACEMENT REQUIRED";
  }
  return "Normal";
}

function readConfigMap_(ss) {
  const cfgSheet = ss.getSheetByName("IT_Config");
  const map = {};
  if (!cfgSheet || cfgSheet.getLastRow() < 1) return map;
  const rows = cfgSheet.getRange(1, 1, cfgSheet.getLastRow(), 2).getValues();
  for (let i = 0; i < rows.length; i++) {
    const key = String(rows[i][0] || "").trim();
    if (key) map[key] = rows[i][1];
  }
  return map;
}

function buildHeaderIndexMap_(headers) {
  const map = {};
  for (let c = 0; c < headers.length; c++) {
    if (headers[c]) map[headers[c]] = c;
  }
  return map;
}

function getVal_(rowArr, colMap, names) {
  for (let i = 0; i < names.length; i++) {
    const idx = colMap[names[i]];
    if (idx !== undefined && idx < rowArr.length) return rowArr[idx];
  }
  return "";
}

function parseValidDate_(val) {
  if (val === null || val === undefined || val === "") return null;
  if (Object.prototype.toString.call(val) === "[object Date]") {
    return isNaN(val.getTime()) ? null : val;
  }
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

function splitEmails_(raw) {
  if (!raw) return [];
  return String(raw).split(",").map(s => s.trim()).filter(Boolean);
}

function escapeHtml_(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function writeAuditLog_(ss, actorEmail, roleLabel, custId, custName, action, reason, beforeState, afterState, rowNum) {
  const logSheet = ss.getSheetByName("_Audit_Log");
  if (!logSheet) return;
  logSheet.appendRow([
    new Date(),
    actorEmail || "",
    roleLabel || "",
    (custId !== undefined && custId !== null) ? String(custId) : "",
    custName || "",
    action || "",
    reason || "",
    beforeState || "",
    afterState || "",
    (rowNum !== undefined && rowNum !== null) ? String(rowNum) : ""
  ]);
}
