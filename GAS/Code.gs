/**
 * Code.gs - Apple-Aligned Desktop Edition + Compatible CodePoint Emoji + Test Mode Enabled
 */

const MAIL_TEST_MODE = false; // Enabled for testing; set to false for production
const SECRET_WARN_DAYS = 30;
const FALLBACK_MAX_BYTES = 175000; // Only used when falling back to MailApp

const GAS_PROJECT_SETTINGS_URL = "https://script.google.com/u/0/home/projects/1QZGZuhuGlPVixUbUVIqqWWuo9SOr3RQitvzCZ_Hl9ETPvfgniErJj_6m/settings";
const ENTRA_CREDENTIALS_URL = "https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Credentials/appId/f8800eba-1305-4266-90eb-3f20101cbadf/objectId/c3423aba-88e2-4baa-a2a7-37f1037e7860/isMSAApp~/false/defaultBlade/Overview/appSignInAudience/AzureADMyOrg/servicePrincipalCreated~/true";

const EMOJI = {
  redDrop: String.fromCodePoint(0x1F4A7),               // 💧
  chart: String.fromCodePoint(0x1F4CA),                 // 📊
  warn: String.fromCodePoint(0x26A0, 0xFE0F),           // ⚠️ (high compatibility across Gmail/Outlook/Apple)
  key: String.fromCodePoint(0x1F511),                   // 🔑
  bell: String.fromCodePoint(0x1F514),                  // 🔔
  check: String.fromCodePoint(0x2713)                   // ✓
};

function sendDailyFilterEmails() {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const redSs = SpreadsheetApp.getActiveSpreadsheet();
    const redConfig = readConfigMap_(redSs);

    checkSecretExpiration_(redSs, redConfig);
    processSingleSheetReminder_(redSs, redConfig, "RED");

    const greenId = extractSpreadsheetId_(redConfig["GREEN_SHEET_ID"]);
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

    const ccList = MAIL_TEST_MODE ? [] : ["lai@sqrtone.com"];
    const expStr = Utilities.formatDate(expDate, tz, "yyyy-MM-dd");
    const twoYearsLater = new Date(now.getFullYear() + 2, now.getMonth(), now.getDate());
    const suggestedExpiryStr = Utilities.formatDate(twoYearsLater, tz, "yyyy-MM-dd");

    const subject = daysLeft < 0
      ? `${EMOJI.warn} MS Graph Client Secret EXPIRED (${expStr})`
      : `${EMOJI.key} MS Graph Client Secret Expires in ${daysLeft}d (${expStr})`;

    const body = `
      <div style="font-family:-apple-system,BlinkMacSystemFont,'SF Pro Text','Segoe UI',Roboto,Arial,sans-serif;background:#f2f2f7;padding:12px;color:#1c1c1e;">
        <table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:16px;padding:22px 24px;border:1px solid #e5e5ea;">
          <tr>
            <td>
              <div style="font-size:12px;font-weight:700;color:#ff9500;text-transform:uppercase;letter-spacing:0.06em;">${EMOJI.key} Entra ID &amp; Apps Script</div>
              <h2 style="margin:6px 0 14px 0;font-size:21px;font-weight:700;color:#1c1c1e;">MS Graph Client Secret Renewal</h2>
              <table width="100%" cellpadding="0" cellspacing="0" style="background:#f2f2f7;border-radius:12px;padding:14px 16px;margin-bottom:18px;font-size:14.5px;">
                <tr>
                  <td style="padding:5px 0;"><b>Current Expiry:</b> ${escapeHtml_(expStr)}</td>
                  <td align="right" style="padding:5px 0;"><b>Status:</b> ${daysLeft < 0 ? `Expired ${Math.abs(daysLeft)}d ago` : `${daysLeft}d remaining`}</td>
                </tr>
                <tr>
                  <td colspan="2" style="padding:8px 0 2px 0;">
                    <b>2-Year Expiry Date (Copy):</b>
                    <code style="background:#ffffff;padding:4px 10px;border-radius:7px;border:1px solid #d1d1d6;font-weight:700;font-size:14px;margin-left:6px;">${escapeHtml_(suggestedExpiryStr)}</code>
                  </td>
                </tr>
              </table>
              <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle" style="font-size:14px;color:#3a3a3c;line-height:1.55;">
                    Generate a new secret in Entra ID, then update <b>MS_CLIENT_SECRET</b> and <b>MS_SECRET_EXPIRES_AT</b> in GAS Script Properties.
                  </td>
                  <td valign="middle" align="right" style="white-space:nowrap;padding-left:16px;">
                    <a href="${escapeHtml_(ENTRA_CREDENTIALS_URL)}" style="display:inline-block;background:#007aff;color:#ffffff;padding:10px 16px;text-decoration:none;border-radius:9px;font-size:13.5px;font-weight:600;margin-right:8px;">${EMOJI.key} Entra ID Credentials</a>
                    <a href="${escapeHtml_(GAS_PROJECT_SETTINGS_URL)}" style="display:inline-block;background:#f2f2f7;color:#1c1c1e;border:1px solid #c7c7cc;padding:10px 16px;text-decoration:none;border-radius:9px;font-size:13.5px;font-weight:600;">GAS Script Properties</a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </div>
    `;
    dispatchEmailWithFallback_(ss, itRecipients, ccList, subject, body, "SECRET_ALERT");
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
  let overdueCount = 0;
  let dueSoonCount = 0;
  let maxOverdueDays = 0;

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
      const usageVal = getVal_(r, colMap, ["Machine Usage", "Machine usage"]);
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
      const idVal = getVal_(r, colMap, ["第 1 列", "ID", "No.", ""]);
      const extraInfo = String(getVal_(r, colMap, ["Filter / Service Name", "Phone Number"]) || "").trim();

      let dueStr = "-";
      let sortTs = Number.MAX_SAFE_INTEGER;
      let overdueDays = 0;
      const isUrgent = (effectiveStatus === "REPLACEMENT REQUIRED");

      if (nextDateObj) {
        dueStr = Utilities.formatDate(nextDateObj, tz, "dd MMM yy");
        const dueStart = new Date(nextDateObj.getFullYear(), nextDateObj.getMonth(), nextDateObj.getDate()).getTime();
        sortTs = dueStart;
        if (isUrgent && todayStart > dueStart) {
          overdueDays = Math.floor((todayStart - dueStart) / 86400000);
          if (overdueDays > maxOverdueDays) maxOverdueDays = overdueDays;
        }
      }

      if (isUrgent) overdueCount++;
      else dueSoonCount++;

      pendingItems.push({
        id: (idVal !== "" && idVal !== null) ? idVal : (rowNum - 1),
        name: name,
        extraInfo: extraInfo,
        dueStr: dueStr,
        sortTs: sortTs,
        status: effectiveStatus,
        overdueDays: overdueDays,
        isUrgent: isUrgent
      });
    }
  }

  pendingItems.sort((a, b) => {
    const urgDiff = Number(b.isUrgent) - Number(a.isUrgent);
    if (urgDiff !== 0) return urgDiff;
    return a.sortTs - b.sortTs;
  });

  const dashboardUrl = String(config["DASHBOARD_URL"] || "https://itsaso.sqrtone.com").trim();
  const dateStamp = Utilities.formatDate(now, tz, "dd MMM yyyy");
  const titleEmoji = (sheetCode === "RED") ? EMOJI.redDrop : EMOJI.chart;

  const subject = (pendingItems.length === 0)
    ? `${titleEmoji} ${sheetLabel} · All Clear (${dateStamp})`
    : `${titleEmoji} ${sheetLabel} · ${overdueCount} Overdue, ${dueSoonCount} Due Soon (${dateStamp})`;

  const fullHtml = buildAppleAlignedReminderHtml_({
    sheetCode: sheetCode,
    sheetLabel: sheetLabel,
    titleEmoji: titleEmoji,
    dateStamp: dateStamp,
    pendingItems: pendingItems,
    overdueCount: overdueCount,
    dueSoonCount: dueSoonCount,
    maxOverdueDays: maxOverdueDays,
    dashboardUrl: dashboardUrl,
    webAppUrl: webAppUrl,
    maxItemsPerCol: null
  });

  const compactFallbackHtml = buildAppleAlignedReminderHtml_({
    sheetCode: sheetCode,
    sheetLabel: sheetLabel,
    titleEmoji: titleEmoji,
    dateStamp: dateStamp,
    pendingItems: pendingItems,
    overdueCount: overdueCount,
    dueSoonCount: dueSoonCount,
    maxOverdueDays: maxOverdueDays,
    dashboardUrl: dashboardUrl,
    webAppUrl: webAppUrl,
    maxItemsPerCol: 25,
    truncatedNote: "Compact fallback view: showing top 25 per column. Open Dashboard for full list."
  });

  const toList = splitEmails_(config["TO_EMAIL"]);
  const ccList = MAIL_TEST_MODE ? [] : splitEmails_(config["CC_EMAILS"]);
  if (toList.length === 0) {
    throw new Error(`No TO_EMAIL configured in IT_Config for ${sheetCode}`);
  }

  dispatchEmailWithFallback_(ss, toList, ccList, subject, fullHtml, sheetCode, compactFallbackHtml);
}

function buildAppleAlignedReminderHtml_(ctx) {
  const dashBtn = ctx.dashboardUrl
    ? `<a href="${escapeHtml_(ctx.dashboardUrl)}" style="display:inline-block;background:#007aff;color:#ffffff;padding:8px 18px;text-decoration:none;border-radius:999px;font-size:13.5px;font-weight:600;">Open Dashboard</a>`
    : "";

  if (ctx.pendingItems.length === 0) {
    return `
      <div style="font-family:-apple-system,BlinkMacSystemFont,'SF Pro Text','Segoe UI',Roboto,Arial,sans-serif;background:#f2f2f7;padding:12px;color:#1c1c1e;">
        <table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:16px;padding:20px 22px;border:1px solid #e5e5ea;">
          <tr>
            <td>
              <div style="font-size:18px;font-weight:700;margin-bottom:6px;">${escapeHtml_(ctx.titleEmoji)} ${escapeHtml_(ctx.sheetLabel)} · All Clear</div>
              <div style="font-size:14px;color:#6e6e73;margin-bottom:14px;">${escapeHtml_(ctx.dateStamp)} · No pending filter reminders today.</div>
              ${dashBtn}
            </td>
          </tr>
        </table>
      </div>
    `;
  }

  const overdueAll = ctx.pendingItems.filter(x => x.isUrgent);
  const dueSoonAll = ctx.pendingItems.filter(x => !x.isUrgent);
  const overdueList = ctx.maxItemsPerCol ? overdueAll.slice(0, ctx.maxItemsPerCol) : overdueAll;
  const dueSoonList = ctx.maxItemsPerCol ? dueSoonAll.slice(0, ctx.maxItemsPerCol) : dueSoonAll;

  const renderColumnRows = (list, isRed) => {
    if (!list || list.length === 0) {
      return `
        <tr>
          <td style="font-size:13.5px;color:#8e8e93;padding:14px;text-align:center;border:1px dashed #e5e5ea;border-radius:11px;">
            No items
          </td>
        </tr>
      `;
    }

    const accent = isRed ? "#ff3b30" : "#ff9500";
    const bg = isRed ? "#fff5f5" : "#fffbeb";
    const border = isRed ? "#ffd6d6" : "#fde68a";
    const tagColor = isRed ? "#d70015" : "#b25000";

    return list.map(item => {
      const tagText = isRed
        ? (item.overdueDays > 0 ? `${item.overdueDays}d overdue` : "Overdue")
        : "Due Soon";

      const notifiedUrl = ctx.webAppUrl
        ? `${ctx.webAppUrl}?action=markNotified&sheet=${encodeURIComponent(ctx.sheetCode)}&id=${encodeURIComponent(item.id)}`
        : "#";
      const doneUrl = ctx.webAppUrl
        ? `${ctx.webAppUrl}?action=confirmDone&sheet=${encodeURIComponent(ctx.sheetCode)}&id=${encodeURIComponent(item.id)}`
        : "#";

      return `
        <tr>
          <td style="padding-bottom:7px;">
            <table width="100%" cellpadding="0" cellspacing="0" style="background:${bg};border:1px solid ${border};border-left:4px solid ${accent};border-radius:10px;padding:8px 11px;">
              <tr>
                <td width="46" valign="middle" style="font-size:12.5px;font-weight:700;color:#6e6e73;font-variant-numeric:tabular-nums;">
                  #${escapeHtml_(String(item.id))}
                </td>
                <td valign="middle" style="line-height:1.35;">
                  <div style="font-size:13.5px;font-weight:700;color:#1c1c1e;">${escapeHtml_(item.name)}</div>
                  <div style="font-size:12px;color:#6e6e73;margin-top:2px;">
                    ${item.extraInfo ? `${escapeHtml_(item.extraInfo)} · ` : ""}Due: <b style="color:#1c1c1e;">${escapeHtml_(item.dueStr)}</b> · <span style="color:${tagColor};font-weight:700;">${escapeHtml_(tagText)}</span>
                  </div>
                </td>
                <td width="86" valign="middle" align="right" style="white-space:nowrap;">
                  <a href="${notifiedUrl}" title="Mark as Notified" aria-label="Mark as Notified" style="display:inline-block;background:#ffffff;color:#007aff;border:1px solid #c7c7cc;padding:6px 10px;text-decoration:none;border-radius:8px;font-size:14px;line-height:1;margin-right:5px;">${EMOJI.bell}</a>
                  <a href="${doneUrl}" title="Confirm Done (Next Cycle)" aria-label="Confirm Done" style="display:inline-block;background:#34c759;color:#ffffff;padding:6px 11px;text-decoration:none;border-radius:8px;font-size:14px;line-height:1;font-weight:700;">${EMOJI.check}</a>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      `;
    }).join("");
  };

  const noteHtml = ctx.truncatedNote
    ? `<div style="font-size:12px;color:#8e8e93;margin-top:10px;text-align:center;">${escapeHtml_(ctx.truncatedNote)}</div>`
    : "";

  return `
    <div style="font-family:-apple-system,BlinkMacSystemFont,'SF Pro Text','Segoe UI',Roboto,Arial,sans-serif;background:#f2f2f7;padding:12px;color:#1c1c1e;">
      <table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:16px;padding:18px 20px;border:1px solid #e5e5ea;">
        <tr>
          <td>
            <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:14px;">
              <tr>
                <td valign="middle">
                  <span style="font-size:18.5px;font-weight:700;color:#1c1c1e;">${escapeHtml_(ctx.titleEmoji)} ${escapeHtml_(ctx.sheetLabel)}</span>
                  <span style="font-size:13px;color:#6e6e73;margin-left:10px;">${escapeHtml_(ctx.dateStamp)} · Max Overdue: <b>${ctx.maxOverdueDays}d</b></span>
                </td>
                <td valign="middle" align="right">
                  ${dashBtn}
                </td>
              </tr>
            </table>

            <table width="100%" cellpadding="0" cellspacing="0" style="table-layout:fixed;">
              <tr>
                <td width="50%" valign="top" style="padding-right:8px;">
                  <table width="100%" cellpadding="0" cellspacing="0">
                    <tr>
                      <td style="padding-bottom:8px;">
                        <div style="font-size:13.5px;font-weight:700;color:#d70015;background:#fff5f5;border:1px solid #ffd6d6;border-radius:10px;padding:7px 12px;">
                          Overdue (${overdueAll.length})
                        </div>
                      </td>
                    </tr>
                    ${renderColumnRows(overdueList, true)}
                  </table>
                </td>
                <td width="50%" valign="top" style="padding-left:8px;">
                  <table width="100%" cellpadding="0" cellspacing="0">
                    <tr>
                      <td style="padding-bottom:8px;">
                        <div style="font-size:13.5px;font-weight:700;color:#b25000;background:#fffbeb;border:1px solid #fde68a;border-radius:10px;padding:7px 12px;">
                          Due Soon (${dueSoonAll.length})
                        </div>
                      </td>
                    </tr>
                    ${renderColumnRows(dueSoonList, false)}
                  </table>
                </td>
              </tr>
            </table>

            <div style="margin-top:12px;padding-top:10px;border-top:1px solid #e5e5ea;font-size:12.5px;color:#6e6e73;text-align:center;">
              ${EMOJI.bell} Mark as Notified · ${EMOJI.check} Confirm Done (Reset to today &amp; roll next cycle)
            </div>
            ${noteHtml}
          </td>
        </tr>
      </table>
    </div>
  `;
}

function dispatchEmailWithFallback_(ss, toList, ccList, subject, htmlBody, tag, compactFallbackHtml) {
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

  let fallbackBody = compactFallbackHtml || htmlBody;
  if (Utilities.newBlob(fallbackBody).getBytes().length > FALLBACK_MAX_BYTES && compactFallbackHtml) {
    fallbackBody = compactFallbackHtml;
  }

  const mimeSubject = "=?UTF-8?B?" + Utilities.base64Encode(subject, Utilities.Charset.UTF_8) + "?=";
  MailApp.sendEmail({
    to: toList.join(","),
    cc: ccList.join(","),
    subject: mimeSubject,
    htmlBody: fallbackBody
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

  // Lightweight E5 developer activity heartbeat (low-frequency read-only Graph calls)
  try {
    UrlFetchApp.fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(senderEmail)}?$select=id,displayName`, {
      headers: { Authorization: "Bearer " + accessToken },
      muteHttpExceptions: true
    });
    UrlFetchApp.fetch("https://graph.microsoft.com/v1.0/organization?$select=id,displayName", {
      headers: { Authorization: "Bearer " + accessToken },
      muteHttpExceptions: true
    });
  } catch (_) {}

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

function extractSpreadsheetId_(raw) {
  const s = String(raw || "").trim();
  const match = s.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : s;
}

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
