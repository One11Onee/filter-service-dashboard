/**
 * webApp.gs - Phase 4 Complete (JSONP + CORS-safe)
 */

const ROLE_LEVELS = {
  "IT": 100,
  "BOSS": 80,
  "ADMIN_H": 60,
  "ADMIN_A": 40,
  "ADMIN_L": 40,
  "GUEST": 0
};

function doGet(e) {
  const params = (e && e.parameter) || {};
  const action = String(params.action || "").trim();
  const callback = String(params.callback || "").trim();

  if (action === "confirmDone" || (params.id && !action)) {
    const sheetCode = String(params.sheet || "RED").toUpperCase();
    const targetId = String(params.id || "").trim();
    const res = handleConfirmDoneById_(sheetCode, targetId, "Email One-Click", "ADMIN_ACTION");
    return HtmlService.createHtmlOutput(`
      <html>
        <body style="font-family: Arial, sans-serif; padding: 28px; max-width: 540px;">
          <h3 style="color: ${res.ok ? "#198754" : "#dc3545"};">
            ${res.ok ? "Updated & Next Cycle Calculated" : "Update Failed"}
          </h3>
          <p>${escapeHtml_(res.message)}</p>
          ${res.nextDateStr ? `<p><b>Next Change Date:</b> ${escapeHtml_(res.nextDateStr)}</p>` : ""}
        </body>
      </html>
    `);
  }

  if (action === "getDashboardData") {
    try {
      const auth = resolveEffectiveAuth_(params.email, params.viewAs);
      if (auth.actualRole === "GUEST") {
        return jsonpOrJsonResponse_({ ok: false, error: "Unauthorized email", role: "GUEST" }, callback);
      }
      const sheetCode = resolveAllowedSheet_(auth.effectiveRole, params.sheet);
      const ss = getTargetSpreadsheet_(sheetCode);
      const includeHidden = (auth.effectiveRole === "IT" || auth.effectiveRole === "BOSS" || auth.effectiveRole === "ADMIN_H");
      const payload = readSheetRecords_(ss, sheetCode, includeHidden);

      return jsonpOrJsonResponse_({
        ok: true,
        actorEmail: auth.actorEmail,
        actualRole: auth.actualRole,
        effectiveRole: auth.effectiveRole,
        canSwitchView: auth.actualRole === "IT",
        activeSheet: sheetCode,
        records: payload.records,
        summary: payload.summary
      }, callback);
    } catch (err) {
      return jsonpOrJsonResponse_({ ok: false, error: err.message }, callback);
    }
  }

  return ContentService.createTextOutput("Filter Service API Ready.");
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const raw = (e && e.postData && e.postData.contents) || "{}";
    const body = JSON.parse(raw);
    const action = String(body.action || "").trim();
    const auth = resolveEffectiveAuth_(body.email, body.viewAs);

    if (auth.actualRole === "GUEST") {
      return jsonResponse_({ ok: false, error: "Unauthorized user" });
    }

    const sheetCode = resolveAllowedSheet_(auth.effectiveRole, body.sheet);
    const ss = getTargetSpreadsheet_(sheetCode);
    const actorLabel = (auth.actualRole === "IT" && auth.effectiveRole !== "IT")
      ? `IT(as ${auth.effectiveRole})`
      : auth.effectiveRole;

    let result;
    switch (action) {
      case "createRecord":
        assertMinRole_(auth.effectiveRole, 40);
        result = handleCreateRecord_(ss, sheetCode, body.payload || {}, auth.actorEmail, actorLabel);
        break;
      case "updateRecord":
        assertMinRole_(auth.effectiveRole, 40);
        result = handleUpdateRecord_(ss, sheetCode, body.id, body.payload || {}, auth.actorEmail, actorLabel);
        break;
      case "confirmDone":
        assertMinRole_(auth.effectiveRole, 40);
        result = handleConfirmDoneById_(sheetCode, body.id, auth.actorEmail, actorLabel);
        break;
      case "directHide":
        assertMinRole_(auth.effectiveRole, 40);
        result = handleSoftHideById_(ss, sheetCode, body.id, body.reason || "Direct soft delete", "DIRECT_HIDE", auth.actorEmail, actorLabel);
        break;
      case "requestDelete":
        assertMinRole_(auth.effectiveRole, 40);
        result = handleRequestDelete_(ss, sheetCode, body.id, body.reason || "Requested deletion", auth.actorEmail, actorLabel);
        break;
      case "approveDelete":
        assertMinRole_(auth.effectiveRole, 60);
        result = handleSoftHideById_(ss, sheetCode, body.id, body.reason || "Approved by Admin H", "APPROVE_HIDE", auth.actorEmail, actorLabel);
        break;
      case "rejectDelete":
        assertMinRole_(auth.effectiveRole, 60);
        result = handleRejectDelete_(ss, sheetCode, body.id, body.reason || "Rejected by Admin H", auth.actorEmail, actorLabel);
        break;
      case "restoreRow":
        assertMinRole_(auth.effectiveRole, 60);
        result = handleRestoreRow_(ss, sheetCode, body.id, body.reason || "Restored record", auth.actorEmail, actorLabel);
        break;
      default:
        return jsonResponse_({ ok: false, error: "Unsupported action: " + action });
    }

    return jsonResponse_(result);
  } catch (err) {
    return jsonResponse_({ ok: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}

/* ================= Actions ================= */

function handleConfirmDoneById_(sheetCode, targetId, actorEmail, roleLabel) {
  const ss = getTargetSpreadsheet_(sheetCode);
  const sheet = ss.getSheetByName("Sheet 1");
  const found = findRowById_(sheet, targetId);
  if (!found) {
    return { ok: false, message: `Record ID=${targetId} not found in ${sheetCode}.` };
  }

  const { rowNum, colMap, rowValues } = found;
  const custName = String(getVal_(rowValues, colMap, ["Name"]) || "");
  const prevLast = parseValidDate_(getVal_(rowValues, colMap, ["Last Changed Date"]));
  const cycleDays = Number(getVal_(rowValues, colMap, ["Change Cycle (Days)"]) || 180);

  const nowDate = new Date();
  const nextDateObj = computeNextChangeDate_(nowDate, cycleDays);
  const nextDateStr = Utilities.formatDate(nextDateObj, Session.getScriptTimeZone(), "dd/MM/yy");

  if (colMap["Last Changed Date"] !== undefined) {
    sheet.getRange(rowNum, colMap["Last Changed Date"] + 1).setValue(nowDate);
  }
  if (colMap["Notification Status"] !== undefined) {
    sheet.getRange(rowNum, colMap["Notification Status"] + 1).setValue(false);
  }
  applyRowFormulasAndFallback_(sheet, rowNum, colMap, nowDate, cycleDays);
  SpreadsheetApp.flush();

  writeAuditLog_(
    ss,
    actorEmail,
    roleLabel,
    targetId,
    custName,
    "CONFIRM_DONE",
    "Completed & auto-calculated next cycle",
    `Last:${prevLast ? formatShortDate_(prevLast) : "-"}`,
    `Last:${formatShortDate_(nowDate)}|Next:${nextDateStr}|Notif:FALSE`,
    rowNum
  );

  return {
    ok: true,
    id: targetId,
    name: custName,
    nextDateStr: nextDateStr,
    message: `Customer "${custName}" updated to today; next due date is ${nextDateStr}.`
  };
}

function handleSoftHideById_(ss, sheetCode, targetId, reason, actionType, actorEmail, roleLabel) {
  const sheet = ss.getSheetByName("Sheet 1");
  const found = findRowById_(sheet, targetId);
  if (!found) throw new Error(`Record ID=${targetId} not found`);

  const { rowNum, colMap, rowValues } = found;
  const custName = String(getVal_(rowValues, colMap, ["Name"]) || "");
  const prevStatus = String(getVal_(rowValues, colMap, ["Status"]) || "");

  if (colMap["Status"] !== undefined) {
    sheet.getRange(rowNum, colMap["Status"] + 1).setValue("Archived");
  }
  if (colMap["Notification Status"] !== undefined) {
    sheet.getRange(rowNum, colMap["Notification Status"] + 1).setValue(true);
  }
  if (colMap["Approval Status"] !== undefined && actionType === "APPROVE_HIDE") {
    sheet.getRange(rowNum, colMap["Approval Status"] + 1).setValue("APPROVED_ARCHIVED");
  }

  sheet.hideRows(rowNum);
  SpreadsheetApp.flush();

  writeAuditLog_(
    ss,
    actorEmail,
    roleLabel,
    targetId,
    custName,
    actionType,
    reason,
    `Status:${prevStatus}|Hidden:false`,
    `Status:Archived|Hidden:true`,
    rowNum
  );

  return { ok: true, id: targetId, message: `Archived and hid "${custName}" (recoverable anytime).` };
}

function handleRequestDelete_(ss, sheetCode, targetId, reason, actorEmail, roleLabel) {
  const sheet = ss.getSheetByName("Sheet 1");
  const found = findRowById_(sheet, targetId);
  if (!found) throw new Error(`Record ID=${targetId} not found`);

  const { rowNum, colMap, rowValues } = found;
  const custName = String(getVal_(rowValues, colMap, ["Name"]) || "");
  if (colMap["Approval Status"] === undefined) {
    throw new Error("Missing Approval Status column");
  }

  sheet.getRange(rowNum, colMap["Approval Status"] + 1).setValue("PENDING_DELETE");
  SpreadsheetApp.flush();

  writeAuditLog_(
    ss,
    actorEmail,
    roleLabel,
    targetId,
    custName,
    "REQ_DELETE",
    reason,
    "Approval:NONE",
    "Approval:PENDING_DELETE",
    rowNum
  );

  return { ok: true, id: targetId, message: `Delete request submitted for "${custName}".` };
}

function handleRejectDelete_(ss, sheetCode, targetId, reason, actorEmail, roleLabel) {
  const sheet = ss.getSheetByName("Sheet 1");
  const found = findRowById_(sheet, targetId);
  if (!found) throw new Error(`Record ID=${targetId} not found`);

  const { rowNum, colMap, rowValues } = found;
  const custName = String(getVal_(rowValues, colMap, ["Name"]) || "");
  if (colMap["Approval Status"] !== undefined) {
    sheet.getRange(rowNum, colMap["Approval Status"] + 1).setValue("REJECTED");
  }
  SpreadsheetApp.flush();

  writeAuditLog_(
    ss,
    actorEmail,
    roleLabel,
    targetId,
    custName,
    "REJECT_DELETE",
    reason,
    "Approval:PENDING_DELETE",
    "Approval:REJECTED",
    rowNum
  );

  return { ok: true, id: targetId, message: `Rejected delete request for "${custName}".` };
}

function handleRestoreRow_(ss, sheetCode, targetId, reason, actorEmail, roleLabel) {
  const sheet = ss.getSheetByName("Sheet 1");
  const found = findRowById_(sheet, targetId);
  if (!found) throw new Error(`Record ID=${targetId} not found`);

  const { rowNum, colMap, rowValues } = found;
  const custName = String(getVal_(rowValues, colMap, ["Name"]) || "");

  sheet.showRows(rowNum);
  if (colMap["Status"] !== undefined) sheet.getRange(rowNum, colMap["Status"] + 1).setValue("Normal");
  if (colMap["Notification Status"] !== undefined) sheet.getRange(rowNum, colMap["Notification Status"] + 1).setValue(false);
  if (colMap["Approval Status"] !== undefined) sheet.getRange(rowNum, colMap["Approval Status"] + 1).setValue("");
  SpreadsheetApp.flush();

  writeAuditLog_(
    ss,
    actorEmail,
    roleLabel,
    targetId,
    custName,
    "RESTORE",
    reason,
    "Status:Archived|Hidden:true",
    "Status:Normal|Hidden:false",
    rowNum
  );

  return { ok: true, id: targetId, message: `Restored "${custName}".` };
}

function handleUpdateRecord_(ss, sheetCode, targetId, payload, actorEmail, roleLabel) {
  const sheet = ss.getSheetByName("Sheet 1");
  const found = findRowById_(sheet, targetId);
  if (!found) throw new Error(`Record ID=${targetId} not found`);

  const { rowNum, colMap, rowValues } = found;
  const custName = String(getVal_(rowValues, colMap, ["Name"]) || "");
  const editable = [
    "Name",
    "Filter / Service Name",
    "Phone Number",
    "Machine usage",
    "Last Changed Date",
    "Change Cycle (Days)",
    "Notification Status"
  ];

  const changed = [];
  for (let i = 0; i < editable.length; i++) {
    const field = editable[i];
    if (Object.prototype.hasOwnProperty.call(payload, field) && colMap[field] !== undefined) {
      let val = payload[field];
      if (field === "Last Changed Date" && val) {
        const d = parseValidDate_(val);
        if (d) val = d;
      } else if (field === "Change Cycle (Days)" && val !== "") {
        val = Number(val);
      } else if (field === "Machine usage" || field === "Notification Status") {
        val = (val === true || String(val).toUpperCase() === "TRUE");
      }
      sheet.getRange(rowNum, colMap[field] + 1).setValue(val);
      changed.push(`${field}=${String(val)}`);
    }
  }

  const updatedRow = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastD = parseValidDate_(getVal_(updatedRow, colMap, ["Last Changed Date"]));
  const cycD = Number(getVal_(updatedRow, colMap, ["Change Cycle (Days)"]) || 180);
  if (lastD && !isNaN(cycD) && cycD > 0) {
    applyRowFormulasAndFallback_(sheet, rowNum, colMap, lastD, cycD);
  }
  SpreadsheetApp.flush();

  writeAuditLog_(
    ss,
    actorEmail,
    roleLabel,
    targetId,
    custName,
    "UPDATE",
    payload.reason || "Dashboard edit",
    "",
    changed.join("; "),
    rowNum
  );

  return { ok: true, id: targetId, message: `Updated "${custName}".` };
}

function handleCreateRecord_(ss, sheetCode, payload, actorEmail, roleLabel) {
  const sheet = ss.getSheetByName("Sheet 1");
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
  const colMap = buildHeaderIndexMap_(headers);

  let maxId = 0;
  if (lastRow >= 2 && colMap["第 1 列"] !== undefined) {
    const idVals = sheet.getRange(2, colMap["第 1 列"] + 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < idVals.length; i++) {
      const n = Number(idVals[i][0]);
      if (!isNaN(n) && n > maxId) maxId = n;
    }
  }

  const newId = maxId + 1;
  const newRowNum = lastRow + 1;
  const lastChangedDate = parseValidDate_(payload["Last Changed Date"]) || new Date();
  const cycleDays = Number(payload["Change Cycle (Days)"] || 180);

  const rowArr = new Array(lastCol).fill("");
  if (colMap["第 1 列"] !== undefined) rowArr[colMap["第 1 列"]] = newId;
  if (colMap["Name"] !== undefined) rowArr[colMap["Name"]] = String(payload["Name"] || "").trim();
  if (colMap["Filter / Service Name"] !== undefined) {
    rowArr[colMap["Filter / Service Name"]] = String(payload["Filter / Service Name"] || "").trim();
  }
  if (colMap["Phone Number"] !== undefined) {
    rowArr[colMap["Phone Number"]] = String(payload["Phone Number"] || "").trim();
  }
  if (colMap["Machine usage"] !== undefined) {
    rowArr[colMap["Machine usage"]] = (payload["Machine usage"] === false || String(payload["Machine usage"]).toUpperCase() === "FALSE") ? false : true;
  }
  if (colMap["Last Changed Date"] !== undefined) rowArr[colMap["Last Changed Date"]] = lastChangedDate;
  if (colMap["Change Cycle (Days)"] !== undefined) rowArr[colMap["Change Cycle (Days)"]] = cycleDays;
  if (colMap["Notification Status"] !== undefined) rowArr[colMap["Notification Status"]] = false;

  sheet.getRange(newRowNum, 1, 1, lastCol).setValues([rowArr]);
  applyRowFormulasAndFallback_(sheet, newRowNum, colMap, lastChangedDate, cycleDays);
  SpreadsheetApp.flush();

  writeAuditLog_(
    ss,
    actorEmail,
    roleLabel,
    newId,
    rowArr[colMap["Name"]] || "",
    "CREATE",
    payload.reason || "Created record",
    "",
    `ID=${newId}|Name=${rowArr[colMap["Name"]] || ""}`,
    newRowNum
  );

  return { ok: true, id: newId, rowNum: newRowNum, message: `Created record ID=${newId}.` };
}

function applyRowFormulasAndFallback_(sheet, rowNum, colMap, lastChangedDate, cycleDays) {
  const colLet = (name) => (colMap[name] !== undefined ? columnToLetter_(colMap[name] + 1) : null);
  const cName = colLet("Name");
  const cLast = colLet("Last Changed Date");
  const cCycle = colLet("Change Cycle (Days)");
  const cNext = colLet("Next Change Date");
  const cNotifDate = colLet("Notification Date");
  const cStatus = colLet("Status");
  const cNotifStat = colLet("Notification Status");
  const cImp = colLet("Importance");

  if (colMap["Next Change Date"] !== undefined && cLast && cCycle) {
    sheet.getRange(rowNum, colMap["Next Change Date"] + 1).setFormula(
      `=IF(WEEKDAY(${cLast}${rowNum}+${cCycle}${rowNum},2)>5,(${cLast}${rowNum}+${cCycle}${rowNum})-(WEEKDAY(${cLast}${rowNum}+${cCycle}${rowNum},2)-5),${cLast}${rowNum}+${cCycle}${rowNum})`
    );
  }
  if (colMap["Notification Date"] !== undefined && cNext) {
    sheet.getRange(rowNum, colMap["Notification Date"] + 1).setFormula(`=${cNext}${rowNum}-14`);
  }
  if (colMap["Status"] !== undefined && cName && cNotifDate && cNext) {
    sheet.getRange(rowNum, colMap["Status"] + 1).setFormula(
      `=IF(${cName}${rowNum}="","",IF(TODAY()>=${cNotifDate}${rowNum},IF(TODAY()<=${cNext}${rowNum},"Pending Notification","REPLACEMENT REQUIRED"),"Normal"))`
    );
  }
  if (colMap["Importance"] !== undefined && cStatus && cNotifStat) {
    sheet.getRange(rowNum, colMap["Importance"] + 1).setFormula(
      `=IF(${cStatus}${rowNum}="REPLACEMENT REQUIRED",1,IF(AND(${cStatus}${rowNum}="Pending Notification",${cNotifStat}${rowNum}=FALSE),2,IF(AND(${cStatus}${rowNum}="Pending Notification",${cNotifStat}${rowNum}=TRUE),3,4)))`
    );
  }
}

/* ================= Auth & Data Readers ================= */

function resolveEffectiveAuth_(emailParam, viewAsParam) {
  const actorEmail = String(emailParam || "").trim().toLowerCase();
  const redSs = SpreadsheetApp.getActiveSpreadsheet();
  const redCfg = readConfigMap_(redSs);

  let actualRole = "GUEST";
  if (actorEmail) {
    if (inCsv_(actorEmail, redCfg["ROLE_IT"])) actualRole = "IT";
    else if (inCsv_(actorEmail, redCfg["ROLE_BOSS"])) actualRole = "BOSS";
    else if (inCsv_(actorEmail, redCfg["ROLE_ADMIN_H"])) actualRole = "ADMIN_H";
    else if (inCsv_(actorEmail, redCfg["ROLE_ADMIN_A"])) actualRole = "ADMIN_A";
    else {
      const greenId = String(redCfg["GREEN_SHEET_ID"] || "").trim();
      if (greenId) {
        const greenCfg = readConfigMap_(SpreadsheetApp.openById(greenId));
        if (inCsv_(actorEmail, greenCfg["ROLE_IT"])) actualRole = "IT";
        else if (inCsv_(actorEmail, greenCfg["ROLE_BOSS"])) actualRole = "BOSS";
        else if (inCsv_(actorEmail, greenCfg["ROLE_ADMIN_L"])) actualRole = "ADMIN_L";
      }
    }
  }

  let effectiveRole = actualRole;
  const viewAs = String(viewAsParam || "").trim().toUpperCase();
  if (actualRole === "IT" && ROLE_LEVELS[viewAs] > 0) {
    effectiveRole = viewAs;
  }

  return { actorEmail: actorEmail, actualRole: actualRole, effectiveRole: effectiveRole };
}

function resolveAllowedSheet_(effectiveRole, requestedSheet) {
  const req = String(requestedSheet || "").toUpperCase();
  if (effectiveRole === "ADMIN_L") return "GREEN";
  return req === "GREEN" ? "GREEN" : "RED";
}

function getTargetSpreadsheet_(sheetCode) {
  if (sheetCode === "GREEN") {
    const redCfg = readConfigMap_(SpreadsheetApp.getActiveSpreadsheet());
    const greenId = String(redCfg["GREEN_SHEET_ID"] || "").trim();
    if (!greenId) throw new Error("GREEN_SHEET_ID not configured");
    return SpreadsheetApp.openById(greenId);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function readSheetRecords_(ss, sheetCode, includeHidden) {
  const sheet = ss.getSheetByName("Sheet 1");
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) {
    return {
      records: [],
      summary: { total: 0, replacementRequired: 0, pendingNotification: 0, pendingApproval: 0, hiddenCount: 0 }
    };
  }

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
  const colMap = buildHeaderIndexMap_(headers);
  const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();

  const records = [];
  let replacementRequired = 0;
  let pendingNotification = 0;
  let pendingApproval = 0;
  let hiddenCount = 0;

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const rowNum = i + 2;
    const name = String(getVal_(r, colMap, ["Name"]) || "").trim();
    if (!name) continue;

    const rawStatus = String(getVal_(r, colMap, ["Status"]) || "Normal").trim();
    const isHidden = sheet.isRowHiddenByUser(rowNum) || rawStatus.toUpperCase() === "ARCHIVED";
    if (isHidden) hiddenCount++;

    const approvalStatus = String(getVal_(r, colMap, ["Approval Status"]) || "").trim();
    if (approvalStatus === "PENDING_DELETE") pendingApproval++;

    const machineUsage = (sheetCode === "RED")
      ? (getVal_(r, colMap, ["Machine usage", "Machine Usage"]) === true || String(getVal_(r, colMap, ["Machine usage", "Machine Usage"])).toUpperCase() === "TRUE")
      : true;

    const lastDateObj = parseValidDate_(getVal_(r, colMap, ["Last Changed Date"]));
    const cycleNum = Number(getVal_(r, colMap, ["Change Cycle (Days)"]) || 0);
    let nextDateObj = parseValidDate_(getVal_(r, colMap, ["Next Change Date"]));
    if (!nextDateObj && lastDateObj && cycleNum > 0) {
      nextDateObj = computeNextChangeDate_(lastDateObj, cycleNum);
    }

    const effStatus = computeEffectiveStatus_(rawStatus, nextDateObj, todayStart);
    let overdueDays = 0;
    if (nextDateObj && effStatus === "REPLACEMENT REQUIRED") {
      const dueStart = new Date(nextDateObj.getFullYear(), nextDateObj.getMonth(), nextDateObj.getDate()).getTime();
      if (todayStart > dueStart) overdueDays = Math.floor((todayStart - dueStart) / 86400000);
    }

    if (!isHidden && (sheetCode !== "RED" || machineUsage)) {
      if (effStatus === "REPLACEMENT REQUIRED") replacementRequired++;
      else if (effStatus === "Pending Notification") pendingNotification++;
    }

    if (isHidden && !includeHidden) continue;

    const notifVal = getVal_(r, colMap, ["Notification Status"]);
    records.push({
      id: getVal_(r, colMap, ["第 1 列", "ID"]),
      rowNum: rowNum,
      name: name,
      filterOrPhone: String(getVal_(r, colMap, ["Filter / Service Name", "Phone Number"]) || ""),
      machineUsage: machineUsage,
      lastChangedDate: lastDateObj ? Utilities.formatDate(lastDateObj, tz, "yyyy-MM-dd") : "",
      cycleDays: cycleNum || "",
      nextChangeDate: nextDateObj ? Utilities.formatDate(nextDateObj, tz, "yyyy-MM-dd") : "",
      status: effStatus,
      overdueDays: overdueDays,
      notified: (notifVal === true || String(notifVal).toUpperCase() === "TRUE"),
      approvalStatus: approvalStatus,
      isHidden: isHidden
    });
  }

  return {
    records: records,
    summary: {
      total: records.length,
      replacementRequired: replacementRequired,
      pendingNotification: pendingNotification,
      pendingApproval: pendingApproval,
      hiddenCount: hiddenCount
    }
  };
}

function findRowById_(sheet, targetId) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return null;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
  const colMap = buildHeaderIndexMap_(headers);
  const idColIdx = (colMap["第 1 列"] !== undefined) ? colMap["第 1 列"] : 0;
  const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][idColIdx]).trim() === String(targetId).trim()) {
      return { rowNum: i + 2, colMap: colMap, rowValues: rows[i] };
    }
  }
  return null;
}

function assertMinRole_(role, minLevel) {
  if ((ROLE_LEVELS[role] || 0) < minLevel) {
    throw new Error("Permission denied for role: " + role);
  }
}

function inCsv_(email, csvVal) {
  if (!csvVal) return false;
  return String(csvVal).split(",").map(s => s.trim().toLowerCase()).indexOf(email) !== -1;
}

function formatShortDate_(d) {
  const parsed = parseValidDate_(d);
  if (!parsed) return "-";
  return Utilities.formatDate(parsed, Session.getScriptTimeZone(), "dd/MM/yy");
}

function columnToLetter_(col) {
  let temp, letter = "";
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonpOrJsonResponse_(obj, callbackParam) {
  const safeCb = String(callbackParam || "").replace(/[^a-zA-Z0-9_$.]/g, "");
  if (safeCb) {
    return ContentService
      .createTextOutput(`${safeCb}(${JSON.stringify(obj)});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonResponse_(obj);
}
