# Filter Service & Dispatch System — Engineering & AI Handoff Context

> **Purpose**: This document provides complete architectural, schema, API, and operational context so any developer or AI assistant can immediately maintain or extend this system without re-discovering past decisions.
> **Security Rule**: Never commit actual secrets, passwords, client secrets, or personal credentials to this repository.
---

## 0. Repository Structure & File Locations
- **Repository Root**:
  - `SYSTEM_CONTEXT.md`: Master handoff context (architecture, schemas, variables, guardrails, preferences, changelog).
  - `README.md`: Concise project overview pointing to `SYSTEM_CONTEXT.md`.
  - `index.html`: Frontend web dashboard deployed to Cloudflare Pages (`https://itsaso.sqrtone.com`).
- **`forGAS/` Folder**:
  - Contains the Google Apps Script source files (`Code.gs`, `webApp.gs`, etc.) maintained in Apps Script and backed up in GitHub under `forGAS/`.

---

## 1. System Architecture & Evolution
The system evolved from a manual Excel/Google Sheet + MS Forms workflow into a single-entry, role-isolated web application backed by Google Sheets and automated daily emails:

1. **Frontend (`index.html`)**: Single-file vanilla HTML/CSS/JS dashboard hosted on **Cloudflare Pages** (`https://itsaso.sqrtone.com`) behind **Cloudflare Zero Trust** authentication. Designed with responsive multi-screen scaling (desktop, split-screen, laptop, mobile), an iOS-inspired UI, Dynamic Island status notifications, `localStorage` snapshot caching, and landscape print formatting.
2. **Backend API & Webhooks (`GAS/webApp.gs`)**: Google Apps Script deployed as a Web App (`Execute as: Me`) serving JSON/JSONP data to the frontend and handling one-click email action links (`markNotified`, `confirmDone`).
3. **Daily Mailer & E5 Keepalive (`GAS/Code.gs`)**: Triggered daily via GAS time-driven trigger. Sends two independent HTML summary emails (`RED` and `GREEN`) via **Microsoft Graph API** (`sendMail` + lightweight read-only `/users` & `/organization` calls to maintain Microsoft 365 E5 developer subscription activity), with automatic fallback to Google `MailApp`.
4. **Database (Google Sheets)**: Two spreadsheets (`RED` = Water Filter Service; `GREEN` = Itsaso Filter Service). Direct user editing on the sheets is restricted; all CRUD operations flow through the web dashboard.

---

## 2. Database Schemas (Strict Column Structures — DO NOT Add/Remove Business Columns)

Every spreadsheet contains three tabs: `Sheet 1` (main records), `IT_Config` (key-value settings), and `_Audit_Log` (append-only audit trail).

### A. `RED` Sheet (`Water Filter Service` — 12 Columns)
Header-mapped dynamically via `buildHeaderIndexMap_()`, expected columns in `Sheet 1`:
1. `第 1 列` (or `ID`): Static integer ID (must remain static values, **never** `=A_prev+1` formulas).
2. `Name`: Customer / company name. Multi-filter customers store **one row per filter** sharing the exact same `Name` (no extra business columns required).
3. `Filter / Service Name`: Filter model/service description.
4. `Machine Usage`: Boolean (`TRUE` / `FALSE`). **Critical**: `Code.gs` skips rows where `Machine Usage` is `FALSE` on the `RED` sheet.
5. `Last Changed Date`: Date of last service (`Date` object).
6. `Change Cycle (Days)`: Integer cycle (e.g., `90`, `180`, `270`, `360`, `540`).
7. `Next Change Date`: Formula rolling weekend due dates back to Friday:
   `=IF(WEEKDAY(E{r}+F{r},2)>5, (E{r}+F{r})-(WEEKDAY(E{r}+F{r},2)-5), E{r}+F{r})`
8. `Notification Date`: `=G{r}-14` (14 days prior to `Next Change Date`).
9. `Status`: Formula computing `"Normal"`, `"Pending Notification"`, or `"REPLACEMENT REQUIRED"`, or static `"Archived"` when soft-deleted.
10. `Notification Status`: Boolean (`TRUE` / `FALSE`). Reset to `FALSE` upon `confirmDone`.
11. `Importance`: Priority rank (`1`–`4`).
12. `Approval Status` (or `Contact` if present).

### B. `GREEN` Sheet (`Itsaso Filter Service` — 10 Columns)
1. `第 1 列` (or `ID`): Static integer ID.
2. `Name`: Customer name.
3. `Phone Number`: Customer contact number (maps to `extraInfo` / `filterOrPhone` in code). **Note**: `GREEN` has no `Service / Filter Plan` column; do not add one to the modal or sheet.
4. `Last Changed Date`: Date of last service.
5. `Change Cycle (Days)`: Integer cycle.
6. `Next Change Date`: Weekend-adjusted due date formula (`=IF(WEEKDAY(D{r}+E{r},2)>5, ...)`).
7. `Notification Date`: `=F{r}-14`.
8. `Status`: `"Normal"` | `"Pending Notification"` | `"REPLACEMENT REQUIRED"` | `"Archived"`.
9. `Notification Status`: Boolean (`TRUE` / `FALSE`).
10. `Importance`: Priority rank (`1`–`4`).

### C. `_Audit_Log` Tab (10 Columns)
`Timestamp` | `Actor_Email` | `Effective_Role` | `Customer_ID` | `Customer_Name` | `Action` | `Reason` | `Before_State` | `After_State` | `Row_Number`

---

## 3. Configuration & Environment Variables

### A. GAS Script Properties (`Project Settings -> Script Properties`)
* `MS_TENANT_ID`: Entra ID Directory (tenant) ID
* `MS_CLIENT_ID`: Entra ID Application (client) ID
* `MS_CLIENT_SECRET`: Entra ID Client Secret **Value** (not Secret ID)
* `MS_SENDER_EMAIL`: Licensed/shared E5 sender address (e.g. `Itsaso-Alert@sqrtone.com`)
* `MS_SECRET_EXPIRES_AT`: Secret expiration date in `YYYY-MM-DD` format (triggers renewal email when `<= 30` days remain)

### B. `IT_Config` Sheet Keys (Col A = Key, Col B = Value)
* **Sheet Identity**: `SHEET_CODE` (`RED` or `GREEN`), `SHEET_LABEL`, `GREEN_SHEET_ID` (stored in RED's `IT_Config` so GAS can open GREEN via `SpreadsheetApp.openById()`), `DASHBOARD_URL` (`https://itsaso.sqrtone.com`).
* **Email Routing**: `TO_EMAIL`, `CC_EMAILS` (comma-separated), `ENABLE_MS_GRAPH` (`TRUE` / `FALSE`).
* **Role Emails**: `ROLE_IT`, `ROLE_BOSS`, `ROLE_ADMIN_H`, `ROLE_ADMIN_A`, `ROLE_ADMIN_L`.
* **Role Display Names**: `NAME_IT`, `NAME_BOSS`, `NAME_ADMIN_H`, `NAME_ADMIN_A`, `NAME_ADMIN_L`.
* **Sheet Permissions**: `SHEETS_ADMIN_A` (e.g. `RED`), `SHEETS_ADMIN_H` (`RED`), `SHEETS_ADMIN_L` (`GREEN`). `IT` and `BOSS` automatically have access to both `RED` and `GREEN`.
* **Form Dropdowns & Filter-Cycle Bindings**:
  * `OPTIONS_RED_FILTER`: Supports `FilterName:DefaultDays` pairs or plain `FilterName` entries separated by commas (e.g., `Sediment:180, Pre Carbon / GAC:270, Post Carbon / Antibacterial:360, Alkaline / Mineral Stone:540, Silver Carbon:540, UF Membrane:540, Mineral Filter:540, PP 5M:90, Carbon Block:180, PP 1M:180, UDF:180`).
    * Selecting a filter with `:Days` auto-fills `Change Cycle (Days)`.
    * Selecting/typing a filter without `:Days` keeps the default preset from `OPTIONS_CYCLES`.
    * Clean filter names (without `:Days`) are displayed in suggestions and saved to `Sheet 1`.
  * `OPTIONS_CYCLES`: Comma-separated preset cycle days (e.g., `90,180,270,360,540`).

### C. Key Code Constants
* **`index.html`**:
  * `GAS_URL`: Deployed GAS Web App `/exec` URL.
  * `COOKIE_KEY` / `STORAGE_KEY` (`wf_zt_email`): Persists verified user email.
  * `VIEW_AS_KEY` (`wf_view_as`): Allows `IT` role to preview the dashboard as `BOSS`, `ADMIN_H`, `ADMIN_A`, or `ADMIN_L`.
  * `CLIENT_CACHE_PREFIX` (`wf_snap_v6_`): `localStorage` snapshot cache key prefix.
  * `SESSION_USER_LABEL_PREFIX` (`wf_user_label_v1_`): Per-email display name cache for faster greeting updates (cleared on `logout()`).
* **`GAS/Code.gs`**:
  * `MAIL_TEST_MODE`: Boolean. When `true`, `CC_EMAILS` are suppressed so only `TO_EMAIL` receives test runs. **Must be `false` in production.**
  * `SECRET_WARN_DAYS`: `30`.
  * `FALLBACK_MAX_BYTES`: `175000` (triggers compact top-25-per-column rendering only when falling back to `MailApp` so Gmail body limits are never exceeded).

---

## 4. API & Frontend-Backend Interaction Contract

* **Read / Email Webhooks (`doGet` in `webApp.gs`)**:
  * `action=getDashboardData` (with `email`, `sheet`, `viewAs`, optional `includeLogs=1`, optional `callback` for JSONP): Returns `{ ok, actorEmail, actualRole, effectiveRole, roleDisplayName, roleLabels, canSwitchView, allowedSheets, activeSheet, sheetHeaders, records, summary, dropdownOptions: { redFilter, filterCycles, greenPlan, cycles }, auditLogs }`.
  * `action=confirmDone&sheet=RED|GREEN&id=<id>`: Sets `Last Changed Date` = today, resets `Notification Status` = `FALSE`, recalculates formulas, logs `CONFIRM_DONE` to `_Audit_Log`, invalidates GAS cache, and returns an HTML confirmation page.
  * `action=markNotified&sheet=RED|GREEN&id=<id>`: Sets `Notification Status` = `TRUE`, invalidates cache, and returns HTML confirmation.
* **Batch Write (`doPost` in `webApp.gs`)**:
  * Frontend stages edits, single/multi-filter creations, soft-deletes (`Status = "Archived"` + row hide), and restores in `currentState.pendingOps` with cell-level visual diff highlights, then commits all staged changes in a single `batchSave` POST request when the user clicks **Save**.

---

## 5. Strict Engineering Guardrails (Must Follow in Future Changes)
1. **Dropdowns Must Strictly Follow `IT_Config`**: Never auto-merge historical values from `Sheet 1` into `<datalist>` / dropdown options; users select from `IT_Config` suggestions or type custom values directly in the single-line combo inputs.
2. **Multi-Filter Customer Architecture (`RED` Only)**: One customer with multiple filters is stored as multiple rows in `Sheet 1` sharing the same `Name`. Each filter card in the modal binds `Filter / Service Name` + `Change Cycle (Days)` + `Last Changed Date`. In the table, active rows with the same `Name` are grouped together on default `ID` ascending sort and display a dynamic visible count pill (`N filters`) computed from the currently filtered result set.
3. **Soft Delete Only**: Never physically delete customer rows (`deleteRow`). Always archive (`Status = "Archived"` + hide row) and log to `_Audit_Log` so `BOSS` / `IT` can audit and restore.
4. **Preserve GAS Web App URL**: When updating `webApp.gs` or `Code.gs`, always deploy via **Deploy -> Manage deployments -> Edit (pencil icon) -> New version**. Never create a brand-new deployment unless intentionally rotating `GAS_URL` in `index.html`.
5. **Cold-Start UX & Logout Cache Wipe**: Keep the fullscreen `Welcome` overlay strictly conditional on the absence of `localStorage` snapshot cache, update the greeting to `roleDisplayName` immediately upon backend response, and purge all `wf_*` keys in `localStorage`, `sessionStorage`, cookies, and Cache Storage on `logout()`.
6. **Print Contract (`Ctrl+P`)**: Intercept `Ctrl+P` / `Cmd+P` on the dashboard to show the Status multi-select print modal first; force landscape orientation (`@page { size: landscape }`), light theme, hidden controls/`Actions` column, group printed rows strictly in status order (`REPLACEMENT REQUIRED` -> `Pending Notification` -> `Normal` -> `Archived` without mixing), show `Printed By: <roleDisplayName> (<email>)` in the fixed print footer, and restore previous filter/sort state on `afterprint`.
7. **Email Rendering**: Use `String.fromCodePoint(...)` for emojis in `Code.gs` to avoid GAS editor encoding loss, and keep two-column (`Overdue` left / `Due Soon` right) table layouts compatible with desktop Outlook and Gmail.

---

## 6. Key Function Map (Fast Navigation)
- **`index.html`**: `initApp`, `loadData`, `applyDashboardPayload`, `syncDatalists`, `addFilterBindRow`, `onBindFilterChange`, `submitModal`, `renderTable`, `openPrintModal`, `executePrintWithSelection`, `clearAllClientCaches`, `showWelcomeScreen`, `hideWelcomeScreen`.
- **`GAS/Code.gs`**: `sendDailyFilterEmails`, `checkSecretExpiration_`, `processSingleSheetReminder_`, `buildAppleAlignedReminderHtml_`, `dispatchEmailWithFallback_`, `sendViaMsGraph_`.
- **`GAS/webApp.gs`**: `doGet`, `doPost`, `buildDropdownOptions_`, `resolveEffectiveAuth_`, `handleConfirmDoneById_`, `handleBatchSave_`.

---

## 7. Common Failure Modes & Quick Diagnosis
- **MS Graph fails (`401` / `invalid_client`)**: Check `MS_CLIENT_SECRET` (must be Secret Value, not Secret ID) and `MS_SECRET_EXPIRES_AT` in GAS Script Properties; verify `dispatchEmailWithFallback_` logs `GRAPH_FALLBACK` in `_Audit_Log` and falls back to `MailApp`.
- **Fallback `MailApp` body size error**: Ensure `compactFallbackHtml` and `FALLBACK_MAX_BYTES` are intact so oversized lists truncate cleanly instead of throwing body-size exceptions.
- **Role shows as `GUEST`**: Verify user email matches `ROLE_IT` / `ROLE_BOSS` / `ROLE_ADMIN_*` in `IT_Config` and Cloudflare Zero Trust header/email parameter.
- **Row update hits wrong customer**: Check that Col A (`第 1 列`) contains static numbers rather than formulas, and verify `Sheet 1` is targeted explicitly rather than active sheet.

---

## 8. Safe Testing & Rollback Protocol
1. **Safe Email Testing**: Set `MAIL_TEST_MODE = true` in `Code.gs` and temporarily set `TO_EMAIL` in `IT_Config` to your test inbox so `CC_EMAILS` are not spammed. Restore both before finishing.
2. **GAS Rollback**: Open Apps Script -> **Deploy -> Manage deployments -> Edit** -> switch **Version** back to the previous stable version.
3. **Sheet Recovery**: Use the backup spreadsheet copies (`_Backup_*`) in Google Drive or check `_Audit_Log` to trace recent changes.

---

## 9. Collaboration & Style Preferences (For Any AI Assisting on This Project)
- **Factuality First**: Stick strictly to provided files, logs, and explicit facts. Never invent unstated roles, backgrounds, or assumptions.
- **Direct & Concise**: Lead with conclusions and actionable steps. Avoid fluff, lecturing, or long-winded filler. When editing code, provide complete drop-in code or unambiguous head-to-tail replacement blocks.
- **Visual & UI Taste**: Prefer clean, restrained, Apple/iOS-inspired clarity—aligned layouts, clear hierarchy, readable font/button sizes, and no gaudy full-row fills or excessive emojis.
- **Scope Discipline**: Favor simple, reliable solutions over over-engineering.

---

## 10. Changelog & Recent Engineering Decisions (2026-11-21)
1. **Multi-Filter Binding per Customer (`RED` Only)**:
   - Implemented multi-row creation (`+ Add Filter`) and filter addition during `Edit` without adding columns to `Sheet 1`.
   - Each filter card binds `Filter / Service Name`, `Change Cycle (Days)`, and `Last Changed Date` using single-line `<input list>` combo boxes (eliminating extra stacked `Other` input boxes).
   - Removed the unused `Service / Filter Plan` field from the `GREEN` modal so it matches `GREEN`'s 10-column schema.
2. **Configurable `IT_Config` Filter-to-Cycle Mapping**:
   - Updated `buildDropdownOptions_()` in `webApp.gs` to parse `FilterName:Days` entries in `OPTIONS_RED_FILTER` and return `filterCycles` alongside clean `redFilter` labels. Filters without `:Days` fall back to the `OPTIONS_CYCLES` preset without forcing an override.
3. **Multi-Select Checkbox Filters & Visible Group Counts**:
   - Converted `Status`, `Notified`, and `Machine Usage` toolbar filters into multi-select checkbox dropdowns with `Select All`.
   - Computed `N filters` group badges and customer clustering on the filtered result set so filtered views never show broken/stale counts.
4. **Landscape Light-Theme Print Workflow (`Ctrl+P`) & Full Logout Cache Wipe**:
   - Added a pre-print Status checkbox selector modal (`Ctrl+P`), non-mixed status ordering (`REPLACEMENT REQUIRED` -> `Pending Notification` -> `Normal` -> `Archived`), forced light landscape layout, `Printed By: <User>` footer, and automatic post-print state restoration.
   - Updated `logout()` to clear all `wf_*` keys in `localStorage`, `sessionStorage`, cookies, and Cache Storage, and updated the Welcome screen to reflect `roleDisplayName` as soon as `getDashboardData` resolves.
