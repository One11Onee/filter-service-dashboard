# Filter Service & Dispatch System

Internal web dashboard and automated daily reminder pipeline for water filter maintenance (`RED` and `GREEN` service sheets).

## Stack
- **Frontend**: `index.html` on Cloudflare Pages (`https://itsaso.sqrtone.com`) with Cloudflare Zero Trust
- **Backend & Scheduler**: Google Apps Script (`GAS/webApp.gs`, `GAS/Code.gs`)
- **Database**: Google Sheets (`Sheet 1`, `IT_Config`, `_Audit_Log`)
- **Mailer & E5 Keepalive**: Microsoft Graph API with automatic `MailApp` fallback

## AI & Maintenance Handoff
For full schemas, variables, API contracts, and guardrails, see [`SYSTEM_CONTEXT.md`](./SYSTEM_CONTEXT.md).
