# LinkedIn Outreach Automation - System Flow

This document outlines the complete, detailed lifecycle of the system, tracking exactly how data moves through the codebase, the Unipile API, and the Google Sheets database. 

---

## 1. Initial Setup & Account Synchronization
**Trigger:** User Action (Clicks `LinkedIn Outreach System -> 🌀 . Sync Unipile Accounts`)
**File:** `src/Accounts.gs`

1. **Read Credentials:** The script reads `UNIPILE_API_KEY` and `UNIPILE_BASE_URL` from the Apps Script script properties.
2. **Fetch Accounts:** Makes a `GET /accounts` request to the Unipile API to retrieve all connected LinkedIn accounts.
3. **Update Sheet:** Updates the `Accounts` sheet. 
   - New accounts are added with default limits (e.g., 100 invites/day).
   - Accounts returning a status of "OK" are marked as **Active**.
   - Missing or disconnected accounts are marked as **Inactive**.

---

## 2. Prospect Enrichment
**Trigger:** User Action (Clicks `LinkedIn Outreach System -> 🔍 . Enrich LinkedIn Prospects`)
**File:** `src/Prospects.gs`

1. **Selection:** The script prompts the user to select an active sending account.
2. **Read Prospects:** Reads raw rows from the `Prospects` sheet containing LinkedIn URLs.
3. **Batch Fetch:** Parses the URLs to extract LinkedIn identifiers and performs a batch of `UrlFetchApp.fetchAll()` queries (`GET /users/{id}?account_id={accountId}`) against Unipile.
4. **Update Sheet:** Writes enriched data (First Name, Last Name, Location, Title, Company) back to the `Prospects` sheet. Most importantly, it logs the Unipile `provider_id` (Column 11) and flags the prospect as enriched (`Yes` in Column 12).

---

## 3. Campaign Creation & Database Entry
**Trigger:** User Action (Clicks `🚀 . Create New Campaign`, then `Start Campaign Now -> 1 - Create database entries`)
**Files:** `src/UI.gs`, `src/Campaigns.gs`

1. **Form Submission:** A user submits the `CampaignForm.html` modal. A new row is appended to the `Campaigns` sheet with status **"Not Started"**, along with connection notes and drip messaging delays.
2. **Create Entries:** When the user initiates the campaign, the script:
   - Selects a "Not Started" campaign.
   - Grabs the requested number of target prospects from the `Prospects` sheet.
   - Checks if they are enriched. If not, it does an inline bulk enrichment.
   - Moves the successfully enriched prospects into the `Database` sheet, assigning the `campaign_id` and the `account_id`.
3. **Status Update:** The prospect's `connection_request_status` is initialized as **Pending**. The Campaign status is updated to **Active**.

---

## 4. Dispatching Connection Requests
**Trigger:** User Action (Clicks `Start Campaign Now -> 2 - Send connection requests`)
**File:** `src/Campaigns.gs`

1. **Daily Limit Check:** Loops through `Database` sheet items marked as **Pending**. Checks the `Accounts` sheet to ensure the selected account has not exceeded its daily connection limits.
2. **API Request:** Issues a `POST /users/invite` to Unipile with the `provider_id` and an optional customized connection note (replacing `$name` with the prospect's First Name).
3. **Response Handling:**
   - **Success (200/201):** Marks `connection_request_status` as **Sent** and logs the `connection_request_time` (Column 14). Logs the `invitation_id` in the `Invitations` sheet to allow future un-inviting. Updates the `Accounts` sent counter.
   - **Already Connected:** Auto-corrects the status to **Accepted**.
   - **Invite Already Sent:** Attempts to fetch the existing `invitation_id` and auto-corrects to **Sent**.
   - **Errors:** Logs the specific Unipile API error in the `failed_reason` column and marks the status as **Failed**.

---

## 5. Webhook Event Handling (Real-Time Updates)
**Trigger:** Webhook from Unipile (`POST` to Apps Script Web App URL)
**File:** `src/Webhook.gs`

1. **Receive Payload:** The script captures incoming JSON payloads via the `doPost(e)` function.
2. **Route Event:**
   - **`new_relation`:** A prospect accepts the connection request.
     - Action: Looks up the row in the `Database` by `account_id` and `provider_id`.
     - Updates: `connection_request_status` -> **Accepted**, `connection_accepted` -> `TRUE`, `connection_accepted_time` -> Current Time.
   - **`message_received`:** A prospect replies.
     - Action: Confirms the message is not sent by the user (`isSelf === false`).
     - Updates: `reply_received` -> `TRUE`, `reply_text` -> Message body, `reply_time` -> Current Time.
3. **Rollup:** Calls `updateGlobalStats()` to immediately sync total counts to the `Campaigns` dashboard.

---

## 6. Background Worker Execution (Automation Loop)
**Trigger:** Time-Driven Apps Script Trigger (e.g., Every 15 minutes)
**File:** `src/Worker.gs`

This worker runs automatically in the background to handle asynchronous tasks and drip messaging.
1. **Filter Active Targets:** Scans the `Database`. It entirely skips any prospect where `reply_received` is `TRUE`. (If a prospect replies, the sequence stops immediately).
2. **Logic 1: Uninvite Stale Requests:**
   - If `connection_request_status` is **Sent** but not accepted, and > 7 days have passed since `connection_request_time`, it issues a `DELETE /users/invite/sent/{invitation_id}`.
   - Updates status to **Failed** with the reason "7 days passed, uninvited".
3. **Logic 2: Drip Messaging Sequence:**
   - If `connection_accepted` is `TRUE`:
     - **Message 1:** If status is **Pending**, sends Msg 1 (`POST /chats`), updates `msg_1_status` to **Sent**, and records `msg_1_time`.
     - **Message 2:** If Msg 1 is **Sent** and Msg 2 is **Pending**, checks if the user-defined time delay (e.g., 24 hours) has passed since Msg 1 was sent. If yes, sends Msg 2.
     - **Message 3:** Same logic applied for Msg 3, based on the delay after Msg 2.
4. **Rate Limit Protection:** Executes API requests sequentially, with random `Utilities.sleep(3000-7000)` delays to mimic human behavior and avoid LinkedIn detection.

---

## 7. Statistics & Dashboard Rollup
**Trigger:** `onOpen`, `onEdit`, Webhooks, or Time-Driven Stats Worker
**File:** `src/Worker.gs` -> `updateGlobalStats()`

1. **Calculate Aggregates:** Scans the entire `Database` sheet.
2. **Campaign Sync:** Tallys `connectionsSent`, `connectionsAccepted`, `messagesSent`, and `repliesReceived` per campaign. Writes these totals to Columns K, L, M, and N in the `Campaigns` sheet.
3. **Account Sync:** Tallys the "Sent Today" volume and updates the `Accounts` sheet to ensure limits are strictly enforced.
4. **Completion Check:** If the `repliesReceived` meets or exceeds the campaign's `targetProspects`, the campaign status is marked as **Completed**.

---

## 8. Force Check (Fallback Mechanism)
**Trigger:** User Action (Clicks `Start Campaign Now -> 3 - Check connection requests`)
**File:** `src/Campaigns.gs` -> `forceCheckRequests()`

If the Unipile webhook fails or is disabled:
1. Loops through all prospects in the `Database` currently marked as **Sent** or **Pending**.
2. Makes a direct API polling query (`GET /users/{provider_id}`) to check the user's profile.
3. If `connected_at` is populated in the profile JSON, it manually forces the `Database` record to **Accepted** and triggers the messaging worker to pick it up on the next cycle.
